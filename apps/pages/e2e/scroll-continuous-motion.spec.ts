import { expect, test } from '@playwright/test';
import { openDemo, switchToScrollMode } from './helpers.js';

/**
 * Regression coverage for two real bugs documented in `packages/to-html5/CLAUDE.md`'s "Key design
 * decision: three behavioral scroll-mode bugs" — a Web Animations composite-order bug (bug 2) and
 * a dead-scroll-zone bug (bug 3/4), plus the cropped-final-slide bug (bug 1). `portrait-slides.pptx`
 * is the deck all three were actually found against: three slides, push throughout, no builds at
 * all — the simplest possible scroll-triggered deck.
 *
 * These drive `PptxScrollPresentationElement.seekTo` directly (not real `wheel`/`scroll` events) —
 * the same directly-testable entry point `scroll-presentation-element.test.ts`'s own unit tests use
 * — since what's under test here is genuine browser layout/WAAPI behavior once a `ms` is chosen,
 * not the scroll-event-to-`ms` conversion itself (already covered by a unit test with a mocked
 * `requestAnimationFrame`).
 */
test.describe('scroll-driven playback: continuous motion, no dead zones or shadowed animations', () => {
  test('scrubbing the whole timeline never produces two consecutive frames with identical geometry', async ({
    page,
  }) => {
    await openDemo(page, 'Portrait Slides');
    await switchToScrollMode(page);

    const { maxDeadRunMs, totalDurationMs } = await page.evaluate(() => {
      const el = document.querySelector('pptx-scroll-presentation')!;
      const viewport = el.shadowRoot!.querySelector('.pptx-scroll-viewport')!;
      const slides = [...viewport.children] as HTMLElement[];
      const total = (el as unknown as { totalDurationMs: number }).totalDurationMs;
      const seekTo = (el as unknown as { seekTo: (ms: number) => void }).seekTo.bind(el);

      const steps = 60;
      let previousKey: string | null = null;
      let runStartMs: number | null = null;
      let maxRun = 0;
      for (let i = 0; i <= steps; i++) {
        const ms = (total * i) / steps;
        seekTo(ms);
        const key = JSON.stringify(
          slides
            .map((s) => ({
              display: getComputedStyle(s).display,
              transform: getComputedStyle(s).transform,
            }))
            .filter((s) => s.display !== 'none'),
        );
        if (key === previousKey) {
          if (runStartMs === null) runStartMs = ms;
        } else {
          if (runStartMs !== null) maxRun = Math.max(maxRun, ms - runStartMs);
          runStartMs = null;
        }
        previousKey = key;
      }
      return { maxDeadRunMs: maxRun, totalDurationMs: total };
    });

    expect(totalDurationMs).toBeGreaterThan(0);
    // Regression test for the dead-scroll-zone bug: a fixed content-phase floor used to reserve a
    // long stretch of scroll distance during which nothing animated at all, on every slide.
    expect(maxDeadRunMs).toBe(0);
  });

  test("a middle slide's own arrival actually animates, not shadowed by the next transition's departure", async ({
    page,
  }) => {
    await openDemo(page, 'Portrait Slides');
    await switchToScrollMode(page);

    const midwayTransform = await page.evaluate(() => {
      const el = document.querySelector('pptx-scroll-presentation')!;
      const seekTo = (el as unknown as { seekTo: (ms: number) => void }).seekTo.bind(el);
      const total = (el as unknown as { totalDurationMs: number }).totalDurationMs;
      // Halfway through the whole timeline lands inside the middle slide's own arrival — a push
      // deck with N equal-speed transitions and no builds spends its first ~1/(N-1) of the total
      // on that arrival.
      seekTo(total * 0.25);
      const slides = [...el.shadowRoot!.querySelector('.pptx-scroll-viewport')!.children];
      const slide1 = slides[1] as HTMLElement;
      return getComputedStyle(slide1).transform;
    });

    // Regression test for the composite-order bug: a shadowed arrival is permanently stuck at its
    // *own* first keyframe — the identity transform — regardless of scroll position, indistinguishable
    // from a hard cut.
    expect(midwayTransform).not.toBe('matrix(1, 0, 0, 1, 0, 0)');
    expect(midwayTransform).not.toBe('none');
  });

  test('scrolling to the very bottom of the page reaches the deck’s own final millisecond, fully settled', async ({
    page,
  }) => {
    await openDemo(page, 'Portrait Slides');
    await switchToScrollMode(page);

    const result = await page.evaluate(() => {
      const el = document.querySelector('pptx-scroll-presentation')!;
      const track = el.shadowRoot!.querySelector('.pptx-scroll-track') as HTMLElement;
      const pixelsPerSecond = (el as unknown as { pixelsPerSecond: number }).pixelsPerSecond;
      const seekTo = (el as unknown as { seekTo: (ms: number) => void }).seekTo.bind(el);
      const totalDurationMs = (el as unknown as { totalDurationMs: number }).totalDurationMs;

      track.scrollTop = track.scrollHeight;
      const msAtMaxScroll = track.scrollTop / (pixelsPerSecond / 1000);
      seekTo(msAtMaxScroll);

      const slides = [...el.shadowRoot!.querySelector('.pptx-scroll-viewport')!.children];
      const last = slides[slides.length - 1] as HTMLElement;
      return { msAtMaxScroll, totalDurationMs, transform: getComputedStyle(last).transform };
    });

    // Regression test for the cropped-final-slide bug: the spacer used to be short by exactly one
    // screen's worth, so scrolling to the bottom of the page never actually reached the deck's own
    // end, leaving the final slide's own arrival permanently mid-flight.
    expect(result.msAtMaxScroll).toBeCloseTo(result.totalDurationMs, 0);
    expect(result.transform).toBe('matrix(1, 0, 0, 1, 0, 0)');
  });
});
