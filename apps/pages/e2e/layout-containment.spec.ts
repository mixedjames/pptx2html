import { expect, test } from '@playwright/test';
import { openDemo, switchToScrollMode } from './helpers.js';

/**
 * Regression coverage for `packages/to-html5/CLAUDE.md`'s "Hard requirement" rules 1–3 (opaque,
 * self-sizing box; click/scroll visual parity; no leakage past the host's own box) and the
 * specific scroll-mode letterbox bug that violated them — see that file's "Bug 1" writeup.
 * `portrait-slides.pptx` is used deliberately: a portrait deck in a landscape viewport is the
 * strongest pillarbox case, exactly where a broken letterbox stands out.
 */
test.describe('layout containment and click/scroll visual parity', () => {
  test.use({ viewport: { width: 1200, height: 800 } });

  test('click mode pillarboxes a portrait deck within the host box, not stretched to full width', async ({
    page,
  }) => {
    await openDemo(page, 'Portrait Slides');
    const host = page.locator('pptx-presentation');
    const hostBox = await host.boundingBox();
    const slideBox = await host.locator('.pptx-slide.pptx-slide--active').boundingBox();
    expect(hostBox).not.toBeNull();
    expect(slideBox).not.toBeNull();

    // Nothing rendered inside the element may exceed the host's own box (Hard requirement rule 3).
    expect(slideBox!.x).toBeGreaterThanOrEqual(hostBox!.x - 1);
    expect(slideBox!.y).toBeGreaterThanOrEqual(hostBox!.y - 1);
    expect(slideBox!.x + slideBox!.width).toBeLessThanOrEqual(hostBox!.x + hostBox!.width + 1);
    expect(slideBox!.y + slideBox!.height).toBeLessThanOrEqual(hostBox!.y + hostBox!.height + 1);

    // A genuinely portrait deck in a landscape viewport must pillarbox — narrower than the host,
    // not stretched to fill its full width.
    expect(slideBox!.width).toBeLessThan(hostBox!.width * 0.9);
  });

  test('scroll mode sizes the deck identically to click mode for the same viewport', async ({
    page,
  }) => {
    await openDemo(page, 'Portrait Slides');
    const clickSlideBox = await page
      .locator('pptx-presentation .pptx-slide.pptx-slide--active')
      .boundingBox();

    await switchToScrollMode(page);
    const scrollSlideBox = await page
      .locator('pptx-scroll-presentation .pptx-scroll-viewport > *:visible')
      .first()
      .boundingBox();

    expect(clickSlideBox).not.toBeNull();
    expect(scrollSlideBox).not.toBeNull();
    // Regression test for the scroll-mode letterbox bug: before the fix, `.pptx-scroll-viewport`'s
    // content ignored its own letterboxed size entirely and rendered at the host's full,
    // unletterboxed width instead (see that file's "Bug, a later session still" paragraph) — this
    // would fail loudly here.
    expect(scrollSlideBox!.width).toBeCloseTo(clickSlideBox!.width, 0);
    expect(scrollSlideBox!.height).toBeCloseTo(clickSlideBox!.height, 0);
  });
});
