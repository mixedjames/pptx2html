import type { Page } from '@playwright/test';

/**
 * Both `PptxPresentationElement` and `PptxScrollPresentationElement` correct their own sizing
 * (`contain-size.ts`'s `ResizeObserver`, and for scroll mode, `#trackResize` too) asynchronously
 * after a fresh render, potentially across more than one callback firing as layout settles in
 * stages — waiting for "the element exists" or even "its size is non-zero" isn't enough, since a
 * caller can still catch an intermediate reading between two firings, not the final one. This is
 * what actually caught a real, reproducible flake in this suite (roughly 1 run in 25) before
 * `openDemo`/`switchToScrollMode` below were changed to wait for two *consecutive* identical
 * readings instead of just one non-zero one.
 */

/**
 * Opens the chooser view and picks a demo deck by its (rendered, title-cased) name — waiting for
 * the click-mode element's own letterbox sizing to have actually settled, not just exist (see the
 * module doc comment above).
 */
export async function openDemo(page: Page, titleSubstring: string): Promise<void> {
  await page.goto('/');
  await page.click(`.demo-button:has-text("${titleSubstring}")`);
  await page.waitForSelector('#output > *');
  await page.waitForFunction(() => {
    const w = window as unknown as { __lastSlideWidth?: number };
    const slide = document
      .querySelector('pptx-presentation')
      ?.shadowRoot?.querySelector('.pptx-slide.pptx-slide--active');
    const width = slide?.getBoundingClientRect().width ?? 0;
    if (width <= 0) return false;
    const stable = w.__lastSlideWidth === width;
    w.__lastSlideWidth = width;
    return stable;
  });
}

/** The presentation view's header (mode toggle included) starts hidden — see index.ts's own
 *  `setUiVisible`. */
async function revealControls(page: Page): Promise<void> {
  await page.click('#ui-toggle');
  await page.waitForSelector('#mode-toggle:visible');
}

/**
 * Switches the currently-open demo from click mode (the default) to scroll mode, waiting for the
 * scroll element's own sizing to have actually settled before returning, not just exist (see the
 * module doc comment above).
 */
export async function switchToScrollMode(page: Page): Promise<void> {
  await revealControls(page);
  await page.click('#mode-toggle');
  await page.waitForSelector('pptx-scroll-presentation');
  await page.waitForFunction(() => {
    const w = window as unknown as { __lastTrackHeight?: number };
    const track = document
      .querySelector('pptx-scroll-presentation')
      ?.shadowRoot?.querySelector('.pptx-scroll-track');
    const height = track?.clientHeight ?? 0;
    if (height <= 0) return false;
    const stable = w.__lastTrackHeight === height;
    w.__lastTrackHeight = height;
    return stable;
  });
}
