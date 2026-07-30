import type { Background, Slide } from '@pptx2html/presentation';

/**
 * Resolves a slide's effective background (§19.3.1.7, p:bg): the slide's own, falling back to
 * its layout's, falling back to the layout's master's. Simpler than placeholder transform/font
 * inheritance (`placeholder.ts`, `text-style.ts`) — there's no per-shape matching step, just
 * "does this slide-like part define one at all", so the first one found in the chain wins outright
 * rather than being merged field-by-field.
 */
export function resolveEffectiveBackground(slide: Slide): Background | undefined {
  return (
    slide.commonSlideData.background ??
    slide.layout.commonSlideData.background ??
    slide.layout.master.commonSlideData.background
  );
}
