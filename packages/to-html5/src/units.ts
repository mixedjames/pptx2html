import type { Emu, FontSize } from '@pptx2html/presentation';

/** EMU per CSS pixel at 96 DPI (914400 EMU/inch ÷ 96 px/inch). */
export const EMU_PER_PX = 9525;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

/** EMU per point (914400 EMU/inch ÷ 72 pt/inch). */
export const EMU_PER_PT = 12700;

/** Converts a run's FontSize (§20.1.10.16, hundredths of a point) to Emu. */
export function fontSizeToEmu(fontSize: FontSize): Emu {
  return (fontSize / 100) * EMU_PER_PT;
}

/**
 * Expresses an EMU magnitude as a percentage of the slide's own width, in CSS container query
 * width units (`cqw`) rather than a fixed `px`/`pt` value — so it scales with `.pptx-slide`'s own
 * rendered size (see `slide.ts`'s `container-type: inline-size`) exactly the way position/size
 * already do via plain percentages (see `shape-tree.ts`'s `positionElement`), with no JS resize
 * handling. Slide *width* is always the reference dimension, even for values that aren't
 * inherently horizontal (font size, border width) — `.pptx-slide`'s `aspect-ratio` keeps both axes
 * scaling by the same factor, so width and height are equally valid references; width is picked
 * for consistency with `positionElement`'s own left/width percentages.
 */
export function emuToCqw(emu: number, slideWidth: Emu): string {
  return `${(emu / slideWidth) * 100}cqw`;
}
