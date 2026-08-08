import type { Color, Fill, Line } from './drawingml/index.js';

/** The 12 named colour slots a theme defines (§20.1.6.2, a:clrScheme). */
export interface ColorScheme {
  readonly name: string;
  readonly dk1: Color;
  readonly lt1: Color;
  readonly dk2: Color;
  readonly lt2: Color;
  readonly accent1: Color;
  readonly accent2: Color;
  readonly accent3: Color;
  readonly accent4: Color;
  readonly accent5: Color;
  readonly accent6: Color;
  readonly hlink: Color;
  readonly folHlink: Color;
}

export interface FontCollection {
  readonly latin: string;
  readonly eastAsian?: string;
  readonly complexScript?: string;
}

/** §20.1.4.1.18, a:fontScheme. */
export interface FontScheme {
  readonly name: string;
  readonly majorFont: FontCollection;
  readonly minorFont: FontCollection;
}

/**
 * §20.1.4.1.19, a:fmtScheme. The fill, line, and background-fill style matrices (`a:fillStyleLst`/
 * `a:lnStyleLst`/`a:bgFillStyleLst`) are modeled, each always exactly 3 entries. `fillStyleLst`/
 * `lineStyles` are the 1-based index a shape's `p:style/fillRef`/`lnRef` (§20.1.4.2.10/2.12,
 * `ShapeStyle` in `presentationml/shape-style.ts`) points into; `bgFillStyles` is what a slide/
 * layout/master's own `p:bg/p:bgRef` (§19.3.1.6 — `CommonSlideData.backgroundRef`, same
 * `StyleMatrixReference` shape as `fillRef`/`lnRef`, just a different `idx` numbering convention —
 * see `resolve/style-matrix.ts`'s `resolveBackgroundStyleFill`) points into instead — PowerPoint
 * writes a background this way whenever a slide/layout/master relies on the theme's own default
 * background rather than an explicitly-picked one, which is the common case, not an edge case.
 * `a:effectStyleLst` remains unmodeled: no consumer needs it yet (no effect rendering).
 */
export interface FormatScheme {
  readonly name: string;
  readonly fillStyles: readonly Fill[];
  readonly lineStyles: readonly Line[];
  readonly bgFillStyles: readonly Fill[];
}

/** A theme part (§14.2.7, theme1.xml). */
export interface Theme {
  readonly name: string;
  readonly colorScheme: ColorScheme;
  readonly fontScheme: FontScheme;
  readonly formatScheme: FormatScheme;
}
