import type { Color } from './drawingml/index.js';

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
 * §20.1.4.1.19, a:fmtScheme. The fill/line/effect style matrices referenced by a shape's
 * style index are unmodeled for the skeleton.
 */
export interface FormatScheme {
  readonly name: string;
}

/** A theme part (§14.2.7, theme1.xml). */
export interface Theme {
  readonly name: string;
  readonly colorScheme: ColorScheme;
  readonly fontScheme: FontScheme;
  readonly formatScheme: FormatScheme;
}
