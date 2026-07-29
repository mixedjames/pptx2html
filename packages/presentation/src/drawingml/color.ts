import type { Angle, Percentage } from './units.js';

/** Theme color slots defined by a ColorScheme (§20.1.4.1.), plus the placeholder/derived names usable in shape properties. */
export type SchemeColorName =
  | 'dk1'
  | 'lt1'
  | 'dk2'
  | 'lt2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink'
  | 'phClr'
  | 'bg1'
  | 'tx1'
  | 'bg2'
  | 'tx2';

/** Colour transform modifiers (§20.1.2.3.*), applied on top of a base colour value. */
export interface ColorTransform {
  readonly alpha?: Percentage;
  readonly lumMod?: Percentage;
  readonly lumOff?: Percentage;
  readonly shade?: Percentage;
  readonly tint?: Percentage;
  readonly satMod?: Percentage;
  readonly hueMod?: Percentage;
}

export interface SrgbColor {
  readonly type: 'srgb';
  /** 6-digit hex, e.g. "FF0000". */
  readonly value: string;
  readonly transforms?: ColorTransform;
}

export interface SchemeColor {
  readonly type: 'scheme';
  readonly value: SchemeColorName;
  readonly transforms?: ColorTransform;
}

export interface SystemColor {
  readonly type: 'system';
  /** OS-defined colour name, e.g. "windowText". */
  readonly value: string;
  /** 6-digit hex the value resolved to at authoring time. */
  readonly lastColor: string;
  readonly transforms?: ColorTransform;
}

export interface PresetColor {
  readonly type: 'preset';
  /** CSS4/VML preset colour name, e.g. "aliceBlue". */
  readonly value: string;
  readonly transforms?: ColorTransform;
}

export interface HslColor {
  readonly type: 'hsl';
  readonly hue: Angle;
  readonly saturation: Percentage;
  readonly luminance: Percentage;
  readonly transforms?: ColorTransform;
}

export type Color = SrgbColor | SchemeColor | SystemColor | PresetColor | HslColor;
