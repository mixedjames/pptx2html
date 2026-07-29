import type { Color } from './color.js';
import type { MediaPart } from './media.js';
import type { Angle, Percentage } from './units.js';

export interface NoFill {
  readonly type: 'none';
}

export interface SolidFill {
  readonly type: 'solid';
  readonly color: Color;
}

export interface GradientStop {
  readonly position: Percentage;
  readonly color: Color;
}

/** §20.1.8.33, a:gradFill. Path gradients are unmodeled for the skeleton. */
export interface GradientFill {
  readonly type: 'gradient';
  readonly stops: readonly GradientStop[];
  readonly angle?: Angle;
}

/** §20.1.8.47, a:pattFill. */
export interface PatternFill {
  readonly type: 'pattern';
  /** Preset pattern name, e.g. "pct25", "diagCross". */
  readonly preset: string;
  readonly foregroundColor: Color;
  readonly backgroundColor: Color;
}

/** §20.1.8.14, a:blipFill. */
export interface BlipFill {
  readonly type: 'blip';
  readonly image: MediaPart;
  readonly opacity?: Percentage;
  readonly stretch?: boolean;
  readonly tile?: boolean;
}

export type Fill = NoFill | SolidFill | GradientFill | PatternFill | BlipFill;
