import type { Fill } from './fill.js';
import type { Emu } from './units.js';

export type LineCap = 'round' | 'square' | 'flat';

export type LineCompound = 'single' | 'double' | 'thickThin' | 'thinThick' | 'triple';

export type DashStyle =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'dashDot'
  | 'lgDash'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot'
  | 'sysDashDot'
  | 'sysDashDotDot';

export type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';

export interface LineEnd {
  readonly type: LineEndType;
  readonly width?: 'sm' | 'med' | 'lg';
  readonly length?: 'sm' | 'med' | 'lg';
}

/** Outline properties (§20.1.2.2.24, a:ln). */
export interface Line {
  readonly width?: Emu;
  readonly fill?: Fill;
  readonly cap?: LineCap;
  readonly compound?: LineCompound;
  readonly dashStyle?: DashStyle;
  readonly headEnd?: LineEnd;
  readonly tailEnd?: LineEnd;
}
