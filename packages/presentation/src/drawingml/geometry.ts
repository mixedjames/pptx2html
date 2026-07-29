import type { Angle, Emu } from './units.js';

export interface Point2D {
  readonly x: Emu;
  readonly y: Emu;
}

export interface Size2D {
  readonly width: Emu;
  readonly height: Emu;
}

/** 2D transform (§20.1.7.6, a:xfrm): position, size, rotation and flipping of a shape or group. */
export interface Transform2D {
  readonly offset: Point2D;
  readonly extents: Size2D;
  readonly rotation?: Angle;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
  /** Coordinate space children are laid out in, for group shapes only (chOff/chExt). */
  readonly childOffset?: Point2D;
  readonly childExtents?: Size2D;
}

/** A named adjustment handle value for a preset geometry (§20.1.9.5, a:gd). */
export interface ShapeGuide {
  readonly name: string;
  readonly value: number;
}

/** Built-in autoshape outline (§20.1.9.18, a:prstGeom). */
export interface PresetGeometry {
  readonly type: 'preset';
  /** Preset shape name, e.g. "rect", "ellipse", "roundRect". */
  readonly preset: string;
  readonly adjustValues?: readonly ShapeGuide[];
}

/** Freeform outline (§20.1.9.8, a:custGeom). Path data is unmodeled for the skeleton. */
export interface CustomGeometry {
  readonly type: 'custom';
}

export type Geometry = PresetGeometry | CustomGeometry;
