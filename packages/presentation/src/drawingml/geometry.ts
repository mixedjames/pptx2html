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

/** A point in an `a:path`'s own local coordinate space (§20.1.9.8) — arbitrary units scaled onto
 * the shape's EMU bounding box by that path's own `w`/`h`, not EMU itself; a consumer does that
 * scaling, this type just carries the raw numbers as authored. */
export interface PathPoint {
  readonly x: number;
  readonly y: number;
}

/** One drawing command within an `a:path`'s outline (§20.1.9.2–9.9), in document order. */
export type PathCommand =
  | { readonly type: 'moveTo'; readonly point: PathPoint }
  | { readonly type: 'lnTo'; readonly point: PathPoint }
  | { readonly type: 'quadBezTo'; readonly control: PathPoint; readonly point: PathPoint }
  | {
      readonly type: 'cubicBezTo';
      readonly control1: PathPoint;
      readonly control2: PathPoint;
      readonly point: PathPoint;
    }
  | {
      readonly type: 'arcTo';
      readonly widthRadius: number;
      readonly heightRadius: number;
      readonly startAngle: Angle;
      readonly swingAngle: Angle;
    }
  | { readonly type: 'close' };

/**
 * One `a:path` (§20.1.9.8) within a `custGeom`'s `pathLst` — its own local coordinate space
 * (`w`/`h`, generally distinct from the shape's own EMU bounding box; a consumer scales
 * `commands`' points onto that box) plus its ordered outline commands. A `custGeom` with more
 * than one `a:path` (e.g. a boolean "Subtract" result: an outer rectangle plus an inner cutout)
 * renders as multiple subpaths within the same outline, punching a hole under the standard
 * nonzero fill rule — no explicit fill-rule concept exists here or in the schema.
 */
export interface CustomGeometryPath {
  readonly width: number;
  readonly height: number;
  readonly commands: readonly PathCommand[];
}

/**
 * Freeform outline (§20.1.9.8, a:custGeom). `pathLst` carries one entry per `a:path` whose every
 * command point is a literal coordinate; a path containing any `gdLst`-guide-referenced
 * coordinate (formula-derived, not modeled — same "literal only" limitation as `ShapeGuide`'s own
 * doc comment) is dropped entirely rather than rendered with a corrupted outline, and `pathLst`
 * is omitted altogether if no path in the shape was fully literal.
 */
export interface CustomGeometry {
  readonly type: 'custom';
  readonly pathLst?: readonly CustomGeometryPath[];
}

export type Geometry = PresetGeometry | CustomGeometry;
