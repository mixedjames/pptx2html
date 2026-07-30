import type { Geometry, PresetGeometry } from '@pptx2html/presentation';

/**
 * Preset shape outlines (§20.1.9.18, a:prstGeom), scoped to the common subset a real deck is
 * overwhelmingly likely to use — the full ST_ShapeType enumeration has ~180 names. `rect`,
 * `roundRect` and `ellipse` are deliberately **not** handled here even though they're presets
 * too: a plain `<div>` with CSS `border-radius` already draws them exactly, including a border
 * that correctly follows the rounded outline — see `nativeBorderRadius` below and its use in
 * `shape-tree.ts`. This module only covers shapes CSS genuinely can't express as a rectangle
 * (however rounded) with a `border`, where an SVG `<path>` overlay is the only option.
 *
 * Adjustment-guide handling is intentionally approximate: `packages/presentation`'s `ShapeGuide`
 * only preserves a guide's literal `val N` override (no formula evaluator — see its own doc
 * comment), and the defaults below are reasonable stand-ins for each preset's usual look, not
 * transcribed from the spec's own `<gdLst>` formulas. Good enough for the common case of an
 * unadjusted or lightly-adjusted shape, in the same spirit as this package's other approximations
 * (pattern fills, colour-transform ordering — see `color.ts`/`fill.ts`).
 */

const CENTER = 50;
const RADIUS = 50;

function round(n: number): number {
  return Math.round(n * 1000) / 1000;
}

/** A point on a circle of the given radius around the shape's own center, 0deg = top, clockwise. */
function pointOnCircle(radius: number, angleDeg: number): readonly [number, number] {
  const angleRad = (angleDeg * Math.PI) / 180;
  return [round(CENTER + radius * Math.sin(angleRad)), round(CENTER - radius * Math.cos(angleRad))];
}

function polygonPath(points: readonly (readonly [number, number])[]): string {
  return `${points.map(([x, y], i) => `${i === 0 ? 'M' : 'L'} ${x} ${y}`).join(' ')} Z`;
}

function regularPolygonPath(sides: number): string {
  const points: (readonly [number, number])[] = [];
  for (let i = 0; i < sides; i++) points.push(pointOnCircle(RADIUS, (360 / sides) * i));
  return polygonPath(points);
}

function starPath(points: number, innerRatio: number): string {
  const innerRadius = RADIUS * innerRatio;
  const vertices: (readonly [number, number])[] = [];
  for (let i = 0; i < points * 2; i++) {
    const angle = (360 / (points * 2)) * i;
    vertices.push(pointOnCircle(i % 2 === 0 ? RADIUS : innerRadius, angle));
  }
  return polygonPath(vertices);
}

/** Default `adj` guide value (raw 0..100000 scale, same as `ShapeGuide.value`) for presets whose
 * `avLst` is absent — one adjustment guide named "adj" per preset, the common case for all
 * presets below. */
const DEFAULT_ADJUST_VALUE: Partial<Record<string, number>> = {
  triangle: 50000,
  parallelogram: 25000,
  trapezoid: 25000,
  hexagon: 25000,
  octagon: 25000,
  star5: 38200, // inner-point radius as a fraction of the outer radius; not a spec-exact constant.
};

function adjustFraction(geometry: PresetGeometry): number {
  const guide = geometry.adjustValues?.find((g) => g.name === 'adj');
  const raw = guide?.value ?? DEFAULT_ADJUST_VALUE[geometry.preset] ?? 25000;
  return Math.min(Math.max(raw, 0), 100000) / 100000;
}

/**
 * Renders a preset geometry as an SVG path's `d` attribute, in a fixed `0 0 100 100` coordinate
 * space (the caller stretches it non-uniformly onto the shape's actual box via
 * `preserveAspectRatio="none"`, mirroring how every position/size elsewhere in this package is a
 * percentage of its box rather than an absolute unit). Returns `undefined` for `rect`/`roundRect`/
 * `ellipse` (handled natively, see `nativeBorderRadius`) and any preset outside the common subset
 * this module models — the caller falls back to rendering a plain rectangle, today's behaviour.
 */
export function presetShapePath(geometry: Geometry): string | undefined {
  if (geometry.type !== 'preset') return undefined;
  const adjustPercent = () => adjustFraction(geometry) * 100;

  switch (geometry.preset) {
    case 'triangle':
      return polygonPath([
        [adjustPercent(), 0],
        [100, 100],
        [0, 100],
      ]);
    case 'rtTriangle':
      return polygonPath([
        [0, 0],
        [0, 100],
        [100, 100],
      ]);
    case 'diamond':
      return polygonPath([
        [50, 0],
        [100, 50],
        [50, 100],
        [0, 50],
      ]);
    case 'parallelogram': {
      const slant = Math.min(adjustPercent(), 50);
      return polygonPath([
        [slant, 0],
        [100, 0],
        [100 - slant, 100],
        [0, 100],
      ]);
    }
    case 'trapezoid': {
      const inset = Math.min(adjustPercent(), 50);
      return polygonPath([
        [inset, 0],
        [100 - inset, 0],
        [100, 100],
        [0, 100],
      ]);
    }
    case 'pentagon':
      return regularPolygonPath(5);
    case 'hexagon': {
      const inset = Math.min(adjustPercent(), 50);
      return polygonPath([
        [inset, 0],
        [100 - inset, 0],
        [100, 50],
        [100 - inset, 100],
        [inset, 100],
        [0, 50],
      ]);
    }
    case 'octagon': {
      const cut = Math.min(adjustPercent(), 50);
      return polygonPath([
        [cut, 0],
        [100 - cut, 0],
        [100, cut],
        [100, 100 - cut],
        [100 - cut, 100],
        [cut, 100],
        [0, 100 - cut],
        [0, cut],
      ]);
    }
    case 'star5':
      return starPath(5, adjustFraction(geometry));
    default:
      return undefined;
  }
}

/** Spec's commonly-cited default `roundRect` corner radius: 1/6 of the shorter side. Not
 * guaranteed bit-exact, but a reasonable, widely-used stand-in absent a formula evaluator. */
const DEFAULT_ROUND_RECT_ADJUST = 16667;

/**
 * CSS `border-radius` for the presets a plain rectangle-with-rounded-corners already renders
 * exactly — `roundRect` (radius = `adj` as a fraction of the box, capped at 50% so `adj`'s max of
 * 50000 yields a full stadium/pill, matching the preset's own behaviour at that extreme) and
 * `ellipse` (a fixed 50%). CSS applies a percentage `border-radius` independently per axis (against
 * width horizontally, height vertically), so this already produces a correctly-elliptical corner
 * on a non-square box without any extra aspect-ratio math — the same approximation tier as
 * `presetShapePath`'s use of a non-uniformly-stretched `0 0 100 100` space. `undefined` for every
 * other preset (including ones `presetShapePath` doesn't cover) — the caller leaves `border-radius`
 * unset, today's plain-rectangle behaviour.
 */
export function nativeBorderRadius(geometry: Geometry): string | undefined {
  if (geometry.type !== 'preset') return undefined;
  switch (geometry.preset) {
    case 'roundRect': {
      const guide = geometry.adjustValues?.find((g) => g.name === 'adj');
      const raw = guide?.value ?? DEFAULT_ROUND_RECT_ADJUST;
      const fraction = Math.min(Math.max(raw, 0), 100000) / 100000;
      return `${Math.min(fraction * 100, 50)}%`;
    }
    case 'ellipse':
      return '50%';
    default:
      return undefined;
  }
}
