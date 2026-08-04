import type {
  CustomGeometryPath,
  Geometry,
  PathPoint,
  PresetGeometry,
} from '@pptx2html/presentation';

/**
 * Preset shape outlines (§20.1.9.18, a:prstGeom), scoped to the common subset a real deck is
 * overwhelmingly likely to use — the full ST_ShapeType enumeration has ~180 names. `rect`,
 * `roundRect` and `ellipse` are deliberately **not** handled here even though they're presets
 * too: a plain `<div>` with CSS `border-radius` already draws them exactly, including a border
 * that correctly follows the rounded outline — see `nativeBorderRadius` below and its use in
 * `shape-tree.ts`. This module only covers shapes CSS genuinely can't express as a rectangle
 * (however rounded) with a `border`, where an SVG `<path>` overlay is the only option — which is
 * also why freeform (`a:custGeom`) outlines live here too, see `customGeometryPath` below.
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

/** A point in an `a:path`'s own local coordinate space, expressed as a percentage of that path's
 * `w`/`h` — the same `0 0 100 100` space `presetShapePath` uses, so a `custGeom` outline stretches
 * onto the shape's actual box the identical non-uniform way a preset one does. */
function scalePoint(point: PathPoint, path: CustomGeometryPath): readonly [number, number] {
  return [round((point.x / path.width) * 100), round((point.y / path.height) * 100)];
}

/**
 * Renders one `a:path`'s commands as an SVG subpath (`M`/`L`/`Q`/`C`/`A`/`Z`), scaled into the
 * `0 0 100 100` space via `scalePoint`. `arcTo` (§20.1.9.9) is the one command whose SVG
 * translation isn't a direct 1:1 mapping: OOXML defines it as a portion of an ellipse (radii
 * `wR`/`hR`) that the current pen position sits on at angle `stAng`, swinging `swAng` further —
 * so the ellipse's centre and end point are derived here from the current pen position before
 * being handed to SVG's own `A rx ry x-axis-rotation large-arc-flag sweep-flag x y` command
 * (rotation is always 0: neither coordinate system rotates the ellipse relative to the path's own
 * axes). Non-uniform axis scaling (a path whose `w` and `h` differ) maps an ellipse to another
 * ellipse under this transform, so scaling the two radii independently by the same per-axis
 * factor as any other point stays correct.
 */
function buildSubpathD(path: CustomGeometryPath): string {
  let current: PathPoint = { x: 0, y: 0 };
  const parts: string[] = [];
  for (const command of path.commands) {
    switch (command.type) {
      case 'moveTo': {
        const [x, y] = scalePoint(command.point, path);
        parts.push(`M ${x} ${y}`);
        current = command.point;
        break;
      }
      case 'lnTo': {
        const [x, y] = scalePoint(command.point, path);
        parts.push(`L ${x} ${y}`);
        current = command.point;
        break;
      }
      case 'quadBezTo': {
        const [cx, cy] = scalePoint(command.control, path);
        const [x, y] = scalePoint(command.point, path);
        parts.push(`Q ${cx} ${cy} ${x} ${y}`);
        current = command.point;
        break;
      }
      case 'cubicBezTo': {
        const [c1x, c1y] = scalePoint(command.control1, path);
        const [c2x, c2y] = scalePoint(command.control2, path);
        const [x, y] = scalePoint(command.point, path);
        parts.push(`C ${c1x} ${c1y} ${c2x} ${c2y} ${x} ${y}`);
        current = command.point;
        break;
      }
      case 'arcTo': {
        const { widthRadius, heightRadius, startAngle, swingAngle } = command;
        const startRad = (startAngle / 60000) * (Math.PI / 180);
        const endRad = ((startAngle + swingAngle) / 60000) * (Math.PI / 180);
        const centerX = current.x - widthRadius * Math.cos(startRad);
        const centerY = current.y - heightRadius * Math.sin(startRad);
        const endPoint: PathPoint = {
          x: centerX + widthRadius * Math.cos(endRad),
          y: centerY + heightRadius * Math.sin(endRad),
        };
        const rx = round((Math.abs(widthRadius) / path.width) * 100);
        const ry = round((Math.abs(heightRadius) / path.height) * 100);
        const [x, y] = scalePoint(endPoint, path);
        const largeArcFlag = Math.abs(swingAngle) > 180 * 60000 ? 1 : 0;
        const sweepFlag = swingAngle >= 0 ? 1 : 0;
        parts.push(`A ${rx} ${ry} 0 ${largeArcFlag} ${sweepFlag} ${x} ${y}`);
        current = endPoint;
        break;
      }
      case 'close':
        parts.push('Z');
        break;
    }
  }
  return parts.join(' ');
}

/**
 * Renders a `custGeom`'s parsed path data (`packages/presentation`'s `CustomGeometry.pathLst`) as
 * an SVG path's `d` attribute, in the same `0 0 100 100` non-uniformly-stretched space
 * `presetShapePath` uses. Multiple `a:path` entries concatenate into multiple subpaths within one
 * `d` string, which SVG's default nonzero fill rule renders as a hole wherever they overlap with
 * opposite winding — exactly what a boolean "Subtract" shape (an outer outline plus an inner
 * cutout, PowerPoint's own `custGeom` output for Merge Shapes operations) needs, with no separate
 * fill-rule handling required. Returns `undefined` when `geometry` carries no usable path data at
 * all (either genuinely absent, or every `a:path` was dropped by the reader — see
 * `CustomGeometryPath`'s own doc comment) or every path present has a zero `w`/`h` (nothing to
 * scale against) — the caller falls back to a plain rectangle in that case, same as an unmodeled
 * preset.
 */
export function customGeometryPath(geometry: Geometry): string | undefined {
  if (geometry.type !== 'custom' || !geometry.pathLst) return undefined;
  const subpaths = geometry.pathLst
    .filter((path) => path.width > 0 && path.height > 0)
    .map(buildSubpathD)
    .filter((d) => d.length > 0);
  return subpaths.length > 0 ? subpaths.join(' ') : undefined;
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
