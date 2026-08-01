import type { Point2D, Size2D, Transform2D } from '../drawingml/index.js';

/**
 * Affine map (translate + scale, no rotation) from a local EMU coordinate space to the slide's
 * root EMU coordinate space. Rotation is intentionally excluded from composition — a renderer
 * applies a shape's (or group's) own rotation/flip separately, around its own resolved box (see
 * `computeBox`), so a group's rotation does not compose into its children's own rotation. Good
 * enough for laying elements out; not spec-accurate for rotated groups.
 */
export interface CoordinateMap {
  readonly offsetX: number;
  readonly offsetY: number;
  readonly scaleX: number;
  readonly scaleY: number;
}

export const IDENTITY_MAP: CoordinateMap = { offsetX: 0, offsetY: 0, scaleX: 1, scaleY: 1 };

export function mapPoint(map: CoordinateMap, point: Point2D): Point2D {
  return { x: map.offsetX + point.x * map.scaleX, y: map.offsetY + point.y * map.scaleY };
}

export function mapSize(map: CoordinateMap, size: Size2D): Size2D {
  return { width: size.width * map.scaleX, height: size.height * map.scaleY };
}

/**
 * Extends a coordinate map with a group shape's transform, producing the map that the group's
 * children's own transforms — expressed in the group's child coordinate space, chOff/chExt —
 * should be read against (§20.1.7.6, a:xfrm).
 */
export function composeGroupMap(map: CoordinateMap, transform: Transform2D): CoordinateMap {
  const childExtents = transform.childExtents ?? transform.extents;
  const childOffset = transform.childOffset ?? { x: 0, y: 0 };
  const groupScaleX = childExtents.width === 0 ? 1 : transform.extents.width / childExtents.width;
  const groupScaleY =
    childExtents.height === 0 ? 1 : transform.extents.height / childExtents.height;

  const scaleX = map.scaleX * groupScaleX;
  const scaleY = map.scaleY * groupScaleY;

  return {
    scaleX,
    scaleY,
    offsetX: map.offsetX + transform.offset.x * map.scaleX - childOffset.x * scaleX,
    offsetY: map.offsetY + transform.offset.y * map.scaleY - childOffset.y * scaleY,
  };
}

/**
 * A shape's position/size in slide-root EMU coordinates (i.e. already run through a
 * `CoordinateMap`, but not yet converted to any renderer-specific unit), plus its own (uncomposed)
 * rotation/flip. Deliberately left unit-less in EMU rather than any renderer's own unit so callers
 * can turn it into whatever's appropriate — e.g. `@pptx2html/to-html5`'s `shape-tree.ts` expresses
 * it as a percentage of the slide's own size so the whole slide scales with its container.
 */
export interface ElementBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly rotationDeg?: number;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}

export function computeBox(map: CoordinateMap, transform: Transform2D): ElementBox {
  const pos = mapPoint(map, transform.offset);
  const size = mapSize(map, transform.extents);
  return {
    left: pos.x,
    top: pos.y,
    width: size.width,
    height: size.height,
    rotationDeg: transform.rotation === undefined ? undefined : transform.rotation / 60000,
    flipHorizontal: transform.flipHorizontal,
    flipVertical: transform.flipVertical,
  };
}
