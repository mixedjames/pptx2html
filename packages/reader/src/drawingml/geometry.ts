import type {
  CustomGeometryPath,
  Geometry,
  PathCommand,
  PathPoint,
  Point2D,
  ShapeGuide,
  Size2D,
  Transform2D,
} from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, findAllChildren, findChild, localName } from '../xml/query.js';
import { parseAngle, parseBoolean, parseEmu, parseIntAttr } from './units.js';

function parsePoint(node: XmlNode | undefined): Point2D | undefined {
  if (!node) return undefined;
  const x = parseEmu(attr(node, 'x'));
  const y = parseEmu(attr(node, 'y'));
  return x !== undefined && y !== undefined ? { x, y } : undefined;
}

function parseSize(node: XmlNode | undefined): Size2D | undefined {
  if (!node) return undefined;
  const width = parseEmu(attr(node, 'cx'));
  const height = parseEmu(attr(node, 'cy'));
  return width !== undefined && height !== undefined ? { width, height } : undefined;
}

/** Parses a:xfrm (§20.1.7.6). */
export function parseTransform(node: XmlNode | undefined): Transform2D | undefined {
  if (!node) return undefined;
  const offset = parsePoint(findChild(node, 'off'));
  const extents = parseSize(findChild(node, 'ext'));
  if (!offset || !extents) return undefined;

  const rotation = parseAngle(attr(node, 'rot'));
  const flipHorizontal = parseBoolean(attr(node, 'flipH'));
  const flipVertical = parseBoolean(attr(node, 'flipV'));
  const childOffset = parsePoint(findChild(node, 'chOff'));
  const childExtents = parseSize(findChild(node, 'chExt'));

  return {
    offset,
    extents,
    ...(rotation !== undefined ? { rotation } : {}),
    ...(flipHorizontal !== undefined ? { flipHorizontal } : {}),
    ...(flipVertical !== undefined ? { flipVertical } : {}),
    ...(childOffset ? { childOffset } : {}),
    ...(childExtents ? { childExtents } : {}),
  };
}

/**
 * Parses a:avLst adjustment guides. Only the literal "val N" formula form is resolvable to a
 * plain number without a formula evaluator (the only form modeled by ShapeGuide); guides using
 * arithmetic formulas are skipped.
 */
function parseAdjustValues(node: XmlNode): readonly ShapeGuide[] | undefined {
  const avLst = findChild(node, 'avLst');
  if (!avLst) return undefined;
  const guides: ShapeGuide[] = [];
  for (const gd of children(avLst)) {
    if (localName(gd) !== 'gd') continue;
    const name = attr(gd, 'name');
    const fmla = attr(gd, 'fmla');
    if (!name || !fmla?.startsWith('val ')) continue;
    const value = Number.parseInt(fmla.slice('val '.length), 10);
    if (!Number.isNaN(value)) guides.push({ name, value });
  }
  return guides.length > 0 ? guides : undefined;
}

/** An `a:*` command node's `a:pt` children's `x`/`y`, in document order. Returns `[]` if any
 * `pt`'s coordinate isn't a plain literal integer (e.g. a `gdLst` guide-name reference) — the
 * caller treats that as "this command can't be represented", not "this command has no points". */
function parsePoints(node: XmlNode): readonly PathPoint[] {
  const points: PathPoint[] = [];
  for (const ptNode of findAllChildren(node, 'pt')) {
    const x = parseIntAttr(attr(ptNode, 'x'));
    const y = parseIntAttr(attr(ptNode, 'y'));
    if (x === undefined || y === undefined) return [];
    points.push({ x, y });
  }
  return points;
}

/**
 * Parses one outline-drawing command within an `a:path` (§20.1.9.2–9.9). Returns `undefined` for
 * a command this package doesn't model, or one whose points/attributes didn't fully resolve to
 * literal numbers (a `gdLst`-guide-referenced coordinate — see `CustomGeometryPath`'s own doc
 * comment) — the caller drops the whole enclosing path rather than render a corrupted outline.
 */
function parsePathCommand(node: XmlNode): PathCommand | undefined {
  switch (localName(node)) {
    case 'moveTo': {
      const points = parsePoints(node);
      return points.length === 1 ? { type: 'moveTo', point: points[0]! } : undefined;
    }
    case 'lnTo': {
      const points = parsePoints(node);
      return points.length === 1 ? { type: 'lnTo', point: points[0]! } : undefined;
    }
    case 'quadBezTo': {
      const points = parsePoints(node);
      return points.length === 2
        ? { type: 'quadBezTo', control: points[0]!, point: points[1]! }
        : undefined;
    }
    case 'cubicBezTo': {
      const points = parsePoints(node);
      return points.length === 3
        ? { type: 'cubicBezTo', control1: points[0]!, control2: points[1]!, point: points[2]! }
        : undefined;
    }
    case 'arcTo': {
      const widthRadius = parseIntAttr(attr(node, 'wR'));
      const heightRadius = parseIntAttr(attr(node, 'hR'));
      const startAngle = parseAngle(attr(node, 'stAng'));
      const swingAngle = parseAngle(attr(node, 'swAng'));
      return widthRadius !== undefined &&
        heightRadius !== undefined &&
        startAngle !== undefined &&
        swingAngle !== undefined
        ? { type: 'arcTo', widthRadius, heightRadius, startAngle, swingAngle }
        : undefined;
    }
    case 'close':
      return { type: 'close' };
    default:
      return undefined;
  }
}

/**
 * Parses one `a:path` within a `custGeom`'s `pathLst`. Returns `undefined` if `w`/`h` are missing
 * or any command within it didn't fully resolve (see `parsePathCommand`) — dropped by the caller
 * rather than folded in as a partial/corrupt outline.
 */
function parsePath(node: XmlNode): CustomGeometryPath | undefined {
  const width = parseIntAttr(attr(node, 'w'));
  const height = parseIntAttr(attr(node, 'h'));
  if (width === undefined || height === undefined) return undefined;
  const commands: PathCommand[] = [];
  for (const child of children(node)) {
    const command = parsePathCommand(child);
    if (!command) return undefined;
    commands.push(command);
  }
  return commands.length > 0 ? { width, height, commands } : undefined;
}

/** Parses a shape's outline geometry: a:prstGeom (§20.1.9.18) or a:custGeom (§20.1.9.8). */
export function parseGeometry(node: XmlNode | undefined): Geometry | undefined {
  if (!node) return undefined;
  switch (localName(node)) {
    case 'prstGeom': {
      const preset = attr(node, 'prst');
      if (!preset) return undefined;
      const adjustValues = parseAdjustValues(node);
      return { type: 'preset', preset, ...(adjustValues ? { adjustValues } : {}) };
    }
    case 'custGeom': {
      const pathLstNode = findChild(node, 'pathLst');
      const pathLst = pathLstNode
        ? findAllChildren(pathLstNode, 'path')
            .map(parsePath)
            .filter((path): path is CustomGeometryPath => path !== undefined)
        : [];
      return pathLst.length > 0 ? { type: 'custom', pathLst } : { type: 'custom' };
    }
    default:
      return undefined;
  }
}

/** Finds and parses a shape properties node's geometry child, whichever variant is present. */
export function findChildGeometry(spPrNode: XmlNode): Geometry | undefined {
  return parseGeometry(findChild(spPrNode, 'prstGeom') ?? findChild(spPrNode, 'custGeom'));
}
