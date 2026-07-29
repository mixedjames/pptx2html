import type { Geometry, Point2D, ShapeGuide, Size2D, Transform2D } from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, findChild, localName } from '../xml/query.js';
import { parseAngle, parseBoolean, parseEmu } from './units.js';

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
    case 'custGeom':
      return { type: 'custom' };
    default:
      return undefined;
  }
}

/** Finds and parses a shape properties node's geometry child, whichever variant is present. */
export function findChildGeometry(spPrNode: XmlNode): Geometry | undefined {
  return parseGeometry(findChild(spPrNode, 'prstGeom') ?? findChild(spPrNode, 'custGeom'));
}
