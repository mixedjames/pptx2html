import type {
  ConnectionShape,
  FontCollectionIndex,
  FontReference,
  GraphicFrame,
  GraphicPlaceholder,
  GroupShape,
  Picture,
  Shape,
  ShapeConnection,
  ShapeStyle,
  ShapeTreeNode,
  StyleMatrixReference,
} from '@pptx2html/presentation';

import { parseChildColor } from '../drawingml/color.js';
import { blipEmbedId, type MediaResolver } from '../drawingml/fill.js';
import { parseTransform } from '../drawingml/geometry.js';
import {
  parseNonVisualDrawingProperties,
  parseShapeProperties,
} from '../drawingml/shape-common.js';
import { parseTextBody } from '../drawingml/text.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, children, findChild, localName } from '../xml/query.js';
import { parseTable } from './table.js';

function parseShapeConnection(node: XmlNode | undefined): ShapeConnection | undefined {
  if (!node) return undefined;
  const shapeId = Number.parseInt(attr(node, 'id') ?? '', 10);
  const connectionSiteIndex = Number.parseInt(attr(node, 'idx') ?? '', 10);
  if (Number.isNaN(shapeId) || Number.isNaN(connectionSiteIndex)) return undefined;
  return { shapeId, connectionSiteIndex };
}

function parseStyleMatrixReference(node: XmlNode | undefined): StyleMatrixReference | undefined {
  if (!node) return undefined;
  const index = Number.parseInt(attr(node, 'idx') ?? '', 10);
  const color = parseChildColor(node);
  if (Number.isNaN(index) || !color) return undefined;
  return { index, color };
}

const FONT_COLLECTION_INDICES: ReadonlySet<string> = new Set(['major', 'minor', 'none']);

function parseFontReference(node: XmlNode | undefined): FontReference | undefined {
  if (!node) return undefined;
  const idx = attr(node, 'idx');
  const color = parseChildColor(node);
  if (!idx || !FONT_COLLECTION_INDICES.has(idx) || !color) return undefined;
  return { collection: idx as FontCollectionIndex, color };
}

/** Parses p:style's fillRef/lnRef/fontRef (§19.3.1.44) — effectRef is unmodeled, see
 * `ShapeStyle`'s own doc comment. */
function parseShapeStyle(node: XmlNode | undefined): ShapeStyle | undefined {
  if (!node) return undefined;
  const fillRef = parseStyleMatrixReference(findChild(node, 'fillRef'));
  const lineRef = parseStyleMatrixReference(findChild(node, 'lnRef'));
  const fontRef = parseFontReference(findChild(node, 'fontRef'));
  if (!fillRef && !lineRef && !fontRef) return undefined;
  return {
    ...(fillRef ? { fillRef } : {}),
    ...(lineRef ? { lineRef } : {}),
    ...(fontRef ? { fontRef } : {}),
  };
}

function parseShape(node: XmlNode, resolveMedia: MediaResolver): Shape {
  const textBody = parseTextBody(findChild(node, 'txBody'), resolveMedia);
  const style = parseShapeStyle(findChild(node, 'style'));
  return {
    kind: 'shape',
    nonVisual: parseNonVisualDrawingProperties(findChild(node, 'nvSpPr')),
    properties: parseShapeProperties(findChild(node, 'spPr'), resolveMedia),
    ...(style ? { style } : {}),
    ...(textBody ? { textBody } : {}),
  };
}

function parsePicture(node: XmlNode, resolveMedia: MediaResolver): Picture | undefined {
  const blipFill = findChild(node, 'blipFill');
  const blip = blipFill ? findChild(blipFill, 'blip') : undefined;
  const embed = blip ? blipEmbedId(blip) : undefined;
  const image = embed ? resolveMedia(embed) : undefined;
  if (!image) return undefined;

  const style = parseShapeStyle(findChild(node, 'style'));
  return {
    kind: 'picture',
    nonVisual: parseNonVisualDrawingProperties(findChild(node, 'nvPicPr')),
    properties: parseShapeProperties(findChild(node, 'spPr'), resolveMedia),
    ...(style ? { style } : {}),
    image,
  };
}

function parseConnectionShape(node: XmlNode, resolveMedia: MediaResolver): ConnectionShape {
  const startConnection = parseShapeConnection(findChild(node, 'stCxn'));
  const endConnection = parseShapeConnection(findChild(node, 'endCxn'));
  const style = parseShapeStyle(findChild(node, 'style'));
  return {
    kind: 'connector',
    nonVisual: parseNonVisualDrawingProperties(findChild(node, 'nvCxnSpPr')),
    properties: parseShapeProperties(findChild(node, 'spPr'), resolveMedia),
    ...(style ? { style } : {}),
    ...(startConnection ? { startConnection } : {}),
    ...(endConnection ? { endConnection } : {}),
  };
}

const GRAPHIC_TYPE_BY_URI_FRAGMENT: ReadonlyArray<readonly [string, GraphicPlaceholder['type']]> = [
  ['/chart', 'chart'],
  ['/diagram', 'smartArt'],
  ['/oleObject', 'oleObject'],
];

function matchGraphicType(uri: string): GraphicPlaceholder['type'] {
  const match = GRAPHIC_TYPE_BY_URI_FRAGMENT.find(([fragment]) => uri.includes(fragment));
  return match ? match[1] : 'unknown';
}

function parseGraphicFrame(node: XmlNode, resolveMedia: MediaResolver): GraphicFrame | undefined {
  const transform = parseTransform(findChild(node, 'xfrm'));
  if (!transform) return undefined;

  const graphicNode = findChild(node, 'graphic');
  const graphicDataNode = graphicNode ? findChild(graphicNode, 'graphicData') : undefined;
  if (!graphicDataNode) return undefined;

  const tbl = findChild(graphicDataNode, 'tbl');
  const graphic: GraphicFrame['graphic'] = tbl
    ? parseTable(tbl, resolveMedia)
    : { type: matchGraphicType(attr(graphicDataNode, 'uri') ?? '') };

  return {
    kind: 'graphicFrame',
    nonVisual: parseNonVisualDrawingProperties(findChild(node, 'nvGraphicFramePr')),
    transform,
    graphic,
  };
}

function parseGroupShape(node: XmlNode, resolveMedia: MediaResolver): GroupShape | undefined {
  const grpSpPr = findChild(node, 'grpSpPr');
  const transform = parseTransform(grpSpPr ? findChild(grpSpPr, 'xfrm') : undefined);
  if (!transform) return undefined;

  return {
    kind: 'group',
    nonVisual: parseNonVisualDrawingProperties(findChild(node, 'nvGrpSpPr')),
    transform,
    children: parseShapeTree(node, resolveMedia),
  };
}

function parseShapeTreeChild(
  node: XmlNode,
  resolveMedia: MediaResolver,
): ShapeTreeNode | undefined {
  switch (localName(node)) {
    case 'sp':
      return parseShape(node, resolveMedia);
    case 'pic':
      return parsePicture(node, resolveMedia);
    case 'cxnSp':
      return parseConnectionShape(node, resolveMedia);
    case 'graphicFrame':
      return parseGraphicFrame(node, resolveMedia);
    case 'grpSp':
      return parseGroupShape(node, resolveMedia);
    default:
      return undefined;
  }
}

/**
 * Parses a p:spTree's or p:grpSp's shape children (§19.3.1.22/1.32), in document order — this
 * order is the shapes' paint/z-order, which is why xml/query.ts preserves child element order.
 */
export function parseShapeTree(
  spTreeOrGrpSpNode: XmlNode,
  resolveMedia: MediaResolver,
): readonly ShapeTreeNode[] {
  return children(spTreeOrGrpSpNode)
    .map((child) => parseShapeTreeChild(child, resolveMedia))
    .filter((node): node is ShapeTreeNode => node !== undefined);
}
