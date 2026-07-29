import type {
  NonVisualDrawingProperties,
  Placeholder,
  PlaceholderType,
  ShapeProperties,
} from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, findChild } from '../xml/query.js';
import { findChildGeometry, parseTransform } from './geometry.js';
import { parseChildFill, type MediaResolver } from './fill.js';
import { parseLine } from './line.js';
import { parseBoolean, parseIntAttr } from './units.js';

const PLACEHOLDER_TYPES: ReadonlySet<string> = new Set<PlaceholderType>([
  'title',
  'body',
  'ctrTitle',
  'subTitle',
  'dt',
  'ftr',
  'sldNum',
  'sldImg',
  'pic',
  'chart',
  'tbl',
  'clipArt',
  'dgm',
  'media',
  'obj',
]);

/** Parses a p:ph element (§19.3.1.36); `type` and `idx` both default to their spec values when omitted. */
function parsePlaceholder(phNode: XmlNode): Placeholder {
  const typeValue = attr(phNode, 'type');
  const type: PlaceholderType =
    typeValue && PLACEHOLDER_TYPES.has(typeValue) ? (typeValue as PlaceholderType) : 'obj';
  const index = parseIntAttr(attr(phNode, 'idx')) ?? 0;
  return { type, index };
}

/**
 * Parses a shape's non-visual identity from its nv*Pr wrapper (e.g. p:nvSpPr), which holds the
 * cNvPr element (§20.1.2.2.8/2.2.20/2.2.29) and, for placeholder shapes, an nvPr/ph child
 * (§19.3.1.36) identifying which layout/master placeholder it corresponds to.
 */
export function parseNonVisualDrawingProperties(
  nvPrNode: XmlNode | undefined,
): NonVisualDrawingProperties {
  const cNvPr = nvPrNode ? findChild(nvPrNode, 'cNvPr') : undefined;
  const id = parseIntAttr(cNvPr ? attr(cNvPr, 'id') : undefined) ?? 0;
  const name = (cNvPr ? attr(cNvPr, 'name') : undefined) ?? '';
  const description = cNvPr ? attr(cNvPr, 'descr') : undefined;
  const hidden = cNvPr ? parseBoolean(attr(cNvPr, 'hidden')) : undefined;

  const nvPr = nvPrNode ? findChild(nvPrNode, 'nvPr') : undefined;
  const phNode = nvPr ? findChild(nvPr, 'ph') : undefined;
  const placeholder = phNode ? parsePlaceholder(phNode) : undefined;

  return {
    id,
    name,
    ...(description ? { description } : {}),
    ...(hidden !== undefined ? { hidden } : {}),
    ...(placeholder ? { placeholder } : {}),
  };
}

/** Parses a shape/picture/connector's visual properties from its spPr (§20.1.2.2.35). */
export function parseShapeProperties(
  spPrNode: XmlNode | undefined,
  resolveMedia: MediaResolver,
): ShapeProperties {
  if (!spPrNode) return {};

  const transform = parseTransform(findChild(spPrNode, 'xfrm'));
  const geometry = findChildGeometry(spPrNode);
  const fill = parseChildFill(spPrNode, resolveMedia);
  const lnNode = findChild(spPrNode, 'ln');
  const line = lnNode ? parseLine(lnNode, resolveMedia) : undefined;

  return {
    ...(transform ? { transform } : {}),
    ...(geometry ? { geometry } : {}),
    ...(fill ? { fill } : {}),
    ...(line ? { line } : {}),
  };
}
