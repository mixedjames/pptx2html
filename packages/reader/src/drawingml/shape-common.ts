import type { NonVisualDrawingProperties, ShapeProperties } from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, findChild } from '../xml/query.js';
import { findChildGeometry, parseTransform } from './geometry.js';
import { parseChildFill, type MediaResolver } from './fill.js';
import { parseLine } from './line.js';
import { parseBoolean, parseIntAttr } from './units.js';

/**
 * Parses a shape's non-visual identity from its nv*Pr wrapper (e.g. p:nvSpPr), which holds the
 * cNvPr element (§20.1.2.2.8/2.2.20/2.2.29).
 */
export function parseNonVisualDrawingProperties(
  nvPrNode: XmlNode | undefined,
): NonVisualDrawingProperties {
  const cNvPr = nvPrNode ? findChild(nvPrNode, 'cNvPr') : undefined;
  const id = parseIntAttr(cNvPr ? attr(cNvPr, 'id') : undefined) ?? 0;
  const name = (cNvPr ? attr(cNvPr, 'name') : undefined) ?? '';
  const description = cNvPr ? attr(cNvPr, 'descr') : undefined;
  const hidden = cNvPr ? parseBoolean(attr(cNvPr, 'hidden')) : undefined;

  return {
    id,
    name,
    ...(description ? { description } : {}),
    ...(hidden !== undefined ? { hidden } : {}),
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
