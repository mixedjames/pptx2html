import type { Background, CommonSlideData } from '@pptx2html/presentation';

import { parseChildFill, type MediaResolver } from '../drawingml/fill.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, findChild } from '../xml/query.js';
import { parseShapeTree } from './shape-tree.js';

function parseBackground(
  bgNode: XmlNode | undefined,
  resolveMedia: MediaResolver,
): Background | undefined {
  if (!bgNode) return undefined;
  // bgPr holds a direct fill; a bgRef (style-matrix reference) is unmodeled for the skeleton.
  const bgPr = findChild(bgNode, 'bgPr');
  const fill = bgPr ? parseChildFill(bgPr, resolveMedia) : undefined;
  return fill ? { fill } : undefined;
}

/** Parses p:cSld (§19.3.1.16), the content shared by every slide-like part. */
export function parseCommonSlideData(
  cSldNode: XmlNode,
  resolveMedia: MediaResolver,
): CommonSlideData {
  const name = attr(cSldNode, 'name');
  const background = parseBackground(findChild(cSldNode, 'bg'), resolveMedia);
  const spTree = findChild(cSldNode, 'spTree');

  return {
    ...(name ? { name } : {}),
    ...(background ? { background } : {}),
    shapeTree: spTree ? parseShapeTree(spTree, resolveMedia) : [],
  };
}

export const EMPTY_COMMON_SLIDE_DATA: CommonSlideData = { shapeTree: [] };
