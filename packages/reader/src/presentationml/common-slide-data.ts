import type { Background, CommonSlideData, StyleMatrixReference } from '@pptx2html/presentation';

import { parseChildFill, type MediaResolver } from '../drawingml/fill.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, findChild } from '../xml/query.js';
import { parseShapeTree, parseStyleMatrixReference } from './shape-tree.js';

/** A part's own `p:bg` (§19.3.1.6, `EG_Background`'s two alternatives) — `bgPr` a direct fill,
 *  `bgRef` a reference into the theme's `bgFillStyleLst` (PowerPoint's own default whenever a
 *  part relies on the theme's background rather than an explicit one — see `CommonSlideData.
backgroundRef`'s own doc comment). At most one is ever actually present on a real part. */
function parseBackground(
  bgNode: XmlNode | undefined,
  resolveMedia: MediaResolver,
): { background?: Background; backgroundRef?: StyleMatrixReference } {
  if (!bgNode) return {};
  const bgPr = findChild(bgNode, 'bgPr');
  const fill = bgPr ? parseChildFill(bgPr, resolveMedia) : undefined;
  if (fill) return { background: { fill } };
  const backgroundRef = parseStyleMatrixReference(findChild(bgNode, 'bgRef'));
  return backgroundRef ? { backgroundRef } : {};
}

/** Parses p:cSld (§19.3.1.16), the content shared by every slide-like part. */
export function parseCommonSlideData(
  cSldNode: XmlNode,
  resolveMedia: MediaResolver,
): CommonSlideData {
  const name = attr(cSldNode, 'name');
  const { background, backgroundRef } = parseBackground(findChild(cSldNode, 'bg'), resolveMedia);
  const spTree = findChild(cSldNode, 'spTree');

  return {
    ...(name ? { name } : {}),
    ...(background ? { background } : {}),
    ...(backgroundRef ? { backgroundRef } : {}),
    shapeTree: spTree ? parseShapeTree(spTree, resolveMedia) : [],
  };
}

export const EMPTY_COMMON_SLIDE_DATA: CommonSlideData = { shapeTree: [] };
