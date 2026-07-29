import type { SlideLayout, SlideLayoutType, SlideMaster } from '@pptx2html/presentation';

import { createMediaResolver } from '../drawingml/media.js';
import { parseBoolean } from '../drawingml/units.js';
import type { ReaderContext } from '../reader-context.js';
import { parseXml } from '../xml/parse.js';
import { attr, findChild, findRoot } from '../xml/query.js';
import { EMPTY_COMMON_SLIDE_DATA, parseCommonSlideData } from './common-slide-data.js';

const SLIDE_LAYOUT_TYPES: ReadonlySet<string> = new Set<SlideLayoutType>([
  'title',
  'tx',
  'twoColTx',
  'tbl',
  'txAndChart',
  'chartAndTx',
  'dgm',
  'chart',
  'txAndClipArt',
  'clipArtAndTx',
  'titleOnly',
  'blank',
  'txAndObj',
  'objAndTx',
  'objOnly',
  'obj',
  'txAndMedia',
  'mediaAndTx',
  'objOverTx',
  'txOverObj',
  'txAndTwoObj',
  'twoObjAndTx',
  'twoObjOverTx',
  'fourObj',
  'vertTx',
  'clipArtAndVertTx',
  'vertTitleAndTx',
  'vertTitleAndTxOverChart',
  'twoObj',
  'objAndTwoObj',
  'twoTxTwoObj',
  'secHead',
  'objTx',
  'picTx',
  'cust',
]);

/**
 * Parses a slide layout part (§19.3.1.39, p:sldLayout). Called while building its owning
 * SlideMaster, which must already exist (the mutual master<->layouts reference is resolved
 * by the caller) — see slide-master.ts.
 */
export function parseSlideLayout(
  context: ReaderContext,
  partName: string,
  master: SlideMaster,
): SlideLayout {
  const root = findRoot(parseXml(context.package.readText(partName)), 'sldLayout');
  const cSld = root && findChild(root, 'cSld');
  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const commonSlideData = cSld ? parseCommonSlideData(cSld, resolveMedia) : EMPTY_COMMON_SLIDE_DATA;

  const typeValue = root ? attr(root, 'type') : undefined;
  const type: SlideLayoutType =
    typeValue && SLIDE_LAYOUT_TYPES.has(typeValue) ? (typeValue as SlideLayoutType) : 'cust';
  const matchingName = root ? attr(root, 'matchingName') : undefined;
  const showMasterShapes = root ? parseBoolean(attr(root, 'showMasterSp')) : undefined;

  return {
    commonSlideData,
    master,
    type,
    ...(matchingName ? { matchingName } : {}),
    ...(showMasterShapes !== undefined ? { showMasterShapes } : {}),
  };
}
