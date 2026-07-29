import type { Slide } from '@pptx2html/presentation';

import { createMediaResolver } from '../drawingml/media.js';
import { parseBoolean } from '../drawingml/units.js';
import type { ReaderContext } from '../reader-context.js';
import { parseXml } from '../xml/parse.js';
import { attr, findRoot, findChild } from '../xml/query.js';
import { EMPTY_COMMON_SLIDE_DATA, parseCommonSlideData } from './common-slide-data.js';
import { RELATIONSHIP_TYPES } from './relationship-types.js';

/**
 * Parses a slide part (§19.3.1.38, p:sld). Requires that `readSlideMaster` has already been
 * called for every slide master listed in presentation.xml, since that is what populates
 * `context.slideLayouts` with the SlideLayout instance this slide's relationship points at.
 */
export function readSlide(context: ReaderContext, partName: string): Slide {
  const cached = context.slides.get(partName);
  if (cached) return cached;

  const root = findRoot(parseXml(context.package.readText(partName)), 'sld');
  const cSld = root && findChild(root, 'cSld');
  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const commonSlideData = cSld ? parseCommonSlideData(cSld, resolveMedia) : EMPTY_COMMON_SLIDE_DATA;

  const relationships = context.package.relationshipsFor(partName);
  const layoutRel = relationships.findByType(RELATIONSHIP_TYPES.slideLayout)[0];
  const layout = layoutRel ? context.slideLayouts.get(layoutRel.target) : undefined;
  if (!layout) {
    throw new Error(`Slide part "${partName}" does not resolve to a known slide layout`);
  }

  const showMasterShapes = root ? parseBoolean(attr(root, 'showMasterSp')) : undefined;

  const slide: Slide = {
    commonSlideData,
    layout,
    ...(showMasterShapes !== undefined ? { showMasterShapes } : {}),
  };
  context.slides.set(partName, slide);
  return slide;
}
