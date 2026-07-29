import type { SlideMaster, Theme } from '@pptx2html/presentation';

import { createMediaResolver } from '../drawingml/media.js';
import type { Mutable } from '../mutable.js';
import type { ReaderContext } from '../reader-context.js';
import { emptyTheme, parseTheme } from '../theme.js';
import { parseXml } from '../xml/parse.js';
import { findChild, findRoot } from '../xml/query.js';
import { EMPTY_COMMON_SLIDE_DATA, parseCommonSlideData } from './common-slide-data.js';
import { RELATIONSHIP_TYPES } from './relationship-types.js';
import { parseSlideLayout } from './slide-layout.js';

function readTheme(context: ReaderContext, partName: string): Theme {
  const cached = context.themes.get(partName);
  if (cached) return cached;
  const theme = parseTheme(context.package.readText(partName));
  context.themes.set(partName, theme);
  return theme;
}

/**
 * Parses a slide master part (§19.3.1.44, p:sldMaster) and all the slide layouts it owns.
 *
 * SlideMaster.layouts and SlideLayout.master are mutually referential and every field involved
 * is readonly, so construction is two-phase: build the master with a temporary empty `layouts`,
 * parse each layout against that same master instance (so `layout.master === master` by
 * identity), then assign the finished layouts array and freeze.
 */
export function readSlideMaster(context: ReaderContext, partName: string): SlideMaster {
  const cached = context.slideMasters.get(partName);
  if (cached) return cached;

  const root = findRoot(parseXml(context.package.readText(partName)), 'sldMaster');
  const cSld = root && findChild(root, 'cSld');
  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const commonSlideData = cSld ? parseCommonSlideData(cSld, resolveMedia) : EMPTY_COMMON_SLIDE_DATA;

  const relationships = context.package.relationshipsFor(partName);
  const themeRel = relationships.findByType(RELATIONSHIP_TYPES.theme)[0];
  const theme = themeRel ? readTheme(context, themeRel.target) : emptyTheme();

  const master = {
    commonSlideData,
    theme,
    layouts: [],
  } as Mutable<SlideMaster>;
  context.slideMasters.set(partName, master);

  const layoutRels = relationships.findByType(RELATIONSHIP_TYPES.slideLayout);
  master.layouts = layoutRels.map((rel) => {
    const layout = parseSlideLayout(context, rel.target, master);
    context.slideLayouts.set(rel.target, layout);
    return layout;
  });

  return Object.freeze(master);
}
