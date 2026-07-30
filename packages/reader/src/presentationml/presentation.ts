import type { Presentation, SlideSize } from '@pptx2html/presentation';

import { createMediaResolver } from '../drawingml/media.js';
import { parseTextListStyle } from '../drawingml/text.js';
import { parseEmu } from '../drawingml/units.js';
import type { ReaderContext } from '../reader-context.js';
import type { XmlNode } from '../xml/parse.js';
import { parseXml } from '../xml/parse.js';
import { attr, findAllChildren, findChild, findRoot } from '../xml/query.js';
import { readNotesMaster, readNotesSlide } from './notes.js';
import { RELATIONSHIP_TYPES } from './relationship-types.js';
import { readSlide } from './slide.js';
import { readSlideMaster } from './slide-master.js';

function parseSlideSize(node: XmlNode | undefined): SlideSize | undefined {
  if (!node) return undefined;
  const width = parseEmu(attr(node, 'cx'));
  const height = parseEmu(attr(node, 'cy'));
  return width !== undefined && height !== undefined ? { width, height } : undefined;
}

/** Parses the presentation part (§19.2.1.26, p:presentation) into the full object graph. */
export function readPresentationPart(context: ReaderContext, partName: string): Presentation {
  const root = findRoot(parseXml(context.package.readText(partName)), 'presentation');
  if (!root) {
    throw new Error(`"${partName}" is missing its <p:presentation> root element`);
  }

  const slideSize = parseSlideSize(findChild(root, 'sldSz'));
  if (!slideSize) {
    throw new Error(`"${partName}" is missing a valid <p:sldSz>`);
  }
  const notesSize = parseSlideSize(findChild(root, 'notesSz'));

  const relationships = context.package.relationshipsFor(partName);

  // Slide masters (and the layouts they own) must be read before slides: slides resolve their
  // layout relationship against context.slideLayouts, which readSlideMaster populates.
  const sldMasterIdLst = findChild(root, 'sldMasterIdLst');
  const masterIds = sldMasterIdLst ? findAllChildren(sldMasterIdLst, 'sldMasterId') : [];
  const slideMasters = masterIds
    .map((node) => attr(node, 'r:id'))
    .map((rId) => (rId ? relationships.get(rId) : undefined))
    .filter((rel): rel is NonNullable<typeof rel> => rel !== undefined)
    .map((rel) => readSlideMaster(context, rel.target));

  const sldIdLst = findChild(root, 'sldIdLst');
  const slideIds = sldIdLst ? findAllChildren(sldIdLst, 'sldId') : [];
  const slideRels = slideIds
    .map((node) => attr(node, 'r:id'))
    .map((rId) => (rId ? relationships.get(rId) : undefined))
    .filter((rel): rel is NonNullable<typeof rel> => rel !== undefined);
  const slides = slideRels.map((rel) => readSlide(context, rel.target));

  // Notes slides aren't listed in presentation.xml; each slide part links to its own via its
  // own relationships.
  const notesSlides = slideRels
    .map((rel) => {
      const notesRel = context.package
        .relationshipsFor(rel.target)
        .findByType(RELATIONSHIP_TYPES.notesSlide)[0];
      return notesRel ? readNotesSlide(context, notesRel.target) : undefined;
    })
    .filter((notesSlide): notesSlide is NonNullable<typeof notesSlide> => notesSlide !== undefined);

  const notesMasterIdLst = findChild(root, 'notesMasterIdLst');
  const notesMasterIdNode = notesMasterIdLst
    ? findChild(notesMasterIdLst, 'notesMasterId')
    : undefined;
  const notesMasterRId = notesMasterIdNode ? attr(notesMasterIdNode, 'r:id') : undefined;
  const notesMasterRel = notesMasterRId ? relationships.get(notesMasterRId) : undefined;
  const notesMaster = notesMasterRel ? readNotesMaster(context, notesMasterRel.target) : undefined;

  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const defaultTextStyle = parseTextListStyle(findChild(root, 'defaultTextStyle'), resolveMedia);

  return {
    slideSize,
    ...(notesSize ? { notesSize } : {}),
    slideMasters,
    slides,
    ...(notesMaster ? { notesMaster } : {}),
    notesSlides,
    ...(defaultTextStyle ? { defaultTextStyle } : {}),
  };
}
