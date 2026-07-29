import type { NotesMaster, NotesSlide } from '@pptx2html/presentation';

import { createMediaResolver } from '../drawingml/media.js';
import type { ReaderContext } from '../reader-context.js';
import { parseXml } from '../xml/parse.js';
import { findChild, findRoot } from '../xml/query.js';
import { EMPTY_COMMON_SLIDE_DATA, parseCommonSlideData } from './common-slide-data.js';
import { RELATIONSHIP_TYPES } from './relationship-types.js';
import { readSlide } from './slide.js';

/** Parses the notes master part (§19.3.1.35, p:notesMaster). */
export function readNotesMaster(context: ReaderContext, partName: string): NotesMaster {
  const cached = context.notesMasters.get(partName);
  if (cached) return cached;

  const root = findRoot(parseXml(context.package.readText(partName)), 'notesMaster');
  const cSld = root && findChild(root, 'cSld');
  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const commonSlideData = cSld ? parseCommonSlideData(cSld, resolveMedia) : EMPTY_COMMON_SLIDE_DATA;

  const notesMaster: NotesMaster = { commonSlideData };
  context.notesMasters.set(partName, notesMaster);
  return notesMaster;
}

/**
 * Parses a notes slide part (§19.3.1.36, p:notes). Reuses the already-built Slide object for
 * the slide it annotates (via context.slides), rather than re-parsing that slide's XML again.
 */
export function readNotesSlide(context: ReaderContext, partName: string): NotesSlide {
  const root = findRoot(parseXml(context.package.readText(partName)), 'notes');
  const cSld = root && findChild(root, 'cSld');
  const resolveMedia = createMediaResolver(context.package, partName, context.media);
  const commonSlideData = cSld ? parseCommonSlideData(cSld, resolveMedia) : EMPTY_COMMON_SLIDE_DATA;

  const relationships = context.package.relationshipsFor(partName);
  const notesMasterRel = relationships.findByType(RELATIONSHIP_TYPES.notesMaster)[0];
  const slideRel = relationships.findByType(RELATIONSHIP_TYPES.slide)[0];
  if (!notesMasterRel || !slideRel) {
    throw new Error(
      `Notes slide part "${partName}" is missing its notesMaster or slide relationship`,
    );
  }

  return {
    commonSlideData,
    master: readNotesMaster(context, notesMasterRel.target),
    slide: readSlide(context, slideRel.target),
  };
}
