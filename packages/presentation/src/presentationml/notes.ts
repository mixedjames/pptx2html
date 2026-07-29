import type { CommonSlideData } from './common-slide-data.js';
import type { Slide } from './slide.js';

/** The notes master part (§19.3.1.35, p:notesMaster): default look for every notes page. */
export interface NotesMaster {
  readonly commonSlideData: CommonSlideData;
}

/** A notes slide part (§19.3.1.36, p:notes): speaker notes for a single slide. */
export interface NotesSlide {
  readonly commonSlideData: CommonSlideData;
  readonly master: NotesMaster;
  readonly slide: Slide;
}
