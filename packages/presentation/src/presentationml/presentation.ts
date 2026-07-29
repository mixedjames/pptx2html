import type { Emu } from '../drawingml/index.js';
import type { NotesMaster, NotesSlide } from './notes.js';
import type { Slide } from './slide.js';
import type { SlideMaster } from './slide-master.js';

export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
}

/**
 * The root of the in-memory DOM, corresponding to the presentation part (§19.2.1.26, p:presentation).
 * Default text styles and custom shows are unmodeled for the skeleton.
 */
export interface Presentation {
  readonly slideSize: SlideSize;
  readonly notesSize?: SlideSize;
  readonly slideMasters: readonly SlideMaster[];
  readonly slides: readonly Slide[];
  readonly notesMaster?: NotesMaster;
  readonly notesSlides: readonly NotesSlide[];
}
