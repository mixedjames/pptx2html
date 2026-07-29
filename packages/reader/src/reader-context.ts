import type {
  MediaPart,
  NotesMaster,
  Slide,
  SlideLayout,
  SlideMaster,
  Theme,
} from '@pptx2html/presentation';

import type { OpcPackage } from './opc/package.js';

/**
 * Threaded through every parser. Caches are keyed by part name so a theme/master/layout/notes
 * master/slide/media part referenced from multiple places is parsed once and the same object
 * instance is reused (in particular, NotesSlide.slide must reuse the Slide already built for
 * Presentation.slides, not re-parse that slide a second time).
 */
export interface ReaderContext {
  readonly package: OpcPackage;
  readonly themes: Map<string, Theme>;
  readonly slideMasters: Map<string, SlideMaster>;
  /** SlideLayout instances built while reading their owning master, keyed by their part name. */
  readonly slideLayouts: Map<string, SlideLayout>;
  readonly notesMasters: Map<string, NotesMaster>;
  readonly slides: Map<string, Slide>;
  readonly media: Map<string, MediaPart>;
}

export function createReaderContext(pkg: OpcPackage): ReaderContext {
  return {
    package: pkg,
    themes: new Map(),
    slideMasters: new Map(),
    slideLayouts: new Map(),
    notesMasters: new Map(),
    slides: new Map(),
    media: new Map(),
  };
}
