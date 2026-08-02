import type { SlideTiming } from './animation.js';
import type { CommonSlideData } from './common-slide-data.js';
import type { SlideLayout } from './slide-layout.js';
import type { SlideTransition } from './transition.js';

/** A slide part (§19.3.1.38, p:sld). */
export interface Slide {
  readonly commonSlideData: CommonSlideData;
  readonly layout: SlideLayout;
  readonly showMasterShapes?: boolean;
  readonly transition?: SlideTransition;
  readonly timing?: SlideTiming;
}
