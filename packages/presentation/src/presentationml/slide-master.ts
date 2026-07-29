import type { Theme } from '../theme.js';
import type { CommonSlideData } from './common-slide-data.js';
import type { SlideLayout } from './slide-layout.js';

/**
 * A slide master part (§19.3.1.44, p:sldMaster): the default look, placeholders and layouts
 * shared by a family of slides. Title/body/other text style defaults are unmodeled for the skeleton.
 */
export interface SlideMaster {
  readonly commonSlideData: CommonSlideData;
  readonly theme: Theme;
  readonly layouts: readonly SlideLayout[];
}
