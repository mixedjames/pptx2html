import type { TextListStyle } from '../drawingml/text.js';
import type { Theme } from '../theme.js';
import type { CommonSlideData } from './common-slide-data.js';
import type { SlideLayout } from './slide-layout.js';

/**
 * The master's default text styles by placeholder category (§19.3.1.53, p:txStyles): title
 * placeholders use `titleStyle`, other placeholder shapes (body, subtitle, etc.) use `bodyStyle`,
 * and non-placeholder shapes fall back to `otherStyle`. Each is a per-outline-level default, one
 * rung below a placeholder shape's own `TextBody.listStyle` in the inheritance chain — see
 * `resolve/text-style.ts`.
 */
export interface TextStyles {
  readonly titleStyle?: TextListStyle;
  readonly bodyStyle?: TextListStyle;
  readonly otherStyle?: TextListStyle;
}

/**
 * A slide master part (§19.3.1.44, p:sldMaster): the default look, placeholders and layouts
 * shared by a family of slides.
 */
export interface SlideMaster {
  readonly commonSlideData: CommonSlideData;
  readonly theme: Theme;
  readonly layouts: readonly SlideLayout[];
  readonly textStyles?: TextStyles;
}
