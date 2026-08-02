import type { SlideSize, TextStyleContext } from '@pptx2html/presentation';
import type { UnsupportedFeatureShapeRef } from './unsupported-features.js';

/**
 * Context threaded through shape-tree rendering that's constant for an entire slide: its size
 * (every element is positioned as a percentage of this, so the whole slide scales responsively —
 * see `shape-tree.ts`'s `positionElement`), plus `layout`/`defaultTextStyle` — inherited from
 * `@pptx2html/presentation`'s `TextStyleContext`, since those two are needed for placeholder/font
 * inheritance resolution (`resolveInheritedTransform`, `resolveEffectiveRunProperties` & co.) and
 * that resolution logic now lives with the presentation model, not this renderer — see that
 * package's `resolve/text-style.ts`. `CoordinateMap`, by contrast, changes at every group nesting
 * level, so it's threaded separately.
 *
 * `reportUnsupported` (optional — tests and other one-off contexts routinely omit it) lets a
 * renderer function flag a feature it's about to silently fall back on, already bound to this
 * slide's own index (see `slide.ts`) so call sites only need to supply the feature itself.
 */
export interface RenderContext extends TextStyleContext {
  readonly slideSize: SlideSize;
  readonly reportUnsupported?: (
    code: string,
    message: string,
    shape?: UnsupportedFeatureShapeRef,
  ) => void;
}
