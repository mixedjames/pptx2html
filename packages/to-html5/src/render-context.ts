import type { SlideSize, TextStyleContext } from '@pptx2html/presentation';

/**
 * Context threaded through shape-tree rendering that's constant for an entire slide: its size
 * (every element is positioned as a percentage of this, so the whole slide scales responsively —
 * see `shape-tree.ts`'s `positionElement`), plus `layout`/`defaultTextStyle` — inherited from
 * `@pptx2html/presentation`'s `TextStyleContext`, since those two are needed for placeholder/font
 * inheritance resolution (`resolveInheritedTransform`, `resolveEffectiveRunProperties` & co.) and
 * that resolution logic now lives with the presentation model, not this renderer — see that
 * package's `resolve/text-style.ts`. `CoordinateMap`, by contrast, changes at every group nesting
 * level, so it's threaded separately.
 */
export interface RenderContext extends TextStyleContext {
  readonly slideSize: SlideSize;
}
