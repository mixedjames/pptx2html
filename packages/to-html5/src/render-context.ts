import type { SlideLayout, SlideSize } from '@pptx2html/presentation';

/**
 * Context threaded through shape-tree rendering that's constant for an entire slide: its size
 * (every element is positioned as a percentage of this, so the whole slide scales responsively —
 * see `shape-tree.ts`'s `positionElement`) and its layout (for placeholder inheritance, see
 * `placeholder.ts`). `CoordinateMap`, by contrast, changes at every group nesting level, so it's
 * threaded separately.
 */
export interface RenderContext {
  readonly slideSize: SlideSize;
  readonly layout: SlideLayout | undefined;
}
