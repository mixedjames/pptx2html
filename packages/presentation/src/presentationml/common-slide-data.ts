import type { Fill } from '../drawingml/index.js';
import type { ShapeTreeNode } from './shape-tree.js';
import type { StyleMatrixReference } from './shape-style.js';

/** §19.3.1.7, bg — the `p:bgPr` alternative: a literal, directly-authored fill. */
export interface Background {
  readonly fill: Fill;
}

/**
 * Content shared by every slide-like part — slides, layouts, masters and notes pages
 * (§19.3.1.16, p:cSld). A part's own background (§19.3.1.6, `EG_Background`) is one of two
 * mutually-exclusive alternatives, modeled as two independent optional fields rather than a
 * union — the same split `ShapeProperties.fill` (a shape's own literal fill) vs. `ShapeStyle.
 * fillRef` (its style-matrix reference fallback) already uses, since they're resolved by
 * different means, not just different shapes of the same value: `background` is `p:bgPr`, a
 * fill authored directly on this part; `backgroundRef` is `p:bgRef`, a reference into the theme's
 * `FormatScheme.bgFillStyles` (PowerPoint's own default — a background inherited from the theme
 * rather than explicitly picked). At most one is ever actually set on a real part. Resolving
 * either into a final `Fill` — including the slide → layout → master fallback chain and the
 * `bgRef` case's own style-matrix lookup — is `resolve/background.ts`'s `resolveEffectiveBackground`
 * job, not this type's.
 */
export interface CommonSlideData {
  readonly name?: string;
  readonly background?: Background;
  readonly backgroundRef?: StyleMatrixReference;
  readonly shapeTree: readonly ShapeTreeNode[];
}
