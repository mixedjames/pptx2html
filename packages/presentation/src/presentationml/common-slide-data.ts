import type { Fill } from '../drawingml/index.js';
import type { ShapeTreeNode } from './shape-tree.js';

/** §19.3.1.7, bg. */
export interface Background {
  readonly fill: Fill;
}

/**
 * Content shared by every slide-like part — slides, layouts, masters and notes pages
 * (§19.3.1.16, p:cSld).
 */
export interface CommonSlideData {
  readonly name?: string;
  readonly background?: Background;
  readonly shapeTree: readonly ShapeTreeNode[];
}
