import type { Presentation } from '@pptx2html/presentation';
import { PptxPresentationElement, definePresentationElement } from './presentation-element.js';

export { resolveEffectiveBackground } from './background.js';
export { resolveColor, resolveFillColor } from './color.js';
export { IDENTITY_MAP, composeGroupMap, computeBox, mapPoint, mapSize } from './coordinate.js';
export type { CoordinateMap, ElementBox } from './coordinate.js';
export { applyFill, applyLine, resolveGradientCss } from './fill.js';
export { findPlaceholderMatch, resolveInheritedTransform } from './placeholder.js';
export { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
export type { RenderContext } from './render-context.js';
export { renderShapeTreeNode } from './shape-tree.js';
export { renderSlide } from './slide.js';
export { renderTable } from './table.js';
export { renderTextBody } from './text.js';
export { resolveEffectiveRunProperties, resolveTypeface } from './text-style.js';
export { EMU_PER_PX, emuToPx } from './units.js';

/**
 * Renders a presentation object graph into a `<pptx-presentation>` custom element with a shadow
 * DOM. Append the returned element to the document to display it.
 */
export function renderPresentation(presentation: Presentation): PptxPresentationElement {
  definePresentationElement();
  const element = document.createElement('pptx-presentation') as PptxPresentationElement;
  element.render(presentation);
  return element;
}
