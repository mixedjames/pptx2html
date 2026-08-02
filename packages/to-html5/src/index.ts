import type { Presentation } from '@pptx2html/presentation';
import { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

export { resolveColor, resolveFillColor } from './color.js';
export { applyFill, applyLine, resolveGradientCss } from './fill.js';
export { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
export type { RenderContext } from './render-context.js';
export { renderShapeTreeNode } from './shape-tree.js';
export { renderSlide } from './slide.js';
export { renderTable } from './table.js';
export { renderTextBody } from './text.js';
export { EMU_PER_PT, EMU_PER_PX, emuToCqw, emuToPx, fontSizeToEmu } from './units.js';
export type { UnsupportedFeature, UnsupportedFeatureShapeRef } from './unsupported-features.js';
export { UnsupportedFeatureCollector } from './unsupported-features.js';

/**
 * Renders a presentation object graph into a `<pptx-presentation>` custom element with a shadow
 * DOM (append `element` to the document to display it), alongside the set of `.pptx`-authored
 * features this render didn't (fully) support (`unsupportedFeatures.all`/`.bySlide` — see
 * `unsupported-features.ts`).
 */
export function renderPresentation(presentation: Presentation): {
  element: PptxPresentationElement;
  unsupportedFeatures: UnsupportedFeatureCollector;
} {
  definePresentationElement();
  const element = document.createElement('pptx-presentation') as PptxPresentationElement;
  const unsupportedFeatures = element.render(presentation);
  return { element, unsupportedFeatures };
}
