import type { Presentation } from '@pptx2html/presentation';
import { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

// Public API is `renderPresentation` below, plus the types needed to use its return value —
// everything else in this package (fill/color/text/shape-tree/slide rendering, unit conversion)
// is an internal implementation detail, not exported. `PptxPresentationElement` and
// `UnsupportedFeature`/`UnsupportedFeatureShapeRef` are exported only because they're part of
// `renderPresentation`'s own return type, not as separately-intended entry points.
export { PptxPresentationElement } from './presentation-element.js';
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
