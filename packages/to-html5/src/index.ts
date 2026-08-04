import type { Presentation } from '@pptx2html/presentation';
import { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
import {
  defineScrollPresentationElement,
  PptxScrollPresentationElement,
} from './scroll-presentation-element.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

// Public API is `renderPresentation`/`renderScrollPresentation` below, plus the types needed to
// use their return values — everything else in this package (fill/color/text/shape-tree/slide
// rendering, unit conversion) is an internal implementation detail, not exported.
// `PptxPresentationElement`/`PptxScrollPresentationElement` and
// `UnsupportedFeature`/`UnsupportedFeatureShapeRef` are exported only because they're part of
// those functions' own return types, not as separately-intended entry points.
export { PptxPresentationElement } from './presentation-element.js';
export { PptxScrollPresentationElement } from './scroll-presentation-element.js';
export type { UnsupportedFeature, UnsupportedFeatureShapeRef } from './unsupported-features.js';
export { UnsupportedFeatureCollector } from './unsupported-features.js';

/**
 * Renders a presentation object graph into a `<pptx-presentation>` custom element with a shadow
 * DOM (append `element` to the document to display it), alongside the set of `.pptx`-authored
 * features this render didn't (fully) support (`unsupportedFeatures.all`/`.bySlide` — see
 * `unsupported-features.ts`). Click/keyboard-driven, slide-granular navigation — see
 * `renderScrollPresentation` for the continuous, scroll-driven alternative.
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

/**
 * Renders a presentation object graph into a `<pptx-scroll-presentation>` custom element — every
 * slide transition and build animation positioned on one scrubbable scroll timeline instead of
 * `renderPresentation`'s click-advanced slideshow. The host page must give the returned `element`
 * an explicit box (e.g. `height: 100vh`) and must not place it inside another scrollable ancestor
 * — see `PptxScrollPresentationElement`'s own doc comment (`scroll-presentation-element.ts`).
 */
export function renderScrollPresentation(presentation: Presentation): {
  element: PptxScrollPresentationElement;
  unsupportedFeatures: UnsupportedFeatureCollector;
} {
  defineScrollPresentationElement();
  const element = document.createElement(
    'pptx-scroll-presentation',
  ) as PptxScrollPresentationElement;
  const unsupportedFeatures = element.render(presentation);
  return { element, unsupportedFeatures };
}
