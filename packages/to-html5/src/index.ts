import type { Presentation } from '@pptx2html/presentation';
import { resolveEffectiveBackground } from '@pptx2html/presentation';

import { resolveFillColor } from './color.js';
import { PptxPresentationElement, definePresentationElement } from './presentation-element.js';
import {
  defineScrollPresentationElement,
  PptxScrollPresentationElement,
} from './scroll-presentation-element.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

// Public API is `renderPresentation`/`renderScrollPresentation`/`resolvePresentationBackgroundCss`
// below, plus the types needed to use their return values — everything else in this package
// (fill/color/text/shape-tree/slide rendering, unit conversion) is an internal implementation
// detail, not exported.
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

/**
 * Resolves a presentation's effective background as a CSS colour string, for a host page to paint
 * behind the rendered element itself — e.g. the letterboxing/pillarboxing space `contain-size.ts`
 * centers the deck within, or a fullscreened container, both of which sit outside any individual
 * `.pptx-slide`'s own background and would otherwise fall back to the browser's own default (an
 * arbitrary, and for the Fullscreen API's UA stylesheet, black) rather than the deck's own
 * background. Walks the same inheritance chain `renderSlide` already uses per-slide
 * (`@pptx2html/presentation`'s `resolveEffectiveBackground` — slide falling back to its layout,
 * then that layout's master) off the presentation's first slide, since that's the one shown first
 * and real decks essentially always share one background across every slide. Only a solid fill
 * resolves to a single CSS colour (`resolveFillColor`, the same solid-only resolver a run's own
 * text colour uses) — an empty deck, a slide with no background anywhere in the chain, or one
 * whose background is a gradient/pattern/image this simple colour-only helper can't reduce to one
 * value, all fall back to white (`'#fff'`) rather than leaving the caller with nothing to paint.
 */
export function resolvePresentationBackgroundCss(presentation: Presentation): string {
  const slide = presentation.slides[0];
  if (!slide) return '#fff';
  const background = resolveEffectiveBackground(slide);
  if (!background) return '#fff';
  const scheme = slide.layout.master.theme.colorScheme;
  return resolveFillColor(background, scheme) ?? '#fff';
}
