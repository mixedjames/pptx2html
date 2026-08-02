import type { Slide, SlideSize, TextListStyle } from '@pptx2html/presentation';
import { IDENTITY_MAP, resolveEffectiveBackground } from '@pptx2html/presentation';

import { applyFill } from './fill.js';
import type { RenderContext } from './render-context.js';
import { renderShapeTreeNode } from './shape-tree.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

/**
 * `slideIndex`/`unsupportedFeatures` are optional (defaulting to 0 and a throwaway collector) so
 * callers that don't care about the unsupported-feature log — chiefly this file's own tests —
 * don't need to supply them.
 */
export function renderSlide(
  doc: Document,
  slide: Slide,
  slideSize: SlideSize,
  defaultTextStyle?: TextListStyle,
  slideIndex = 0,
  unsupportedFeatures: UnsupportedFeatureCollector = new UnsupportedFeatureCollector(),
): HTMLElement {
  const el = doc.createElement('div');
  el.className = 'pptx-slide';
  el.style.position = 'relative';
  el.style.overflow = 'hidden';
  // Fills whatever width its container gives it; height follows from the slide's own aspect
  // ratio. This — plus every descendant being positioned as a percentage of `slideSize` (see
  // shape-tree.ts's positionElement) — is what makes the whole slide scale with its container
  // with no JS involved.
  el.style.width = '100%';
  el.style.aspectRatio = `${slideSize.width} / ${slideSize.height}`;
  // Establishes `.pptx-slide` as a query container so descendants can size themselves in `cqw`
  // (a percentage of *this* element's own width) rather than a fixed px/pt value — see
  // units.ts's emuToCqw for why that's how font-size/border-width scale with the slide.
  el.style.setProperty('container-type', 'inline-size');

  const background = resolveEffectiveBackground(slide);
  if (background) applyFill(el, background.fill, slide.layout.master.theme.colorScheme);

  const context: RenderContext = {
    slideSize,
    layout: slide.layout,
    defaultTextStyle,
    reportUnsupported: (code, message, shape) =>
      unsupportedFeatures.report({ code, message, slideIndex, shape }),
  };
  for (const node of slide.commonSlideData.shapeTree) {
    el.appendChild(renderShapeTreeNode(doc, node, IDENTITY_MAP, context));
  }

  return el;
}
