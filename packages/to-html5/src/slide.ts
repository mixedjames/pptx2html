import type { Slide, SlideSize } from '@pptx2html/presentation';
import { IDENTITY_MAP } from './coordinate.js';
import type { RenderContext } from './render-context.js';
import { renderShapeTreeNode } from './shape-tree.js';

export function renderSlide(doc: Document, slide: Slide, slideSize: SlideSize): HTMLElement {
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

  const context: RenderContext = { slideSize, layout: slide.layout };
  for (const node of slide.commonSlideData.shapeTree) {
    el.appendChild(renderShapeTreeNode(doc, node, IDENTITY_MAP, context));
  }

  return el;
}
