// @vitest-environment happy-dom
import type { Presentation } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { definePresentationElement, PptxPresentationElement } from './presentation-element.js';

describe('PptxPresentationElement', () => {
  it('registers as <pptx-presentation> and renders one .pptx-slide per slide into its shadow root', () => {
    definePresentationElement();

    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: {} as never },
        { commonSlideData: { shapeTree: [] }, layout: {} as never },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    expect(el.shadowRoot).not.toBeNull();
    expect(el.shadowRoot?.querySelectorAll('.pptx-slide')).toHaveLength(2);
  });
});
