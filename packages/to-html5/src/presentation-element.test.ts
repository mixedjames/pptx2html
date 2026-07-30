// @vitest-environment happy-dom
import type { Presentation, SlideLayout } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { definePresentationElement, PptxPresentationElement } from './presentation-element.js';

const LAYOUT: SlideLayout = {
  commonSlideData: { shapeTree: [] },
  master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
  type: 'blank',
};

describe('PptxPresentationElement', () => {
  it('registers as <pptx-presentation> and renders one .pptx-slide per slide into its shadow root', () => {
    definePresentationElement();

    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    expect(el.shadowRoot).not.toBeNull();
    expect(el.shadowRoot?.querySelectorAll('.pptx-slide')).toHaveLength(2);
  });
});
