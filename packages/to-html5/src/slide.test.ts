// @vitest-environment happy-dom
import type { GroupShape, Shape, Slide, SlideSize } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { renderSlide } from './slide.js';

const SLIDE_SIZE: SlideSize = { width: 12192000, height: 6858000 };

/** Mirrors shape-tree.ts's positionElement formula exactly, so expectations can't drift from it. */
function pct(value: number, total: number): string {
  return `${(value / total) * 100}%`;
}

function shape(id: number, x: number, y: number, w: number, h: number, text: string): Shape {
  return {
    kind: 'shape',
    nonVisual: { id, name: `Shape ${id}` },
    properties: { transform: { offset: { x, y }, extents: { width: w, height: h } } },
    textBody: { paragraphs: [{ runs: [{ kind: 'run', text }] }] },
  };
}

describe('renderSlide', () => {
  it('scales to fill its container width via aspect-ratio, and positions shapes as percentages', () => {
    const slide: Slide = {
      commonSlideData: {
        shapeTree: [shape(1, 914400, 914400, 1828800, 914400, 'Hello')],
      },
      layout: {
        commonSlideData: { shapeTree: [] },
        master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
        type: 'blank',
      },
    };

    const el = renderSlide(document, slide, SLIDE_SIZE);

    expect(el.className).toBe('pptx-slide');
    expect(el.style.width).toBe('100%');
    expect(el.style.aspectRatio).toBe('12192000 / 6858000');

    const shapeEl = el.querySelector('.pptx-shape') as HTMLElement;
    expect(shapeEl.style.left).toBe(pct(914400, SLIDE_SIZE.width));
    expect(shapeEl.style.top).toBe(pct(914400, SLIDE_SIZE.height));
    expect(shapeEl.style.width).toBe(pct(1828800, SLIDE_SIZE.width));
    expect(shapeEl.style.height).toBe(pct(914400, SLIDE_SIZE.height));
    expect(shapeEl.textContent).toBe('Hello');
  });

  it('remaps a group child through chOff/chExt into slide-relative percentages', () => {
    const group: GroupShape = {
      kind: 'group',
      nonVisual: { id: 2, name: 'Group 1' },
      transform: {
        offset: { x: 914400, y: 0 },
        extents: { width: 1828800, height: 914400 },
        childOffset: { x: 0, y: 0 },
        childExtents: { width: 914400, height: 457200 },
      },
      children: [shape(3, 0, 0, 457200, 228600, 'Nested')],
    };

    const slide: Slide = {
      commonSlideData: { shapeTree: [group] },
      layout: {
        commonSlideData: { shapeTree: [] },
        master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
        type: 'blank',
      },
    };

    const el = renderSlide(document, slide, SLIDE_SIZE);

    const groupEl = el.querySelector('.pptx-group') as HTMLElement;
    expect(groupEl.style.left).toBe('0px');
    expect(groupEl.style.top).toBe('0px');
    expect(groupEl.style.right).toBe('0px');
    expect(groupEl.style.bottom).toBe('0px');

    // Group scales its 1x0.5in child space up 2x into its 2x1in own box, placed at (1in, 0) —
    // still expressed as a percentage of the whole slide, same as any other element.
    const nested = groupEl.querySelector('.pptx-shape') as HTMLElement;
    expect(nested.style.left).toBe(pct(914400, SLIDE_SIZE.width)); // 1in
    expect(nested.style.top).toBe(pct(0, SLIDE_SIZE.height));
    expect(nested.style.width).toBe(pct(914400, SLIDE_SIZE.width)); // 0.5in * 2
    expect(nested.style.height).toBe(pct(457200, SLIDE_SIZE.height)); // 0.25in * 2
  });

  it('positions a placeholder shape with no own transform via its layout placeholder', () => {
    const titlePlaceholder: Shape = {
      kind: 'shape',
      nonVisual: { id: 1, name: 'Title 1', placeholder: { type: 'title', index: 0 } },
      properties: {},
      textBody: { paragraphs: [{ runs: [{ kind: 'run', text: 'Title' }] }] },
    };

    const layoutTitlePlaceholder: Shape = {
      kind: 'shape',
      nonVisual: { id: 2, name: 'Title Placeholder 1', placeholder: { type: 'title', index: 0 } },
      properties: {
        transform: {
          offset: { x: 457200, y: 285750 },
          extents: { width: 8229600, height: 1143000 },
        },
      },
    };

    const slide: Slide = {
      commonSlideData: { shapeTree: [titlePlaceholder] },
      layout: {
        commonSlideData: { shapeTree: [layoutTitlePlaceholder] },
        master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
        type: 'title',
      },
    };

    const el = renderSlide(document, slide, SLIDE_SIZE);
    const shapeEl = el.querySelector('.pptx-shape') as HTMLElement;

    expect(shapeEl.style.left).toBe(pct(457200, SLIDE_SIZE.width));
    expect(shapeEl.style.top).toBe(pct(285750, SLIDE_SIZE.height));
    expect(shapeEl.style.width).toBe(pct(8229600, SLIDE_SIZE.width));
    expect(shapeEl.textContent).toBe('Title');
  });

  it("applies the slide's own background, falling back to the master's when the slide has none", () => {
    const master: Slide['layout']['master'] = {
      commonSlideData: {
        shapeTree: [],
        background: { fill: { type: 'solid', color: { type: 'srgb', value: '0000FF' } } },
      },
      theme: {} as never,
      layouts: [],
    };
    const layout: Slide['layout'] = { commonSlideData: { shapeTree: [] }, master, type: 'blank' };

    const withOwnBackground: Slide = {
      commonSlideData: {
        shapeTree: [],
        background: { fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } } },
      },
      layout,
    };
    expect(renderSlide(document, withOwnBackground, SLIDE_SIZE).style.backgroundColor).toBe(
      'rgb(255, 0, 0)',
    );

    const inheritingSlide: Slide = { commonSlideData: { shapeTree: [] }, layout };
    expect(renderSlide(document, inheritingSlide, SLIDE_SIZE).style.backgroundColor).toBe(
      'rgb(0, 0, 255)',
    );
  });
});
