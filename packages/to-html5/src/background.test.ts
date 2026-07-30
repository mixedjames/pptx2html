import type { Background, Slide, SlideLayout, SlideMaster } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { resolveEffectiveBackground } from './background.js';

const SOLID_RED: Background = { fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } } };
const SOLID_BLUE: Background = {
  fill: { type: 'solid', color: { type: 'srgb', value: '0000FF' } },
};
const SOLID_GREEN: Background = {
  fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
};

function master(background?: Background): SlideMaster {
  return {
    commonSlideData: { shapeTree: [], ...(background ? { background } : {}) },
    theme: {} as never,
    layouts: [],
  };
}

function layout(masterValue: SlideMaster, background?: Background): SlideLayout {
  return {
    commonSlideData: { shapeTree: [], ...(background ? { background } : {}) },
    master: masterValue,
    type: 'blank',
  };
}

function slide(layoutValue: SlideLayout, background?: Background): Slide {
  return {
    commonSlideData: { shapeTree: [], ...(background ? { background } : {}) },
    layout: layoutValue,
  };
}

describe('resolveEffectiveBackground', () => {
  it("prefers the slide's own background", () => {
    const l = layout(master(SOLID_GREEN), SOLID_BLUE);
    expect(resolveEffectiveBackground(slide(l, SOLID_RED))).toBe(SOLID_RED);
  });

  it("falls back to the layout's background when the slide has none", () => {
    const l = layout(master(SOLID_GREEN), SOLID_BLUE);
    expect(resolveEffectiveBackground(slide(l))).toBe(SOLID_BLUE);
  });

  it("falls back to the master's background when neither slide nor layout has one", () => {
    const l = layout(master(SOLID_GREEN));
    expect(resolveEffectiveBackground(slide(l))).toBe(SOLID_GREEN);
  });

  it('returns undefined when nothing in the chain defines a background', () => {
    const l = layout(master());
    expect(resolveEffectiveBackground(slide(l))).toBeUndefined();
  });
});
