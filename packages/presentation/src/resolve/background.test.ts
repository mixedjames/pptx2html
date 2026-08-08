import type { Background, Fill, Slide, SlideLayout, SlideMaster, Theme } from '../index.js';
import type { StyleMatrixReference } from '../presentationml/shape-style.js';
import { describe, expect, it } from 'vitest';
import { resolveEffectiveBackground } from './background.js';

const RED: Fill = { type: 'solid', color: { type: 'srgb', value: 'FF0000' } };
const BLUE: Fill = { type: 'solid', color: { type: 'srgb', value: '0000FF' } };
const GREEN: Fill = { type: 'solid', color: { type: 'srgb', value: '00FF00' } };
const SOLID_RED: Background = { fill: RED };
const SOLID_BLUE: Background = { fill: BLUE };
const SOLID_GREEN: Background = { fill: GREEN };

function theme(bgFillStyles: readonly Fill[] = []): Theme {
  return {
    name: 'theme',
    colorScheme: {} as never,
    fontScheme: {} as never,
    formatScheme: { name: 'fmt', fillStyles: [], lineStyles: [], bgFillStyles },
  };
}

function master(
  options: { background?: Background; backgroundRef?: StyleMatrixReference; theme?: Theme } = {},
): SlideMaster {
  return {
    commonSlideData: {
      shapeTree: [],
      ...(options.background ? { background: options.background } : {}),
      ...(options.backgroundRef ? { backgroundRef: options.backgroundRef } : {}),
    },
    theme: options.theme ?? theme(),
    layouts: [],
  };
}

function layout(
  masterValue: SlideMaster,
  options: { background?: Background; backgroundRef?: StyleMatrixReference } = {},
): SlideLayout {
  return {
    commonSlideData: {
      shapeTree: [],
      ...(options.background ? { background: options.background } : {}),
      ...(options.backgroundRef ? { backgroundRef: options.backgroundRef } : {}),
    },
    master: masterValue,
    type: 'blank',
  };
}

function slide(
  layoutValue: SlideLayout,
  options: { background?: Background; backgroundRef?: StyleMatrixReference } = {},
): Slide {
  return {
    commonSlideData: {
      shapeTree: [],
      ...(options.background ? { background: options.background } : {}),
      ...(options.backgroundRef ? { backgroundRef: options.backgroundRef } : {}),
    },
    layout: layoutValue,
  };
}

describe('resolveEffectiveBackground', () => {
  it("prefers the slide's own background", () => {
    const l = layout(master({ background: SOLID_GREEN }), { background: SOLID_BLUE });
    expect(resolveEffectiveBackground(slide(l, { background: SOLID_RED }))).toBe(RED);
  });

  it("falls back to the layout's background when the slide has none", () => {
    const l = layout(master({ background: SOLID_GREEN }), { background: SOLID_BLUE });
    expect(resolveEffectiveBackground(slide(l))).toBe(BLUE);
  });

  it("falls back to the master's background when neither slide nor layout has one", () => {
    const l = layout(master({ background: SOLID_GREEN }));
    expect(resolveEffectiveBackground(slide(l))).toBe(GREEN);
  });

  it('returns undefined when nothing in the chain defines a background', () => {
    const l = layout(master());
    expect(resolveEffectiveBackground(slide(l))).toBeUndefined();
  });

  describe('backgroundRef (p:bgRef) — a background inherited from the theme rather than an explicit fill', () => {
    const PALE = {
      type: 'solid',
      color: { type: 'scheme', value: 'phClr' },
    } as const satisfies Fill;
    const bgTheme = theme([PALE]);
    const ref: StyleMatrixReference = { index: 1001, color: { type: 'scheme', value: 'bg1' } };

    it("resolves a part's own backgroundRef against the theme's bgFillStyles, substituting phClr", () => {
      const m = master({ backgroundRef: ref, theme: bgTheme });
      const l = layout(m);
      expect(resolveEffectiveBackground(slide(l))).toEqual({
        type: 'solid',
        color: { type: 'scheme', value: 'bg1' },
      });
    });

    it("a part's own backgroundRef counts as that part defining a background — it doesn't fall through to a deeper part", () => {
      // Regression test: before this was fixed, a part whose only background was a backgroundRef
      // silently resolved as if it had none at all, incorrectly falling through the chain — see
      // this package's CLAUDE.md for the real deck (Presentation1.pptx) this was caught against.
      const m = master({ background: SOLID_GREEN, theme: bgTheme });
      const l = layout(m, { backgroundRef: ref });
      expect(resolveEffectiveBackground(slide(l))).toEqual({
        type: 'solid',
        color: { type: 'scheme', value: 'bg1' },
      });
    });

    it('returns undefined for an out-of-range or missing-theme backgroundRef, same as an unresolved fillRef', () => {
      const m = master({ backgroundRef: { index: 9999, color: { type: 'scheme', value: 'bg1' } } });
      const l = layout(m);
      expect(resolveEffectiveBackground(slide(l))).toBeUndefined();
    });
  });
});
