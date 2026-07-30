// @vitest-environment happy-dom
import type { ColorScheme, Fill, GradientFill, Line } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { applyFill, applyLine, resolveGradientCss } from './fill.js';

const SCHEME: ColorScheme = {
  name: 'Office',
  dk1: { type: 'srgb', value: '000000' },
  lt1: { type: 'srgb', value: 'FFFFFF' },
  dk2: { type: 'srgb', value: '1F497D' },
  lt2: { type: 'srgb', value: 'EEECE1' },
  accent1: { type: 'srgb', value: '4F81BD' },
  accent2: { type: 'srgb', value: 'C0504D' },
  accent3: { type: 'srgb', value: '9BBB59' },
  accent4: { type: 'srgb', value: '8064A2' },
  accent5: { type: 'srgb', value: '4BACC6' },
  accent6: { type: 'srgb', value: 'F79646' },
  hlink: { type: 'srgb', value: '0000FF' },
  folHlink: { type: 'srgb', value: '800080' },
};

describe('resolveGradientCss', () => {
  it('converts an OOXML shade-path angle (clockwise from east) to a CSS gradient angle (clockwise from north)', () => {
    const fill: GradientFill = {
      type: 'gradient',
      angle: 0,
      stops: [
        { position: 0, color: { type: 'srgb', value: 'FF0000' } },
        { position: 100000, color: { type: 'srgb', value: '0000FF' } },
      ],
    };
    expect(resolveGradientCss(fill, undefined)).toBe(
      'linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)',
    );
  });

  it('sorts stops by position regardless of source order', () => {
    const fill: GradientFill = {
      type: 'gradient',
      stops: [
        { position: 100000, color: { type: 'srgb', value: '0000FF' } },
        { position: 0, color: { type: 'srgb', value: 'FF0000' } },
      ],
    };
    expect(resolveGradientCss(fill, undefined)).toBe(
      'linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)',
    );
  });

  it('returns undefined with no stops', () => {
    expect(resolveGradientCss({ type: 'gradient', stops: [] }, undefined)).toBeUndefined();
  });
});

describe('applyFill', () => {
  it('sets background-color for a solid fill', () => {
    const el = document.createElement('div');
    applyFill(el, { type: 'solid', color: { type: 'scheme', value: 'accent1' } }, SCHEME);
    expect(el.style.backgroundColor).toBe('rgb(79, 129, 189)');
  });

  it('sets background-image for a gradient fill', () => {
    const el = document.createElement('div');
    const fill: Fill = {
      type: 'gradient',
      stops: [
        { position: 0, color: { type: 'srgb', value: 'FF0000' } },
        { position: 100000, color: { type: 'srgb', value: '0000FF' } },
      ],
    };
    applyFill(el, fill, undefined);
    expect(el.style.backgroundImage).toContain('linear-gradient');
  });

  it('approximates a pattern fill with a background color plus hatch overlay', () => {
    const el = document.createElement('div');
    const fill: Fill = {
      type: 'pattern',
      preset: 'pct25',
      foregroundColor: { type: 'srgb', value: '000000' },
      backgroundColor: { type: 'srgb', value: 'FFFFFF' },
    };
    applyFill(el, fill, undefined);
    expect(el.style.backgroundColor).toBe('rgb(255, 255, 255)');
    expect(el.style.backgroundImage).toContain('repeating-linear-gradient');
  });

  it('does nothing for an explicit noFill', () => {
    const el = document.createElement('div');
    applyFill(el, { type: 'none' }, undefined);
    expect(el.style.backgroundColor).toBe('');
    expect(el.style.backgroundImage).toBe('');
  });

  it('does nothing when fill is undefined', () => {
    const el = document.createElement('div');
    applyFill(el, undefined, undefined);
    expect(el.style.cssText).toBe('');
  });

  it('sets a stretched background-image for a blip fill', () => {
    const el = document.createElement('div');
    const fill: Fill = {
      type: 'blip',
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
      stretch: true,
    };
    applyFill(el, fill, undefined);
    expect(el.style.backgroundImage).toMatch(/^url\(.*\)$/);
    expect(el.style.backgroundSize).toBe('100% 100%');
    expect(el.style.backgroundRepeat).toBe('no-repeat');
  });

  it('tiles a blip fill when requested', () => {
    const el = document.createElement('div');
    const fill: Fill = {
      type: 'blip',
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
      tile: true,
    };
    applyFill(el, fill, undefined);
    expect(el.style.backgroundRepeat).toBe('repeat');
  });
});

describe('applyLine', () => {
  const SLIDE_WIDTH = 9144000; // 10in, a standard 4:3 slide width.

  // NOTE: `applyLine` sets `border-width` in `cqw` (container query width units, see
  // units.ts's `emuToCqw`) so it scales with the slide — but happy-dom's CSSOM doesn't
  // recognize `cqw` as a valid length yet and silently drops the assignment (confirmed: real
  // browsers with Container Query Unit support, e.g. Chrome 105+, accept it fine). The actual
  // `emuToCqw` conversion math is covered directly, without a DOM, in units.test.ts; the tests
  // below can only assert on the properties `applyLine` sets that don't hit this gap.

  it('sets style/color for a solid line', () => {
    const el = document.createElement('div');
    const line: Line = {
      width: 12700,
      fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
    };
    applyLine(el, line, undefined, SLIDE_WIDTH);
    expect(el.style.borderStyle).toBe('solid');
    expect(el.style.borderColor).toBe('rgb(255, 0, 0)');
  });

  it('maps dot/sysDot to dotted and other dash styles to dashed', () => {
    const dotted = document.createElement('div');
    applyLine(dotted, { width: 12700, dashStyle: 'sysDot' }, undefined, SLIDE_WIDTH);
    expect(dotted.style.borderStyle).toBe('dotted');

    const dashed = document.createElement('div');
    applyLine(dashed, { width: 12700, dashStyle: 'lgDash' }, undefined, SLIDE_WIDTH);
    expect(dashed.style.borderStyle).toBe('dashed');
  });

  it('maps a double compound line to border-style: double', () => {
    const el = document.createElement('div');
    applyLine(el, { width: 12700, compound: 'double' }, undefined, SLIDE_WIDTH);
    expect(el.style.borderStyle).toBe('double');
  });

  it('renders no border for an explicit line noFill', () => {
    const el = document.createElement('div');
    applyLine(el, { width: 12700, fill: { type: 'none' } }, undefined, SLIDE_WIDTH);
    expect(el.style.cssText).toBe('');
  });

  it('renders no border when line is undefined', () => {
    const el = document.createElement('div');
    applyLine(el, undefined, undefined, SLIDE_WIDTH);
    expect(el.style.cssText).toBe('');
  });
});
