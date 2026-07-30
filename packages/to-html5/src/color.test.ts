import type { Color, ColorScheme, Fill } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { resolveColor, resolveFillColor } from './color.js';

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

describe('resolveColor', () => {
  it('resolves an srgb colour straight to rgb()', () => {
    expect(resolveColor({ type: 'srgb', value: 'FF0000' }, undefined)).toBe('rgb(255, 0, 0)');
  });

  it('applies alpha as rgba()', () => {
    const color: Color = { type: 'srgb', value: 'FF0000', transforms: { alpha: 50000 } };
    expect(resolveColor(color, undefined)).toBe('rgba(255, 0, 0, 0.5)');
  });

  it('resolves a scheme colour by its named slot', () => {
    expect(resolveColor({ type: 'scheme', value: 'accent1' }, SCHEME)).toBe('rgb(79, 129, 189)');
  });

  it('resolves the bg1/tx1/bg2/tx2 aliases against the default clrMap', () => {
    expect(resolveColor({ type: 'scheme', value: 'tx1' }, SCHEME)).toBe('rgb(0, 0, 0)');
    expect(resolveColor({ type: 'scheme', value: 'bg1' }, SCHEME)).toBe('rgb(255, 255, 255)');
  });

  it('returns undefined for a scheme colour with no theme available', () => {
    expect(resolveColor({ type: 'scheme', value: 'accent1' }, undefined)).toBeUndefined();
  });

  it('returns undefined for phClr (no fixed value outside a style-matrix context)', () => {
    expect(resolveColor({ type: 'scheme', value: 'phClr' }, SCHEME)).toBeUndefined();
  });

  it('uses lastColor for a system colour', () => {
    const color: Color = { type: 'system', value: 'windowText', lastColor: '123456' };
    expect(resolveColor(color, undefined)).toBe('rgb(18, 52, 86)');
  });

  it('converts hsl to rgb (red at hue 0, full saturation, mid luminance)', () => {
    const color: Color = { type: 'hsl', hue: 0, saturation: 100000, luminance: 50000 };
    expect(resolveColor(color, undefined)).toBe('rgb(255, 0, 0)');
  });

  it('passes a preset colour name straight through to CSS when untransformed', () => {
    expect(resolveColor({ type: 'preset', value: 'aliceBlue' }, undefined)).toBe('aliceBlue');
  });

  it('ignores a transform stacked on a preset colour (documented limitation)', () => {
    const color: Color = { type: 'preset', value: 'aliceBlue', transforms: { shade: 50000 } };
    expect(resolveColor(color, undefined)).toBe('aliceBlue');
  });

  it('lightens via lumMod/lumOff (a common "tx1, 65% lighter" style pattern)', () => {
    const color: Color = {
      type: 'scheme',
      value: 'tx1',
      transforms: { lumMod: 65000, lumOff: 35000 },
    };
    // Black (l=0) scaled by 0.65 then offset by 0.35 lands at l=0.35 -> mid-grey.
    expect(resolveColor(color, SCHEME)).toBe('rgb(89, 89, 89)');
  });

  it('darkens via shade', () => {
    const color: Color = { type: 'srgb', value: 'FFFFFF', transforms: { shade: 50000 } };
    expect(resolveColor(color, undefined)).toBe('rgb(128, 128, 128)');
  });

  it('lightens via tint', () => {
    const color: Color = { type: 'srgb', value: '000000', transforms: { tint: 50000 } };
    expect(resolveColor(color, undefined)).toBe('rgb(128, 128, 128)');
  });
});

describe('resolveFillColor', () => {
  it('resolves a solid fill', () => {
    const fill: Fill = { type: 'solid', color: { type: 'srgb', value: '00FF00' } };
    expect(resolveFillColor(fill, undefined)).toBe('rgb(0, 255, 0)');
  });

  it('returns undefined for non-solid fills', () => {
    expect(resolveFillColor({ type: 'none' }, undefined)).toBeUndefined();
    expect(resolveFillColor({ type: 'gradient', stops: [] }, undefined)).toBeUndefined();
  });
});
