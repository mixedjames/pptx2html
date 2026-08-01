import type { Color, ColorScheme, Fill } from '../index.js';
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
  it('resolves an srgb colour to an unrounded rgb + alpha:1', () => {
    expect(resolveColor({ type: 'srgb', value: 'FF0000' }, undefined)).toEqual({
      type: 'rgb',
      rgb: [255, 0, 0],
      alpha: 1,
    });
  });

  it('resolves alpha from the transform', () => {
    const color: Color = { type: 'srgb', value: 'FF0000', transforms: { alpha: 50000 } };
    expect(resolveColor(color, undefined)).toEqual({ type: 'rgb', rgb: [255, 0, 0], alpha: 0.5 });
  });

  it('resolves a scheme colour by its named slot', () => {
    expect(resolveColor({ type: 'scheme', value: 'accent1' }, SCHEME)).toEqual({
      type: 'rgb',
      rgb: [79, 129, 189],
      alpha: 1,
    });
  });

  it('resolves the bg1/tx1/bg2/tx2 aliases against the default clrMap', () => {
    expect(resolveColor({ type: 'scheme', value: 'tx1' }, SCHEME)).toEqual({
      type: 'rgb',
      rgb: [0, 0, 0],
      alpha: 1,
    });
    expect(resolveColor({ type: 'scheme', value: 'bg1' }, SCHEME)).toEqual({
      type: 'rgb',
      rgb: [255, 255, 255],
      alpha: 1,
    });
  });

  it('returns undefined for a scheme colour with no theme available', () => {
    expect(resolveColor({ type: 'scheme', value: 'accent1' }, undefined)).toBeUndefined();
  });

  it('returns undefined for phClr (no fixed value outside a style-matrix context)', () => {
    expect(resolveColor({ type: 'scheme', value: 'phClr' }, SCHEME)).toBeUndefined();
  });

  it('uses lastColor for a system colour', () => {
    const color: Color = { type: 'system', value: 'windowText', lastColor: '123456' };
    expect(resolveColor(color, undefined)).toEqual({
      type: 'rgb',
      rgb: [18, 52, 86],
      alpha: 1,
    });
  });

  it('converts hsl to rgb (red at hue 0, full saturation, mid luminance)', () => {
    const color: Color = { type: 'hsl', hue: 0, saturation: 100000, luminance: 50000 };
    expect(resolveColor(color, undefined)).toEqual({ type: 'rgb', rgb: [255, 0, 0], alpha: 1 });
  });

  it('passes a preset colour through as an opaque ResolvedPresetColor when untransformed', () => {
    expect(resolveColor({ type: 'preset', value: 'aliceBlue' }, undefined)).toEqual({
      type: 'preset',
      value: 'aliceBlue',
    });
  });

  it('ignores a transform stacked on a preset colour (documented limitation)', () => {
    const color: Color = { type: 'preset', value: 'aliceBlue', transforms: { shade: 50000 } };
    expect(resolveColor(color, undefined)).toEqual({ type: 'preset', value: 'aliceBlue' });
  });

  it('lightens via lumMod/lumOff (a common "tx1, 65% lighter" style pattern)', () => {
    const color: Color = {
      type: 'scheme',
      value: 'tx1',
      transforms: { lumMod: 65000, lumOff: 35000 },
    };
    // Black (l=0) scaled by 0.65 then offset by 0.35 lands at l=0.35 -> mid-grey.
    const resolved = resolveColor(color, SCHEME);
    expect(resolved?.type).toBe('rgb');
    expect(resolved && resolved.type === 'rgb' ? resolved.rgb.map(Math.round) : undefined).toEqual([
      89, 89, 89,
    ]);
  });

  it('darkens via shade', () => {
    const color: Color = { type: 'srgb', value: 'FFFFFF', transforms: { shade: 50000 } };
    expect(resolveColor(color, undefined)).toEqual({
      type: 'rgb',
      rgb: [127.5, 127.5, 127.5],
      alpha: 1,
    });
  });

  it('lightens via tint', () => {
    const color: Color = { type: 'srgb', value: '000000', transforms: { tint: 50000 } };
    expect(resolveColor(color, undefined)).toEqual({
      type: 'rgb',
      rgb: [127.5, 127.5, 127.5],
      alpha: 1,
    });
  });
});

describe('resolveFillColor', () => {
  it('resolves a solid fill', () => {
    const fill: Fill = { type: 'solid', color: { type: 'srgb', value: '00FF00' } };
    expect(resolveFillColor(fill, undefined)).toEqual({ type: 'rgb', rgb: [0, 255, 0], alpha: 1 });
  });

  it('returns undefined for non-solid fills', () => {
    expect(resolveFillColor({ type: 'none' }, undefined)).toBeUndefined();
    expect(resolveFillColor({ type: 'gradient', stops: [] }, undefined)).toBeUndefined();
  });
});
