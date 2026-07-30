import type { PresetGeometry } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { nativeBorderRadius, presetShapePath } from './shape-geometry.js';

function preset(
  name: string,
  adjustValues?: readonly { name: string; value: number }[],
): PresetGeometry {
  return { type: 'preset', preset: name, ...(adjustValues ? { adjustValues } : {}) };
}

describe('presetShapePath', () => {
  it('returns undefined for rect/roundRect/ellipse — handled natively via border-radius', () => {
    expect(presetShapePath(preset('rect'))).toBeUndefined();
    expect(presetShapePath(preset('roundRect'))).toBeUndefined();
    expect(presetShapePath(preset('ellipse'))).toBeUndefined();
  });

  it('returns undefined for a preset outside the modeled common subset', () => {
    expect(presetShapePath(preset('cloud'))).toBeUndefined();
  });

  it('returns undefined for custom geometry', () => {
    expect(presetShapePath({ type: 'custom' })).toBeUndefined();
  });

  it('renders a centered isosceles triangle by default', () => {
    expect(presetShapePath(preset('triangle'))).toBe('M 50 0 L 100 100 L 0 100 Z');
  });

  it('shifts the triangle apex per an explicit adj guide', () => {
    expect(presetShapePath(preset('triangle', [{ name: 'adj', value: 0 }]))).toBe(
      'M 0 0 L 100 100 L 0 100 Z',
    );
  });

  it('renders a right triangle with the right angle at the bottom-left', () => {
    expect(presetShapePath(preset('rtTriangle'))).toBe('M 0 0 L 0 100 L 100 100 Z');
  });

  it('renders a diamond as a fixed rhombus (no adjustment guide)', () => {
    expect(presetShapePath(preset('diamond'))).toBe('M 50 0 L 100 50 L 50 100 L 0 50 Z');
  });

  it('renders a parallelogram, capping slant at 50%', () => {
    expect(presetShapePath(preset('parallelogram', [{ name: 'adj', value: 100000 }]))).toBe(
      'M 50 0 L 100 0 L 50 100 L 0 100 Z',
    );
  });

  it('renders a trapezoid, capping inset at 50%', () => {
    expect(presetShapePath(preset('trapezoid', [{ name: 'adj', value: 100000 }]))).toBe(
      'M 50 0 L 50 0 L 100 100 L 0 100 Z',
    );
  });

  it('renders a regular pentagon', () => {
    const path = presetShapePath(preset('pentagon'));
    expect(path).toMatch(/^M 50 0 L .* Z$/);
    // 5 vertices: 1 M + 4 L + Z.
    expect(path?.split(' L ').length).toBe(5);
  });

  it('renders a hexagon with flat left/right sides', () => {
    const path = presetShapePath(preset('hexagon'));
    expect(path).toBe('M 25 0 L 75 0 L 100 50 L 75 100 L 25 100 L 0 50 Z');
  });

  it('renders an octagon with cut corners', () => {
    const path = presetShapePath(preset('octagon'));
    expect(path).toBe('M 25 0 L 75 0 L 100 25 L 100 75 L 75 100 L 25 100 L 0 75 L 0 25 Z');
  });

  it('renders a 5-point star with an inner radius ratio from adj', () => {
    const path = presetShapePath(preset('star5', [{ name: 'adj', value: 50000 }]));
    // 10 vertices alternating outer/inner.
    expect(path?.split(' L ').length).toBe(10);
    expect(path).toMatch(/^M 50 0 L/);
  });
});

describe('nativeBorderRadius', () => {
  it('returns undefined for a plain rect', () => {
    expect(nativeBorderRadius(preset('rect'))).toBeUndefined();
  });

  it('returns a percentage radius for roundRect using its adj guide', () => {
    expect(nativeBorderRadius(preset('roundRect', [{ name: 'adj', value: 25000 }]))).toBe('25%');
  });

  it('falls back to the spec default adj (1/6) when roundRect has no explicit guide', () => {
    expect(nativeBorderRadius(preset('roundRect'))).toBe(`${(16667 / 100000) * 100}%`);
  });

  it('caps roundRect radius at 50% (a full stadium/pill)', () => {
    expect(nativeBorderRadius(preset('roundRect', [{ name: 'adj', value: 100000 }]))).toBe('50%');
  });

  it('returns a fixed 50% for ellipse', () => {
    expect(nativeBorderRadius(preset('ellipse'))).toBe('50%');
  });

  it('returns undefined for presets covered by presetShapePath instead', () => {
    expect(nativeBorderRadius(preset('triangle'))).toBeUndefined();
  });

  it('returns undefined for custom geometry', () => {
    expect(nativeBorderRadius({ type: 'custom' })).toBeUndefined();
  });
});
