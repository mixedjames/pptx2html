import type { CustomGeometry, CustomGeometryPath, PresetGeometry } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { customGeometryPath, nativeBorderRadius, presetShapePath } from './shape-geometry.js';

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

function customGeometry(pathLst?: readonly CustomGeometryPath[]): CustomGeometry {
  return { type: 'custom', ...(pathLst ? { pathLst } : {}) };
}

describe('customGeometryPath', () => {
  it('returns undefined for a preset geometry', () => {
    expect(customGeometryPath({ type: 'preset', preset: 'rect' })).toBeUndefined();
  });

  it('returns undefined when custGeom has no path data at all', () => {
    expect(customGeometryPath(customGeometry())).toBeUndefined();
  });

  it('renders a single subpath of moveTo/lnTo/close as M/L/Z', () => {
    const geometry = customGeometry([
      {
        width: 100,
        height: 50,
        commands: [
          { type: 'moveTo', point: { x: 0, y: 0 } },
          { type: 'lnTo', point: { x: 100, y: 0 } },
          { type: 'lnTo', point: { x: 100, y: 50 } },
          { type: 'close' },
        ],
      },
    ]);
    // x scales against w=100 (1:1), y scales against h=50 (2x) — both land in 0..100.
    expect(customGeometryPath(geometry)).toBe('M 0 0 L 100 0 L 100 100 Z');
  });

  it('concatenates multiple a:path entries into multiple subpaths (a boolean-subtract cutout)', () => {
    const geometry = customGeometry([
      {
        width: 100,
        height: 100,
        commands: [
          { type: 'moveTo', point: { x: 0, y: 0 } },
          { type: 'lnTo', point: { x: 100, y: 0 } },
          { type: 'lnTo', point: { x: 100, y: 100 } },
          { type: 'lnTo', point: { x: 0, y: 100 } },
          { type: 'close' },
        ],
      },
      {
        width: 100,
        height: 100,
        commands: [
          { type: 'moveTo', point: { x: 25, y: 25 } },
          { type: 'lnTo', point: { x: 75, y: 25 } },
          { type: 'lnTo', point: { x: 50, y: 75 } },
          { type: 'close' },
        ],
      },
    ]);
    expect(customGeometryPath(geometry)).toBe(
      'M 0 0 L 100 0 L 100 100 L 0 100 Z M 25 25 L 75 25 L 50 75 Z',
    );
  });

  it('renders quadBezTo and cubicBezTo as Q/C', () => {
    const geometry = customGeometry([
      {
        width: 100,
        height: 100,
        commands: [
          { type: 'moveTo', point: { x: 0, y: 0 } },
          { type: 'quadBezTo', control: { x: 10, y: 0 }, point: { x: 10, y: 10 } },
          {
            type: 'cubicBezTo',
            control1: { x: 20, y: 10 },
            control2: { x: 20, y: 20 },
            point: { x: 30, y: 20 },
          },
        ],
      },
    ]);
    expect(customGeometryPath(geometry)).toBe('M 0 0 Q 10 0 10 10 C 20 10 20 20 30 20');
  });

  it('renders arcTo as an SVG elliptical arc, deriving the end point from the current pen position', () => {
    const geometry = customGeometry([
      {
        width: 100,
        height: 100,
        commands: [
          { type: 'moveTo', point: { x: 50, y: 0 } },
          { type: 'arcTo', widthRadius: 50, heightRadius: 50, startAngle: 0, swingAngle: 5400000 },
        ],
      },
    ]);
    // A 90deg (5,400,000 / 60,000) sweep from (50,0) around a 50-radius circle centered at (0,0)
    // ends at (0,50); radii scale 1:1 since w=h=100.
    expect(customGeometryPath(geometry)).toBe('M 50 0 A 50 50 0 0 1 0 50');
  });

  it('sets the large-arc-flag when the swing exceeds 180deg and sweep-flag 0 for a negative swing', () => {
    const geometry = customGeometry([
      {
        width: 100,
        height: 100,
        commands: [
          { type: 'moveTo', point: { x: 50, y: 0 } },
          {
            type: 'arcTo',
            widthRadius: 50,
            heightRadius: 50,
            startAngle: 0,
            swingAngle: -5400000,
          },
        ],
      },
    ]);
    expect(customGeometryPath(geometry)).toBe('M 50 0 A 50 50 0 0 0 0 -50');
  });

  it('drops a path with a zero w/h, falling back to undefined if nothing else is renderable', () => {
    const geometry = customGeometry([
      {
        width: 0,
        height: 0,
        commands: [{ type: 'moveTo', point: { x: 0, y: 0 } }],
      },
    ]);
    expect(customGeometryPath(geometry)).toBeUndefined();
  });
});
