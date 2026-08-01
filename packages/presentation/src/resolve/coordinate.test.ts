import { describe, expect, it } from 'vitest';
import { IDENTITY_MAP, composeGroupMap, computeBox, mapPoint, mapSize } from './coordinate.js';

describe('IDENTITY_MAP', () => {
  it('leaves points and sizes unchanged', () => {
    expect(mapPoint(IDENTITY_MAP, { x: 10, y: 20 })).toEqual({ x: 10, y: 20 });
    expect(mapSize(IDENTITY_MAP, { width: 30, height: 40 })).toEqual({ width: 30, height: 40 });
  });
});

describe('composeGroupMap', () => {
  it('scales and translates child coordinates when chOff/chExt differ from off/ext', () => {
    // Group occupies (100, 200) sized 300x400 in parent space, but its children are laid out
    // in a 150x200 child space starting at (0, 0) — i.e. a 2x scale-up.
    const childMap = composeGroupMap(IDENTITY_MAP, {
      offset: { x: 100, y: 200 },
      extents: { width: 300, height: 400 },
      childOffset: { x: 0, y: 0 },
      childExtents: { width: 150, height: 200 },
    });

    expect(mapPoint(childMap, { x: 10, y: 10 })).toEqual({ x: 120, y: 220 });
    expect(mapSize(childMap, { width: 20, height: 20 })).toEqual({ width: 40, height: 40 });
  });

  it('accounts for a non-zero child origin', () => {
    const childMap = composeGroupMap(IDENTITY_MAP, {
      offset: { x: 100, y: 100 },
      extents: { width: 100, height: 100 },
      childOffset: { x: 50, y: 50 },
      childExtents: { width: 100, height: 100 },
    });

    // No scaling (1:1), but the child space's origin (50, 50) must map back to (100, 100).
    expect(mapPoint(childMap, { x: 50, y: 50 })).toEqual({ x: 100, y: 100 });
    expect(mapPoint(childMap, { x: 60, y: 60 })).toEqual({ x: 110, y: 110 });
  });

  it('composes across nested groups', () => {
    const outer = composeGroupMap(IDENTITY_MAP, {
      offset: { x: 1000, y: 0 },
      extents: { width: 2000, height: 1000 },
      childOffset: { x: 0, y: 0 },
      childExtents: { width: 1000, height: 500 },
    });
    const inner = composeGroupMap(outer, {
      offset: { x: 0, y: 0 },
      extents: { width: 1000, height: 500 },
      childOffset: { x: 0, y: 0 },
      childExtents: { width: 500, height: 250 },
    });

    // Overall scale is 2x (outer) * 2x (inner) = 4x, offset by the outer group's placement.
    expect(mapPoint(inner, { x: 10, y: 10 })).toEqual({ x: 1000 + 40, y: 40 });
  });

  it('defaults to no remapping when chOff/chExt are absent', () => {
    const childMap = composeGroupMap(IDENTITY_MAP, {
      offset: { x: 5, y: 5 },
      extents: { width: 100, height: 100 },
    });

    expect(mapPoint(childMap, { x: 10, y: 10 })).toEqual({ x: 15, y: 15 });
  });
});

describe('computeBox', () => {
  it('maps offset/extents through the coordinate map (staying in EMU) and passes through rotation/flip', () => {
    const box = computeBox(IDENTITY_MAP, {
      offset: { x: 9525, y: 19050 },
      extents: { width: 95250, height: 190500 },
      rotation: 5400000,
      flipHorizontal: true,
    });

    expect(box).toEqual({
      left: 9525,
      top: 19050,
      width: 95250,
      height: 190500,
      rotationDeg: 90,
      flipHorizontal: true,
      flipVertical: undefined,
    });
  });
});
