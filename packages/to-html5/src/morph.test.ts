import type { Shape, Slide } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
import { resolveSlideMorphMatch } from './morph.js';

function shape(id: number, name: string): Shape {
  return { kind: 'shape', nonVisual: { id, name }, properties: {} };
}

function slide(shapes: readonly Shape[]): Slide {
  return { commonSlideData: { shapeTree: shapes }, layout: {} as never };
}

describe('resolveSlideMorphMatch', () => {
  it('reports and falls back when there is no previous slide', () => {
    const report = vi.fn();
    const result = resolveSlideMorphMatch(undefined, slide([shape(1, 'Title 1')]), report);
    expect(result).toBeUndefined();
    expect(report).toHaveBeenCalledWith('morph-match-degraded', expect.any(String));
  });

  it('reduces a confident match to plain shape ids, reporting nothing', () => {
    const report = vi.fn();
    const previous = slide([shape(2, 'Title 1'), shape(5, 'Body 2')]);
    const current = slide([shape(2, 'Title 1'), shape(5, 'Body 2')]);

    const result = resolveSlideMorphMatch(previous, current, report);
    expect(result).toEqual({
      matched: [
        { outgoingShapeId: 2, incomingShapeId: 2 },
        { outgoingShapeId: 5, incomingShapeId: 5 },
      ],
      disappearingShapeIds: [],
      appearingShapeIds: [],
    });
    expect(report).not.toHaveBeenCalled();
  });

  it('includes appearing/disappearing shape ids in an otherwise-confident match', () => {
    const report = vi.fn();
    const previous = slide([shape(2, 'Title 1'), shape(3, 'Old Icon')]);
    const current = slide([shape(2, 'Title 1'), shape(9, 'New Icon')]);

    const result = resolveSlideMorphMatch(previous, current, report);
    expect(result).toEqual({
      matched: [{ outgoingShapeId: 2, incomingShapeId: 2 }],
      disappearingShapeIds: [3],
      appearingShapeIds: [9],
    });
    expect(report).not.toHaveBeenCalled();
  });

  it('falls back and reports when the match rate is below the confidence threshold', () => {
    const report = vi.fn();
    // Only 1 of 4 shapes matches on the larger side (25%, below the 1/3 threshold).
    const previous = slide([shape(2, 'Title 1'), shape(3, 'A'), shape(4, 'B')]);
    const current = slide([shape(2, 'Title 1'), shape(9, 'C'), shape(10, 'D'), shape(11, 'E')]);

    const result = resolveSlideMorphMatch(previous, current, report);
    expect(result).toBeUndefined();
    expect(report).toHaveBeenCalledWith('morph-match-degraded', expect.any(String));
  });

  it('falls back and reports when nothing matches at all', () => {
    const report = vi.fn();
    const previous = slide([shape(2, 'Alpha')]);
    const current = slide([shape(9, 'Beta')]);

    const result = resolveSlideMorphMatch(previous, current, report);
    expect(result).toBeUndefined();
    expect(report).toHaveBeenCalledWith('morph-match-degraded', expect.any(String));
  });

  it('plays a morph where every shape matches exactly (right at the confidence ceiling)', () => {
    const report = vi.fn();
    const previous = slide([shape(2, 'Title 1')]);
    const current = slide([shape(2, 'Title 1')]);

    const result = resolveSlideMorphMatch(previous, current, report);
    expect(result?.matched).toHaveLength(1);
    expect(report).not.toHaveBeenCalled();
  });
});
