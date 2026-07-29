import type { PlaceholderType, Shape, SlideLayout, SlideMaster } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { resolveInheritedTransform } from './placeholder.js';

function placeholderShape(
  id: number,
  type: PlaceholderType,
  index: number,
  hasTransform: boolean,
): Shape {
  return {
    kind: 'shape',
    nonVisual: { id, name: `Placeholder ${id}`, placeholder: { type, index } },
    properties: hasTransform
      ? { transform: { offset: { x: 1, y: 2 }, extents: { width: 3, height: 4 } } }
      : {},
  };
}

function master(shapes: readonly Shape[]): SlideMaster {
  return { commonSlideData: { shapeTree: shapes }, theme: {} as never, layouts: [] };
}

function layout(shapes: readonly Shape[], masterShapes: readonly Shape[] = []): SlideLayout {
  return {
    commonSlideData: { shapeTree: shapes },
    master: master(masterShapes),
    type: 'blank',
  };
}

describe('resolveInheritedTransform', () => {
  it('matches a layout placeholder by exact type+index', () => {
    const l = layout([placeholderShape(10, 'body', 1, true)]);
    const transform = resolveInheritedTransform({ type: 'body', index: 1 }, l);
    expect(transform).toEqual({ offset: { x: 1, y: 2 }, extents: { width: 3, height: 4 } });
  });

  it('falls back to a type-only match when index differs', () => {
    const l = layout([placeholderShape(10, 'title', 0, true)]);
    const transform = resolveInheritedTransform({ type: 'title', index: 5 }, l);
    expect(transform).toBeDefined();
  });

  it('falls through to the master when the layout placeholder has no transform of its own', () => {
    const l = layout(
      [placeholderShape(10, 'ftr', 3, false)],
      [placeholderShape(20, 'ftr', 3, true)],
    );
    const transform = resolveInheritedTransform({ type: 'ftr', index: 3 }, l);
    expect(transform).toEqual({ offset: { x: 1, y: 2 }, extents: { width: 3, height: 4 } });
  });

  it('returns undefined when nothing in the chain matches', () => {
    const l = layout([placeholderShape(10, 'body', 1, true)]);
    expect(resolveInheritedTransform({ type: 'pic', index: 9 }, l)).toBeUndefined();
  });
});
