// @vitest-environment happy-dom
import type { Picture, Shape } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { IDENTITY_MAP } from './coordinate.js';
import type { RenderContext } from './render-context.js';
import { renderShapeTreeNode } from './shape-tree.js';

const CONTEXT: RenderContext = {
  slideSize: { width: 1, height: 1 },
  layout: undefined,
  defaultTextStyle: undefined,
};

describe('renderShapeTreeNode fill/line wiring', () => {
  it("applies a shape's own fill and line as CSS background/border", () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 1, name: 'Rect 1' },
      properties: {
        fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
        line: { width: 25400, fill: { type: 'solid', color: { type: 'srgb', value: '000000' } } },
      },
    };

    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.backgroundColor).toBe('rgb(0, 255, 0)');
    expect(el.style.borderColor).toBe('rgb(0, 0, 0)');
    expect(el.style.borderStyle).toBe('solid');
  });

  it('applies fill/line to a picture alongside its image src', () => {
    const picture: Picture = {
      kind: 'picture',
      nonVisual: { id: 2, name: 'Picture 1' },
      properties: {
        line: { width: 12700, fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } } },
      },
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
    };

    const el = renderShapeTreeNode(document, picture, IDENTITY_MAP, CONTEXT) as HTMLImageElement;
    expect(el.tagName).toBe('IMG');
    expect(el.src).toMatch(/^blob:|^data:/);
    expect(el.style.borderColor).toBe('rgb(255, 0, 0)');
  });

  it('leaves a shape with no fill/line unstyled', () => {
    const shape: Shape = { kind: 'shape', nonVisual: { id: 3, name: 'Plain' }, properties: {} };
    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.backgroundColor).toBe('');
    expect(el.style.borderStyle).toBe('');
  });
});
