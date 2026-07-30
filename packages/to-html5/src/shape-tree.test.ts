// @vitest-environment happy-dom
import type { FormatScheme, Picture, Shape, SlideLayout } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { IDENTITY_MAP } from './coordinate.js';
import type { RenderContext } from './render-context.js';
import { renderShapeTreeNode } from './shape-tree.js';

const CONTEXT: RenderContext = {
  slideSize: { width: 1, height: 1 },
  layout: undefined,
  defaultTextStyle: undefined,
};

function contextWithFormatScheme(formatScheme: FormatScheme): RenderContext {
  const layout: SlideLayout = {
    commonSlideData: { shapeTree: [] },
    type: 'blank',
    master: {
      commonSlideData: { shapeTree: [] },
      layouts: [],
      theme: {
        name: 'Office',
        colorScheme: {
          name: 'Office',
          dk1: { type: 'srgb', value: '000000' },
          lt1: { type: 'srgb', value: 'FFFFFF' },
          dk2: { type: 'srgb', value: '000000' },
          lt2: { type: 'srgb', value: 'FFFFFF' },
          accent1: { type: 'srgb', value: '4F81BD' },
          accent2: { type: 'srgb', value: '000000' },
          accent3: { type: 'srgb', value: '000000' },
          accent4: { type: 'srgb', value: '000000' },
          accent5: { type: 'srgb', value: '000000' },
          accent6: { type: 'srgb', value: '000000' },
          hlink: { type: 'srgb', value: '000000' },
          folHlink: { type: 'srgb', value: '000000' },
        },
        fontScheme: {
          name: 'Office',
          majorFont: { latin: 'Calibri Light' },
          minorFont: { latin: 'Calibri' },
        },
        formatScheme,
      },
    },
  };
  return { slideSize: { width: 1, height: 1 }, layout, defaultTextStyle: undefined };
}

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

  it('positions a shape with box-sizing: border-box, so its line does not grow past its width/height', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 4, name: 'Bordered' },
      properties: {
        transform: { offset: { x: 0, y: 0 }, extents: { width: 1, height: 1 } },
        line: { width: 25400, fill: { type: 'solid', color: { type: 'srgb', value: '000000' } } },
      },
    };

    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.boxSizing).toBe('border-box');
  });
});

describe('renderShapeTreeNode preset geometry', () => {
  it('renders ellipse/roundRect via native border-radius, keeping CSS background/border', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 5, name: 'Oval' },
      properties: {
        geometry: { type: 'preset', preset: 'ellipse' },
        fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
      },
    };
    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.borderRadius).toBe('50%');
    expect(el.style.backgroundColor).toBe('rgb(0, 255, 0)');
    expect(el.querySelector('svg')).toBeNull();
  });

  it('renders a non-rectangular preset (triangle) as an SVG path instead of a CSS box', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 6, name: 'Triangle' },
      properties: {
        geometry: { type: 'preset', preset: 'triangle' },
        fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
        line: { width: 12700, fill: { type: 'solid', color: { type: 'srgb', value: '000000' } } },
      },
    };
    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.backgroundColor).toBe('');
    expect(el.style.borderStyle).toBe('');
    const svg = el.querySelector('svg');
    expect(svg).not.toBeNull();
    const path = svg?.querySelector('path');
    expect(path?.getAttribute('d')).toBe('M 50 0 L 100 100 L 0 100 Z');
    expect((path as unknown as SVGPathElement)?.style.fill).toBe('rgb(255, 0, 0)');
    expect((path as unknown as SVGPathElement)?.style.stroke).toBe('rgb(0, 0, 0)');
  });

  it('falls back to a plain rectangle for an unsupported/absent preset', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 7, name: 'Cloud' },
      properties: {
        geometry: { type: 'preset', preset: 'cloud' },
        fill: { type: 'solid', color: { type: 'srgb', value: '0000FF' } },
      },
    };
    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.backgroundColor).toBe('rgb(0, 0, 255)');
    expect(el.querySelector('svg')).toBeNull();
  });

  it('applies native border-radius to a picture for roundRect/ellipse crops', () => {
    const picture: Picture = {
      kind: 'picture',
      nonVisual: { id: 8, name: 'Cropped' },
      properties: { geometry: { type: 'preset', preset: 'ellipse' } },
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
    };
    const el = renderShapeTreeNode(document, picture, IDENTITY_MAP, CONTEXT) as HTMLImageElement;
    expect(el.style.borderRadius).toBe('50%');
  });
});

describe('renderShapeTreeNode style-matrix (p:style) fallback', () => {
  const FORMAT_SCHEME: FormatScheme = {
    name: 'Office',
    fillStyles: [{ type: 'solid', color: { type: 'scheme', value: 'phClr' } }],
    lineStyles: [
      { width: 25400, fill: { type: 'solid', color: { type: 'scheme', value: 'phClr' } } },
    ],
  };

  it("resolves a shape's fillRef/lineRef when it has no explicit spPr fill/line (the PowerPoint Shape Styles gallery case)", () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 9, name: 'Oval 1' },
      properties: {},
      style: {
        fillRef: { index: 1, color: { type: 'scheme', value: 'accent1' } },
        lineRef: { index: 1, color: { type: 'scheme', value: 'accent1' } },
      },
    };
    const el = renderShapeTreeNode(
      document,
      shape,
      IDENTITY_MAP,
      contextWithFormatScheme(FORMAT_SCHEME),
    );
    expect(el.style.backgroundColor).toBe('rgb(79, 129, 189)');
    expect(el.style.borderColor).toBe('rgb(79, 129, 189)');
  });

  it("prefers a shape's own explicit spPr fill/line over its style reference", () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 10, name: 'Explicit' },
      properties: { fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } } },
      style: { fillRef: { index: 1, color: { type: 'scheme', value: 'accent1' } } },
    };
    const el = renderShapeTreeNode(
      document,
      shape,
      IDENTITY_MAP,
      contextWithFormatScheme(FORMAT_SCHEME),
    );
    expect(el.style.backgroundColor).toBe('rgb(0, 255, 0)');
  });

  it('feeds an SVG preset outline from a style reference too, not just an explicit fill/line', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 11, name: '5-point Star' },
      properties: { geometry: { type: 'preset', preset: 'star5' } },
      style: { fillRef: { index: 1, color: { type: 'scheme', value: 'accent1' } } },
    };
    const el = renderShapeTreeNode(
      document,
      shape,
      IDENTITY_MAP,
      contextWithFormatScheme(FORMAT_SCHEME),
    );
    const path = el.querySelector('svg path') as unknown as SVGPathElement;
    expect(path.style.fill).toBe('rgb(79, 129, 189)');
  });

  it("resolves a picture's style reference the same way", () => {
    const picture: Picture = {
      kind: 'picture',
      nonVisual: { id: 12, name: 'Picture 1' },
      properties: {},
      style: { fillRef: { index: 1, color: { type: 'scheme', value: 'accent1' } } },
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
    };
    const el = renderShapeTreeNode(
      document,
      picture,
      IDENTITY_MAP,
      contextWithFormatScheme(FORMAT_SCHEME),
    ) as HTMLImageElement;
    expect(el.style.backgroundColor).toBe('rgb(79, 129, 189)');
  });

  it('leaves a shape unstyled when neither spPr fill/line nor a style reference resolve', () => {
    const shape: Shape = { kind: 'shape', nonVisual: { id: 13, name: 'Plain' }, properties: {} };
    const el = renderShapeTreeNode(
      document,
      shape,
      IDENTITY_MAP,
      contextWithFormatScheme(FORMAT_SCHEME),
    );
    expect(el.style.backgroundColor).toBe('');
  });
});
