// @vitest-environment happy-dom
import type {
  ConnectionShape,
  FormatScheme,
  GraphicFrame,
  Picture,
  Shape,
  SlideLayout,
} from '@pptx2html/presentation';
import { IDENTITY_MAP } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
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

  it('renders custGeom path data as an SVG path (a boolean-subtract cutout, two subpaths)', () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 9, name: 'Freeform 1' },
      properties: {
        geometry: {
          type: 'custom',
          pathLst: [
            {
              width: 100,
              height: 100,
              commands: [
                { type: 'moveTo', point: { x: 0, y: 0 } },
                { type: 'lnTo', point: { x: 100, y: 0 } },
                { type: 'lnTo', point: { x: 100, y: 100 } },
                { type: 'close' },
              ],
            },
            {
              width: 100,
              height: 100,
              commands: [
                { type: 'moveTo', point: { x: 25, y: 25 } },
                { type: 'lnTo', point: { x: 75, y: 25 } },
                { type: 'close' },
              ],
            },
          ],
        },
        fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
      },
    };
    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);
    expect(el.style.backgroundColor).toBe('');
    const svg = el.querySelector('svg');
    expect(svg).not.toBeNull();
    const path = svg?.querySelector('path');
    expect(path?.getAttribute('d')).toBe('M 0 0 L 100 0 L 100 100 Z M 25 25 L 75 25 Z');
    expect((path as unknown as SVGPathElement)?.style.fill).toBe('rgb(255, 0, 0)');
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
    bgFillStyles: [],
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

describe('renderShapeTreeNode text body anchor and fontRef fallback', () => {
  function shapeWithText(anchor: 't' | 'ctr' | 'b' | undefined): Shape {
    return {
      kind: 'shape',
      nonVisual: { id: 14, name: 'Oval 1' },
      properties: {},
      style: { fontRef: { collection: 'minor', color: { type: 'scheme', value: 'lt1' } } },
      textBody: {
        ...(anchor ? { properties: { anchor } } : {}),
        paragraphs: [{ runs: [{ kind: 'run', text: '42' }] }],
      },
    };
  }

  it('vertically centers the text body via flexbox when bodyPr anchor="ctr"', () => {
    const el = renderShapeTreeNode(document, shapeWithText('ctr'), IDENTITY_MAP, CONTEXT);
    expect(el.style.display).toBe('flex');
    expect(el.style.flexDirection).toBe('column');
    expect(el.style.justifyContent).toBe('center');
  });

  it('defaults to top-anchored (flex-start) when bodyPr has no anchor', () => {
    const el = renderShapeTreeNode(document, shapeWithText(undefined), IDENTITY_MAP, CONTEXT);
    expect(el.style.justifyContent).toBe('flex-start');
  });

  it('bottom-anchors via flex-end for anchor="b"', () => {
    const el = renderShapeTreeNode(document, shapeWithText('b'), IDENTITY_MAP, CONTEXT);
    expect(el.style.justifyContent).toBe('flex-end');
  });

  it("falls back to the shape's own p:style fontRef colour for a run with none of its own", () => {
    const el = renderShapeTreeNode(
      document,
      shapeWithText('ctr'),
      IDENTITY_MAP,
      contextWithFormatScheme({ name: 'Office', fillStyles: [], lineStyles: [], bgFillStyles: [] }),
    );
    const run = el.querySelector('.pptx-run') as HTMLElement;
    expect(run.textContent).toBe('42');
    expect(run.style.color).toBe('rgb(255, 255, 255)'); // lt1, per contextWithFormatScheme's theme
  });
});

describe('renderShapeTreeNode unsupported-feature reporting', () => {
  function contextWithReport(): {
    context: RenderContext;
    reportUnsupported: ReturnType<typeof vi.fn>;
  } {
    const reportUnsupported = vi.fn();
    return {
      context: { ...CONTEXT, reportUnsupported },
      reportUnsupported,
    };
  }

  it("reports an unmodeled preset falling back to a rectangle, with the shape's own id/name", () => {
    const { context, reportUnsupported } = contextWithReport();
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 20, name: 'Cloud 1' },
      properties: { geometry: { type: 'preset', preset: 'cloud' } },
    };

    renderShapeTreeNode(document, shape, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith(
      'shape-geometry-unmodeled',
      expect.stringContaining('"cloud"'),
      { id: 20, name: 'Cloud 1' },
    );
  });

  it('reports custom geometry (custGeom) distinctly from an unmodeled preset', () => {
    const { context, reportUnsupported } = contextWithReport();
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 21, name: 'Freeform 1' },
      properties: { geometry: { type: 'custom' } },
    };

    renderShapeTreeNode(document, shape, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith(
      'shape-geometry-unmodeled',
      expect.stringContaining('custGeom'),
      { id: 21, name: 'Freeform 1' },
    );
  });

  it('does not report for a plain rect, or a preset covered by an SVG path/border-radius', () => {
    const { context, reportUnsupported } = contextWithReport();
    for (const geometry of [
      { type: 'preset' as const, preset: 'rect' },
      { type: 'preset' as const, preset: 'ellipse' },
      { type: 'preset' as const, preset: 'triangle' },
      undefined,
    ]) {
      const shape: Shape = {
        kind: 'shape',
        nonVisual: { id: 22, name: 'Shape' },
        properties: { geometry },
      };
      renderShapeTreeNode(document, shape, IDENTITY_MAP, context);
    }

    expect(reportUnsupported).not.toHaveBeenCalled();
  });

  it('reports an unrendered graphic (chart/smartArt/oleObject) placeholder', () => {
    const { context, reportUnsupported } = contextWithReport();
    const frame: GraphicFrame = {
      kind: 'graphicFrame',
      nonVisual: { id: 23, name: 'Chart 1' },
      transform: { offset: { x: 0, y: 0 }, extents: { width: 1, height: 1 } },
      graphic: { type: 'chart' },
    };

    renderShapeTreeNode(document, frame, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith(
      'graphic-placeholder-unmodeled',
      expect.stringContaining('chart'),
      { id: 23, name: 'Chart 1' },
    );
  });

  it('does not report for a table graphicFrame', () => {
    const { context, reportUnsupported } = contextWithReport();
    const frame: GraphicFrame = {
      kind: 'graphicFrame',
      nonVisual: { id: 24, name: 'Table 1' },
      transform: { offset: { x: 0, y: 0 }, extents: { width: 1, height: 1 } },
      graphic: { type: 'table', columns: [], rows: [] },
    };

    renderShapeTreeNode(document, frame, IDENTITY_MAP, context);

    expect(reportUnsupported).not.toHaveBeenCalled();
  });

  it("reports that a connector's line is never rendered", () => {
    const { context, reportUnsupported } = contextWithReport();
    const connector: ConnectionShape = {
      kind: 'connector',
      nonVisual: { id: 25, name: 'Straight Connector 1' },
      properties: {},
    };

    renderShapeTreeNode(document, connector, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith('connector-line-unmodeled', expect.any(String), {
      id: 25,
      name: 'Straight Connector 1',
    });
  });

  it('reports a gradient fill / non-solid outline colour dropped on an SVG-path preset', () => {
    const { context, reportUnsupported } = contextWithReport();
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 26, name: 'Triangle 1' },
      properties: {
        geometry: { type: 'preset', preset: 'triangle' },
        fill: {
          type: 'gradient',
          stops: [
            { position: 0, color: { type: 'srgb', value: 'FF0000' } },
            { position: 100, color: { type: 'srgb', value: '0000FF' } },
          ],
        },
        line: {
          width: 12700,
          fill: {
            type: 'gradient',
            stops: [
              { position: 0, color: { type: 'srgb', value: 'FF0000' } },
              { position: 100, color: { type: 'srgb', value: '0000FF' } },
            ],
          },
        },
      },
    };

    renderShapeTreeNode(document, shape, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith(
      'svg-preset-fill-unmodeled',
      expect.stringContaining('"gradient"'),
      { id: 26, name: 'Triangle 1' },
    );
    expect(reportUnsupported).toHaveBeenCalledWith(
      'svg-preset-line-unmodeled',
      expect.stringContaining('"gradient"'),
      { id: 26, name: 'Triangle 1' },
    );
  });

  it('does not report a solid fill/line on an SVG-path preset', () => {
    const { context, reportUnsupported } = contextWithReport();
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 27, name: 'Triangle 2' },
      properties: {
        geometry: { type: 'preset', preset: 'triangle' },
        fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
        line: { width: 12700, fill: { type: 'solid', color: { type: 'srgb', value: '000000' } } },
      },
    };

    renderShapeTreeNode(document, shape, IDENTITY_MAP, context);

    expect(reportUnsupported).not.toHaveBeenCalled();
  });

  it('reports that a picture crop is not applied outside rect/roundRect/ellipse', () => {
    const { context, reportUnsupported } = contextWithReport();
    const picture: Picture = {
      kind: 'picture',
      nonVisual: { id: 28, name: 'Picture 1' },
      properties: { geometry: { type: 'preset', preset: 'triangle' } },
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
    };

    renderShapeTreeNode(document, picture, IDENTITY_MAP, context);

    expect(reportUnsupported).toHaveBeenCalledWith(
      'picture-crop-unmodeled',
      expect.stringContaining('"triangle"'),
      { id: 28, name: 'Picture 1' },
    );
  });

  it('does not report a picture crop for rect/roundRect/ellipse, or no geometry at all', () => {
    const { context, reportUnsupported } = contextWithReport();
    for (const geometry of [
      { type: 'preset' as const, preset: 'rect' },
      { type: 'preset' as const, preset: 'roundRect' },
      { type: 'preset' as const, preset: 'ellipse' },
      undefined,
    ]) {
      const picture: Picture = {
        kind: 'picture',
        nonVisual: { id: 29, name: 'Picture 2' },
        properties: { geometry },
        image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
      };
      renderShapeTreeNode(document, picture, IDENTITY_MAP, context);
    }

    expect(reportUnsupported).not.toHaveBeenCalled();
  });
});

describe('renderShapeTreeNode data-pptx-shape-id', () => {
  it("tags the rendered element with the node's own nonVisual id, for animation.ts to target later", () => {
    const shape: Shape = {
      kind: 'shape',
      nonVisual: { id: 99, name: 'Oval 1' },
      properties: {},
    };

    const el = renderShapeTreeNode(document, shape, IDENTITY_MAP, CONTEXT);

    expect(el.dataset.pptxShapeId).toBe('99');
  });

  it('tags a group (its wrapper carries the id an AnimationTarget can address)', () => {
    const el = renderShapeTreeNode(
      document,
      {
        kind: 'group',
        nonVisual: { id: 5, name: 'Group 1' },
        transform: { offset: { x: 0, y: 0 }, extents: { width: 1, height: 1 } },
        children: [],
      },
      IDENTITY_MAP,
      CONTEXT,
    );

    expect(el.dataset.pptxShapeId).toBe('5');
  });
});
