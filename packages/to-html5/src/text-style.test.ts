import type {
  FontScheme,
  Paragraph,
  PlaceholderType,
  Shape,
  SlideLayout,
  SlideMaster,
  TextBody,
  TextRunElement,
} from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import type { RenderContext } from './render-context.js';
import {
  resolveEffectiveAlignment,
  resolveEffectiveRunProperties,
  resolveTypeface,
} from './text-style.js';

function placeholderShape(
  id: number,
  type: PlaceholderType,
  index: number,
  textBody?: TextBody,
): Shape {
  return {
    kind: 'shape',
    nonVisual: { id, name: `Placeholder ${id}`, placeholder: { type, index } },
    properties: {},
    ...(textBody ? { textBody } : {}),
  };
}

function master(shapes: readonly Shape[], textStyles?: SlideMaster['textStyles']): SlideMaster {
  return {
    commonSlideData: { shapeTree: shapes },
    theme: {} as never,
    layouts: [],
    ...(textStyles ? { textStyles } : {}),
  };
}

function layout(shapes: readonly Shape[], masterShapes: readonly Shape[] = []): SlideLayout {
  return {
    commonSlideData: { shapeTree: shapes },
    master: master(masterShapes),
    type: 'blank',
  };
}

function contextFor(
  layoutValue: SlideLayout | undefined,
  defaultTextStyle?: RenderContext['defaultTextStyle'],
): RenderContext {
  return {
    slideSize: { width: 1, height: 1 },
    layout: layoutValue,
    defaultTextStyle,
  };
}

const run: TextRunElement = { kind: 'run', text: 'Hi' };
const paragraph: Paragraph = { runs: [run] };
const emptyTextBody: TextBody = { paragraphs: [paragraph] };

describe('resolveEffectiveRunProperties', () => {
  it('returns nothing when the chain is entirely empty', () => {
    const context = contextFor(undefined);
    expect(
      resolveEffectiveRunProperties(run, paragraph, emptyTextBody, undefined, context),
    ).toEqual({});
  });

  it("falls back to the presentation's default text style", () => {
    const context = contextFor(undefined, { levels: [{ runProperties: { fontSize: 1800 } }] });
    expect(
      resolveEffectiveRunProperties(run, paragraph, emptyTextBody, undefined, context),
    ).toEqual({
      fontSize: 1800,
    });
  });

  it("uses the master's titleStyle for a title placeholder, bodyStyle otherwise", () => {
    const m = master([], {
      titleStyle: { levels: [{ runProperties: { fontSize: 4400, bold: true } }] },
      bodyStyle: { levels: [{ runProperties: { fontSize: 2800 } }] },
    });
    const l: SlideLayout = { commonSlideData: { shapeTree: [] }, master: m, type: 'title' };

    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        emptyTextBody,
        { type: 'title', index: 0 },
        contextFor(l),
      ),
    ).toEqual({ fontSize: 4400, bold: true });
    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        emptyTextBody,
        { type: 'body', index: 1 },
        contextFor(l),
      ),
    ).toEqual({ fontSize: 2800 });
  });

  it('uses otherStyle for a non-placeholder shape', () => {
    const m = master([], { otherStyle: { levels: [{ runProperties: { italic: true } }] } });
    const l: SlideLayout = { commonSlideData: { shapeTree: [] }, master: m, type: 'blank' };
    expect(
      resolveEffectiveRunProperties(run, paragraph, emptyTextBody, undefined, contextFor(l)),
    ).toEqual({
      italic: true,
    });
  });

  it("falls through layout -> master for a matching placeholder's own list style", () => {
    const masterPlaceholder = placeholderShape(20, 'ftr', 3, {
      paragraphs: [],
      listStyle: { levels: [{ runProperties: { fontSize: 1200 } }] },
    });
    const l = layout([], [masterPlaceholder]);
    const context = contextFor(l);
    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        emptyTextBody,
        { type: 'ftr', index: 3 },
        context,
      ),
    ).toEqual({ fontSize: 1200 });
  });

  it("prefers the layout placeholder's list style over the master's", () => {
    const masterPlaceholder = placeholderShape(20, 'ftr', 3, {
      paragraphs: [],
      listStyle: { levels: [{ runProperties: { fontSize: 1200, bold: true } }] },
    });
    const layoutPlaceholder = placeholderShape(10, 'ftr', 3, {
      paragraphs: [],
      listStyle: { levels: [{ runProperties: { fontSize: 1400 } }] },
    });
    const l = layout([layoutPlaceholder], [masterPlaceholder]);
    const context = contextFor(l);
    // fontSize comes from the layout (closer), bold still falls through from the master.
    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        emptyTextBody,
        { type: 'ftr', index: 3 },
        context,
      ),
    ).toEqual({ fontSize: 1400, bold: true });
  });

  it("prefers the shape's own list style over placeholder inheritance", () => {
    const layoutPlaceholder = placeholderShape(10, 'body', 1, {
      paragraphs: [],
      listStyle: { levels: [{ runProperties: { fontSize: 1400 } }] },
    });
    const l = layout([layoutPlaceholder]);
    const shapeTextBody: TextBody = {
      paragraphs: [paragraph],
      listStyle: { levels: [{ runProperties: { fontSize: 2000 } }] },
    };
    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        shapeTextBody,
        { type: 'body', index: 1 },
        contextFor(l),
      ),
    ).toEqual({ fontSize: 2000 });
  });

  it("prefers the paragraph's defRPr over list styles, and the run's own rPr over everything", () => {
    const shapeTextBody: TextBody = {
      paragraphs: [],
      listStyle: { levels: [{ runProperties: { fontSize: 2000, italic: true } }] },
    };
    const p: Paragraph = { properties: { defaultRunProperties: { fontSize: 2400 } }, runs: [] };
    const styledRun: TextRunElement = { kind: 'run', text: 'Hi', properties: { fontSize: 3000 } };

    expect(
      resolveEffectiveRunProperties(styledRun, p, shapeTextBody, undefined, contextFor(undefined)),
    ).toEqual({ fontSize: 3000, italic: true });
  });

  it('reads the level-specific defaults for the paragraph outline level', () => {
    const shapeTextBody: TextBody = {
      paragraphs: [],
      listStyle: {
        levels: [{ runProperties: { fontSize: 2000 } }, { runProperties: { fontSize: 1600 } }],
      },
    };
    const level1Paragraph: Paragraph = { properties: { level: 1 }, runs: [] };
    expect(
      resolveEffectiveRunProperties(
        run,
        level1Paragraph,
        shapeTextBody,
        undefined,
        contextFor(undefined),
      ),
    ).toEqual({ fontSize: 1600 });
  });
});

describe('resolveEffectiveAlignment', () => {
  it('returns undefined when the chain is entirely empty', () => {
    expect(
      resolveEffectiveAlignment(paragraph, emptyTextBody, undefined, contextFor(undefined)),
    ).toBeUndefined();
  });

  it("a paragraph's own alignment always wins, even over its shape's list style", () => {
    const shapeTextBody: TextBody = {
      paragraphs: [],
      listStyle: { levels: [{ alignment: 'right' }] },
    };
    const centered: Paragraph = { properties: { alignment: 'center' }, runs: [] };
    expect(
      resolveEffectiveAlignment(centered, shapeTextBody, undefined, contextFor(undefined)),
    ).toBe('center');
  });

  it("inherits a title placeholder's centered alignment from the master's titleStyle", () => {
    const m = master([], { titleStyle: { levels: [{ alignment: 'center' }] } });
    const l: SlideLayout = { commonSlideData: { shapeTree: [] }, master: m, type: 'title' };
    expect(
      resolveEffectiveAlignment(
        paragraph,
        emptyTextBody,
        { type: 'title', index: 0 },
        contextFor(l),
      ),
    ).toBe('center');
  });

  it("prefers the layout placeholder's list style over the master's", () => {
    const masterPlaceholder = placeholderShape(20, 'body', 1, {
      paragraphs: [],
      listStyle: { levels: [{ alignment: 'right' }] },
    });
    const layoutPlaceholder = placeholderShape(10, 'body', 1, {
      paragraphs: [],
      listStyle: { levels: [{ alignment: 'center' }] },
    });
    const l = layout([layoutPlaceholder], [masterPlaceholder]);
    expect(
      resolveEffectiveAlignment(
        paragraph,
        emptyTextBody,
        { type: 'body', index: 1 },
        contextFor(l),
      ),
    ).toBe('center');
  });

  it("falls back to the presentation's default text style", () => {
    const context = contextFor(undefined, { levels: [{ alignment: 'justify' }] });
    expect(resolveEffectiveAlignment(paragraph, emptyTextBody, undefined, context)).toBe('justify');
  });

  it('reads the level-specific alignment for the paragraph outline level', () => {
    const shapeTextBody: TextBody = {
      paragraphs: [],
      listStyle: { levels: [{ alignment: 'left' }, { alignment: 'right' }] },
    };
    const level1Paragraph: Paragraph = { properties: { level: 1 }, runs: [] };
    expect(
      resolveEffectiveAlignment(level1Paragraph, shapeTextBody, undefined, contextFor(undefined)),
    ).toBe('right');
  });

  it('is independent of run-property resolution on the same list-style level', () => {
    const shapeTextBody: TextBody = {
      paragraphs: [],
      listStyle: { levels: [{ alignment: 'center', runProperties: { bold: true } }] },
    };
    expect(
      resolveEffectiveAlignment(paragraph, shapeTextBody, undefined, contextFor(undefined)),
    ).toBe('center');
    expect(
      resolveEffectiveRunProperties(
        run,
        paragraph,
        shapeTextBody,
        undefined,
        contextFor(undefined),
      ),
    ).toEqual({ bold: true });
  });
});

describe('resolveTypeface', () => {
  const fontScheme: FontScheme = {
    name: 'Office',
    majorFont: { latin: 'Calibri Light', eastAsian: 'MS PGothic' },
    minorFont: { latin: 'Calibri' },
  };

  it('resolves theme font tokens against the font scheme', () => {
    expect(resolveTypeface('+mj-lt', fontScheme)).toBe('Calibri Light');
    expect(resolveTypeface('+mn-lt', fontScheme)).toBe('Calibri');
    expect(resolveTypeface('+mj-ea', fontScheme)).toBe('MS PGothic');
  });

  it('falls back to latin when the requested script is absent from the font collection', () => {
    expect(resolveTypeface('+mn-ea', fontScheme)).toBe('Calibri');
  });

  it('passes a literal typeface straight through', () => {
    expect(resolveTypeface('Georgia', fontScheme)).toBe('Georgia');
  });

  it('returns a theme token unresolved when no font scheme is available', () => {
    expect(resolveTypeface('+mj-lt', undefined)).toBe('+mj-lt');
  });

  it('returns undefined for an absent typeface', () => {
    expect(resolveTypeface(undefined, fontScheme)).toBeUndefined();
  });
});
