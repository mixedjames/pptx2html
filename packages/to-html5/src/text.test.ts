// @vitest-environment happy-dom
import type { Shape, SlideLayout, SlideMaster, TextBody } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import type { RenderContext } from './render-context.js';
import { renderTextBody } from './text.js';

const BARE_CONTEXT: RenderContext = {
  slideSize: { width: 1, height: 1 },
  layout: undefined,
  defaultTextStyle: undefined,
};

describe('renderTextBody', () => {
  it('renders paragraphs as <p>, runs/fields as <span>, breaks as <br>', () => {
    const textBody: TextBody = {
      paragraphs: [
        {
          runs: [
            { kind: 'run', text: 'Hello ' },
            { kind: 'break' },
            { kind: 'run', text: 'world' },
            { kind: 'field', fieldType: 'slidenum', cachedText: '1' },
          ],
        },
        { runs: [] },
      ],
    };

    const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
    const paragraphs = el.querySelectorAll('p');
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.querySelectorAll('br')).toHaveLength(1);
    expect(paragraphs[0]?.querySelectorAll('span.pptx-run')).toHaveLength(3);
    expect(paragraphs[0]?.textContent).toBe('Hello world1');
    expect(paragraphs[1]?.querySelector('br')).not.toBeNull();
  });

  it("applies a run's own character formatting as inline CSS", () => {
    const textBody: TextBody = {
      paragraphs: [
        {
          runs: [
            {
              kind: 'run',
              text: 'Styled',
              properties: {
                bold: true,
                italic: true,
                underline: true,
                strikethrough: true,
                fontSize: 1800,
                typeface: 'Georgia',
                fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
              },
            },
          ],
        },
      ],
    };

    const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
    const span = el.querySelector('span.pptx-run') as HTMLElement;
    // The CSSOM re-serializes a quoted single-token family name without the quotes.
    expect(span.style.fontFamily).toBe('Georgia');
    expect(span.style.fontSize).toBe('18pt');
    expect(span.style.fontWeight).toBe('bold');
    expect(span.style.fontStyle).toBe('italic');
    expect(span.style.textDecoration).toBe('underline line-through');
    expect(span.style.color).toBe('rgb(255, 0, 0)');
  });

  it('resolves a theme font token against the layout master theme', () => {
    const masterWithTheme: SlideMaster = {
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
          accent1: { type: 'srgb', value: '000000' },
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
        formatScheme: { name: 'Office' },
      },
    };
    const layout: SlideLayout = {
      commonSlideData: { shapeTree: [] },
      master: masterWithTheme,
      type: 'blank',
    };
    const context: RenderContext = {
      slideSize: { width: 1, height: 1 },
      layout,
      defaultTextStyle: undefined,
    };

    const textBody: TextBody = {
      paragraphs: [{ runs: [{ kind: 'run', text: 'Title', properties: { typeface: '+mj-lt' } }] }],
    };

    const el = renderTextBody(document, textBody, undefined, context);
    const span = el.querySelector('span.pptx-run') as HTMLElement;
    expect(span.style.fontFamily).toBe('"Calibri Light"');
  });

  it("inherits font size from the shape's own placeholder via the master's title style", () => {
    const master: SlideMaster = {
      commonSlideData: { shapeTree: [] },
      layouts: [],
      theme: {} as never,
      textStyles: { titleStyle: { levels: [{ fontSize: 4400, bold: true }] } },
    };
    const layout: SlideLayout = { commonSlideData: { shapeTree: [] }, master, type: 'title' };
    const context: RenderContext = {
      slideSize: { width: 1, height: 1 },
      layout,
      defaultTextStyle: undefined,
    };

    const titlePlaceholder: Shape['nonVisual']['placeholder'] = { type: 'title', index: 0 };
    const textBody: TextBody = { paragraphs: [{ runs: [{ kind: 'run', text: 'Title' }] }] };

    const el = renderTextBody(document, textBody, titlePlaceholder, context);
    const span = el.querySelector('span.pptx-run') as HTMLElement;
    expect(span.style.fontSize).toBe('44pt');
    expect(span.style.fontWeight).toBe('bold');
  });
});
