// @vitest-environment happy-dom
import type { Shape, SlideLayout, SlideMaster, TextBody } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
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

  it("applies a paragraph's alignment as text-align, approximating distributed via text-align-last", () => {
    const textBody: TextBody = {
      paragraphs: [
        { properties: { alignment: 'center' }, runs: [{ kind: 'run', text: 'Centered' }] },
        { properties: { alignment: 'distributed' }, runs: [{ kind: 'run', text: 'Distributed' }] },
        { runs: [{ kind: 'run', text: 'Default' }] },
      ],
    };

    const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
    const paragraphs = el.querySelectorAll('p');
    expect(paragraphs[0]?.style.textAlign).toBe('center');
    expect(paragraphs[1]?.style.textAlign).toBe('justify');
    expect(paragraphs[1]?.style.textAlignLast).toBe('justify');
    expect(paragraphs[2]?.style.textAlign).toBe('');
  });

  // NOTE: font-size is set in `cqw` (container query width units, see units.ts's `emuToCqw`) so
  // it scales with the slide — but happy-dom's CSSOM doesn't recognize `cqw` as a valid length
  // yet and silently drops the assignment (confirmed: real browsers with Container Query Unit
  // support, e.g. Chrome 105+, accept it fine). The `emuToCqw`/`fontSizeToEmu` conversion math is
  // covered directly, without a DOM, in units.test.ts; the tests below can't assert
  // `style.fontSize` itself as a result.

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
        formatScheme: { name: 'Office', fillStyles: [], lineStyles: [], bgFillStyles: [] },
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

  it("inherits run properties from the shape's own placeholder via the master's title style", () => {
    const master: SlideMaster = {
      commonSlideData: { shapeTree: [] },
      layouts: [],
      theme: {} as never,
      textStyles: { titleStyle: { levels: [{ runProperties: { fontSize: 4400, bold: true } }] } },
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
    // fontSize isn't independently checkable here — see the NOTE above — but it comes from the
    // same resolved list-style level object as bold, so a correct bold proves the level was
    // found and merged correctly.
    expect(span.style.fontWeight).toBe('bold');
  });

  it("inherits centered alignment from the master's title style, same as a placeholder's font size", () => {
    const master: SlideMaster = {
      commonSlideData: { shapeTree: [] },
      layouts: [],
      theme: {} as never,
      textStyles: { titleStyle: { levels: [{ alignment: 'center' }] } },
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
    expect(el.querySelector('p')?.style.textAlign).toBe('center');
  });

  describe('bullets and numbering', () => {
    it('renders a char bullet as a span before the paragraph text, styled from its own font/colour', () => {
      const textBody: TextBody = {
        paragraphs: [
          {
            properties: {
              bullet: {
                type: 'char',
                char: '•',
                font: 'Arial',
                color: { type: 'srgb', value: 'FF0000' },
              },
            },
            runs: [{ kind: 'run', text: 'Item' }],
          },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const bullet = el.querySelector('span.pptx-bullet') as HTMLElement;
      expect(bullet.textContent).toBe('•');
      expect(bullet.style.fontFamily).toBe('Arial');
      expect(bullet.style.color).toBe('rgb(255, 0, 0)');
      expect(el.querySelector('p')?.textContent).toBe('• Item');
    });

    it('numbers consecutive autoNum paragraphs at the same level sequentially', () => {
      const bullet = { type: 'autoNum' as const, scheme: 'arabicPeriod' as const };
      const textBody: TextBody = {
        paragraphs: [
          { properties: { bullet }, runs: [{ kind: 'run', text: 'One' }] },
          { properties: { bullet }, runs: [{ kind: 'run', text: 'Two' }] },
          { properties: { bullet }, runs: [{ kind: 'run', text: 'Three' }] },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const labels = [...el.querySelectorAll('span.pptx-bullet')].map((b) => b.textContent);
      expect(labels).toEqual(['1.', '2.', '3.']);
    });

    it('restarts numbering for a nested sub-list and resumes the outer count after', () => {
      const outer = { type: 'autoNum' as const, scheme: 'arabicPeriod' as const };
      const inner = { type: 'autoNum' as const, scheme: 'romanLcPeriod' as const };
      const textBody: TextBody = {
        paragraphs: [
          { properties: { bullet: outer }, runs: [{ kind: 'run', text: 'A' }] },
          { properties: { level: 1, bullet: inner }, runs: [{ kind: 'run', text: 'A-i' }] },
          { properties: { level: 1, bullet: inner }, runs: [{ kind: 'run', text: 'A-ii' }] },
          { properties: { bullet: outer }, runs: [{ kind: 'run', text: 'B' }] },
          { properties: { level: 1, bullet: inner }, runs: [{ kind: 'run', text: 'B-i' }] },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const labels = [...el.querySelectorAll('span.pptx-bullet')].map((b) => b.textContent);
      expect(labels).toEqual(['1.', 'i.', 'ii.', '2.', 'i.']);
    });

    it('an explicit buNone suppresses a bullet inherited from the shape list style', () => {
      const shapeTextBody: TextBody = {
        listStyle: { levels: [{ bullet: { type: 'char', char: '-' } }] },
        paragraphs: [
          { runs: [{ kind: 'run', text: 'Bulleted' }] },
          {
            properties: { bullet: { type: 'none' } },
            runs: [{ kind: 'run', text: 'Not bulleted' }],
          },
        ],
      };

      const el = renderTextBody(document, shapeTextBody, undefined, BARE_CONTEXT);
      const paragraphs = el.querySelectorAll('p');
      expect(paragraphs[0]?.querySelector('span.pptx-bullet')?.textContent).toBe('-');
      expect(paragraphs[1]?.querySelector('span.pptx-bullet')).toBeNull();
    });

    it("falls back to the paragraph's ambient run colour when the bullet has none of its own", () => {
      const textBody: TextBody = {
        paragraphs: [
          {
            properties: {
              bullet: { type: 'char', char: '•' },
              defaultRunProperties: {
                fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
              },
            },
            runs: [{ kind: 'run', text: 'Item' }],
          },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const bullet = el.querySelector('span.pptx-bullet') as HTMLElement;
      expect(bullet.style.color).toBe('rgb(0, 255, 0)');
    });

    it('renders no bullet span when a paragraph has none', () => {
      const textBody: TextBody = { paragraphs: [{ runs: [{ kind: 'run', text: 'Plain' }] }] };
      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      expect(el.querySelector('span.pptx-bullet')).toBeNull();
    });

    it('renders no bullet for a trailing empty paragraph, even with an inherited char bullet', () => {
      const shapeTextBody: TextBody = {
        listStyle: { levels: [{ bullet: { type: 'char', char: '•' } }] },
        paragraphs: [{ runs: [{ kind: 'run', text: 'Item' }] }, { runs: [] }],
      };

      const el = renderTextBody(document, shapeTextBody, undefined, BARE_CONTEXT);
      const paragraphs = el.querySelectorAll('p');
      expect(paragraphs[0]?.querySelector('span.pptx-bullet')?.textContent).toBe('•');
      expect(paragraphs[1]?.querySelector('span.pptx-bullet')).toBeNull();
      expect(paragraphs[1]?.querySelector('br')).not.toBeNull();
    });

    it('does not consume a number for a trailing empty paragraph', () => {
      const bullet = { type: 'autoNum' as const, scheme: 'arabicPeriod' as const };
      const textBody: TextBody = {
        paragraphs: [
          { properties: { bullet }, runs: [{ kind: 'run', text: 'One' }] },
          { properties: { bullet }, runs: [] },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const labels = [...el.querySelectorAll('span.pptx-bullet')].map((b) => b.textContent);
      expect(labels).toEqual(['1.']);
    });

    it('a blank line between two autoNum paragraphs does not break the running count', () => {
      const bullet = { type: 'autoNum' as const, scheme: 'arabicPeriod' as const };
      const textBody: TextBody = {
        paragraphs: [
          { properties: { bullet }, runs: [{ kind: 'run', text: 'One' }] },
          { runs: [] },
          { properties: { bullet }, runs: [{ kind: 'run', text: 'Two' }] },
        ],
      };

      const el = renderTextBody(document, textBody, undefined, BARE_CONTEXT);
      const labels = [...el.querySelectorAll('span.pptx-bullet')].map((b) => b.textContent);
      expect(labels).toEqual(['1.', '2.']);
    });
  });

  describe('unsupported-feature reporting', () => {
    it('reports wrap="none", with the passed-through shape ref', () => {
      const reportUnsupported = vi.fn();
      const textBody: TextBody = { properties: { wrap: 'none' }, paragraphs: [] };

      renderTextBody(
        document,
        textBody,
        undefined,
        { ...BARE_CONTEXT, reportUnsupported },
        undefined,
        { id: 30, name: 'Text Box 1' },
      );

      expect(reportUnsupported).toHaveBeenCalledWith('text-wrap-unmodeled', expect.any(String), {
        id: 30,
        name: 'Text Box 1',
      });
    });

    it('does not report wrap="square" or an absent wrap', () => {
      const reportUnsupported = vi.fn();
      for (const wrap of ['square' as const, undefined]) {
        const textBody: TextBody = { properties: wrap ? { wrap } : {}, paragraphs: [] };
        renderTextBody(document, textBody, undefined, { ...BARE_CONTEXT, reportUnsupported });
      }

      expect(reportUnsupported).not.toHaveBeenCalled();
    });
  });
});
