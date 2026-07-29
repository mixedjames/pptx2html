import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import type { MediaResolver } from './fill.js';
import { parseTextBody } from './text.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

const noMedia: MediaResolver = () => undefined;

describe('drawingml/text', () => {
  it('preserves run/break/field interleaving within a paragraph', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:bodyPr wrap="square" anchor="ctr"/>
        <a:p>
          <a:pPr algn="ctr" lvl="1"/>
          <a:r><a:t>Hello</a:t></a:r>
          <a:br/>
          <a:fld type="slidenum"><a:t>1</a:t></a:fld>
          <a:r><a:t>world</a:t></a:r>
        </a:p>
      </p:txBody>`,
    );

    const body = parseTextBody(node, noMedia);
    expect(body?.properties).toEqual({ wrap: 'square', anchor: 'ctr' });
    expect(body?.paragraphs).toHaveLength(1);
    const paragraph = body!.paragraphs[0]!;
    expect(paragraph.properties).toEqual({ alignment: 'center', level: 1 });
    expect(paragraph.runs).toEqual([
      { kind: 'run', text: 'Hello' },
      { kind: 'break' },
      { kind: 'field', fieldType: 'slidenum', cachedText: '1' },
      { kind: 'run', text: 'world' },
    ]);
  });

  it('parses run character formatting', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p>
          <a:r>
            <a:rPr b="1" i="0" u="sng" strike="sngStrike" sz="1800" lang="en-US">
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
              <a:latin typeface="Calibri"/>
            </a:rPr>
            <a:t>Styled</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    );

    const run = parseTextBody(node, noMedia)!.paragraphs[0]!.runs[0];
    expect(run).toEqual({
      kind: 'run',
      text: 'Styled',
      properties: {
        bold: true,
        italic: false,
        underline: true,
        strikethrough: true,
        fontSize: 1800,
        fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
        typeface: 'Calibri',
        language: 'en-US',
      },
    });
  });

  it('returns undefined for a missing text body', () => {
    expect(parseTextBody(undefined, noMedia)).toBeUndefined();
  });
});
