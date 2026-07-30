import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import type { MediaResolver } from './fill.js';
import { parseTextBody, parseTextListStyle } from './text.js';

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

  it('parses a paragraph defRPr as ParagraphProperties.defaultRunProperties', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p>
          <a:pPr><a:defRPr b="1" sz="2400"/></a:pPr>
          <a:r><a:t>Hi</a:t></a:r>
        </a:p>
      </p:txBody>`,
    );

    const paragraph = parseTextBody(node, noMedia)!.paragraphs[0]!;
    expect(paragraph.properties?.defaultRunProperties).toEqual({ bold: true, fontSize: 2400 });
  });

  it('parses a buChar bullet with its font/colour/size and marL/indent', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p>
          <a:pPr marL="457200" indent="-457200">
            <a:buFont typeface="Arial"/>
            <a:buClr><a:srgbClr val="FF0000"/></a:buClr>
            <a:buSzPct val="80000"/>
            <a:buChar char="•"/>
          </a:pPr>
          <a:r><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>`,
    );

    const properties = parseTextBody(node, noMedia)!.paragraphs[0]!.properties;
    expect(properties?.marginLeft).toBe(457200);
    expect(properties?.indent).toBe(-457200);
    expect(properties?.bullet).toEqual({
      type: 'char',
      char: '•',
      font: 'Arial',
      color: { type: 'srgb', value: 'FF0000' },
      sizePercent: 80000,
    });
  });

  it('parses a buAutoNum bullet with its scheme and startAt', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p>
          <a:pPr><a:buAutoNum type="arabicParenR" startAt="3"/></a:pPr>
          <a:r><a:t>Item</a:t></a:r>
        </a:p>
      </p:txBody>`,
    );

    const bullet = parseTextBody(node, noMedia)!.paragraphs[0]!.properties?.bullet;
    expect(bullet).toEqual({ type: 'autoNum', scheme: 'arabicParenR', startAt: 3 });
  });

  it('falls back to arabicPeriod for an unrecognized buAutoNum type', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p><a:pPr><a:buAutoNum type="bogusScheme"/></a:pPr><a:r><a:t>Item</a:t></a:r></a:p>
      </p:txBody>`,
    );

    const bullet = parseTextBody(node, noMedia)!.paragraphs[0]!.properties?.bullet;
    expect(bullet).toEqual({ type: 'autoNum', scheme: 'arabicPeriod' });
  });

  it('parses buNone as an explicit no-bullet', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>Item</a:t></a:r></a:p>
      </p:txBody>`,
    );

    const bullet = parseTextBody(node, noMedia)!.paragraphs[0]!.properties?.bullet;
    expect(bullet).toEqual({ type: 'none' });
  });

  it('parses a shape-level lstStyle as TextBody.listStyle', () => {
    const node = firstNode(
      `<p:txBody xmlns:p="p" xmlns:a="a">
        <a:lstStyle>
          <a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr>
          <a:lvl2pPr><a:defRPr sz="1600" i="1"/></a:lvl2pPr>
        </a:lstStyle>
        <a:p><a:r><a:t>Hi</a:t></a:r></a:p>
      </p:txBody>`,
    );

    const body = parseTextBody(node, noMedia);
    expect(body?.listStyle?.levels[0]).toEqual({ runProperties: { fontSize: 1800 } });
    expect(body?.listStyle?.levels[1]).toEqual({
      runProperties: { fontSize: 1600, italic: true },
    });
    expect(body?.listStyle?.levels[2]).toBeUndefined();
  });
});

describe('parseTextListStyle', () => {
  it('parses each lvlNpPr child’s algn and defRPr, indexed 0-based, skipping absent levels', () => {
    const node = firstNode(
      `<a:lstStyle xmlns:a="a">
        <a:lvl1pPr algn="ctr"><a:defRPr sz="3200" b="1"/></a:lvl1pPr>
        <a:lvl3pPr><a:defRPr><a:latin typeface="Georgia"/></a:defRPr></a:lvl3pPr>
      </a:lstStyle>`,
    );

    const style = parseTextListStyle(node, noMedia);
    expect(style?.levels).toHaveLength(9);
    expect(style?.levels[0]).toEqual({
      alignment: 'center',
      runProperties: { fontSize: 3200, bold: true },
    });
    expect(style?.levels[1]).toBeUndefined();
    expect(style?.levels[2]).toEqual({ runProperties: { typeface: 'Georgia' } });
  });

  it('parses a level with only an algn and no defRPr', () => {
    const node = firstNode(`<a:lstStyle xmlns:a="a"><a:lvl1pPr algn="r"/></a:lstStyle>`);
    expect(parseTextListStyle(node, noMedia)?.levels[0]).toEqual({ alignment: 'right' });
  });

  it('parses a level with a bullet and marL/indent', () => {
    const node = firstNode(
      `<a:lstStyle xmlns:a="a">
        <a:lvl1pPr marL="228600" indent="-228600"><a:buChar char="-"/></a:lvl1pPr>
      </a:lstStyle>`,
    );

    expect(parseTextListStyle(node, noMedia)?.levels[0]).toEqual({
      marginLeft: 228600,
      indent: -228600,
      bullet: { type: 'char', char: '-' },
    });
  });

  it('returns undefined when no level carries an algn or defRPr', () => {
    const node = firstNode(`<a:lstStyle xmlns:a="a"><a:lvl1pPr/></a:lstStyle>`);
    expect(parseTextListStyle(node, noMedia)).toBeUndefined();
  });

  it('returns undefined for a missing node', () => {
    expect(parseTextListStyle(undefined, noMedia)).toBeUndefined();
  });
});
