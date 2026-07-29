import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/parse.js';
import { parseNonVisualDrawingProperties } from './shape-common.js';

describe('parseNonVisualDrawingProperties', () => {
  it('parses id/name/description/hidden from cNvPr', () => {
    const [nvSpPr] = parseXml(
      `<p:nvSpPr xmlns:p="p"><p:cNvPr id="2" name="Title 1" descr="a title" hidden="1"/></p:nvSpPr>`,
    );

    expect(parseNonVisualDrawingProperties(nvSpPr)).toEqual({
      id: 2,
      name: 'Title 1',
      description: 'a title',
      hidden: true,
    });
  });

  it('parses a placeholder type and index from nvPr/ph', () => {
    const [nvSpPr] = parseXml(
      `<p:nvSpPr xmlns:p="p">
        <p:cNvPr id="3" name="Content Placeholder 2"/>
        <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
      </p:nvSpPr>`,
    );

    expect(parseNonVisualDrawingProperties(nvSpPr)?.placeholder).toEqual({
      type: 'body',
      index: 1,
    });
  });

  it('defaults placeholder type to "obj" and idx to 0 when omitted, per the spec', () => {
    const [nvSpPr] = parseXml(
      `<p:nvSpPr xmlns:p="p"><p:cNvPr id="2" name="Title 1"/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr>`,
    );

    expect(parseNonVisualDrawingProperties(nvSpPr)?.placeholder).toEqual({
      type: 'obj',
      index: 0,
    });
  });

  it('leaves placeholder undefined when there is no ph element', () => {
    const [nvSpPr] = parseXml(`<p:nvSpPr xmlns:p="p"><p:cNvPr id="2" name="Title 1"/></p:nvSpPr>`);

    expect(parseNonVisualDrawingProperties(nvSpPr)?.placeholder).toBeUndefined();
  });
});
