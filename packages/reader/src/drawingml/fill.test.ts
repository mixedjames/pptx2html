import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import { parseChildFill, type MediaResolver } from './fill.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

const noMedia: MediaResolver = () => undefined;

describe('drawingml/fill', () => {
  it('parses noFill', () => {
    const node = firstNode('<a:spPr xmlns:a="a"><a:noFill/></a:spPr>');
    expect(parseChildFill(node, noMedia)).toEqual({ type: 'none' });
  });

  it('parses solidFill', () => {
    const node = firstNode(
      '<a:spPr xmlns:a="a"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:spPr>',
    );
    expect(parseChildFill(node, noMedia)).toEqual({
      type: 'solid',
      color: { type: 'srgb', value: '00FF00' },
    });
  });

  it('parses a linear gradFill with stops in document order', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:gradFill>
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000"/>
        </a:gradFill>
      </a:spPr>`,
    );
    expect(parseChildFill(node, noMedia)).toEqual({
      type: 'gradient',
      angle: 5400000,
      stops: [
        { position: 0, color: { type: 'srgb', value: 'FFFFFF' } },
        { position: 100000, color: { type: 'srgb', value: '000000' } },
      ],
    });
  });

  it('parses pattFill', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:pattFill prst="pct25">
          <a:fgClr><a:srgbClr val="FF0000"/></a:fgClr>
          <a:bgClr><a:srgbClr val="0000FF"/></a:bgClr>
        </a:pattFill>
      </a:spPr>`,
    );
    expect(parseChildFill(node, noMedia)).toEqual({
      type: 'pattern',
      preset: 'pct25',
      foregroundColor: { type: 'srgb', value: 'FF0000' },
      backgroundColor: { type: 'srgb', value: '0000FF' },
    });
  });

  it('resolves blipFill media via the injected resolver and reads opacity/stretch', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a" xmlns:r="r">
        <a:blipFill>
          <a:blip r:embed="rId2"><a:alphaModFix amt="50000"/></a:blip>
          <a:stretch/>
        </a:blipFill>
      </a:spPr>`,
    );
    const resolveMedia: MediaResolver = (id) =>
      id === 'rId2' ? { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) } : undefined;

    expect(parseChildFill(node, resolveMedia)).toEqual({
      type: 'blip',
      image: { contentType: 'image/png', data: new Uint8Array([1, 2, 3]) },
      opacity: 50000,
      stretch: true,
    });
  });

  it('omits blipFill entirely when the relationship cannot be resolved', () => {
    const node = firstNode(
      '<a:spPr xmlns:a="a" xmlns:r="r"><a:blipFill><a:blip r:embed="rIdX"/></a:blipFill></a:spPr>',
    );
    expect(parseChildFill(node, noMedia)).toBeUndefined();
  });
});
