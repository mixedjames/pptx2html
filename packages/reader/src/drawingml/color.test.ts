import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import { parseChildColor } from './color.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

describe('drawingml/color', () => {
  it('parses a solidFill srgbClr', () => {
    const node = firstNode('<a:solidFill xmlns:a="a"><a:srgbClr val="FF0000"/></a:solidFill>');
    expect(parseChildColor(node)).toEqual({ type: 'srgb', value: 'FF0000' });
  });

  it('parses a schemeClr with transforms', () => {
    const node = firstNode(
      `<a:solidFill xmlns:a="a">
        <a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>
      </a:solidFill>`,
    );
    expect(parseChildColor(node)).toEqual({
      type: 'scheme',
      value: 'accent1',
      transforms: { lumMod: 60000, lumOff: 40000 },
    });
  });

  it('rejects an unrecognized scheme colour name', () => {
    const node = firstNode('<a:solidFill xmlns:a="a"><a:schemeClr val="bogus"/></a:solidFill>');
    expect(parseChildColor(node)).toBeUndefined();
  });

  it('parses a sysClr with its lastClr fallback', () => {
    const node = firstNode(
      '<a:solidFill xmlns:a="a"><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>',
    );
    expect(parseChildColor(node)).toEqual({
      type: 'system',
      value: 'windowText',
      lastColor: '000000',
    });
  });

  it('parses an hslClr', () => {
    const node = firstNode(
      '<a:solidFill xmlns:a="a"><a:hslClr hue="14400000" sat="50%" lum="50%"/></a:solidFill>',
    );
    expect(parseChildColor(node)).toEqual({
      type: 'hsl',
      hue: 14400000,
      saturation: 50000,
      luminance: 50000,
    });
  });
});
