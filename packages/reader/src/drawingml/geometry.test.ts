import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/parse.js';
import { findChildGeometry, parseTransform } from './geometry.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

describe('drawingml/geometry', () => {
  it('parses a full xfrm including rotation, flips and child coordinate space', () => {
    const node = firstNode(
      `<a:xfrm xmlns:a="a" rot="5400000" flipH="1" flipV="0">
        <a:off x="100" y="200"/>
        <a:ext cx="300" cy="400"/>
        <a:chOff x="0" y="0"/>
        <a:chExt cx="300" cy="400"/>
      </a:xfrm>`,
    );
    expect(parseTransform(node)).toEqual({
      offset: { x: 100, y: 200 },
      extents: { width: 300, height: 400 },
      rotation: 5400000,
      flipHorizontal: true,
      flipVertical: false,
      childOffset: { x: 0, y: 0 },
      childExtents: { width: 300, height: 400 },
    });
  });

  it('returns undefined when off/ext are missing', () => {
    expect(parseTransform(firstNode('<a:xfrm xmlns:a="a"/>'))).toBeUndefined();
    expect(parseTransform(undefined)).toBeUndefined();
  });

  it('parses prstGeom with literal-valued adjustment guides, skipping formula guides', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:prstGeom prst="roundRect">
          <a:avLst>
            <a:gd name="adj" fmla="val 16667"/>
            <a:gd name="derived" fmla="*/ adj 2 1"/>
          </a:avLst>
        </a:prstGeom>
      </a:spPr>`,
    );
    expect(findChildGeometry(node)).toEqual({
      type: 'preset',
      preset: 'roundRect',
      adjustValues: [{ name: 'adj', value: 16667 }],
    });
  });

  it('parses custGeom as a kind-only marker', () => {
    const node = firstNode('<a:spPr xmlns:a="a"><a:custGeom/></a:spPr>');
    expect(findChildGeometry(node)).toEqual({ type: 'custom' });
  });
});
