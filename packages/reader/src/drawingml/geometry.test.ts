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

  it('parses custGeom with no pathLst as a bare kind marker', () => {
    const node = firstNode('<a:spPr xmlns:a="a"><a:custGeom/></a:spPr>');
    expect(findChildGeometry(node)).toEqual({ type: 'custom' });
  });

  it('parses a custGeom pathLst with two subpaths (a boolean-subtract rectangle-minus-star)', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:custGeom>
          <a:pathLst>
            <a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
              <a:lnTo><a:pt x="100" y="100"/></a:lnTo>
              <a:close/>
            </a:path>
            <a:path w="100" h="100">
              <a:moveTo><a:pt x="10" y="10"/></a:moveTo>
              <a:lnTo><a:pt x="20" y="10"/></a:lnTo>
              <a:close/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
      </a:spPr>`,
    );
    expect(findChildGeometry(node)).toEqual({
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
            { type: 'moveTo', point: { x: 10, y: 10 } },
            { type: 'lnTo', point: { x: 20, y: 10 } },
            { type: 'close' },
          ],
        },
      ],
    });
  });

  it('parses quadBezTo, cubicBezTo and arcTo commands', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:custGeom>
          <a:pathLst>
            <a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:quadBezTo>
                <a:pt x="10" y="0"/>
                <a:pt x="10" y="10"/>
              </a:quadBezTo>
              <a:cubicBezTo>
                <a:pt x="20" y="10"/>
                <a:pt x="20" y="20"/>
                <a:pt x="30" y="20"/>
              </a:cubicBezTo>
              <a:arcTo wR="5" hR="5" stAng="0" swAng="5400000"/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
      </a:spPr>`,
    );
    expect(findChildGeometry(node)).toEqual({
      type: 'custom',
      pathLst: [
        {
          width: 100,
          height: 100,
          commands: [
            { type: 'moveTo', point: { x: 0, y: 0 } },
            { type: 'quadBezTo', control: { x: 10, y: 0 }, point: { x: 10, y: 10 } },
            {
              type: 'cubicBezTo',
              control1: { x: 20, y: 10 },
              control2: { x: 20, y: 20 },
              point: { x: 30, y: 20 },
            },
            { type: 'arcTo', widthRadius: 5, heightRadius: 5, startAngle: 0, swingAngle: 5400000 },
          ],
        },
      ],
    });
  });

  it('drops a path containing a gdLst-guide-referenced coordinate rather than a corrupted outline', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:custGeom>
          <a:pathLst>
            <a:path w="100" h="100">
              <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
              <a:lnTo><a:pt x="csX0" y="10"/></a:lnTo>
              <a:close/>
            </a:path>
            <a:path w="50" h="50">
              <a:moveTo><a:pt x="1" y="1"/></a:moveTo>
              <a:close/>
            </a:path>
          </a:pathLst>
        </a:custGeom>
      </a:spPr>`,
    );
    expect(findChildGeometry(node)).toEqual({
      type: 'custom',
      pathLst: [
        {
          width: 50,
          height: 50,
          commands: [{ type: 'moveTo', point: { x: 1, y: 1 } }, { type: 'close' }],
        },
      ],
    });
  });

  it('falls back to a bare kind marker when every path was dropped', () => {
    const node = firstNode(
      `<a:spPr xmlns:a="a">
        <a:custGeom>
          <a:pathLst>
            <a:path w="100" h="100">
              <a:moveTo><a:pt x="csX0" y="0"/></a:moveTo>
            </a:path>
          </a:pathLst>
        </a:custGeom>
      </a:spPr>`,
    );
    expect(findChildGeometry(node)).toEqual({ type: 'custom' });
  });
});
