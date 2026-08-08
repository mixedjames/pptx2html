import { describe, expect, it } from 'vitest';

import type { MediaResolver } from '../drawingml/fill.js';
import { parseXml } from '../xml/parse.js';
import { parseCommonSlideData } from './common-slide-data.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

const NS = 'xmlns:p="p" xmlns:a="a"';
const noMedia: MediaResolver = () => undefined;

describe('parseCommonSlideData background parsing', () => {
  it('has neither background nor backgroundRef when p:bg is absent', () => {
    const node = firstNode(`<p:cSld ${NS}><p:spTree/></p:cSld>`);
    const result = parseCommonSlideData(node, noMedia);
    expect(result.background).toBeUndefined();
    expect(result.backgroundRef).toBeUndefined();
  });

  it('parses a literal p:bgPr fill into background', () => {
    const node = firstNode(
      `<p:cSld ${NS}><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:bgPr></p:bg><p:spTree/></p:cSld>`,
    );
    const result = parseCommonSlideData(node, noMedia);
    expect(result.background).toEqual({
      fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
    });
    expect(result.backgroundRef).toBeUndefined();
  });

  it('parses a p:bgRef style-matrix reference into backgroundRef — the real gap this covers', () => {
    // Regression test: p:bgRef used to be silently dropped ("unmodeled for the skeleton"), so any
    // slide/layout/master relying on the theme's own default background (PowerPoint's own common
    // case, not an edge case) parsed as having no background at all — see this package's own
    // CLAUDE.md and packages/presentation/CLAUDE.md for the real deck this was caught against.
    const node = firstNode(
      `<p:cSld ${NS}><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree/></p:cSld>`,
    );
    const result = parseCommonSlideData(node, noMedia);
    expect(result.background).toBeUndefined();
    expect(result.backgroundRef).toEqual({
      index: 1001,
      color: { type: 'scheme', value: 'bg1' },
    });
  });

  it('prefers bgPr over bgRef if a real file somehow had both (schema disallows it, but be defensive)', () => {
    const node = firstNode(
      `<p:cSld ${NS}><p:bg><p:bgPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:bgPr><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree/></p:cSld>`,
    );
    const result = parseCommonSlideData(node, noMedia);
    expect(result.background).toEqual({
      fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
    });
    expect(result.backgroundRef).toBeUndefined();
  });

  it('has neither field for an empty p:bg (no bgPr or bgRef child)', () => {
    const node = firstNode(`<p:cSld ${NS}><p:bg/><p:spTree/></p:cSld>`);
    const result = parseCommonSlideData(node, noMedia);
    expect(result.background).toBeUndefined();
    expect(result.backgroundRef).toBeUndefined();
  });
});
