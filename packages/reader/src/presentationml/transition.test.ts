import { describe, expect, it } from 'vitest';

import type { MediaResolver } from '../drawingml/fill.js';
import { findAlternateContentChild } from '../xml/query.js';
import { parseXml } from '../xml/parse.js';
import { parseSlideTransition, pickTransitionNode } from './transition.js';

function firstNode(xml: string) {
  const [root] = parseXml(xml);
  return root!;
}

const NS = 'xmlns:p="p" xmlns:r="r"';
const noMedia: MediaResolver = () => undefined;

describe('parseSlideTransition', () => {
  it('returns undefined when there is no p:transition element', () => {
    expect(parseSlideTransition(undefined, noMedia)).toBeUndefined();
  });

  it('returns undefined for an empty p:transition', () => {
    const node = firstNode(`<p:transition ${NS}/>`);
    expect(parseSlideTransition(node, noMedia)).toBeUndefined();
  });

  it('parses speed and advance-on-click/advance-after attributes with no effect', () => {
    const node = firstNode(`<p:transition ${NS} spd="slow" advClick="0" advTm="3000"/>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      speed: 'slow',
      advanceOnClick: false,
      advanceAfter: 3000,
    });
  });

  it('parses an orientation effect (blinds) with its dir attribute', () => {
    const node = firstNode(`<p:transition ${NS}><p:blinds dir="vert"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'blinds', orientation: 'vert' },
    });
  });

  it('parses a parameterless effect (dissolve)', () => {
    const node = firstNode(`<p:transition ${NS}><p:dissolve/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({ effect: { kind: 'dissolve' } });
  });

  it('parses an eight-direction effect (cover) with its dir attribute', () => {
    const node = firstNode(`<p:transition ${NS}><p:cover dir="ru"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'cover', direction: 'ru' },
    });
  });

  it('parses cut with thruBlk', () => {
    const node = firstNode(`<p:transition ${NS}><p:cut thruBlk="1"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'cut', throughBlack: true },
    });
  });

  it('parses fade with thruBlk', () => {
    const node = firstNode(`<p:transition ${NS}><p:fade thruBlk="1"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'fade', throughBlack: true },
    });
  });

  it('parses a side-direction effect (wipe) with its dir attribute', () => {
    const node = firstNode(`<p:transition ${NS}><p:wipe dir="d"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'wipe', direction: 'd' },
    });
  });

  it('parses split with orient and dir', () => {
    const node = firstNode(`<p:transition ${NS}><p:split orient="vert" dir="in"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'split', orientation: 'vert', direction: 'in' },
    });
  });

  it('parses strips with its dir attribute', () => {
    const node = firstNode(`<p:transition ${NS}><p:strips dir="rd"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'strips', direction: 'rd' },
    });
  });

  it('parses wheel with a spoke count', () => {
    const node = firstNode(`<p:transition ${NS}><p:wheel spokes="8"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'wheel', spokes: 8 },
    });
  });

  it('parses zoom with a direction', () => {
    const node = firstNode(`<p:transition ${NS}><p:zoom dir="out"/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'zoom', direction: 'out' },
    });
  });

  it('parses morph with its option attribute', () => {
    const node = firstNode(
      `<p:transition ${NS} xmlns:p159="p159"><p159:morph option="byWord"/></p:transition>`,
    );
    expect(parseSlideTransition(node, noMedia)).toEqual({
      effect: { kind: 'morph', option: 'byWord' },
    });
  });

  it('parses morph with no option attribute (PowerPoint default is byObject, left absent here)', () => {
    const node = firstNode(`<p:transition ${NS} xmlns:p159="p159"><p159:morph/></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({ effect: { kind: 'morph' } });
  });

  it('parses an explicit p14:dur duration alongside spd', () => {
    const node = firstNode(
      `<p:transition ${NS} xmlns:p14="p14" spd="slow" p14:dur="2000"><p:fade/></p:transition>`,
    );
    expect(parseSlideTransition(node, noMedia)).toEqual({
      speed: 'slow',
      durationMs: 2000,
      effect: { kind: 'fade' },
    });
  });

  it('resolves a stSnd sound action via the injected resolver, with loop', () => {
    const node = firstNode(
      `<p:transition ${NS}>
        <p:fade/>
        <p:sndAc>
          <p:stSnd loop="1"><p:snd r:embed="rId7"/></p:stSnd>
        </p:sndAc>
      </p:transition>`,
    );
    const resolveMedia: MediaResolver = (id) =>
      id === 'rId7' ? { contentType: 'audio/wav', data: new Uint8Array([1, 2, 3]) } : undefined;

    expect(parseSlideTransition(node, resolveMedia)).toEqual({
      effect: { kind: 'fade' },
      sound: {
        kind: 'play',
        sound: { contentType: 'audio/wav', data: new Uint8Array([1, 2, 3]) },
        loop: true,
      },
    });
  });

  it('omits the sound action entirely when the relationship cannot be resolved', () => {
    const node = firstNode(
      `<p:transition ${NS}><p:sndAc><p:stSnd><p:snd r:embed="rIdX"/></p:stSnd></p:sndAc></p:transition>`,
    );
    expect(parseSlideTransition(node, noMedia)).toBeUndefined();
  });

  it('parses an endSnd sound action as stop', () => {
    const node = firstNode(`<p:transition ${NS}><p:sndAc><p:endSnd/></p:sndAc></p:transition>`);
    expect(parseSlideTransition(node, noMedia)).toEqual({ sound: { kind: 'stop' } });
  });
});

describe('pickTransitionNode', () => {
  it('picks the Choice branch when its effect is recognized (morph)', () => {
    const choice = firstNode(`<p:transition ${NS} xmlns:p159="p159"><p159:morph/></p:transition>`);
    const resolved = firstNode(`<p:transition ${NS}><p:fade/></p:transition>`);
    expect(pickTransitionNode({ choice, resolved })).toBe(choice);
  });

  it('falls back to the resolved branch when the Choice effect is unrecognized', () => {
    const choice = firstNode(`<p:transition ${NS} xmlns:p188="p188"><p188:ripple/></p:transition>`);
    const resolved = firstNode(`<p:transition ${NS}><p:fade/></p:transition>`);
    expect(pickTransitionNode({ choice, resolved })).toBe(resolved);
  });

  it('returns the resolved branch when there is no Choice at all', () => {
    const resolved = firstNode(`<p:transition ${NS}><p:push dir="u"/></p:transition>`);
    expect(pickTransitionNode({ resolved })).toBe(resolved);
  });

  it('returns undefined when neither branch is present', () => {
    expect(pickTransitionNode({})).toBeUndefined();
  });

  it(
    'end-to-end: a real Morph-authored p:transition (mc:AlternateContent wrapping the whole ' +
      'element, matching apps/web-demo/src/Presentation1.pptx slide 4) parses as morph, not the fade fallback',
    () => {
      const [root] = parseXml(
        `<p:sld xmlns:p="p" xmlns:mc="mc">
          <mc:AlternateContent>
            <mc:Choice xmlns:p159="p159" Requires="p159">
              <p:transition spd="slow" xmlns:p14="p14" p14:dur="2000">
                <p159:morph option="byObject"/>
              </p:transition>
            </mc:Choice>
            <mc:Fallback>
              <p:transition spd="slow"><p:fade/></p:transition>
            </mc:Fallback>
          </mc:AlternateContent>
        </p:sld>`,
      );
      const lookup = findAlternateContentChild(root!, 'transition');
      const transition = parseSlideTransition(pickTransitionNode(lookup), noMedia);
      expect(transition).toEqual({
        speed: 'slow',
        durationMs: 2000,
        effect: { kind: 'morph', option: 'byObject' },
      });
    },
  );
});
