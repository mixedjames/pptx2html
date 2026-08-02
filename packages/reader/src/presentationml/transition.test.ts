import { describe, expect, it } from 'vitest';

import type { MediaResolver } from '../drawingml/fill.js';
import { parseXml } from '../xml/parse.js';
import { parseSlideTransition } from './transition.js';

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
