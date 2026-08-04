// @vitest-environment happy-dom
import type { Presentation, Shape, SlideLayout, SlideTransition } from '@pptx2html/presentation';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  defineScrollPresentationElement,
  PptxScrollPresentationElement,
} from './scroll-presentation-element.js';

const LAYOUT: SlideLayout = {
  commonSlideData: { shapeTree: [] },
  master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
  type: 'blank',
};

// happy-dom has no Web Animations API at all (see presentation-element.test.ts's own note on
// this) — mocked here too, but simpler: scroll-mode playback never awaits `.finished` (every
// Animation is paused and scrubbed via `currentTime`, not fired-and-forgotten), so this mock only
// needs enough surface for that: `.pause()`, `.cancel()`, and a plain settable `currentTime`.
class FakeAnimation {
  playState: AnimationPlayState = 'running';
  currentTime: number | null = 0;
  pause(): void {
    this.playState = 'paused';
  }
  cancel(): void {
    this.playState = 'idle';
  }
}

interface RecordedAnimation {
  readonly target: HTMLElement;
  readonly keyframes: Keyframe[];
  readonly options: KeyframeAnimationOptions;
  readonly animation: FakeAnimation;
}

let recordedAnimations: RecordedAnimation[] = [];

/** The most recently created animation for `target` (mirrors presentation-element.test.ts's own
 *  helper of the same name/purpose). */
function latestAnimation(target: HTMLElement): RecordedAnimation | undefined {
  for (let i = recordedAnimations.length - 1; i >= 0; i--) {
    if (recordedAnimations[i]!.target === target) return recordedAnimations[i];
  }
  return undefined;
}

beforeEach(() => {
  recordedAnimations = [];
  HTMLElement.prototype.animate = vi.fn(function (
    this: HTMLElement,
    keyframes: Keyframe[] | PropertyIndexedKeyframes | null,
    options?: number | KeyframeAnimationOptions,
  ) {
    const animation = new FakeAnimation();
    recordedAnimations.push({
      target: this,
      keyframes: keyframes as Keyframe[],
      options: options as KeyframeAnimationOptions,
      animation,
    });
    return animation as unknown as Animation;
  }) as typeof HTMLElement.prototype.animate;
});

afterEach(() => {
  delete (HTMLElement.prototype as { animate?: unknown }).animate;
});

function buildPresentation(transitions: (SlideTransition | undefined)[]): Presentation {
  return {
    slideSize: { width: 12192000, height: 6858000 },
    slideMasters: [],
    slides: transitions.map((transition) => ({
      commonSlideData: { shapeTree: [] },
      layout: LAYOUT,
      ...(transition ? { transition } : {}),
    })),
    notesSlides: [],
  };
}

function slideEls(el: PptxScrollPresentationElement): HTMLElement[] {
  return [...(el.shadowRoot?.querySelectorAll<HTMLElement>('.pptx-scroll-viewport > *') ?? [])];
}

function newElement(): PptxScrollPresentationElement {
  return document.createElement('pptx-scroll-presentation') as PptxScrollPresentationElement;
}

describe('PptxScrollPresentationElement', () => {
  it('registers as <pptx-scroll-presentation> and renders one slide per slide into the viewport', () => {
    defineScrollPresentationElement();
    const el = newElement();
    el.render(buildPresentation([undefined, undefined, undefined]));

    expect(el.shadowRoot).not.toBeNull();
    expect(slideEls(el)).toHaveLength(3);
  });

  it('starts at ms 0 with only the first slide visible', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, undefined]));

    const [first, second] = slideEls(el);
    expect(first?.style.display).toBe('block');
    expect(second?.style.display).toBe('none');
  });

  it("seekTo during a static slide's content phase keeps only that slide visible", () => {
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'fade' } }, undefined]));
    // Segment 0 content: [0, 1200); segment 1 transition: [1200, 1600); segment 1 content:
    // [1600, 2800); segment 2 transition: [2800, 3200); segment 2 content: [3200, 4400).
    const [slide0, slide1, slide2] = slideEls(el);

    el.seekTo(2000); // slide 1's own content phase
    expect(slide0?.style.display).toBe('none');
    expect(slide1?.style.display).toBe('block');
    expect(slide2?.style.display).toBe('none');
  });

  it('seekTo mid-push-transition shows both participating slides and sets each Animation.currentTime', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'push', direction: 'r' } }]));
    const [slide0, slide1] = slideEls(el);

    el.seekTo(1400); // 200ms into slide 1's transition (starts at 1200, "fast" = 400ms)
    expect(slide0?.style.display).toBe('block');
    expect(slide1?.style.display).toBe('block');
    expect(latestAnimation(slide0!)?.animation.currentTime).toBe(200);
    expect(latestAnimation(slide1!)?.animation.currentTime).toBe(200);
    expect(latestAnimation(slide0!)?.keyframes).toEqual([
      { transform: 'translate(0, 0)' },
      { transform: 'translate(100%, 0)' },
    ]);
  });

  it('scrubbing backward through the same transition is just a smaller currentTime, no re-authored animation', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'push', direction: 'l' } }]));
    const [outgoing, incoming] = slideEls(el);

    el.seekTo(1400); // 200ms into the transition, forward
    const outgoingAnim = latestAnimation(outgoing!)?.animation;
    const incomingAnim = latestAnimation(incoming!)?.animation;

    el.seekTo(1250); // scrub back to 50ms in — same Animation instances, smaller currentTime
    expect(latestAnimation(outgoing!)?.animation).toBe(outgoingAnim);
    expect(outgoingAnim?.currentTime).toBe(50);
    expect(incomingAnim?.currentTime).toBe(50);
  });

  it('every Animation is paused immediately after creation, never auto-playing', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'fade' } }]));
    expect(recordedAnimations.length).toBeGreaterThan(0);
    for (const { animation } of recordedAnimations) {
      expect(animation.playState).toBe('paused');
    }
  });

  it('clamps seekTo to [0, totalDurationMs]', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, undefined]));
    // slide 0 content (1200ms dwell) + synthetic push-up transition ("fast" = 400ms) + slide 1
    // content (1200ms dwell) — even an unauthored transition still gets a segment, see above.
    expect(el.totalDurationMs).toBe(2800);

    el.seekTo(-500);
    expect(slideEls(el)[0]?.style.display).toBe('block');

    el.seekTo(999_999);
    expect(slideEls(el)[1]?.style.display).toBe('block');
  });

  describe('fade builds within a slide', () => {
    function shape(id: number): Shape {
      return { kind: 'shape', nonVisual: { id, name: `Shape ${id}` }, properties: {} };
    }

    function buildWithFade(fadeShapeId: number): Presentation {
      return {
        slideSize: { width: 12192000, height: 6858000 },
        slideMasters: [],
        slides: [
          {
            commonSlideData: { shapeTree: [shape(fadeShapeId)] },
            layout: LAYOUT,
            timing: {
              timeNodeTree: {
                kind: 'animEffect',
                common: { id: 0, duration: 500, startConditions: [{ delay: 200 }] },
                target: { kind: 'shape', shapeId: fadeShapeId },
                filter: 'fade',
                transition: 'in',
              },
            },
          },
        ],
        notesSlides: [],
      };
    }

    it("sets the fade's currentTime relative to its own slide's content start, clamped to its own domain", () => {
      const el = newElement();
      el.render(buildWithFade(42));
      const target = el.shadowRoot?.querySelector<HTMLElement>('[data-pptx-shape-id="42"]');
      expect(target).not.toBeNull();

      el.seekTo(0);
      expect(latestAnimation(target!)?.animation.currentTime).toBe(0);

      el.seekTo(400); // 200ms delay + halfway through the 500ms fade
      expect(latestAnimation(target!)?.animation.currentTime).toBe(400);

      el.seekTo(2000); // past the fade's own [0, 700] domain, but still within the slide's dwell
      expect(latestAnimation(target!)?.animation.currentTime).toBe(700);
    });
  });

  describe('morph transitions', () => {
    function morphShape(id: number, name: string, x = 0, y = 0): Shape {
      return {
        kind: 'shape',
        nonVisual: { id, name },
        properties: {
          transform: { offset: { x, y }, extents: { width: 1000000, height: 1000000 } },
        },
      };
    }

    function buildMorphPresentation(
      slides: readonly { shapes?: readonly Shape[]; transition?: SlideTransition }[],
    ): Presentation {
      return {
        slideSize: { width: 12192000, height: 6858000 },
        slideMasters: [],
        slides: slides.map(({ shapes, transition }) => ({
          commonSlideData: { shapeTree: shapes ?? [] },
          layout: LAYOUT,
          ...(transition ? { transition } : {}),
        })),
        notesSlides: [],
      };
    }

    it('hides the departing copy for the transition window and restores it once scrubbed back into its own slide', () => {
      const presentation = buildMorphPresentation([
        { shapes: [morphShape(1, 'Title 1', 0, 0)] },
        {
          shapes: [morphShape(1, 'Title 1', 1_000_000, 1_000_000)],
          transition: { effect: { kind: 'morph' } },
        },
      ]);
      const el = newElement();
      el.render(presentation);
      const [slide0, slide1] = slideEls(el);
      const departing = slide0!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;
      const arriving = slide1!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;

      el.seekTo(1400); // mid-transition (starts at 1200, "fast" = 400ms)
      expect(departing.style.opacity).toBe('0');
      expect(latestAnimation(arriving)?.animation.currentTime).toBe(200);

      el.seekTo(500); // scrub back into slide 0's own content phase
      expect(departing.style.opacity).toBe('');
      expect(slide0?.style.display).toBe('block');
      expect(slide1?.style.display).toBe('none');
    });

    it('falls back to a plain crossfade and still scrubs it for a low-confidence match', () => {
      const presentation = buildMorphPresentation([
        { shapes: [morphShape(1, 'Alpha')] },
        { shapes: [morphShape(2, 'Beta')], transition: { effect: { kind: 'morph' } } },
      ]);
      const el = newElement();
      el.render(presentation);
      const [slide0, slide1] = slideEls(el);

      el.seekTo(1400);
      expect(latestAnimation(slide0!)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
      expect(latestAnimation(slide1!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
    });
  });

  describe('pixelsPerSecond', () => {
    it('sizes the scroll track from totalDurationMs * pixelsPerSecond / 1000', () => {
      const el = newElement();
      el.render(buildPresentation([undefined])); // 1200ms dwell, default 600px/s
      const spacer = el.shadowRoot?.querySelector<HTMLElement>('.pptx-scroll-spacer');
      expect(spacer?.style.height).toBe('720px'); // 1200ms * 0.6px/ms

      el.pixelsPerSecond = 1200;
      expect(spacer?.style.height).toBe('1440px');
    });
  });

  describe('scroll-driven playback', () => {
    let rafCallback: FrameRequestCallback | undefined;

    beforeEach(() => {
      rafCallback = undefined;
      vi.stubGlobal(
        'requestAnimationFrame',
        vi.fn((cb: FrameRequestCallback) => {
          rafCallback = cb;
          return 1;
        }),
      );
      vi.stubGlobal('cancelAnimationFrame', vi.fn());
    });

    afterEach(() => {
      vi.unstubAllGlobals();
    });

    it('coalesces multiple scroll events within one frame into a single seekTo call', () => {
      const el = newElement();
      el.render(buildPresentation([undefined, undefined]));
      document.body.appendChild(el);
      const seekToSpy = vi.spyOn(el, 'seekTo');
      const track = el.shadowRoot!.querySelector<HTMLElement>('.pptx-scroll-track')!;

      Object.defineProperty(track, 'scrollTop', { value: 100, configurable: true });
      track.dispatchEvent(new Event('scroll'));
      Object.defineProperty(track, 'scrollTop', { value: 250, configurable: true });
      track.dispatchEvent(new Event('scroll'));

      expect(seekToSpy).not.toHaveBeenCalled(); // nothing runs synchronously off a scroll event

      rafCallback?.(0);
      expect(seekToSpy).toHaveBeenCalledTimes(1);
      expect(seekToSpy).toHaveBeenCalledWith(250 / (600 / 1000)); // latest scrollTop only, in ms

      el.remove();
    });
  });
});
