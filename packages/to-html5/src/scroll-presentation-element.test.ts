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
    // No build animations anywhere in this deck, so every content phase is zero-width (see
    // scroll-timeline.ts's DEFAULT_MIN_DWELL_MS) — segment 0 content: [0, 0); segment 1 transition:
    // [0, 400); segment 1 content: [400, 400); segment 2 transition: [400, 800); segment 2 content:
    // [800, 800).
    const [slide0, slide1, slide2] = slideEls(el);

    el.seekTo(400); // slide 1's own (zero-width) content phase
    expect(slide0?.style.display).toBe('none');
    expect(slide1?.style.display).toBe('block');
    expect(slide2?.style.display).toBe('none');
  });

  it('seekTo mid-push-transition shows both participating slides and sets each Animation.currentTime', () => {
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'push', direction: 'r' } }]));
    const [slide0, slide1] = slideEls(el);

    el.seekTo(200); // 200ms into slide 1's transition (starts at 0, "fast" = 400ms)
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

    el.seekTo(200); // 200ms into the transition, forward
    const outgoingAnim = latestAnimation(outgoing!)?.animation;
    const incomingAnim = latestAnimation(incoming!)?.animation;

    el.seekTo(50); // scrub back to 50ms in — same Animation instances, smaller currentTime
    expect(latestAnimation(outgoing!)?.animation).toBe(outgoingAnim);
    expect(outgoingAnim?.currentTime).toBe(50);
    expect(incomingAnim?.currentTime).toBe(50);
  });

  it("cancels the conflicting sibling animation so a middle slide's own arrival is not silently shadowed by the next transition's departure off the same element+property", () => {
    // Regression test for a real bug found via portrait-slides.pptx (3 slides, push throughout):
    // slide 1 is both segment 1's `incoming` target and segment 2's `outgoing` source, so it ends
    // up with two permanently fill:'both' Animations targeting its own `transform` — WAAPI's
    // default composite lets only the more-recently-created one win, which (since segments build
    // in order) is always segment 2's `outgoing`, silently shadowing segment 1's own arrival for
    // the deck's entire lifetime, not just while segment 2 is nominally active. See
    // #claimTransitionAnimation's own doc comment for the full mechanism.
    const el = newElement();
    el.render(
      buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'r' } },
        { effect: { kind: 'push', direction: 'l' } },
      ]),
    );
    const [, slide1] = slideEls(el);
    const slide1Anims = recordedAnimations.filter((r) => r.target === slide1);
    expect(slide1Anims).toHaveLength(2); // segment 1's incoming + segment 2's outgoing

    const incoming = slide1Anims.find(
      (r) =>
        JSON.stringify(r.keyframes) ===
        JSON.stringify([{ transform: 'translate(-100%, 0)' }, { transform: 'translate(0, 0)' }]),
    )?.animation;
    const outgoing = slide1Anims.find(
      (r) =>
        JSON.stringify(r.keyframes) ===
        JSON.stringify([{ transform: 'translate(0, 0)' }, { transform: 'translate(-100%, 0)' }]),
    )?.animation;
    expect(incoming).toBeDefined();
    expect(outgoing).toBeDefined();

    el.seekTo(200); // 200ms into segment 1's transition — slide 1 is arriving
    expect(incoming?.currentTime).toBe(200);
    // The conflicting sibling (segment 2's own departure, not due to start until ms 400) must be
    // cancelled — left alive, it would permanently paint over slide 1 with its own first keyframe
    // (identity), which is exactly what an unanimated "hard cut" looks like.
    expect(outgoing?.playState).toBe('idle');
  });

  it("settles a slide's own arrival to its fully-arrived frame once scrubbed into its content phase, even when the transition window itself was never visited", () => {
    // A fast scroll (a big wheel/trackpad jump, or a direct seekTo) can skip straight from before a
    // transition's window to after it, without ever calling seekTo with a ms inside that window —
    // #scrubTransition's own currentTime writes would then never run for it at all, leaving the
    // Animation exactly wherever it started (currentTime 0, i.e. its *first* keyframe) instead of
    // its settled, fully-arrived one.
    const el = newElement();
    el.render(buildPresentation([undefined, { effect: { kind: 'push', direction: 'r' } }]));
    const [, slide1] = slideEls(el);
    const incoming = latestAnimation(slide1!)?.animation; // slide1's only Animation here
    expect(incoming?.currentTime).toBe(0);

    el.seekTo(400); // straight to segment 1's own (zero-width) content phase — its transition
    // window ([0, 400)) is never visited by any seekTo call in between.
    expect(incoming?.currentTime).toBe(400); // settled to the full "fast" (400ms) duration
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
    // Zero-width content phases (no builds anywhere) + a synthetic push-up transition ("fast" =
    // 400ms) — even an unauthored transition still gets a segment, see above.
    expect(el.totalDurationMs).toBe(400);

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

      el.seekTo(2000); // past the fade's own [0, 700] domain — clamped to totalDurationMs (700),
      // itself exactly the fade's own end (its content phase has no floor beyond what it needs)
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

      el.seekTo(200); // mid-transition (starts at 0, "fast" = 400ms)
      expect(departing.style.opacity).toBe('0');
      expect(latestAnimation(arriving)?.animation.currentTime).toBe(200);

      el.seekTo(0); // scrub back into slide 0's own (zero-width) content phase
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

      el.seekTo(200);
      expect(latestAnimation(slide0!)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
      expect(latestAnimation(slide1!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
    });
  });

  describe('pixelsPerSecond', () => {
    it('sizes the scroll track from totalDurationMs * pixelsPerSecond / 1000', () => {
      const el = newElement();
      // A single slide with no builds has zero inherent duration (see scroll-timeline.ts's
      // DEFAULT_MIN_DWELL_MS) — a second slide with an authored transition gives this deck a real,
      // non-zero duration to size the track from: 400ms ("fast"), default 600px/s.
      el.render(buildPresentation([undefined, { effect: { kind: 'fade' } }]));
      const spacer = el.shadowRoot?.querySelector<HTMLElement>('.pptx-scroll-spacer');
      expect(spacer?.style.height).toBe('240px'); // 400ms * 0.6px/ms

      el.pixelsPerSecond = 1200;
      expect(spacer?.style.height).toBe('480px');
    });

    it("pads the spacer by the track's own clientHeight, since CSS's scrollable range is spacer.scrollHeight minus that, not the spacer's height alone", () => {
      // Regression test for a real bug found via portrait-slides.pptx: without this padding, the
      // deck's own final millisecond of scroll-timeline was never reachable by actually scrolling
      // to the bottom of the page — short by exactly one screen's worth — so the last slide's own
      // transition-in was always caught mid-flight, looking cropped, however far the user scrolled.
      const el = newElement();
      const track = el.shadowRoot!.querySelector<HTMLElement>('.pptx-scroll-track')!;
      Object.defineProperty(track, 'clientHeight', { value: 500, configurable: true });

      el.render(buildPresentation([undefined, { effect: { kind: 'fade' } }])); // 400ms content, default 600px/s -> 240px content
      const spacer = el.shadowRoot?.querySelector<HTMLElement>('.pptx-scroll-spacer');
      expect(spacer?.style.height).toBe('740px'); // 240px content + 500px track height
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
