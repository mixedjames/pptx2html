// @vitest-environment happy-dom
import type { Presentation, Shape, SlideLayout, SlideTransition } from '@pptx2html/presentation';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { definePresentationElement, PptxPresentationElement } from './presentation-element.js';

const LAYOUT: SlideLayout = {
  commonSlideData: { shapeTree: [] },
  master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
  type: 'blank',
};

function activeSlides(el: PptxPresentationElement): Element[] {
  return [...(el.shadowRoot?.querySelectorAll('.pptx-slide--active') ?? [])];
}

describe('PptxPresentationElement', () => {
  it('registers as <pptx-presentation> and renders one .pptx-slide per slide into its shadow root', () => {
    definePresentationElement();

    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    expect(el.shadowRoot).not.toBeNull();
    expect(el.shadowRoot?.querySelectorAll('.pptx-slide')).toHaveLength(3);
  });

  it('shows only the current slide, starting at the first', () => {
    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    expect(el.currentSlideIndex).toBe(0);
    expect(el.slideCount).toBe(3);
    expect(activeSlides(el)).toHaveLength(1);
    expect(el.shadowRoot?.querySelectorAll('.pptx-slide')[0]).toBe(activeSlides(el)[0]);
  });

  it('next()/previous() advance and retreat the active slide, clamped to the slide range', () => {
    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    el.previous(); // no-op: already at the first slide
    expect(el.currentSlideIndex).toBe(0);

    el.next();
    expect(el.currentSlideIndex).toBe(1);
    expect(el.shadowRoot?.querySelectorAll('.pptx-slide')[1]).toBe(activeSlides(el)[0]);

    el.next(); // no-op: already at the last slide
    expect(el.currentSlideIndex).toBe(1);

    el.previous();
    expect(el.currentSlideIndex).toBe(0);
  });

  it('goToSlide() jumps directly to an index, clamped to the valid range', () => {
    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    el.goToSlide(2);
    expect(el.currentSlideIndex).toBe(2);

    el.goToSlide(99);
    expect(el.currentSlideIndex).toBe(2);

    el.goToSlide(-5);
    expect(el.currentSlideIndex).toBe(0);
  });

  it('clicking the element advances to the next slide', () => {
    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    el.click();

    expect(el.currentSlideIndex).toBe(1);
  });

  it('ArrowRight/space advance and ArrowLeft/backspace retreat on keydown', () => {
    const presentation: Presentation = {
      slideSize: { width: 12192000, height: 6858000 },
      slideMasters: [],
      slides: [
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
        { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
      ],
      notesSlides: [],
    };

    const el = document.createElement('pptx-presentation') as PptxPresentationElement;
    el.render(presentation);

    el.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }));
    expect(el.currentSlideIndex).toBe(1);

    el.dispatchEvent(new KeyboardEvent('keydown', { key: ' ' }));
    expect(el.currentSlideIndex).toBe(2);

    el.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft' }));
    expect(el.currentSlideIndex).toBe(1);

    el.dispatchEvent(new KeyboardEvent('keydown', { key: 'End' }));
    expect(el.currentSlideIndex).toBe(2);

    el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home' }));
    expect(el.currentSlideIndex).toBe(0);
  });

  describe('push/fade transitions', () => {
    // happy-dom (this file's DOM environment) has no Web Animations API implementation at all —
    // `Element.prototype.animate` doesn't exist — so it's mocked here. `FakeAnimation` mirrors just
    // enough of the real `Animation` interface for production code's needs (`.finished`, `.cancel()`,
    // `.playState`); its `.finished` promise resolves via a `setTimeout` matching the requested
    // duration, which is what makes it driveable by `vi.useFakeTimers()`. `onfinish` is deliberately
    // not implemented — production code only ever reads `.finished`.
    class FakeAnimation {
      playState: AnimationPlayState = 'running';
      readonly finished: Promise<Animation>;
      #reject!: (reason: unknown) => void;
      #timeoutId!: ReturnType<typeof setTimeout>;

      constructor(durationMs: number) {
        this.finished = new Promise((resolve, reject) => {
          this.#reject = reject;
          this.#timeoutId = setTimeout(() => {
            this.playState = 'finished';
            resolve(this as unknown as Animation);
          }, durationMs);
        });
      }

      cancel(): void {
        clearTimeout(this.#timeoutId);
        this.playState = 'idle';
        this.#reject(new DOMException('The animation was canceled', 'AbortError'));
      }
    }

    interface RecordedAnimation {
      readonly target: HTMLElement;
      readonly keyframes: Keyframe[];
      readonly options: KeyframeAnimationOptions;
      readonly animation: FakeAnimation;
    }

    let recordedAnimations: RecordedAnimation[] = [];

    /** The most recently queued animation for `target`, regardless of how many earlier
     *  transitions already ran in the same test. */
    function latestAnimation(target: HTMLElement): RecordedAnimation | undefined {
      for (let i = recordedAnimations.length - 1; i >= 0; i--) {
        if (recordedAnimations[i]!.target === target) return recordedAnimations[i];
      }
      return undefined;
    }

    beforeEach(() => {
      vi.useFakeTimers();
      recordedAnimations = [];
      HTMLElement.prototype.animate = vi.fn(function (
        this: HTMLElement,
        keyframes: Keyframe[] | PropertyIndexedKeyframes | null,
        options?: number | KeyframeAnimationOptions,
      ) {
        const durationMs =
          typeof options === 'number' ? options : ((options?.duration as number) ?? 0);
        const animation = new FakeAnimation(durationMs);
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
      vi.useRealTimers();
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

    function slideEls(el: PptxPresentationElement): HTMLElement[] {
      return [...(el.shadowRoot?.querySelectorAll<HTMLElement>('.pptx-slide') ?? [])];
    }

    it('push forward: outgoing exits left, incoming enters from the right', async () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'l' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(el.currentSlideIndex).toBe(1);
      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(true);
      expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(true);
      expect(outgoing?.classList.contains('pptx-slide--active')).toBe(true);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(outgoing?.style.position).toBe('absolute');
      expect(incoming?.style.position).toBe('absolute');
      expect(latestAnimation(outgoing!)?.keyframes).toEqual([
        { transform: 'translate(0, 0)' },
        { transform: 'translate(-100%, 0)' },
      ]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([
        { transform: 'translate(100%, 0)' },
        { transform: 'translate(0, 0)' },
      ]);
      expect(latestAnimation(outgoing!)?.options).toMatchObject({ duration: 400 });

      await vi.advanceTimersByTimeAsync(400);

      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(outgoing?.classList.contains('pptx-slide--active')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(outgoing?.style.position).toBe('relative');
      expect(incoming?.style.position).toBe('relative');
    });

    it('push direction defaults to "l" when the effect omits it', () => {
      const presentation = buildPresentation([undefined, { effect: { kind: 'push' } }]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(latestAnimation(outgoing!)?.keyframes).toEqual([
        { transform: 'translate(0, 0)' },
        { transform: 'translate(-100%, 0)' },
      ]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([
        { transform: 'translate(100%, 0)' },
        { transform: 'translate(0, 0)' },
      ]);
    });

    it('push on the vertical axis (direction "u")', () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'u' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(latestAnimation(outgoing!)?.keyframes).toEqual([
        { transform: 'translate(0, 0)' },
        { transform: 'translate(0, -100%)' },
      ]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([
        { transform: 'translate(0, 100%)' },
        { transform: 'translate(0, 0)' },
      ]);
    });

    it("push reversed: backward navigation undoes the outgoing slide's own transition", async () => {
      // Slide 1's own transition (push-left) is what brought it into view; slide 0 has none of its
      // own. Backward navigation must replay slide 1's push-left in reverse (push-right) to undo
      // it — not consult slide 0's (nonexistent) transition, which would instant-swap instead.
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'l' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next(); // forward onto slide 1: plays push-left, as authored
      expect(el.currentSlideIndex).toBe(1);
      await vi.advanceTimersByTimeAsync(400);

      el.previous(); // backward into slide 0: undoes slide 1's push-left => push-right
      const [incoming, outgoing] = slideEls(el);

      expect(el.currentSlideIndex).toBe(0);
      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(true);
      expect(latestAnimation(outgoing!)?.keyframes).toEqual([
        { transform: 'translate(0, 0)' },
        { transform: 'translate(100%, 0)' },
      ]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([
        { transform: 'translate(-100%, 0)' },
        { transform: 'translate(0, 0)' },
      ]);

      await vi.advanceTimersByTimeAsync(400);
      expect(el.currentSlideIndex).toBe(0);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(outgoing?.classList.contains('pptx-slide--active')).toBe(false);
    });

    it("backward navigation uses the outgoing slide's own effect kind, not the destination's", async () => {
      // Slide 1 authors push-left (its own arrival transition); slide 2 authors fade (its own).
      // Forward 1->2 must play fade (slide 2's own). Backward 2->1 must undo slide 2's *fade* —
      // not slide 1's push — even though slide 1 has a transition of its own too. This is a
      // regression test for a bug where backward navigation looked up the destination's
      // transition instead of the outgoing slide's, playing the wrong effect entirely.
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'l' } },
        { effect: { kind: 'fade' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.goToSlide(1);
      await vi.advanceTimersByTimeAsync(400);
      el.goToSlide(2); // forward onto slide 2: plays fade, as authored
      const [, slide1, slide2] = slideEls(el);
      expect(latestAnimation(slide2!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
      await vi.advanceTimersByTimeAsync(400);

      el.previous(); // backward into slide 1: must undo slide 2's fade, not play a push

      expect(el.currentSlideIndex).toBe(1);
      expect(latestAnimation(slide2!)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
      expect(latestAnimation(slide1!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
    });

    it('fade forward: outgoing fades out, incoming fades in', async () => {
      const presentation = buildPresentation([undefined, { effect: { kind: 'fade' } }]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(latestAnimation(outgoing!)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);

      await vi.advanceTimersByTimeAsync(400);

      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
    });

    it('fade with throughBlack renders identically to a plain fade this round', () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'fade', throughBlack: true } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(latestAnimation(outgoing!)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
      expect(latestAnimation(incoming!)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
    });

    it('no transition: instant swap unchanged, no animation started', async () => {
      const presentation = buildPresentation([undefined, undefined]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(el.currentSlideIndex).toBe(1);
      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(outgoing?.classList.contains('pptx-slide--active')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(outgoing?.style.position).toBe('relative');
      expect(recordedAnimations).toHaveLength(0);

      await vi.advanceTimersByTimeAsync(10_000); // nothing was scheduled — advancing changes nothing
      expect(el.currentSlideIndex).toBe(1);
    });

    it('an unsupported effect kind (e.g. wipe) falls back to the instant swap the same way', () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'wipe', direction: 'l' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [outgoing, incoming] = slideEls(el);

      expect(outgoing?.classList.contains('pptx-slide--transitioning')).toBe(false);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(recordedAnimations).toHaveLength(0);
    });

    it('reports an unsupported effect kind to the unsupported-features log, keyed to that slide', () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'wipe', direction: 'l' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      expect(unsupportedFeatures.all).toEqual([
        {
          code: 'transition-effect-unmodeled',
          message:
            'Slide transition effect "wipe" is not animated; falling back to an instant swap.',
          slideIndex: 1,
        },
      ]);
      expect(unsupportedFeatures.bySlide.get(1)).toHaveLength(1);
    });

    it('does not report push/fade transitions, or a slide with no transition at all', () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'push', direction: 'l' } },
        { effect: { kind: 'fade' } },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      expect(unsupportedFeatures.all).toHaveLength(0);
    });

    it('reports nothing for an empty Slide.timing (no timeNodeTree or buildList)', () => {
      const presentation: Presentation = {
        slideSize: { width: 12192000, height: 6858000 },
        slideMasters: [],
        slides: [
          { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
          { commonSlideData: { shapeTree: [] }, layout: LAYOUT, timing: {} },
        ],
        notesSlides: [],
      };
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      expect(unsupportedFeatures.all).toEqual([]);
    });

    it('reports an unplayed animation behavior kind, attributed to the right slide', () => {
      const presentation: Presentation = {
        slideSize: { width: 12192000, height: 6858000 },
        slideMasters: [],
        slides: [
          { commonSlideData: { shapeTree: [] }, layout: LAYOUT },
          {
            commonSlideData: { shapeTree: [] },
            layout: LAYOUT,
            timing: { timeNodeTree: { kind: 'animClr', common: { id: 0 } } },
          },
        ],
        notesSlides: [],
      };
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      expect(unsupportedFeatures.all).toEqual([
        {
          code: 'animation-behavior-unmodeled',
          message: 'Animation behavior kind(s) animClr are not played.',
          slideIndex: 1,
        },
      ]);
    });

    it('reports fade throughBlack, advanceOnClick=false, advanceAfter, and an authored sound action', () => {
      const presentation = buildPresentation([
        {
          effect: { kind: 'fade', throughBlack: true },
          advanceOnClick: false,
          advanceAfter: 5000,
          sound: { kind: 'stop' },
        },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      const codes = unsupportedFeatures.all.map((feature) => feature.code);
      expect(codes).toEqual(
        expect.arrayContaining([
          'transition-through-black-unmodeled',
          'transition-advance-on-click-unmodeled',
          'transition-advance-after-unmodeled',
          'transition-sound-unmodeled',
        ]),
      );
      // fade itself IS supported, so no transition-effect-unmodeled entry should be present.
      expect(codes).not.toContain('transition-effect-unmodeled');
      expect(unsupportedFeatures.all).toHaveLength(4);
    });

    it('ignores navigation entirely while a transition is in flight', async () => {
      const presentation = buildPresentation([undefined, { effect: { kind: 'fade' } }, undefined]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next(); // starts the fade into slide 1
      expect(el.currentSlideIndex).toBe(1);

      el.next();
      expect(el.currentSlideIndex).toBe(1);
      el.previous();
      expect(el.currentSlideIndex).toBe(1);
      el.goToSlide(2);
      expect(el.currentSlideIndex).toBe(1);
      el.click();
      expect(el.currentSlideIndex).toBe(1);
      el.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }));
      expect(el.currentSlideIndex).toBe(1);

      await vi.advanceTimersByTimeAsync(400); // fade's "fast"-default duration elapses

      el.next();
      expect(el.currentSlideIndex).toBe(2);
    });

    it('maps "med"/"slow" speed to their own duration, defaulting absent speed to "fast"', async () => {
      const cases: { speed: SlideTransition['speed']; durationMs: number }[] = [
        { speed: undefined, durationMs: 400 },
        { speed: 'med', durationMs: 700 },
        { speed: 'slow', durationMs: 1000 },
      ];

      for (const { speed, durationMs } of cases) {
        recordedAnimations = [];
        const presentation = buildPresentation([
          undefined,
          { effect: { kind: 'fade' }, ...(speed ? { speed } : {}) },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        el.render(presentation);

        el.next();
        const [, incoming] = slideEls(el);
        expect(latestAnimation(incoming!)?.options).toMatchObject({ duration: durationMs });

        await vi.advanceTimersByTimeAsync(durationMs - 1);
        expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(true);

        await vi.advanceTimersByTimeAsync(1);
        expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(false);
      }
    });

    it('lets an explicit durationMs (p14:dur) override speed entirely', async () => {
      const presentation = buildPresentation([
        undefined,
        { effect: { kind: 'fade' }, speed: 'slow', durationMs: 2000 },
      ]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      el.next();
      const [, incoming] = slideEls(el);
      expect(latestAnimation(incoming!)?.options).toMatchObject({ duration: 2000 });

      await vi.advanceTimersByTimeAsync(2000);
    });

    it('an unmodeled effect kind (e.g. wipe) falls back to the instant swap and is reported unsupported', () => {
      const presentation = buildPresentation([undefined, { effect: { kind: 'wipe' } }]);
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      const unsupportedFeatures = el.render(presentation);

      el.next();
      expect(el.currentSlideIndex).toBe(1);
      const [, incoming] = slideEls(el);
      expect(incoming?.classList.contains('pptx-slide--active')).toBe(true);
      expect(incoming?.classList.contains('pptx-slide--transitioning')).toBe(false);

      expect(unsupportedFeatures.all).toContainEqual(
        expect.objectContaining({ code: 'transition-effect-unmodeled', slideIndex: 1 }),
      );
    });

    describe('morph transitions', () => {
      function morphShape(id: number, name: string, x: number, y: number): Shape {
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

      it('moves the arriving copy opaquely between the two boxes and hides the departing copy instantly (not a fade)', async () => {
        const presentation = buildMorphPresentation([
          { shapes: [morphShape(1, 'Title 1', 0, 0)] },
          {
            shapes: [morphShape(1, 'Title 1', 1000000, 2000000)],
            transition: { effect: { kind: 'morph' } },
          },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        el.render(presentation);

        const [outgoingSlide, incomingSlide] = slideEls(el);
        const departing = outgoingSlide!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;
        const arriving = incomingSlide!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;

        el.next();

        // Crossfading two overlapping opaque copies would make both translucent mid-transition
        // (the bug this design avoids) — the departing copy is switched off instantly instead.
        const departingAnim = latestAnimation(departing);
        expect(departingAnim?.keyframes).toEqual([{ opacity: 0 }]);
        expect(departingAnim?.options).toMatchObject({ duration: 0 });

        // The arriving copy alone carries the "same object moving" illusion: no opacity field at
        // all, so it stays fully opaque for its whole journey from the departing box to its own.
        const arrivingKeyframes = latestAnimation(arriving)?.keyframes;
        expect(arrivingKeyframes).toBeDefined();
        expect(arrivingKeyframes![0]).toMatchObject({ left: departing.style.left });
        expect(arrivingKeyframes![1]).toMatchObject({ left: arriving.style.left });
        expect(arrivingKeyframes![0]!.opacity).toBeUndefined();
        expect(arrivingKeyframes![1]!.opacity).toBeUndefined();
        expect(incomingSlide?.classList.contains('pptx-slide--transitioning')).toBe(true);

        await vi.advanceTimersByTimeAsync(400);
      });

      it('fades a disappearing shape out in place and an appearing shape in in place', async () => {
        const presentation = buildMorphPresentation([
          { shapes: [morphShape(1, 'Title 1', 0, 0), morphShape(2, 'Old Icon', 500000, 500000)] },
          {
            shapes: [morphShape(1, 'Title 1', 0, 0), morphShape(3, 'New Icon', 700000, 700000)],
            transition: { effect: { kind: 'morph' } },
          },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        el.render(presentation);

        const [outgoingSlide, incomingSlide] = slideEls(el);
        const oldIcon = outgoingSlide!.querySelector<HTMLElement>('[data-pptx-shape-id="2"]')!;
        const newIcon = incomingSlide!.querySelector<HTMLElement>('[data-pptx-shape-id="3"]')!;

        el.next();

        expect(latestAnimation(oldIcon)?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
        expect(latestAnimation(newIcon)?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);

        await vi.advanceTimersByTimeAsync(400);
      });

      it('falls back to a plain crossfade and reports morph-match-degraded for a low-confidence match', async () => {
        const presentation = buildMorphPresentation([
          { shapes: [morphShape(1, 'Alpha', 0, 0)] },
          { shapes: [morphShape(2, 'Beta', 0, 0)], transition: { effect: { kind: 'morph' } } },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        const unsupportedFeatures = el.render(presentation);

        expect(unsupportedFeatures.all).toContainEqual(
          expect.objectContaining({ code: 'morph-match-degraded', slideIndex: 1 }),
        );

        const [outgoingSlide, incomingSlide] = slideEls(el);
        el.next();

        // Falls back to a plain whole-slide crossfade (#animateFade), not a per-shape morph tween.
        expect(latestAnimation(outgoingSlide!)?.keyframes).toEqual([
          { opacity: 1 },
          { opacity: 0 },
        ]);
        expect(latestAnimation(incomingSlide!)?.keyframes).toEqual([
          { opacity: 0 },
          { opacity: 1 },
        ]);
        expect(incomingSlide?.classList.contains('pptx-slide--transitioning')).toBe(true);

        await vi.advanceTimersByTimeAsync(400);
      });

      it('reports morph-match-degraded for a morph transition on the very first slide', () => {
        const presentation = buildMorphPresentation([
          { shapes: [morphShape(1, 'Title 1', 0, 0)], transition: { effect: { kind: 'morph' } } },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        const unsupportedFeatures = el.render(presentation);

        expect(unsupportedFeatures.all).toContainEqual(
          expect.objectContaining({ code: 'morph-match-degraded', slideIndex: 0 }),
        );
      });

      it('reverses departing/arriving roles correctly for backward navigation', async () => {
        const presentation = buildMorphPresentation([
          { shapes: [morphShape(1, 'Title 1', 0, 0)] },
          {
            shapes: [morphShape(1, 'Title 1', 1000000, 1000000)],
            transition: { effect: { kind: 'morph' } },
          },
        ]);
        const el = document.createElement('pptx-presentation') as PptxPresentationElement;
        el.render(presentation);

        const [slide0, slide1] = slideEls(el);
        const shapeOnSlide0 = slide0!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;
        const shapeOnSlide1 = slide1!.querySelector<HTMLElement>('[data-pptx-shape-id="1"]')!;

        el.next();
        await vi.advanceTimersByTimeAsync(400);
        recordedAnimations = []; // isolate the backward navigation's own animate() calls

        el.previous();

        // Now slide1 (currently shown) is departing (hidden instantly) and slide0 (the target) is
        // arriving (the one that moves, opaquely) — the opposite of the forward case, even though
        // it's the exact same authored transition.
        const departingAnim = latestAnimation(shapeOnSlide1);
        expect(departingAnim?.keyframes).toEqual([{ opacity: 0 }]);
        expect(departingAnim?.options).toMatchObject({ duration: 0 });

        const arrivingKeyframes = latestAnimation(shapeOnSlide0)?.keyframes;
        expect(arrivingKeyframes![0]).toMatchObject({ left: shapeOnSlide1.style.left });
        expect(arrivingKeyframes![1]).toMatchObject({ left: shapeOnSlide0.style.left });

        await vi.advanceTimersByTimeAsync(400);
      });
    });
  });

  describe('Slide.timing fade animations', () => {
    // Only needs to record calls, not drive a `.finished` promise to resolution — unlike
    // push/fade slide transitions (see above), #playSlideAnimations fires-and-forgets each
    // Animation, so there's nothing to await here.
    let recordedCalls: { target: HTMLElement; keyframes: Keyframe[]; options: unknown }[] = [];

    beforeEach(() => {
      recordedCalls = [];
      HTMLElement.prototype.animate = vi.fn(function (
        this: HTMLElement,
        keyframes: Keyframe[] | PropertyIndexedKeyframes | null,
        options?: number | KeyframeAnimationOptions,
      ) {
        recordedCalls.push({ target: this, keyframes: keyframes as Keyframe[], options });
        return {} as Animation;
      }) as typeof HTMLElement.prototype.animate;
    });

    afterEach(() => {
      delete (HTMLElement.prototype as { animate?: unknown }).animate;
    });

    function shape(id: number): Shape {
      return { kind: 'shape', nonVisual: { id, name: `Shape ${id}` }, properties: {} };
    }

    function buildPresentation(fadeShapeId: number, transition: 'in' | 'out' = 'in'): Presentation {
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
                common: { id: 0, duration: 500 },
                target: { kind: 'shape', shapeId: fadeShapeId },
                filter: 'fade',
                transition,
              },
            },
          },
        ],
        notesSlides: [],
      };
    }

    it('plays a fade-in on the target shape as soon as its slide is shown', () => {
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(buildPresentation(42));

      const target = el.shadowRoot?.querySelector<HTMLElement>('[data-pptx-shape-id="42"]');
      expect(target).not.toBeNull();
      const call = recordedCalls.find((c) => c.target === target);
      expect(call).toBeDefined();
      expect(call?.keyframes).toEqual([{ opacity: 0 }, { opacity: 1 }]);
      expect(call?.options).toMatchObject({ duration: 500, delay: 0, fill: 'forwards' });
    });

    it('plays a fade-out with reversed keyframes', () => {
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(buildPresentation(42, 'out'));

      const target = el.shadowRoot?.querySelector<HTMLElement>('[data-pptx-shape-id="42"]');
      const call = recordedCalls.find((c) => c.target === target);
      expect(call?.keyframes).toEqual([{ opacity: 1 }, { opacity: 0 }]);
    });

    it("replays a slide's fade animations each time it becomes active again", () => {
      const base = buildPresentation(42);
      const presentation: Presentation = {
        ...base,
        // A second, untimed slide to navigate to and back from.
        slides: [...base.slides, { commonSlideData: { shapeTree: [] }, layout: LAYOUT }],
      };
      const el = document.createElement('pptx-presentation') as PptxPresentationElement;
      el.render(presentation);

      const target = el.shadowRoot?.querySelector<HTMLElement>('[data-pptx-shape-id="42"]');
      const callsOnTarget = () => recordedCalls.filter((c) => c.target === target);
      expect(callsOnTarget()).toHaveLength(1); // from the initial render

      el.next();
      el.previous();
      expect(callsOnTarget()).toHaveLength(2); // replayed on re-entering slide 0
    });
  });
});
