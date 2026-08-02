// @vitest-environment happy-dom
import type { Presentation, SlideLayout, SlideTransition } from '@pptx2html/presentation';
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
  });
});
