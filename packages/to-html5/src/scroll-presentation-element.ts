import type { Presentation } from '@pptx2html/presentation';
import type { ShapeFadeAnimation } from './animation.js';
import { resolveScrollTimeline, type ScrollSegment } from './scroll-timeline.js';
import { renderSlide } from './slide.js';
import {
  findShapeElement,
  IDENTITY_TRANSFORM,
  morphKeyframes,
  offscreenTransform,
  readShapeBox,
  REVERSE_SIDE_DIRECTION,
} from './transition-keyframes.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

const STYLES = `
  :host {
    display: block;
    position: relative;
    overflow: hidden;
  }
  .pptx-scroll-track {
    position: absolute;
    inset: 0;
    overflow-y: auto;
    overflow-x: hidden;
    z-index: 1;
  }
  .pptx-scroll-spacer {
    width: 100%;
  }
  .pptx-scroll-viewport {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    z-index: 0;
  }
`;

const DEFAULT_PIXELS_PER_SECOND = 600;

/** One transition segment's pre-built, paused `Animation`s — see `#buildTransitionAnimations`. */
type TransitionAnimations =
  | { readonly kind: 'push' | 'fade'; readonly outgoing: Animation; readonly incoming: Animation }
  | {
      readonly kind: 'morph';
      readonly matched: readonly { arrivingAnimation: Animation; departingEl: HTMLElement }[];
      readonly fadeOut: readonly Animation[];
      readonly fadeIn: readonly Animation[];
    };

type ActiveState =
  | { readonly kind: 'transition'; readonly segment: ScrollSegment; readonly elapsedMs: number }
  | { readonly kind: 'content'; readonly segment: ScrollSegment; readonly elapsedMs: number };

/**
 * Which segment (and how far into it) `ms` falls in, given the deck's contiguous, sequentially
 * chained segments (see `scroll-timeline.ts`). A pure function, kept outside the class so it's
 * directly testable without any DOM: the first segment whose content phase hasn't ended yet (or,
 * failing that, the last segment — covers `ms` clamped to exactly `totalDurationMs`) is the active
 * one; whether that's its own transition-in or its content phase depends only on whether `ms` still
 * precedes that segment's own `transition.endMs`.
 */
function locateSegment(segments: readonly ScrollSegment[], ms: number): ActiveState | undefined {
  for (let i = 0; i < segments.length; i++) {
    const segment = segments[i]!;
    const isLast = i === segments.length - 1;
    if (ms > segment.content.endMs && !isLast) continue;
    if (segment.transition && ms < segment.transition.endMs) {
      return { kind: 'transition', segment, elapsedMs: ms - segment.transition.startMs };
    }
    return { kind: 'content', segment, elapsedMs: Math.max(0, ms - segment.content.startMs) };
  }
  return undefined;
}

/**
 * A `<pptx-scroll-presentation>` element: renders a `Presentation` as a continuously scrubbable
 * "scrollytelling" timeline instead of a click-advanced slideshow (`PptxPresentationElement`,
 * `presentation-element.ts`) — every slide transition and build animation this package already
 * knows how to play is instead positioned on one absolute-millisecond timeline
 * (`resolveScrollTimeline`, `scroll-timeline.ts`) and scrubbed via each `Animation`'s own
 * `currentTime` as the user scrolls, rather than fired-and-forgotten in real time. See
 * `packages/to-html5/CLAUDE.md`'s scroll-mode design decision for the full rationale, in
 * particular why this is a *separate* element rather than a mode on `PptxPresentationElement`
 * (click-driven and scroll-driven navigation are different enough state machines).
 *
 * **The visible content (`.pptx-scroll-viewport`) is a sibling of the scrolling element
 * (`.pptx-scroll-track`), never a descendant of it — deliberately, to avoid the browser's own
 * scroll-linked compositing/layout fighting with this class's `seekTo` writes.** An earlier version
 * nested the viewport inside the track (compensating for scroll with a manual `transform:
 * translateY(scrollTop)` on every frame, the way `position: sticky` would do natively) and this
 * visibly jittered: the rendered slide subtree is real, non-trivial DOM, and having it live inside
 * the actively-scrolling element meant the browser's native scroll compositing of that subtree and
 * this class's once-per-frame JS correction were two independent, not-quite-synchronized sources of
 * truth for the same box. Since `.pptx-scroll-track` here holds nothing but an empty spacer (no
 * visible content at all), moving the real content out to a sibling removes that race entirely —
 * `.pptx-scroll-viewport` never moves due to scrolling, because it was never part of the scrolling
 * subtree to begin with, no compensating transform needed. `.pptx-scroll-track` still visually
 * covers the same box (`position: absolute; inset: 0`, on top via `z-index`) purely to keep
 * capturing scroll/wheel/touch input over the whole visible area — see the constructor.
 *
 * **The host page must give this element an explicit box (e.g. `height: 100vh`) and must not
 * place it inside another scrollable ancestor** — the scrolling happens on a child inside this
 * element's own shadow DOM (`.pptx-scroll-track`), not `window`/`document`, and this element
 * assumes it's the only scroll region involved in its own area of the page.
 *
 * **Known trade-off**: `.pptx-scroll-track` sits on top (`z-index: 1`) so it can hit-test scroll
 * gestures anywhere over the visible area — it has to be topmost, since wheel/touch scroll
 * delegation only walks up the ancestor chain from whatever's hit, never sideways to a sibling.
 * That means it also captures clicks, which is harmless today (nothing this package renders is
 * interactive) but would need revisiting (e.g. `pointer-events: none` toggled off during a
 * click/tap) if that ever changes.
 */
export class PptxScrollPresentationElement extends HTMLElement {
  readonly #track: HTMLElement;
  readonly #spacer: HTMLElement;
  readonly #viewport: HTMLElement;

  #slideEls: HTMLElement[] = [];
  #segments: readonly ScrollSegment[] = [];
  #totalDurationMs = 0;
  #transitionAnimations = new Map<number, TransitionAnimations>();
  #contentAnimations = new Map<number, readonly (Animation | undefined)[]>();
  /** Shapes given a manual `opacity: 0` override for a Morph transition's departing copy — reset
   *  at the start of every `seekTo` call before this call's own active state re-applies it, since
   *  which shapes (if any) need it changes as the user scrubs. See `seekTo`. */
  #hiddenMorphShapes: HTMLElement[] = [];

  #pixelsPerSecond = DEFAULT_PIXELS_PER_SECOND;
  #pendingScrollTop: number | null = null;
  #rafId: number | null = null;

  constructor() {
    super();
    const shadow = this.attachShadow({ mode: 'open' });

    const style = this.ownerDocument.createElement('style');
    style.textContent = STYLES;
    shadow.appendChild(style);

    this.#viewport = this.ownerDocument.createElement('div');
    this.#viewport.className = 'pptx-scroll-viewport';

    this.#track = this.ownerDocument.createElement('div');
    this.#track.className = 'pptx-scroll-track';
    this.#spacer = this.ownerDocument.createElement('div');
    this.#spacer.className = 'pptx-scroll-spacer';
    this.#track.appendChild(this.#spacer);

    // Siblings, not nested — see the class doc comment on why the viewport must not live inside
    // the scrolling track.
    shadow.appendChild(this.#viewport);
    shadow.appendChild(this.#track);

    this.#track.addEventListener('scroll', this.#handleScroll, { passive: true });
  }

  disconnectedCallback(): void {
    if (this.#rafId !== null) cancelAnimationFrame(this.#rafId);
    this.#rafId = null;
    this.#pendingScrollTop = null;
  }

  /** Scroll distance, in CSS pixels, per millisecond of timeline — controls the track's total
   *  scroll height (`totalDurationMs * pixelsPerSecond / 1000`). Tunable; not spec-derived. */
  get pixelsPerSecond(): number {
    return this.#pixelsPerSecond;
  }

  set pixelsPerSecond(value: number) {
    this.#pixelsPerSecond = value;
    this.#applyTrackHeight();
  }

  get totalDurationMs(): number {
    return this.#totalDurationMs;
  }

  /** Returns the set of `.pptx`-authored features this render didn't (fully) support. */
  render(presentation: Presentation): UnsupportedFeatureCollector {
    this.#cancelAllAnimations();

    const unsupportedFeatures = new UnsupportedFeatureCollector();
    this.#slideEls = presentation.slides.map((slide, slideIndex) => {
      const el = renderSlide(
        this.ownerDocument,
        slide,
        presentation.slideSize,
        presentation.defaultTextStyle,
        slideIndex,
        unsupportedFeatures,
      );
      // Every slide is permanently stacked in the viewport — unlike PptxPresentationElement,
      // which only goes position: absolute transiently during a transition, this element never
      // has a single-slide "normal flow" state to return to.
      el.style.position = 'absolute';
      el.style.top = '0';
      el.style.left = '0';
      el.style.width = '100%';
      el.style.display = 'none';
      return el;
    });
    this.#viewport.replaceChildren(...this.#slideEls);
    this.#viewport.style.aspectRatio = `${presentation.slideSize.width} / ${presentation.slideSize.height}`;

    const report = (code: string, message: string, slideIndex: number) =>
      unsupportedFeatures.report({ code, message, slideIndex });
    const { segments, totalDurationMs } = resolveScrollTimeline(presentation, report);
    this.#segments = segments;
    this.#totalDurationMs = totalDurationMs;

    for (const segment of segments) {
      if (segment.transition) {
        this.#transitionAnimations.set(
          segment.slideIndex,
          this.#buildTransitionAnimations(segment),
        );
      }
      this.#contentAnimations.set(
        segment.slideIndex,
        segment.content.fades.map((fade) =>
          this.#buildFadeAnimation(this.#slideEls[segment.slideIndex]!, fade),
        ),
      );
    }

    this.#applyTrackHeight();
    this.#track.scrollTop = 0;
    this.seekTo(0);

    return unsupportedFeatures;
  }

  /**
   * Sets the deck to look exactly as it would `ms` milliseconds into the timeline — the core,
   * directly testable entry point (no real scrolling or layout involved): clamps to
   * `[0, totalDurationMs]`, finds the active segment via `locateSegment`, shows only the slide(s)
   * that segment's state calls for, and sets every relevant `Animation`'s `currentTime`. The
   * `scroll`-driven rAF loop below is the only caller that derives `ms` from real scroll state;
   * everything else (tests, that loop) shares this one code path.
   */
  seekTo(ms: number): void {
    if (this.#segments.length === 0) return;
    const clamped = Math.min(Math.max(ms, 0), this.#totalDurationMs);
    const state = locateSegment(this.#segments, clamped);
    if (!state) return;

    for (const el of this.#hiddenMorphShapes) el.style.opacity = '';
    this.#hiddenMorphShapes = [];

    if (state.kind === 'transition') {
      this.#applyVisibility(new Set([state.segment.slideIndex - 1, state.segment.slideIndex]));
      const anims = this.#transitionAnimations.get(state.segment.slideIndex);
      if (anims) this.#scrubTransition(anims, state.elapsedMs);
    } else {
      this.#applyVisibility(new Set([state.segment.slideIndex]));
      const fadeAnims = this.#contentAnimations.get(state.segment.slideIndex) ?? [];
      state.segment.content.fades.forEach((fade, index) => {
        const animation = fadeAnims[index];
        if (!animation) return;
        animation.currentTime = Math.min(
          Math.max(state.elapsedMs, 0),
          fade.delayMs + fade.durationMs,
        );
      });
    }
  }

  #applyVisibility(visibleIndices: ReadonlySet<number>): void {
    this.#slideEls.forEach((el, index) => {
      el.style.display = visibleIndices.has(index) ? 'block' : 'none';
    });
  }

  #scrubTransition(anims: TransitionAnimations, elapsedMs: number): void {
    if (anims.kind === 'morph') {
      for (const { arrivingAnimation, departingEl } of anims.matched) {
        arrivingAnimation.currentTime = elapsedMs;
        // Constant for the whole transition window (see morphKeyframes's own doc comment on why
        // the departing copy is hidden instantly rather than faded) — a direct style write
        // rather than a third Animation, reset every seekTo call via #hiddenMorphShapes above.
        departingEl.style.opacity = '0';
        this.#hiddenMorphShapes.push(departingEl);
      }
      for (const animation of anims.fadeOut) animation.currentTime = elapsedMs;
      for (const animation of anims.fadeIn) animation.currentTime = elapsedMs;
      return;
    }
    anims.outgoing.currentTime = elapsedMs;
    anims.incoming.currentTime = elapsedMs;
  }

  #buildTransitionAnimations(segment: ScrollSegment): TransitionAnimations {
    const transition = segment.transition!;
    const outgoing = this.#slideEls[segment.slideIndex - 1]!;
    const incoming = this.#slideEls[segment.slideIndex]!;
    const duration = transition.endMs - transition.startMs;
    const options: KeyframeAnimationOptions = { duration, fill: 'both' };
    const effect = transition.effect;

    if (effect.kind === 'push') {
      const direction = effect.direction ?? 'l';
      const outgoingAnim = outgoing.animate(
        [{ transform: IDENTITY_TRANSFORM }, { transform: offscreenTransform(direction) }],
        options,
      );
      const incomingAnim = incoming.animate(
        [
          { transform: offscreenTransform(REVERSE_SIDE_DIRECTION[direction]) },
          { transform: IDENTITY_TRANSFORM },
        ],
        options,
      );
      outgoingAnim.pause();
      incomingAnim.pause();
      return { kind: 'push', outgoing: outgoingAnim, incoming: incomingAnim };
    }

    if (effect.kind === 'fade') {
      const outgoingAnim = outgoing.animate([{ opacity: 1 }, { opacity: 0 }], options);
      const incomingAnim = incoming.animate([{ opacity: 0 }, { opacity: 1 }], options);
      outgoingAnim.pause();
      incomingAnim.pause();
      return { kind: 'fade', outgoing: outgoingAnim, incoming: incomingAnim };
    }

    // 'morph' — resolveScrollTimeline only ever sets effect.kind to 'morph' when morphMatch is
    // defined (a degraded/unmatched morph already becomes a plain 'fade' there instead).
    const match = transition.morphMatch!;
    const matched: { arrivingAnimation: Animation; departingEl: HTMLElement }[] = [];
    for (const pair of match.matched) {
      const departingEl = findShapeElement(outgoing, pair.outgoingShapeId);
      const arrivingEl = findShapeElement(incoming, pair.incomingShapeId);
      if (!departingEl || !arrivingEl) continue;
      const departingBox = readShapeBox(departingEl);
      const arrivingBox = readShapeBox(arrivingEl);
      if (!departingBox || !arrivingBox) continue;
      const arrivingAnimation = arrivingEl.animate(
        morphKeyframes(departingBox, arrivingBox),
        options,
      );
      arrivingAnimation.pause();
      matched.push({ arrivingAnimation, departingEl });
    }
    const fadeOut = match.disappearingShapeIds
      .map((id) => findShapeElement(outgoing, id))
      .filter((el): el is HTMLElement => el !== null)
      .map((el) => {
        const anim = el.animate([{ opacity: 1 }, { opacity: 0 }], options);
        anim.pause();
        return anim;
      });
    const fadeIn = match.appearingShapeIds
      .map((id) => findShapeElement(incoming, id))
      .filter((el): el is HTMLElement => el !== null)
      .map((el) => {
        const anim = el.animate([{ opacity: 0 }, { opacity: 1 }], options);
        anim.pause();
        return anim;
      });
    return { kind: 'morph', matched, fadeOut, fadeIn };
  }

  #buildFadeAnimation(slideEl: HTMLElement, fade: ShapeFadeAnimation): Animation | undefined {
    const target = slideEl.querySelector<HTMLElement>(`[data-pptx-shape-id="${fade.shapeId}"]`);
    if (!target) return undefined;
    const keyframes: Keyframe[] =
      fade.direction === 'in' ? [{ opacity: 0 }, { opacity: 1 }] : [{ opacity: 1 }, { opacity: 0 }];
    const animation = target.animate(keyframes, {
      duration: fade.durationMs,
      delay: fade.delayMs,
      fill: 'both',
    });
    animation.pause();
    return animation;
  }

  #applyTrackHeight(): void {
    const pixelsPerMs = this.#pixelsPerSecond / 1000;
    this.#spacer.style.height = `${this.#totalDurationMs * pixelsPerMs}px`;
  }

  #cancelAllAnimations(): void {
    for (const anims of this.#transitionAnimations.values()) {
      if (anims.kind === 'morph') {
        for (const { arrivingAnimation } of anims.matched) arrivingAnimation.cancel();
        for (const animation of anims.fadeOut) animation.cancel();
        for (const animation of anims.fadeIn) animation.cancel();
      } else {
        anims.outgoing.cancel();
        anims.incoming.cancel();
      }
    }
    for (const animations of this.#contentAnimations.values()) {
      for (const animation of animations) animation?.cancel();
    }
    this.#transitionAnimations = new Map();
    this.#contentAnimations = new Map();
    for (const el of this.#hiddenMorphShapes) el.style.opacity = '';
    this.#hiddenMorphShapes = [];
  }

  #handleScroll = (): void => {
    this.#pendingScrollTop = this.#track.scrollTop;
    if (this.#rafId === null) {
      this.#rafId = requestAnimationFrame(this.#flushScroll);
    }
  };

  /**
   * Runs once per animation frame regardless of how many `scroll` events fired within it —
   * standard scroll-perf coalescing, and independently useful for a future "perfect pinning" goal
   * (one clock, read once per frame). `.pptx-scroll-viewport` needs no positioning write here at
   * all — see the class doc comment on why living outside the scrolling subtree already keeps it
   * visually put — so this only ever calls `seekTo`.
   */
  #flushScroll = (): void => {
    this.#rafId = null;
    if (this.#pendingScrollTop === null) return;
    const scrollTop = this.#pendingScrollTop;
    this.#pendingScrollTop = null;
    this.seekTo(scrollTop / (this.#pixelsPerSecond / 1000));
  };
}

export function defineScrollPresentationElement(tagName = 'pptx-scroll-presentation'): void {
  if (!customElements.get(tagName)) {
    customElements.define(tagName, PptxScrollPresentationElement);
  }
}
