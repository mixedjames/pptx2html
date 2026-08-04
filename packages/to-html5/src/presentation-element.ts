import type { Presentation, SideDirection, Slide, SlideTransition } from '@pptx2html/presentation';
import { resolveTransitionDurationMs } from '@pptx2html/presentation';
import { collectFadeAnimations, type ShapeFadeAnimation } from './animation.js';
import { resolveSlideMorphMatch, type MorphMatchSummary } from './morph.js';
import { renderSlide } from './slide.js';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

/** `reportSlideLevelFeatures`'s return value — see its own doc comment. */
interface SlideLevelFeatures {
  readonly fades: readonly ShapeFadeAnimation[];
  readonly morphMatch: MorphMatchSummary | undefined;
}

/**
 * Everything unsupported that's a property of a whole slide, not tied to one shape — a slide's
 * own `timing`/`transition` — checked once per `render()` up front (see the call site) rather
 * than as each is actually used, so the log is complete even for slides/effects never exercised
 * during this session (a transition not navigated into, timing that's never played at all).
 * Returns the slide's playable fade animations (see `animation.ts`) and its resolved Morph
 * shape-match summary if it has one (see `morph.ts`) — `render()` stores both, alongside
 * `#transitions`, for `#playSlideAnimations`/`#animateMorph` to use once relevant.
 */
function reportSlideLevelFeatures(
  slide: Slide,
  previousSlide: Slide | undefined,
  slideIndex: number,
  unsupportedFeatures: UnsupportedFeatureCollector,
): SlideLevelFeatures {
  const report = (code: string, message: string) =>
    unsupportedFeatures.report({ code, message, slideIndex });
  const fades = collectFadeAnimations(slide.timing, report);

  const transition = slide.transition;
  const effect = transition?.effect;
  const morphMatch =
    effect?.kind === 'morph' ? resolveSlideMorphMatch(previousSlide, slide, report) : undefined;
  if (effect && effect.kind !== 'push' && effect.kind !== 'fade' && effect.kind !== 'morph') {
    unsupportedFeatures.report({
      code: 'transition-effect-unmodeled',
      message: `Slide transition effect "${effect.kind}" is not animated; falling back to an instant swap.`,
      slideIndex,
    });
  }
  if (effect?.kind === 'fade' && effect.throughBlack) {
    unsupportedFeatures.report({
      code: 'transition-through-black-unmodeled',
      message:
        'The fade transition\'s "through black" two-stage variant renders as a plain crossfade.',
      slideIndex,
    });
  }
  if (transition?.advanceOnClick === false) {
    unsupportedFeatures.report({
      code: 'transition-advance-on-click-unmodeled',
      message: 'This slide is authored not to advance on click, but clicking still advances it.',
      slideIndex,
    });
  }
  if (transition?.advanceAfter !== undefined) {
    unsupportedFeatures.report({
      code: 'transition-advance-after-unmodeled',
      message: `This slide is authored to auto-advance after ${transition.advanceAfter}ms; auto-advance is not implemented.`,
      slideIndex,
    });
  }
  if (transition?.sound) {
    unsupportedFeatures.report({
      code: 'transition-sound-unmodeled',
      message: "This slide transition's authored sound action is not played.",
      slideIndex,
    });
  }

  return { fades, morphMatch };
}

const STYLES = `
  :host {
    display: block;
  }
  .pptx-presentation {
    display: block;
    position: relative;
    overflow: hidden;
  }
  .pptx-slide {
    display: none;
  }
  .pptx-slide.pptx-slide--active {
    display: block;
  }
  .pptx-slide.pptx-slide--transitioning {
    top: 0;
    left: 0;
  }
`;

const NEXT_KEYS = new Set(['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter']);
const PREVIOUS_KEYS = new Set(['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace']);

const REVERSE_SIDE_DIRECTION: Record<SideDirection, SideDirection> = {
  l: 'r',
  r: 'l',
  u: 'd',
  d: 'u',
};

/** The translate() a slide sits at when fully off-screen in `direction`, the way `push` uses it. */
function offscreenTransform(direction: SideDirection): string {
  switch (direction) {
    case 'l':
      return 'translate(-100%, 0)';
    case 'r':
      return 'translate(100%, 0)';
    case 'u':
      return 'translate(0, -100%)';
    case 'd':
      return 'translate(0, 100%)';
  }
}

const IDENTITY_TRANSFORM = 'translate(0, 0)';

/** A shape element's own resolved box, read straight off its inline style — see `readShapeBox`. */
interface ShapeBox {
  readonly left: string;
  readonly top: string;
  readonly width: string;
  readonly height: string;
  readonly transform: string;
}

/** Finds a slide subtree's own element for `shapeId`, scoped to that slide — see `#animateMorph`. */
function findShapeElement(slideEl: HTMLElement, shapeId: number): HTMLElement | null {
  return slideEl.querySelector<HTMLElement>(`[data-pptx-shape-id="${shapeId}"]`);
}

/**
 * Reads a shape element's own already-resolved `left`/`top`/`width`/`height`/`transform` inline
 * style — set once by `shape-tree.ts`'s `positionElement` at render time — rather than
 * recomputing anything via `coordinate.ts`. Returns `undefined` if `positionElement` never ran
 * for this element at all (no resolved `left`, meaning it's unpositioned/static-flow — a
 * pre-existing rendering gap, see `shape-tree.ts`'s own scope notes), since animating with an
 * empty-string keyframe value isn't a valid CSS length.
 */
function readShapeBox(el: HTMLElement): ShapeBox | undefined {
  if (!el.style.left) return undefined;
  return {
    left: el.style.left,
    top: el.style.top,
    width: el.style.width,
    height: el.style.height,
    // "none" rather than "" so a keyframe interpolates from/to a well-defined identity transform
    // instead of an ambiguous empty string when only one side of the pair rotates/flips.
    transform: el.style.transform || 'none',
  };
}

/**
 * The two-keyframe `left`/`top`/`width`/`height`/`transform` tween a matched pair's *arriving*
 * copy uses in `#animateMorph` — deliberately no `opacity` field at all, so the element stays
 * fully opaque for its entire journey. See `#animateMorph`'s own doc comment for why: crossfading
 * two overlapping opaque copies (the technique whole-slide `#animateFade` correctly uses) makes
 * *both* visibly translucent for the middle of the transition, wrongly revealing whatever sits
 * behind them — wrong for a shape that's genuinely the same object, just moved/resized/rotated.
 */
function morphKeyframes(from: ShapeBox, to: ShapeBox): Keyframe[] {
  return [{ ...from }, { ...to }];
}

/**
 * A `<pptx-presentation>` element: renders a `Presentation` object graph into a shadow DOM as a
 * slideshow — one slide visible at a time, advanced by click, spacebar, or arrow/paging keys —
 * rather than every slide stacked one below the next. All slides are rendered into the DOM up
 * front and hidden via CSS (`display: none`) rather than re-rendered on navigation, which is what
 * lets a `push`/`fade` transition (see below) briefly show both the outgoing and incoming slide
 * at once instead of creating either on demand mid-transition.
 *
 * A slide's own `Slide.transition` (§19.3.1.49, `p:transition`) describes the effect played when
 * the presentation *arrives* at that slide. Navigating forward plays the destination's own
 * transition, as authored; navigating backward instead replays the *outgoing* slide's own
 * transition in reverse — undoing the animation that brought it into view — since OOXML has no
 * concept of "backward" for this to describe itself (see `goToSlide` below). `morph` (new) plays
 * too, when `@pptx2html/presentation`'s `resolveMorphMatch` found a confident-enough shape
 * correspondence between the two slides (see `morph.ts`/`#animateMorph`) — otherwise it degrades
 * to a plain crossfade (`#animateFade`) instead. Every *other* unmodeled `TransitionEffect.kind`
 * (wipe, cut, dissolve, wheel, ...) and a slide with no `transition` at all still fall back to
 * today's instant `display: none`/`block` swap — see `#animatePush`/`#animateFade`/
 * `#animateMorph` vs. `#updateActiveSlide`.
 *
 * `push`/`fade` play via the Web Animations API (`Element.animate()`), not a plain CSS
 * `transition` — a WAAPI `Animation`'s `currentTime` is directly seekable, which a future
 * scroll-driven playback mode will need (see `docs/scroll-driven-playback.md`); a plain CSS
 * transition can only be started and left to run. This also drops the reflow-forcing dance a
 * CSS-transition version needs (`.animate()` takes both endpoints as data in one call).
 */
export class PptxPresentationElement extends HTMLElement {
  readonly #slidesContainer: HTMLElement;
  #slides: HTMLElement[] = [];
  #transitions: (SlideTransition | undefined)[] = [];
  #animations: readonly (readonly ShapeFadeAnimation[])[] = [];
  #morphMatches: (MorphMatchSummary | undefined)[] = [];
  #currentIndex = 0;
  /** Non-empty while a transition is in flight — this *is* the "in flight" flag. */
  #currentAnimations: Animation[] = [];

  constructor() {
    super();
    const shadow = this.attachShadow({ mode: 'open' });

    const style = this.ownerDocument.createElement('style');
    style.textContent = STYLES;
    shadow.appendChild(style);

    this.#slidesContainer = this.ownerDocument.createElement('div');
    this.#slidesContainer.className = 'pptx-presentation';
    shadow.appendChild(this.#slidesContainer);

    this.addEventListener('click', () => this.next());
    this.addEventListener('keydown', (event) => this.#handleKeydown(event));
  }

  connectedCallback(): void {
    // Focusable so keydown navigation works once the user has clicked/tabbed into the element; a
    // click both focuses it (native behaviour for a tabindex-bearing element) and advances. Set
    // here rather than the constructor: the Custom Elements spec forbids a constructor from adding
    // attributes to the element (tabIndex reflects to the `tabindex` attribute) — doing it there
    // throws in strict implementations (WebKit) and aborts the upgrade entirely. Guarded so it
    // doesn't clobber an explicit `tabindex` the host page may have set.
    if (!this.hasAttribute('tabindex')) this.tabIndex = 0;
  }

  /** Returns the set of `.pptx`-authored features this render didn't (fully) support. */
  render(presentation: Presentation): UnsupportedFeatureCollector {
    for (const animation of this.#currentAnimations) animation.cancel();
    this.#currentAnimations = [];
    const unsupportedFeatures = new UnsupportedFeatureCollector();
    this.#slides = presentation.slides.map((slide, slideIndex) =>
      renderSlide(
        this.ownerDocument,
        slide,
        presentation.slideSize,
        presentation.defaultTextStyle,
        slideIndex,
        unsupportedFeatures,
      ),
    );
    this.#transitions = presentation.slides.map((slide) => slide.transition);
    // Detected here, not in goToSlide: goToSlide only runs a slide's transition once the user
    // actually navigates into it, but the log should surface every unsupported feature up front,
    // whether or not that slide is ever visited during this session. Also collects each slide's
    // playable fade animations (see animation.ts) for #playSlideAnimations below, and its Morph
    // shape-match summary (see morph.ts) for #animateMorph — computed here, from the live
    // `presentation.slides` array, and nowhere else: once this map finishes, nothing retained by
    // this element references the `Presentation`/`Slide` object graph any more (see morph.ts's own
    // doc comment for why that matters).
    const slideLevelFeatures = presentation.slides.map((slide, slideIndex) =>
      reportSlideLevelFeatures(
        slide,
        presentation.slides[slideIndex - 1],
        slideIndex,
        unsupportedFeatures,
      ),
    );
    this.#animations = slideLevelFeatures.map((features) => features.fades);
    this.#morphMatches = slideLevelFeatures.map((features) => features.morphMatch);
    this.#slidesContainer.replaceChildren(...this.#slides);
    // Needed because a transition briefly makes both participating slides `position: absolute`
    // (see #beginTransitionFrame), at which point neither contributes to this container's
    // normal-flow height — without this it would collapse to zero height for the transition's
    // duration. Mirrors slide.ts's own per-slide aspect-ratio line.
    this.#slidesContainer.style.aspectRatio = `${presentation.slideSize.width} / ${presentation.slideSize.height}`;
    this.#currentIndex = 0;
    this.#updateActiveSlide();
    return unsupportedFeatures;
  }

  /** Index of the currently visible slide. */
  get currentSlideIndex(): number {
    return this.#currentIndex;
  }

  get slideCount(): number {
    return this.#slides.length;
  }

  /** Advances to the next slide; a no-op on the last slide (or while a transition is playing). */
  next(): void {
    this.goToSlide(this.#currentIndex + 1);
  }

  /** Returns to the previous slide; a no-op on the first slide (or while a transition is playing). */
  previous(): void {
    this.goToSlide(this.#currentIndex - 1);
  }

  /**
   * Jumps directly to `index`, clamped to the valid slide range. A no-op while a transition from a
   * previous call is still playing — navigation is ignored entirely until it settles, rather than
   * interrupting or queuing.
   */
  goToSlide(index: number): void {
    if (this.#slides.length === 0 || this.#currentAnimations.length > 0) return;
    const clamped = Math.min(Math.max(index, 0), this.#slides.length - 1);
    if (clamped === this.#currentIndex) return;

    const fromIndex = this.#currentIndex;
    const toIndex = clamped;
    const forward = toIndex > fromIndex;
    this.#currentIndex = toIndex;

    // A slide's transition describes arriving at *that* slide (§19.3.1.49). Forward navigation
    // plays the destination's own transition; backward navigation instead replays the *outgoing*
    // slide's own transition — the one that brought it into view in the first place — since
    // that's the animation being undone, not whatever (if anything) the destination separately
    // authors for its own forward arrival.
    const transition = forward ? this.#transitions[toIndex] : this.#transitions[fromIndex];
    const morphMatch = forward ? this.#morphMatches[toIndex] : this.#morphMatches[fromIndex];
    const outgoing = this.#slides[fromIndex]!;
    const incoming = this.#slides[toIndex]!;

    if (transition?.effect?.kind === 'push') {
      this.#animatePush(
        outgoing,
        incoming,
        transition.effect.direction ?? 'l',
        forward,
        resolveTransitionDurationMs(transition.speed, transition.durationMs),
      );
    } else if (transition?.effect?.kind === 'fade') {
      this.#animateFade(
        outgoing,
        incoming,
        resolveTransitionDurationMs(transition.speed, transition.durationMs),
      );
    } else if (transition?.effect?.kind === 'morph' && morphMatch) {
      this.#animateMorph(
        outgoing,
        incoming,
        morphMatch,
        forward,
        resolveTransitionDurationMs(transition.speed, transition.durationMs),
      );
    } else if (transition?.effect?.kind === 'morph') {
      // resolveSlideMorphMatch already reported why (no previous slide, or a low-confidence
      // match) — this is just the fallback playback, same duration resolution as a real morph.
      this.#animateFade(
        outgoing,
        incoming,
        resolveTransitionDurationMs(transition.speed, transition.durationMs),
      );
    } else {
      this.#updateActiveSlide();
    }
  }

  /**
   * `direction` is always the transition's own authored direction — `goToSlide` already resolved
   * *which* slide's transition this is (destination going forward, outgoing going backward, see
   * above). Reversed here for backward navigation (`forward === false`) so the motion undoes the
   * original animation frame-for-frame, rather than repeating it. `duration` is already resolved
   * (`resolveTransitionDurationMs`, folding in an explicit `p14:dur` override when present) by the
   * caller, not recomputed here.
   */
  #animatePush(
    outgoing: HTMLElement,
    incoming: HTMLElement,
    direction: SideDirection,
    forward: boolean,
    duration: number,
  ): void {
    const effectiveDirection = forward ? direction : REVERSE_SIDE_DIRECTION[direction];
    this.#beginTransitionFrame(outgoing, incoming);

    const options: KeyframeAnimationOptions = { duration, easing: 'ease-in-out', fill: 'forwards' };
    const outgoingAnimation = outgoing.animate(
      [{ transform: IDENTITY_TRANSFORM }, { transform: offscreenTransform(effectiveDirection) }],
      options,
    );
    const incomingAnimation = incoming.animate(
      [
        { transform: offscreenTransform(REVERSE_SIDE_DIRECTION[effectiveDirection]) },
        { transform: IDENTITY_TRANSFORM },
      ],
      options,
    );

    this.#awaitTransition(outgoing, incoming, [outgoingAnimation, incomingAnimation]);
  }

  #animateFade(outgoing: HTMLElement, incoming: HTMLElement, duration: number): void {
    this.#beginTransitionFrame(outgoing, incoming);

    const options: KeyframeAnimationOptions = { duration, easing: 'ease-in-out', fill: 'forwards' };
    const outgoingAnimation = outgoing.animate([{ opacity: 1 }, { opacity: 0 }], options);
    const incomingAnimation = incoming.animate([{ opacity: 0 }, { opacity: 1 }], options);

    this.#awaitTransition(outgoing, incoming, [outgoingAnimation, incomingAnimation]);
  }

  /**
   * `outgoing`/`incoming` always mean "currently shown, about to leave"/"about to become shown"
   * here, exactly like `#animatePush`/`#animateFade` — what flips for backward navigation is only
   * *which shape id* to look up in each: `match` was computed as `resolveSlideMorphMatch(earlier-
   * in-deck slide, later-in-deck slide, ...)` (see `morph.ts`), so its own `outgoingShapeId`/
   * `incomingShapeId`/`disappearingShapeIds`/`appearingShapeIds` always mean "earlier"/"later" in
   * the deck, not "departing"/"arriving" — those only coincide for forward navigation. Going
   * backward, the *later*-in-deck slide is the one departing (`outgoing`), so it's `match`'s
   * `incomingShapeId`/`appearingShapeIds` that need to be found there instead, and vice versa for
   * `incoming` — `departingShapeId`/`arrivingShapeId`/`fadeOutIds`/`fadeInIds` below make this
   * swap explicit rather than leaving it implicit in which DOM element gets queried.
   *
   * Reads each matched shape's own already-rendered `left`/`top`/`width`/`height`/`transform`
   * directly off its element's inline style — deliberately not recomputed via `coordinate.ts`.
   *
   * **The departing copy is hidden instantly, not faded — a matched pair does *not* crossfade the
   * way `#animateFade` crossfades a whole slide.** An earlier version of this method animated both
   * copies' opacity (1→0 departing, 0→1 arriving) the same way `#animateFade` does for a whole
   * slide, on the theory that two crossfading copies read as one object moving. That's true for a
   * whole-slide crossfade (the point of that transition), but wrong here: alpha-compositing two
   * *overlapping, opaque* copies makes both visibly translucent for the middle of the tween — a
   * real bug a user caught by seeing text through a large solid shape mid-morph. A shape that
   * genuinely persists between two slides (the common case: same shape, just moved/resized/
   * rotated) should look like one continuous, fully opaque object gliding — not two ghosts
   * blending. So only the *arriving* copy animates its box (`morphKeyframes`, no `opacity` field
   * at all — stays opaque throughout); the *departing* copy is switched to `opacity: 0` via a
   * zero-duration `Animation` (`duration: 0`), an instant, not a fade, timed to coincide with the
   * arriving copy's very first frame already covering the exact same box — so nothing is ever
   * visibly uncovered. A zero-duration `Animation` is still a real `Animation` (`.finished` still
   * resolves, `.cancel()` still works), so it needs no special-casing in `#awaitTransition`/
   * `#finalizeTransition` below.
   *
   * A pair is skipped outright if either side has no resolved position (`positionElement` never
   * ran for it — an unpositioned/static-flow shape, a pre-existing, documented rendering gap
   * unrelated to Morph) rather than animating with invalid keyframe values.
   *
   * Every `Animation` this produces — matched pairs and the plain disappearing/appearing fades
   * alike — feeds into the same `#awaitTransition`/`#finalizeTransition` bookkeeping `#animatePush`/
   * `#animateFade` already use, not a separate tracking mechanism: `#finalizeTransition`'s existing
   * `for (const animation of this.#currentAnimations) animation.cancel()` loop is what reverts
   * every shape's inline-style override back to its own base (correct, resting) style once the
   * transition settles — without that, a departing shape would stay stuck at the opacity/box its
   * finished animation left it at the next time its own slide becomes active again.
   */
  #animateMorph(
    outgoing: HTMLElement,
    incoming: HTMLElement,
    match: MorphMatchSummary,
    forward: boolean,
    duration: number,
  ): void {
    this.#beginTransitionFrame(outgoing, incoming);

    const options: KeyframeAnimationOptions = { duration, easing: 'ease-in-out', fill: 'forwards' };
    const animations: Animation[] = [];

    for (const pair of match.matched) {
      const departingShapeId = forward ? pair.outgoingShapeId : pair.incomingShapeId;
      const arrivingShapeId = forward ? pair.incomingShapeId : pair.outgoingShapeId;
      const departingEl = findShapeElement(outgoing, departingShapeId);
      const arrivingEl = findShapeElement(incoming, arrivingShapeId);
      if (!departingEl || !arrivingEl) continue;
      const departingBox = readShapeBox(departingEl);
      const arrivingBox = readShapeBox(arrivingEl);
      if (!departingBox || !arrivingBox) continue;

      animations.push(departingEl.animate([{ opacity: 0 }], { duration: 0, fill: 'forwards' }));
      animations.push(arrivingEl.animate(morphKeyframes(departingBox, arrivingBox), options));
    }

    const fadeOutIds = forward ? match.disappearingShapeIds : match.appearingShapeIds;
    const fadeInIds = forward ? match.appearingShapeIds : match.disappearingShapeIds;
    for (const shapeId of fadeOutIds) {
      const el = findShapeElement(outgoing, shapeId);
      if (el) animations.push(el.animate([{ opacity: 1 }, { opacity: 0 }], options));
    }
    for (const shapeId of fadeInIds) {
      const el = findShapeElement(incoming, shapeId);
      if (el) animations.push(el.animate([{ opacity: 0 }, { opacity: 1 }], options));
    }

    this.#awaitTransition(outgoing, incoming, animations);
  }

  /** Shared setup: both slides become simultaneously visible and independently animatable. */
  #beginTransitionFrame(outgoing: HTMLElement, incoming: HTMLElement): void {
    for (const slide of [outgoing, incoming]) {
      slide.classList.add('pptx-slide--transitioning');
      slide.style.position = 'absolute';
    }
    incoming.classList.add('pptx-slide--active');
    incoming.style.zIndex = '1';
    outgoing.style.zIndex = '0';
  }

  /** Tracks the in-flight animations and schedules cleanup once both have finished. */
  #awaitTransition(
    outgoing: HTMLElement,
    incoming: HTMLElement,
    animations: readonly Animation[],
  ): void {
    this.#currentAnimations = [...animations];
    Promise.all(animations.map((animation) => animation.finished))
      .then(() => this.#finalizeTransition(outgoing, incoming))
      .catch(() => {
        // Rejects only when render() cancels these animations mid-flight (a new presentation was
        // rendered mid-transition) — that path already does its own cleanup.
      });
  }

  /** Cleans up once both animations have finished and restores the one-active-slide state. */
  #finalizeTransition(outgoing: HTMLElement, incoming: HTMLElement): void {
    // Only one transition is ever in flight at a time (goToSlide's guard), so this always
    // releases exactly the animations #awaitTransition just tracked. Cancelling matters for
    // correctness, not just hygiene: the same slide elements are reused across every future
    // transition, and leaving a finished fill:'forwards' Animation attached would stack
    // ambiguously against the next one.
    for (const animation of this.#currentAnimations) animation.cancel();
    this.#currentAnimations = [];
    for (const slide of [outgoing, incoming]) {
      slide.classList.remove('pptx-slide--transitioning');
      slide.style.position = 'relative'; // restores renderSlide's own inline default
      slide.style.zIndex = '';
    }
    this.#updateActiveSlide();
  }

  #updateActiveSlide(): void {
    this.#slides.forEach((slide, index) => {
      slide.classList.toggle('pptx-slide--active', index === this.#currentIndex);
    });
    this.#playSlideAnimations(this.#currentIndex);
  }

  /**
   * Plays the newly-active slide's fade animations (see `animation.ts`) — called from the one
   * place a slide actually becomes visible, whether via the initial `render()`, an instant
   * no-transition swap, or after `#finalizeTransition` settles a push/fade slide transition, so
   * entrance/exit fades play every time a slide is (re-)entered, matching how PowerPoint replays
   * a slide's own animations each time you navigate into it.
   */
  #playSlideAnimations(index: number): void {
    const slide = this.#slides[index];
    const animations = this.#animations[index];
    if (!slide || !animations) return;
    for (const fade of animations) {
      const target = slide.querySelector<HTMLElement>(`[data-pptx-shape-id="${fade.shapeId}"]`);
      if (!target) continue;
      const keyframes: Keyframe[] =
        fade.direction === 'in'
          ? [{ opacity: 0 }, { opacity: 1 }]
          : [{ opacity: 1 }, { opacity: 0 }];
      target.animate(keyframes, {
        duration: fade.durationMs,
        delay: fade.delayMs,
        easing: 'linear',
        fill: 'forwards',
      });
    }
  }

  #handleKeydown(event: KeyboardEvent): void {
    if (NEXT_KEYS.has(event.key)) {
      event.preventDefault();
      this.next();
    } else if (PREVIOUS_KEYS.has(event.key)) {
      event.preventDefault();
      this.previous();
    } else if (event.key === 'Home') {
      event.preventDefault();
      this.goToSlide(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      this.goToSlide(this.#slides.length - 1);
    }
  }
}

export function definePresentationElement(tagName = 'pptx-presentation'): void {
  if (!customElements.get(tagName)) {
    customElements.define(tagName, PptxPresentationElement);
  }
}
