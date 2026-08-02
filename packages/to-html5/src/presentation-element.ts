import type {
  Presentation,
  SideDirection,
  SlideTransition,
  TransitionSpeed,
} from '@pptx2html/presentation';
import { resolveTransitionDurationMs } from '@pptx2html/presentation';
import { renderSlide } from './slide.js';

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
 * concept of "backward" for this to describe itself (see `goToSlide` below). Every other
 * `TransitionEffect.kind` (wipe, cut, dissolve, wheel, ...) and a slide with no `transition` at
 * all fall back to today's instant `display: none`/`block` swap — see `#animatePush`/
 * `#animateFade` vs. `#updateActiveSlide`.
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

  render(presentation: Presentation): void {
    for (const animation of this.#currentAnimations) animation.cancel();
    this.#currentAnimations = [];
    this.#slides = presentation.slides.map((slide) =>
      renderSlide(this.ownerDocument, slide, presentation.slideSize, presentation.defaultTextStyle),
    );
    this.#transitions = presentation.slides.map((slide) => slide.transition);
    this.#slidesContainer.replaceChildren(...this.#slides);
    // Needed because a transition briefly makes both participating slides `position: absolute`
    // (see #beginTransitionFrame), at which point neither contributes to this container's
    // normal-flow height — without this it would collapse to zero height for the transition's
    // duration. Mirrors slide.ts's own per-slide aspect-ratio line.
    this.#slidesContainer.style.aspectRatio = `${presentation.slideSize.width} / ${presentation.slideSize.height}`;
    this.#currentIndex = 0;
    this.#updateActiveSlide();
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
    const outgoing = this.#slides[fromIndex]!;
    const incoming = this.#slides[toIndex]!;

    if (transition?.effect?.kind === 'push') {
      this.#animatePush(
        outgoing,
        incoming,
        transition.effect.direction ?? 'l',
        forward,
        transition.speed,
      );
    } else if (transition?.effect?.kind === 'fade') {
      this.#animateFade(outgoing, incoming, transition.speed);
    } else {
      this.#updateActiveSlide();
    }
  }

  /**
   * `direction` is always the transition's own authored direction — `goToSlide` already resolved
   * *which* slide's transition this is (destination going forward, outgoing going backward, see
   * above). Reversed here for backward navigation (`forward === false`) so the motion undoes the
   * original animation frame-for-frame, rather than repeating it.
   */
  #animatePush(
    outgoing: HTMLElement,
    incoming: HTMLElement,
    direction: SideDirection,
    forward: boolean,
    speed: TransitionSpeed | undefined,
  ): void {
    const effectiveDirection = forward ? direction : REVERSE_SIDE_DIRECTION[direction];
    const duration = resolveTransitionDurationMs(speed);
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

  #animateFade(
    outgoing: HTMLElement,
    incoming: HTMLElement,
    speed: TransitionSpeed | undefined,
  ): void {
    const duration = resolveTransitionDurationMs(speed);
    this.#beginTransitionFrame(outgoing, incoming);

    const options: KeyframeAnimationOptions = { duration, easing: 'ease-in-out', fill: 'forwards' };
    const outgoingAnimation = outgoing.animate([{ opacity: 1 }, { opacity: 0 }], options);
    const incomingAnimation = incoming.animate([{ opacity: 0 }, { opacity: 1 }], options);

    this.#awaitTransition(outgoing, incoming, [outgoingAnimation, incomingAnimation]);
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
