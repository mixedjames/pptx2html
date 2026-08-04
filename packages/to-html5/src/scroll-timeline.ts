import type { Presentation, TransitionEffect } from '@pptx2html/presentation';
import { resolveTransitionDurationMs } from '@pptx2html/presentation';
import { collectFadeAnimations, type ShapeFadeAnimation } from './animation.js';
import { resolveSlideMorphMatch, type MorphMatchSummary } from './morph.js';

type ReportUnsupported = (code: string, message: string, slideIndex: number) => void;

/** How long a slide's transition-in plays, in absolute scroll-timeline milliseconds. */
export interface ScrollTransitionSegment {
  readonly startMs: number;
  readonly endMs: number;
  readonly effect: TransitionEffect;
  /** Present only for a `morph` effect with a confident-enough shape match — see `morph.ts`. */
  readonly morphMatch?: MorphMatchSummary;
}

/** How long a slide sits fully visible, in absolute scroll-timeline milliseconds. */
export interface ScrollContentSegment {
  readonly startMs: number;
  readonly endMs: number;
  readonly fades: readonly ShapeFadeAnimation[];
}

/** One slide's place on the scroll timeline — see `resolveScrollTimeline`. */
export interface ScrollSegment {
  readonly slideIndex: number;
  /** Absent only for the first slide, which is simply shown at ms 0 — see below. */
  readonly transition?: ScrollTransitionSegment;
  readonly content: ScrollContentSegment;
}

export interface ScrollTimeline {
  readonly segments: readonly ScrollSegment[];
  readonly totalDurationMs: number;
}

/**
 * A slide with no build animations at all (or one whose builds finish quickly) still needs some
 * scroll distance to sit in front of the user before the next slide's transition starts — without
 * a floor, a fully static slide would collapse to zero scroll-timeline length and effectively
 * disappear. Not spec-derived — a documented approximation, same tier as this package's other
 * best-effort magnitude defaults (`shape-geometry.ts`'s adjustment-guide defaults, `fill.ts`'s
 * pattern-hatch spacing).
 */
const DEFAULT_MIN_DWELL_MS = 1200;

/**
 * Scroll mode's own fallback when a slide has no `transition` at all, or one whose `effect.kind`
 * isn't among the three this project actually animates (`push`/`fade`/`morph`) — a **different**
 * fallback than real-time playback's instant `display: none`/`block` swap
 * (`presentation-element.ts`'s `#updateActiveSlide`). An abrupt cut reads badly mid-scrub, and a
 * vertical push matches scrollytelling's own mental model directly: scrolling down brings the next
 * slide up into view. Every other real fallback (a low-confidence Morph match, still degrading to
 * a plain crossfade — see `morph.ts`) is left exactly as real-time mode already resolves it.
 */
const DEFAULT_SCROLL_TRANSITION_EFFECT: TransitionEffect = { kind: 'push', direction: 'u' };

/**
 * Assembles a `Presentation`'s entire deck — every slide's transition-in plus its own build
 * animations — into one absolute-millisecond scroll timeline. Reuses the exact same duration/start
 * resolution this package's real-time renderer already uses (`resolveTransitionDurationMs`,
 * `collectFadeAnimations`, `resolveSlideMorphMatch`) rather than inventing new resolution logic —
 * see `packages/to-html5/CLAUDE.md`'s scroll-mode design decision for why no new resolvers were
 * needed in `@pptx2html/presentation` for this.
 *
 * **Deliberately has no "is this deck scroll-eligible" gate.** A deck's authored
 * `advanceOnClick`/`advanceAfter` and any `onClick`/`onNext` build gating are simply not consulted
 * here at all — reaching a point on the scroll axis *is* the trigger, the same way a click is in
 * real-time mode, so click-gating that would make a deck "not fully time-resolved" for a
 * hypothetical auto-playing kiosk mode is irrelevant to whether scroll mode can render it. Every
 * slide plays; whatever this renderer can't faithfully animate is reported via `report` (an
 * unsupported effect kind, an unplayed animation behavior/build, ...) using the exact same codes
 * `presentation-element.ts`'s `reportSlideLevelFeatures`/`collectFadeAnimations` already use, plus
 * one new code (`transition-effect-approximated-for-scroll`) for the push-up substitution above,
 * which isn't really "unsupported" so much as "approximated differently than real-time mode."
 *
 * The first slide never gets a `transition` segment, matching real-time playback's own behavior —
 * `goToSlide` is the only place a transition ever plays there too, and the initial slide is simply
 * shown, never navigated into.
 */
export function resolveScrollTimeline(
  presentation: Presentation,
  report: ReportUnsupported,
  options?: { readonly minDwellMs?: number },
): ScrollTimeline {
  const minDwellMs = options?.minDwellMs ?? DEFAULT_MIN_DWELL_MS;
  const segments: ScrollSegment[] = [];
  let cursorMs = 0;

  presentation.slides.forEach((slide, slideIndex) => {
    const localReport = (code: string, message: string) => report(code, message, slideIndex);

    let transition: ScrollTransitionSegment | undefined;
    if (slideIndex > 0) {
      const previousSlide = presentation.slides[slideIndex - 1];
      const authoredTransition = slide.transition;
      const authoredEffect = authoredTransition?.effect;
      const morphMatch =
        authoredEffect?.kind === 'morph'
          ? resolveSlideMorphMatch(previousSlide, slide, localReport)
          : undefined;

      let effect: TransitionEffect;
      if (authoredEffect?.kind === 'push' || authoredEffect?.kind === 'fade') {
        effect = authoredEffect;
      } else if (authoredEffect?.kind === 'morph' && morphMatch) {
        effect = authoredEffect;
      } else if (authoredEffect?.kind === 'morph') {
        // resolveSlideMorphMatch already reported morph-match-degraded — same fallback real-time
        // mode's #animateMorph uses when it has no confident match: a plain crossfade.
        effect = { kind: 'fade' };
      } else {
        effect = DEFAULT_SCROLL_TRANSITION_EFFECT;
        if (authoredEffect) {
          localReport(
            'transition-effect-approximated-for-scroll',
            `Slide transition effect "${authoredEffect.kind}" is not animated in scroll mode; substituting a vertical push.`,
          );
        }
      }

      const durationMs = resolveTransitionDurationMs(
        authoredTransition?.speed,
        authoredTransition?.durationMs,
      );
      transition = { startMs: cursorMs, endMs: cursorMs + durationMs, effect, morphMatch };
      cursorMs += durationMs;
    }

    const fades = collectFadeAnimations(slide.timing, localReport);
    const contentDurationMs = Math.max(minDwellMs, ...fades.map((f) => f.delayMs + f.durationMs));
    const content = { startMs: cursorMs, endMs: cursorMs + contentDurationMs, fades };
    cursorMs += contentDurationMs;

    segments.push({ slideIndex, transition, content });
  });

  return { segments, totalDurationMs: cursorMs };
}
