import type { Slide } from '@pptx2html/presentation';
import { resolveMorphMatch } from '@pptx2html/presentation';

/** One shape pair to interpolate between the outgoing and incoming slide of a Morph transition. */
export interface MorphShapeIdPair {
  readonly outgoingShapeId: number;
  readonly incomingShapeId: number;
}

/**
 * A Morph transition's shape correspondence, reduced to plain shape ids — never the `Slide`
 * object graph itself (`@pptx2html/presentation`'s `resolveMorphMatch` returns actual
 * `ShapeTreeNode` references, which this function immediately reduces away). This is
 * deliberate, not just a style choice: `PptxPresentationElement` stores exactly this instead of
 * the `Slide`s it was computed from, so the rendered `<pptx-presentation>` has no dependency on
 * `@pptx2html/presentation`'s object model once `render()` returns — the same property its other
 * retained per-slide state (`SlideTransition`, `ShapeFadeAnimation`) already has, and a real
 * requirement for a future freestanding-HTML-file export target, not a nice-to-have.
 */
export interface MorphMatchSummary {
  readonly matched: readonly MorphShapeIdPair[];
  readonly disappearingShapeIds: readonly number[];
  readonly appearingShapeIds: readonly number[];
}

type ReportUnsupported = (code: string, message: string) => void;

/**
 * How much of the larger side's shapes `resolveMorphMatch` needs to pair up before this renderer
 * trusts the result enough to actually play it, rather than falling back to a plain crossfade.
 * PowerPoint gives no equivalent threshold to match against — this heuristic doesn't exist in any
 * spec — chosen so a slide with mostly-matching shapes (the common case: duplicate-slide-then-
 * tweak) still morphs, while a slide sharing almost no shape names with its predecessor (the
 * name-matching heuristic having essentially nothing to go on) degrades gracefully instead of
 * animating an arbitrary, likely-nonsensical correspondence. A tunable, documented approximation,
 * the same tier as this project's other best-effort stand-ins for underspecified behaviour.
 */
const MIN_MORPH_MATCH_RATIO = 1 / 3;

/**
 * Resolves and reduces a Morph transition's shape correspondence for `slide`, matched against
 * `previousSlide` — the slide immediately before it in the deck, since a Morph transition
 * describes arriving at `slide` (§19.3.1.49) regardless of which direction the presentation
 * actually navigates; `presentation-element.ts` handles reversing the roles for backward
 * navigation itself, this function always matches forward-in-the-deck order.
 *
 * Returns `undefined` — meaning "fall back to a plain crossfade" — when there's no previous slide
 * to morph from at all (this slide is first), or when `resolveMorphMatch`'s match rate falls below
 * `MIN_MORPH_MATCH_RATIO`, reporting `morph-match-degraded` via `report` in both cases so deck
 * authors learn their Morph didn't play as authored, rather than getting an unexplained plain
 * crossfade with no indication anything was approximated — see root `CLAUDE.md`'s Todo 7 for why
 * this reporting matters as much as the animation itself.
 */
export function resolveSlideMorphMatch(
  previousSlide: Slide | undefined,
  slide: Slide,
  report: ReportUnsupported,
): MorphMatchSummary | undefined {
  if (!previousSlide) {
    report(
      'morph-match-degraded',
      'This slide is authored with a Morph transition but is the first slide, so there is ' +
        'nothing to morph from; falling back to a crossfade.',
    );
    return undefined;
  }

  const match = resolveMorphMatch(previousSlide, slide);
  const outgoingTotal = match.matched.length + match.disappearing.length;
  const incomingTotal = match.matched.length + match.appearing.length;
  const largerSide = Math.max(outgoingTotal, incomingTotal);
  const matchRatio = largerSide === 0 ? 0 : match.matched.length / largerSide;

  if (matchRatio < MIN_MORPH_MATCH_RATIO) {
    report(
      'morph-match-degraded',
      `Morph matched only ${match.matched.length} of up to ${largerSide} shape(s) between this ` +
        'slide and its predecessor; falling back to a crossfade instead of playing a ' +
        'low-confidence morph.',
    );
    return undefined;
  }

  return {
    matched: match.matched.map((pair) => ({
      outgoingShapeId: pair.outgoing.nonVisual.id,
      incomingShapeId: pair.incoming.nonVisual.id,
    })),
    disappearingShapeIds: match.disappearing.map((shape) => shape.nonVisual.id),
    appearingShapeIds: match.appearing.map((shape) => shape.nonVisual.id),
  };
}
