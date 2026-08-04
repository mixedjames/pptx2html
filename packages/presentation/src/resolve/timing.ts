import type {
  CommonTimeNodeData,
  ExclusiveTimeNode,
  ParallelTimeNode,
  SequenceTimeNode,
  SlideTiming,
  TimeCondition,
  TimeNode,
  TimeNodeDuration,
} from '../presentationml/animation.js';
import type { TransitionSpeed } from '../presentationml/transition.js';

/**
 * OOXML's `TransitionSpeed` (§19.3.1.49) is qualitative — no spec-mandated millisecond value —
 * a deliberate, documented approximation, the same tier as this package's other best-effort
 * stand-ins for underspecified OOXML magnitudes. Moved here from `to-html5` (previously private
 * to `presentation-element.ts`) since any renderer driving a transition needs the identical
 * answer — see this package's own "resolution logic lives with the model" design decision.
 */
const TRANSITION_SPEED_DURATIONS_MS: Record<TransitionSpeed, number> = {
  fast: 400,
  med: 700,
  slow: 1000,
};

/**
 * `durationMs` (`SlideTransition.durationMs`, sourced from the `p14:dur` extension attribute) is
 * an explicit, spec-exact value when present, so it wins outright over `speed`'s three-tier
 * approximation rather than being blended with it — the same "first defined value wins outright"
 * pattern this package's other resolvers use for a whole-value fallback.
 */
export function resolveTransitionDurationMs(
  speed: TransitionSpeed | undefined,
  durationMs?: number,
): number {
  return durationMs ?? TRANSITION_SPEED_DURATIONS_MS[speed ?? 'fast']; // absent speed means "fast"
}

/**
 * The earliest a node with these start (or end) conditions could possibly fire, in milliseconds,
 * or 'indefinite' if every listed condition requires external input (a click, a mouseover, ...)
 * with no numeric delay fallback. Multiple conditions are OR'd (§19.7.4) — the node fires as soon
 * as any one of them is satisfied — so this is the *minimum* across all numeric delays; absent
 * entirely, a node starts immediately (0).
 */
function earliestMs(conditions: readonly TimeCondition[] | undefined): number | 'indefinite' {
  if (!conditions || conditions.length === 0) return 0;
  let earliest: number | 'indefinite' = 'indefinite';
  for (const condition of conditions) {
    const delay = conditionDelayMs(condition);
    if (delay !== 'indefinite' && (earliest === 'indefinite' || delay < earliest)) {
      earliest = delay;
    }
  }
  return earliest;
}

function conditionDelayMs(condition: TimeCondition): number | 'indefinite' {
  if (condition.delay === 'indefinite') return 'indefinite';
  if (condition.delay !== undefined) return condition.delay;
  // No explicit delay: an event-gated condition (onClick/onNext/...) waits for that event
  // indefinitely; a condition with neither a delay nor an event fires immediately.
  return condition.event !== undefined ? 'indefinite' : 0;
}

/**
 * The earliest a node could start playing on its own account, in milliseconds — 0 if it has no
 * start conditions at all, or 'indefinite' if every one needs external input (a click, a
 * mouseover, ...) with no numeric delay fallback (§19.7.4's OR semantics — the node fires as soon
 * as any one condition is satisfied). Deliberately doesn't account for a parent container's own
 * start offset the way `resolveTimeNodeDuration`'s internal `childContribution` does for
 * total-duration purposes — this answers a different, node-local question: can *this* node play
 * unattended, or is it waiting on interaction nothing here models yet? Exported for exactly that
 * use — see `@pptx2html/to-html5`'s fade-animation playback, which plays every fade node it finds
 * at its own local delay (ignoring ancestor click-gating/sequencing) rather than modeling
 * PowerPoint's full click-driven build system.
 */
export function resolveTimeNodeStartMs(node: TimeNode): number | 'indefinite' {
  return earliestMs(node.common.startConditions);
}

/**
 * Applies `common`'s repeat/auto-reverse adjustment to an already-computed base duration.
 * `repeatDuration` overrides outright (it's already an explicit total); any finite number of
 * repeats of an 'indefinite' base is still 'indefinite'; `autoReverse` is approximated as exactly
 * doubling one cycle (a forward pass plus a same-length reverse pass); `repeatCount` is in the
 * same 1000ths units `CommonTimeNodeData.repeatCount`'s own doc comment specifies (e.g. 5000
 * means 5 repeats), absent meaning exactly one (no repeat).
 */
function applyRepeat(base: TimeNodeDuration, common: CommonTimeNodeData): TimeNodeDuration {
  if (common.repeatDuration !== undefined) return common.repeatDuration;
  if (base === 'indefinite') return 'indefinite';
  const cycle = common.autoReverse ? base * 2 : base;
  if (common.repeatCount === 'indefinite') return 'indefinite';
  const repeats = common.repeatCount !== undefined ? common.repeatCount / 1000 : 1;
  return cycle * repeats;
}

/**
 * The node's own effective duration once started — excluding its *own* start delay, which a
 * parent container is responsible for adding in (see `childContribution`) — or 'indefinite' if it,
 * or any descendant, can't complete without external input (an `onClick`/`onNext`/... wait with no
 * numeric delay fallback, or an explicit `dur="indefinite"`/`repeatCount="indefinite"`). This is
 * exactly the "is this presentation fully time-resolved" question from
 * `docs/scroll-driven-playback.md`: 'indefinite' anywhere in a slide's tree means that slide (and
 * hence the whole deck) doesn't have a computable total duration.
 *
 * Approximated, not spec-exact: a container with no explicit `common.duration` derives its
 * duration from its children (`par`/concurrent-`seq` as the max across children, non-concurrent
 * `seq`/`excl` as their sum) rather than modeling `excl`'s real "at most one active" scheduling or
 * an interactive sequence's own advance conditions; a leaf behavior with no explicit `duration` is
 * treated as instantaneous (0), since OOXML gives no real default to fall back on.
 */
export function resolveTimeNodeDuration(node: TimeNode): TimeNodeDuration {
  if (node.common.duration !== undefined) return applyRepeat(node.common.duration, node.common);

  const own =
    node.kind === 'par' || node.kind === 'seq' || node.kind === 'excl'
      ? resolveContainerDuration(node)
      : 0;
  return applyRepeat(own, node.common);
}

function resolveContainerDuration(
  node: ParallelTimeNode | SequenceTimeNode | ExclusiveTimeNode,
): TimeNodeDuration {
  const contributions = node.children.map(childContribution);
  if (contributions.some((contribution) => contribution === 'indefinite')) return 'indefinite';
  const numeric = contributions as number[];
  if (numeric.length === 0) return 0;

  const parallel = node.kind === 'par' || (node.kind === 'seq' && node.concurrent === true);
  return parallel ? Math.max(...numeric) : numeric.reduce((sum, value) => sum + value, 0);
}

function childContribution(child: TimeNode): TimeNodeDuration {
  const start = earliestMs(child.common.startConditions);
  if (start === 'indefinite') return 'indefinite';
  const duration = resolveTimeNodeDuration(child);
  return duration === 'indefinite' ? 'indefinite' : start + duration;
}

/** A slide's own animation-tree duration, from its `timeNodeTree` — 0 if it has no timing tree at all. */
export function resolveSlideTimingDuration(timing: SlideTiming | undefined): TimeNodeDuration {
  return timing?.timeNodeTree ? resolveTimeNodeDuration(timing.timeNodeTree) : 0;
}
