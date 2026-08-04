import type { SlideTiming, TimeNode } from '@pptx2html/presentation';
import { resolveTimeNodeDuration, resolveTimeNodeStartMs } from '@pptx2html/presentation';

/** One resolved fade animation this renderer knows how to play — see `collectFadeAnimations`. */
export interface ShapeFadeAnimation {
  readonly shapeId: number;
  readonly direction: 'in' | 'out';
  readonly delayMs: number;
  readonly durationMs: number;
}

type ReportUnsupported = (code: string, message: string) => void;

/**
 * Walks a slide's animation timing tree (`Slide.timing`, §19.3.1.48) collecting the one behavior
 * this renderer plays so far — a `p:animEffect` (§19.7.9) whose `filter` is `"fade"` and whose
 * target is a whole shape — and reports everything else it finds as unsupported via `report`:
 * other `animEffect` filters (wipe, blinds, ...) or non-shape targets, every other behavior kind
 * (`set`/`anim`/`animClr`/`animMotion`/`animRot`/`animScale`/`cmd`/`audio`/`video`), and an
 * implicit paragraph/graphic build (`buildList`).
 *
 * Deliberately **not** a faithful playback of PowerPoint's click-driven build system: a fade
 * node's own `delayMs` comes from `resolveTimeNodeStartMs` applied to *that node alone*, not
 * composed with its ancestor `par`/`seq` containers' own start offsets the way a real Animation
 * Pane sequence works (an "After Previous" effect nested three levels deep would, in PowerPoint,
 * wait for everything before it — here it just uses its own local delay, defaulting to 0). A node
 * gated on `onClick`/`onNext`/etc. with no numeric delay fallback still plays, immediately, rather
 * than waiting for that trigger — `to-html5` has no concept of an in-slide "build step" yet, only
 * slide-granular navigation (see `presentation-element.ts`), so waiting isn't an option; playing
 * immediately at least surfaces the content instead of leaving it invisible forever. Both
 * approximations are reported once per slide (`animation-trigger-unmodeled`) rather than
 * per-node, and are exactly the kind of gap a future scroll-driven-playback feature (root
 * `CLAUDE.md`'s design note) would need to resolve properly — this pass doesn't attempt that.
 */
export function collectFadeAnimations(
  timing: SlideTiming | undefined,
  report: ReportUnsupported,
): ShapeFadeAnimation[] {
  if (!timing) return [];

  const fades: ShapeFadeAnimation[] = [];
  const unsupportedBehaviorKinds = new Set<string>();
  let unsupportedEffectCount = 0;
  let triggerGatedCount = 0;

  function visit(node: TimeNode): void {
    switch (node.kind) {
      case 'par':
      case 'seq':
      case 'excl':
        for (const child of node.children) visit(child);
        return;
      case 'animEffect':
        if (node.filter === 'fade' && node.target?.kind === 'shape') {
          const duration = resolveTimeNodeDuration(node);
          if (duration === 'indefinite') {
            // A repeatCount/dur of "indefinite" has no finite animation to play at all.
            unsupportedEffectCount++;
            return;
          }
          const start = resolveTimeNodeStartMs(node);
          if (start === 'indefinite') triggerGatedCount++;
          fades.push({
            shapeId: node.target.shapeId,
            direction: node.transition ?? 'in', // §19.7.9's own default
            delayMs: start === 'indefinite' ? 0 : start,
            durationMs: duration,
          });
        } else {
          unsupportedEffectCount++;
        }
        return;
      default:
        unsupportedBehaviorKinds.add(node.kind);
    }
  }

  if (timing.timeNodeTree) visit(timing.timeNodeTree);

  if (unsupportedEffectCount > 0) {
    report(
      'animation-effect-unmodeled',
      `${unsupportedEffectCount} animation effect(s) other than a plain "fade" on a whole shape are not played.`,
    );
  }
  if (unsupportedBehaviorKinds.size > 0) {
    report(
      'animation-behavior-unmodeled',
      `Animation behavior kind(s) ${[...unsupportedBehaviorKinds].sort().join(', ')} are not played.`,
    );
  }
  if (triggerGatedCount > 0) {
    report(
      'animation-trigger-unmodeled',
      `${triggerGatedCount} fade animation(s) are gated on a click or other trigger this renderer doesn't wait for; they play immediately when the slide is shown instead.`,
    );
  }
  if (timing.buildList && timing.buildList.length > 0) {
    report(
      'animation-build-unmodeled',
      `${timing.buildList.length} paragraph/graphic build(s) are not staged; affected content shows all at once.`,
    );
  }

  return fades;
}
