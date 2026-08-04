import type { SideDirection } from '@pptx2html/presentation';

/**
 * Pure keyframe/transform helpers shared by every renderer that plays a `push`/`fade`/`morph`
 * slide transition — today `presentation-element.ts` (real-time, click-driven) and
 * `scroll-presentation-element.ts` (scroll-driven). Extracted so both produce pixel-identical
 * visuals from one source instead of two copies drifting apart; nothing here is DOM-mutating
 * (`findShapeElement`/`readShapeBox` only read), so both callers can also use these for scrubbed
 * (not just fire-and-forget) playback.
 */

export const REVERSE_SIDE_DIRECTION: Record<SideDirection, SideDirection> = {
  l: 'r',
  r: 'l',
  u: 'd',
  d: 'u',
};

export const IDENTITY_TRANSFORM = 'translate(0, 0)';

/** The translate() a slide sits at when fully off-screen in `direction`, the way `push` uses it. */
export function offscreenTransform(direction: SideDirection): string {
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

/** A shape element's own resolved box, read straight off its inline style — see `readShapeBox`. */
export interface ShapeBox {
  readonly left: string;
  readonly top: string;
  readonly width: string;
  readonly height: string;
  readonly transform: string;
}

/** Finds a slide subtree's own element for `shapeId`, scoped to that slide. */
export function findShapeElement(slideEl: HTMLElement, shapeId: number): HTMLElement | null {
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
export function readShapeBox(el: HTMLElement): ShapeBox | undefined {
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
 * copy uses for a Morph transition — deliberately no `opacity` field at all, so the element stays
 * fully opaque for its entire journey. See `presentation-element.ts`'s `#animateMorph` doc comment
 * for why: crossfading two overlapping opaque copies (the technique a whole-slide fade correctly
 * uses) makes both visibly translucent for the middle of the transition, wrongly revealing
 * whatever sits behind them — wrong for a shape that's genuinely the same object, just moved/
 * resized/rotated.
 */
export function morphKeyframes(from: ShapeBox, to: ShapeBox): Keyframe[] {
  return [{ ...from }, { ...to }];
}
