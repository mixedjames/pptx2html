/** Returned by `observeContainSize` — see its own doc comment. */
export interface ContainSizeController {
  /**
   * Re-runs the sizing decision immediately, without waiting for the next incidental resize —
   * call whenever `target`'s own `aspect-ratio` might have just changed (e.g. a new `render()`
   * with a differently-sized deck), so the new deck's sizing doesn't lag a resize behind.
   */
  reapply(): void;
  /** Stops observing — call from `disconnectedCallback`. */
  disconnect(): void;
}

/**
 * Keeps `target` (which must have its own inline `aspect-ratio` already set by the caller) sized
 * to "contain" itself within `host`'s own box — width-driven (`width: 100%; height: auto`) by
 * default, switching to height-driven (`width: auto; height: 100%`) whenever the width-driven
 * default would actually overflow `host`'s own available height.
 *
 * **Why this needs JS at all, not just `max-width`/`max-height` in CSS.** A first attempt used
 * `max-width: 100%; max-height: 100%` on `target` with no explicit width/height of its own,
 * relying on `aspect-ratio` to fill in a size satisfying both caps — this is exactly how
 * `object-fit: contain` behaves, but `object-fit` only applies to *replaced* elements (`img`,
 * `video`, ...), not a plain `div`. For a non-replaced box, `max-width`/`max-height` are only
 * ever a ceiling on whatever size the normal sizing algorithm would otherwise produce — they
 * don't themselves supply a *preferred* size to size toward. With no explicit width/height at
 * all, a block box falls back to shrink-to-fit sizing based on its content's own intrinsic
 * size — and `target`'s content (`.pptx-slide`, itself `width: 100%` of `target`) is entirely
 * percentage-based, which contributes nothing (effectively zero) to a shrink-to-fit computation,
 * since percentages can't resolve without already knowing the size they're a percentage *of*.
 * Net effect: `target` collapsed to near-zero size instead of filling the available space — an
 * actual regression caught after shipping, not a hypothetical.
 *
 * **Why measured overflow, not a numeric aspect-ratio comparison computed ahead of time.**
 * Comparing `host`'s own width/height ratio against the deck's aspect ratio numerically would
 * need to know `host`'s *available* height up front — but `host` (the custom element itself)
 * commonly has no independent height at all: the traditional embedding (just dropped into a
 * page's own flow, no fixed box) relies on `host`'s height being *derived from* `target`'s own
 * content, not the other way around. Measuring `target`'s actual rendered height under the
 * width-driven baseline and comparing it against `host.clientHeight` self-adapts to both cases
 * without needing to know *why* `host` does or doesn't have an independent height: when `host`'s
 * height is itself content-derived, it always exactly equals `target`'s own height (never
 * "overflowing" by definition), so this never switches modes and behaves exactly as the simple
 * width-only version always did; only a `host` with a genuinely independent, smaller height
 * (imposed by the page, e.g. a full-viewport demo layout) can ever trigger the switch.
 */
export function observeContainSize(host: Element, target: HTMLElement): ContainSizeController {
  const apply = (): void => {
    target.style.width = '100%';
    target.style.height = 'auto';
    const availableHeight = host.clientHeight;
    if (availableHeight > 0 && target.getBoundingClientRect().height > availableHeight) {
      target.style.width = 'auto';
      target.style.height = '100%';
    }
  };

  const observer = new ResizeObserver(apply);
  observer.observe(host);
  apply();

  return { reapply: apply, disconnect: () => observer.disconnect() };
}
