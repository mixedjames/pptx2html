import type { Fill } from '../drawingml/index.js';
import type { CommonSlideData, Slide } from '../presentationml/index.js';
import type { FormatScheme } from '../theme.js';
import { resolveBackgroundStyleFill } from './style-matrix.js';

/**
 * A slide-like part's own background (§19.3.1.6), resolved to a final `Fill` — its literal
 * `p:bgPr` fill if it has one, else its `p:bgRef` style-matrix reference resolved against the
 * theme, else `undefined` if this part defines no background of its own at all. Both are
 * "this part's own" checks, not the slide → layout → master fallback itself (`resolveEffectiveBackground`
 * below composes this per part, in order) — see `CommonSlideData.backgroundRef`'s own doc comment
 * for why a part's background is two independent optional fields rather than one.
 */
function ownBackgroundFill(
  part: CommonSlideData,
  formatScheme: FormatScheme | undefined,
): Fill | undefined {
  return part.background?.fill ?? resolveBackgroundStyleFill(part.backgroundRef, formatScheme);
}

/**
 * Resolves a slide's effective background (§19.3.1.7, p:bg) to a final `Fill`: the slide's own,
 * falling back to its layout's, falling back to the layout's master's. Simpler than placeholder
 * transform/font inheritance (`placeholder.ts`, `text-style.ts`) — there's no per-shape matching
 * step, just "does this slide-like part define one at all", so the first part found in the chain
 * wins outright rather than being merged field-by-field — and, new, "does this part define one at
 * all" now checks *either* of a part's two background fields (`ownBackgroundFill`), not just the
 * literal one: a part whose only background is a `p:bgRef` (PowerPoint's own default whenever a
 * slide/layout/master relies on the theme's background rather than an explicit fill — the common
 * case, not an edge case) used to resolve as if that part had no background at all, silently
 * falling through to the next part in the chain (or `undefined`) instead — see this package's own
 * CLAUDE.md for the real deck that caught this. `formatScheme` (needed only for the `bgRef` case)
 * comes from the slide's own theme (`slide.layout.master.theme.formatScheme`), not a caller-supplied
 * parameter — every part in the chain shares one theme, so there's no ambiguity about whose to use.
 */
export function resolveEffectiveBackground(slide: Slide): Fill | undefined {
  const formatScheme = slide.layout.master.theme.formatScheme;
  return (
    ownBackgroundFill(slide.commonSlideData, formatScheme) ??
    ownBackgroundFill(slide.layout.commonSlideData, formatScheme) ??
    ownBackgroundFill(slide.layout.master.commonSlideData, formatScheme)
  );
}
