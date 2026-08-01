import type { Color, ColorScheme, Fill, ResolvedColor } from '@pptx2html/presentation';
import {
  resolveColor as resolveColorComponents,
  resolveFillColor as resolveFillColorComponents,
} from '@pptx2html/presentation';

function rgbToCss(rgb: readonly [number, number, number], alpha: number): string {
  const [r, g, b] = rgb.map((c) => Math.round(c));
  return alpha < 1 ? `rgba(${r}, ${g}, ${b}, ${alpha})` : `rgb(${r}, ${g}, ${b})`;
}

/** Formats a `ResolvedColor` (see `@pptx2html/presentation`'s `resolve/color.ts`) as a CSS colour
 * string — the only step of colour resolution that's actually CSS-specific; everything else
 * (scheme aliasing, HSL conversion, colour-transform math) lives with the presentation model now,
 * since it's needed identically by any renderer. */
function toCss(resolved: ResolvedColor | undefined): string | undefined {
  if (!resolved) return undefined;
  return resolved.type === 'preset' ? resolved.value : rgbToCss(resolved.rgb, resolved.alpha);
}

/** Resolves a DrawingML colour (§20.1.2.3) to a CSS colour string, applying its transforms. */
export function resolveColor(color: Color, scheme: ColorScheme | undefined): string | undefined {
  return toCss(resolveColorComponents(color, scheme));
}

/**
 * Resolves a Fill to a single CSS colour string, for the common case of a solid fill (e.g. a run's
 * text colour). Gradient/pattern/blip fills need more than one CSS colour to render and are
 * unmodeled here — that's shape-fill rendering's job in a later pass, not this one.
 */
export function resolveFillColor(fill: Fill, scheme: ColorScheme | undefined): string | undefined {
  return toCss(resolveFillColorComponents(fill, scheme));
}
