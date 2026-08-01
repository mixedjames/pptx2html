import type { Color, ColorTransform, Fill, Line } from '../drawingml/index.js';
import type { FormatScheme } from '../theme.js';
import type { StyleMatrixReference } from '../presentationml/index.js';

function mergeTransforms(
  base: ColorTransform | undefined,
  override: ColorTransform | undefined,
): ColorTransform | undefined {
  return base || override ? { ...base, ...override } : undefined;
}

/**
 * Substitutes every `a:schemeClr val="phClr"` placeholder inside a Color with `replacement`
 * (§20.1.2.3.31) — merging the placeholder's own local colour transforms (e.g. the theme's
 * fillStyleLst commonly stacks `lumMod`/`tint`/etc. on its `phClr` entries) underneath whatever
 * transforms `replacement` itself carries (rare in practice — a `fillRef`'s own `schemeClr` almost
 * never has one). Not a fully spec-accurate composition of two ordered transform stacks
 * (`ColorTransform` is a flat field set, not an ordered list — see `color.ts`'s own doc comment for
 * the same caveat elsewhere), but correct for the overwhelming majority of real decks.
 */
function substitutePlaceholder(color: Color, replacement: Color): Color {
  if (color.type !== 'scheme' || color.value !== 'phClr') return color;
  const transforms = mergeTransforms(color.transforms, replacement.transforms);
  return { ...replacement, ...(transforms ? { transforms } : {}) };
}

function substituteFill(fill: Fill, replacement: Color): Fill {
  switch (fill.type) {
    case 'solid':
      return { ...fill, color: substitutePlaceholder(fill.color, replacement) };
    case 'gradient':
      return {
        ...fill,
        stops: fill.stops.map((stop) => ({
          ...stop,
          color: substitutePlaceholder(stop.color, replacement),
        })),
      };
    case 'pattern':
      return {
        ...fill,
        foregroundColor: substitutePlaceholder(fill.foregroundColor, replacement),
        backgroundColor: substitutePlaceholder(fill.backgroundColor, replacement),
      };
    case 'none':
    case 'blip':
      return fill;
  }
}

/**
 * Resolves a shape's `p:style/fillRef` (§20.1.4.2.10) against the theme's format-scheme fill style
 * matrix — the 1-based `index` selects `FormatScheme.fillStyles`, and the reference's own `color`
 * substitutes for that style's `phClr` placeholder(s), see `substitutePlaceholder`. `undefined` if
 * there's no reference, no theme, or the index is out of range — a renderer falls back to the
 * shape's own `spPr` fill either way, so this only matters for a shape that omits one (PowerPoint's
 * Shape Styles gallery writes shapes exactly this way).
 */
export function resolveStyleFill(
  ref: StyleMatrixReference | undefined,
  formatScheme: FormatScheme | undefined,
): Fill | undefined {
  const style = ref && formatScheme?.fillStyles[ref.index - 1];
  return style ? substituteFill(style, ref.color) : undefined;
}

/** `resolveStyleFill`'s `p:style/lnRef` (§20.1.4.2.12) equivalent, against `FormatScheme.lineStyles`. */
export function resolveStyleLine(
  ref: StyleMatrixReference | undefined,
  formatScheme: FormatScheme | undefined,
): Line | undefined {
  const style = ref && formatScheme?.lineStyles[ref.index - 1];
  if (!style) return undefined;
  return style.fill ? { ...style, fill: substituteFill(style.fill, ref.color) } : style;
}
