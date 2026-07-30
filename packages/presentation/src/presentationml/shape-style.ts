import type { Color } from '../drawingml/index.js';

/**
 * A reference into the theme's format-scheme style matrix (§20.1.4.2.10/2.12, a:fillRef/a:lnRef),
 * by 1-based `index` — selects `FormatScheme.fillStyles`/`.lineStyles` (see `theme.ts`).
 * `effectRef`/`fontRef` share this same element shape (§20.1.4.2.8/2.9) but aren't modeled, since
 * nothing here renders effects or falls back to a style's font colour yet. `color` substitutes for
 * every `phClr` ("placeholder colour") scheme-colour reference inside the referenced Fill/Line —
 * resolving that substitution is a consumer's job (e.g. `@pptx2html/to-html5`), not this package's,
 * the same division of labour as `Placeholder`'s inherited-transform resolution.
 */
export interface StyleMatrixReference {
  readonly index: number;
  readonly color: Color;
}

/**
 * A shape's quick-style reference (§19.3.1.44, p:style — present on `p:sp`/`p:pic`/`p:cxnSp`).
 * PowerPoint's Shape Styles gallery writes shapes this way by default (a bare `fillRef`/`lnRef`,
 * no explicit `spPr` fill/line at all) — resolving these is necessary for such a shape to render
 * any fill/outline. `fillRef`/`lnRef` are the two consumed so far; `effectRef`/`fontRef` are
 * unmodeled, see `StyleMatrixReference`.
 */
export interface ShapeStyle {
  readonly fillRef?: StyleMatrixReference;
  readonly lineRef?: StyleMatrixReference;
}
