import type { Color } from '../drawingml/index.js';

/**
 * A reference into the theme's format-scheme style matrix (§20.1.4.2.10/2.12, a:fillRef/a:lnRef),
 * by 1-based `index` — selects `FormatScheme.fillStyles`/`.lineStyles` (see `theme.ts`).
 * `effectRef` shares this same element shape (§20.1.4.2.8) but isn't modeled, since nothing here
 * renders effects yet. `color` substitutes for every `phClr` ("placeholder colour") scheme-colour
 * reference inside the referenced Fill/Line — resolving that substitution is a consumer's job
 * (e.g. `@pptx2html/to-html5`), not this package's, the same division of labour as `Placeholder`'s
 * inherited-transform resolution.
 */
export interface StyleMatrixReference {
  readonly index: number;
  readonly color: Color;
}

/** Which of the theme's font collections a `FontReference` falls back to (§20.1.10.20, ST_FontCollectionIndex). */
export type FontCollectionIndex = 'major' | 'minor' | 'none';

/**
 * A reference to the theme's font scheme plus a fallback colour (§20.1.4.1.17, a:fontRef) — a
 * shape's default *text* formatting when a run has no colour/typeface of its own anywhere in its
 * own inheritance chain (unlike `fillRef`/`lnRef`, this isn't a style-matrix index; `idx` instead
 * picks the theme's major or minor font collection directly, or `none` for no font fallback at
 * all). `color` is the same `phClr`-substituted scheme colour `StyleMatrixReference` carries.
 */
export interface FontReference {
  readonly collection: FontCollectionIndex;
  readonly color: Color;
}

/**
 * A shape's quick-style reference (§19.3.1.44, p:style — present on `p:sp`/`p:pic`/`p:cxnSp`).
 * PowerPoint's Shape Styles gallery writes shapes this way by default (a bare `fillRef`/`lnRef`/
 * `fontRef`, no explicit `spPr` fill/line or run-level colour/typeface at all) — resolving these
 * is necessary for such a shape's fill/outline/text to render as authored. `fillRef`/`lnRef`/
 * `fontRef` are the three consumed so far; `effectRef` is unmodeled, see `StyleMatrixReference`.
 */
export interface ShapeStyle {
  readonly fillRef?: StyleMatrixReference;
  readonly lineRef?: StyleMatrixReference;
  readonly fontRef?: FontReference;
}
