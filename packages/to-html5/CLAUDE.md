# @pptx2html/to-html5

Renders a `@pptx2html/presentation` object graph (the in-memory `.pptx` DOM, produced by
`@pptx2html/reader`) into an actual HTML5 DOM. Public API is a single synchronous function,
mirroring the reader's `readPresentation`:

```ts
import { renderPresentation } from '@pptx2html/to-html5';
const el = renderPresentation(presentation); // <pptx-presentation>, shadow DOM inside
document.body.appendChild(el);
```

## Status: layout, plus formatting passes for fonts, shape fill/line, and slide backgrounds

Every slide and shape lands in the right place at the right size, including placeholder shapes
that inherit their position from the slide layout/master rather than declaring their own (very
common in real decks — see `placeholder.ts`). Run-level character formatting (typeface, size,
bold, italic, underline, strikethrough, text colour) renders too, fully resolved through
OOXML's text-property inheritance chain — run → paragraph → shape → placeholder layout/master →
master's title/body/other style → presentation default → theme font scheme (see "Key design
decision: font inheritance chain" below). Shape/picture `fill`/`.line` (§20.1.2.2.35's spPr) also
render, as CSS background/border (`fill.ts`) — solid fills, linear gradients, and (approximated)
patterns/images for fill; solid-colored, dashed/dotted/double borders for line. A slide's own
background renders too, falling back through layout then master (`background.ts`), reusing the
same `fill.ts` machinery. Paragraph alignment and table cell/table styling are still unrendered —
deliberately deferred to a later pass. See "Scope boundary" below for the precise line.

## Layout

- `units.ts` — `EMU_PER_PX` (9525, i.e. 96 CSS px/inch) and `emuToPx`. Exported for consumers;
  no longer used internally by the positioning code below (see the percentage-based-layout
  design decision).
- `coordinate.ts` — pure math, no DOM, fully unit-testable in Node: `CoordinateMap` (an affine
  translate+scale from a local EMU space to the slide's root EMU space), `IDENTITY_MAP`,
  `composeGroupMap` (extends a map with a group's own transform so its children's transforms —
  expressed against `chOff`/`chExt`, §20.1.7.6 — resolve correctly), `computeBox` (map + a
  `Transform2D` → slide-relative offset/extents, **still in EMU**, plus rotation/flip — converting
  to a CSS unit is left to the caller, see below).
- `text.ts` / `table.ts` — pure content renderers (`TextBody` → `<div><p>…</p></div>`,
  `Table` → `<table>`), no absolute positioning of their own; a paragraph is `<p>`, a run or field
  is a `<span class="pptx-run">` (text content plus its resolved inline font styling — see below),
  a break is `<br>`. `renderTable` sizes itself to `100%`/`100%` of its containing `graphicFrame`
  div, `table-layout: fixed`, with `<col>` widths and `<tr>` heights expressed as a percentage of
  the column-width/row-height totals (so column proportions survive scaling; row-height
  percentages are best-effort — browsers that ignore them just size rows by content instead).
- `text-style.ts` — pure logic, no DOM: `resolveEffectiveRunProperties` walks the run-property
  inheritance chain (see the design decision below) and returns one merged `RunProperties`;
  `resolveTypeface` resolves a `+mj-lt`/`+mn-lt`/etc. theme font token against a `FontScheme`,
  passing any literal font name straight through.
- `color.ts` — pure logic, no DOM: `resolveColor` turns a DrawingML `Color` (srgb/scheme/system/
  preset/hsl, plus its `ColorTransform` modifiers — lumMod/lumOff/shade/tint/satMod/hueMod/alpha)
  into a CSS colour string; `resolveFillColor` is the `Fill`-level wrapper for the common
  single-colour case (only `SolidFill` resolves to a colour — gradient/pattern/blip fills need
  more than one CSS colour), used for both run text colour (`text.ts`) and a shape's `line`
  colour (`fill.ts`). Scheme colour resolution assumes the default clrMap (bg1→lt1, tx1→dk1,
  bg2→lt2, tx2→dk2), since slide/layout clrMap overrides are unmodeled in `packages/presentation`.
- `fill.ts` — DOM-touching (unlike `color.ts`): `applyFill` sets a shape/picture's CSS
  `background-color`/`background-image` from its `ShapeProperties.fill`, and `applyLine` sets
  `border-width`/`-style`/`-color` from its `ShapeProperties.line`. `resolveGradientCss` (the one
  pure piece, exported and separately tested) converts a `GradientFill` to a CSS
  `linear-gradient(...)`, including translating DrawingML's shade-path angle convention (clockwise
  from east) to CSS's (clockwise from north). A `blip` fill creates an object URL exactly like
  `renderPicture` does for a picture's own image — same never-revoked caveat, see below.
- `placeholder.ts` — pure logic, no DOM: `resolveInheritedTransform` walks the OOXML placeholder
  inheritance chain (§19.3.1.36) — slide placeholder → matching layout placeholder (by
  type+index, falling back to index-only then type-only) → matching master placeholder, same
  rule — and returns the first transform found. `shape`/`picture`/`connector` almost always rely
  on this: real decks routinely omit `xfrm` on placeholder shapes entirely.
- `render-context.ts` — `RenderContext { slideSize, layout, defaultTextStyle }`, threaded
  alongside `CoordinateMap` through every `renderShapeTreeNode` call. Unlike the map (which
  changes at every group nesting level), this is constant for the whole slide — `slideSize` is
  what every element's EMU box gets divided by to become a percentage, `layout` is for
  placeholder transform _and_ font resolution (via `layout.master`), `defaultTextStyle` is the
  presentation's own `p:defaultTextStyle`, the bottom rung of the font inheritance chain.
- `shape-tree.ts` — `renderShapeTreeNode`, the one recursive dispatcher over `ShapeTreeNode`'s
  five kinds. `shape`/`picture`/`connector` position from their own transform if they have one,
  else from `resolveInheritedTransform`; `graphicFrame` always has its own transform (mandatory in
  the schema) so it skips that path. All four get positioned by `positionElement`, which converts
  `computeBox`'s EMU box to `left`/`top`/`width`/`height` **percentages of `context.slideSize`**.
  `renderShape`/`renderPicture` also call `fill.ts`'s `applyFill`/`applyLine` with their own
  `ShapeProperties.fill`/`.line` (a picture's fill shows through any transparent pixels in the
  image itself, since `Picture` shares `ShapeProperties` with `Shape`) — `renderConnector` does
  not yet, see the scope boundary below. `group` renders as an anchor `<div>` stretched to exactly
  cover the slide (`left/top/right/bottom:
0`, i.e. 100%×100% of its own containing block) — its own transform only feeds
  `composeGroupMap` for its children, it draws nothing itself, and groups never carry a
  placeholder (no `nvPr`/`ph` in the schema for `grpSp`). Percentages need a well-defined
  containing block to resolve against; stretching every nesting level to exactly the slide's size
  is what makes a doubly-nested child's percentage position still resolve correctly (see the
  design decision below).
- `background.ts` — pure logic, no DOM: `resolveEffectiveBackground` picks a slide's own
  background if it has one, else its layout's, else the layout's master's — the first one found
  wins outright (unlike the font/transform chains above, a background isn't merged field-by-field,
  since `Background` is just a single `Fill`).
- `slide.ts` — `renderSlide`: one `.pptx-slide` div, `position: relative; overflow: hidden`,
  `width: 100%` with `aspect-ratio: <slideWidth> / <slideHeight>` (raw EMU numbers — `aspect-ratio`
  only cares about the ratio) so it fills whatever width its container gives it and its height
  follows automatically. Applies `resolveEffectiveBackground`'s result via `fill.ts`'s `applyFill`
  directly on the slide div, before its `shapeTree` (walked from `IDENTITY_MAP` with a
  `RenderContext` built from `slide.layout`, the passed-in `slideSize`, and an optional
  `defaultTextStyle` — the presentation's `p:defaultTextStyle`, passed down from
  `presentation-element.ts` since `Presentation` itself isn't otherwise threaded this deep).
- `presentation-element.ts` — `PptxPresentationElement`, a `<pptx-presentation>` custom element.
  Shadow DOM is attached in the constructor; `.render(presentation)` replaces its slide children.
  `definePresentationElement()` registers the tag (idempotent).
- `index.ts` — barrel + `renderPresentation`, which registers the element and returns one
  instance with `.render()` already called.

## Key design decision: every element's position is a CSS percentage of the slide, not a px value

Requirement: slides scale to fill the width of whatever container `<pptx-presentation>` sits in,
without JS resize handling. We get this for free from the CSS box model rather than reaching for
a `ResizeObserver` + recomputed `transform: scale(...)`: every element's `left`/`top`/`width`/
`height` is written as `emuValue / slideSize.{width,height} * 100` + `%`, and `.pptx-slide` itself
is `width: 100%` with `aspect-ratio` locking its height — so resizing the slide's container just
changes what those percentages resolve to; the browser recomputes layout on every resize with no
listener code on our side. This is also why `computeBox` deliberately stopped converting to px
(`ElementBox` is now plain EMU) — percentage math needs the _ratio_, so unit conversion would only
add and then cancel out an `EMU_PER_PX` factor for no benefit.

The one wrinkle this creates: CSS percentages for an absolutely positioned element resolve against
its _containing block_ (nearest positioned ancestor), not the slide root — so a shape nested two
groups deep, whose percentage was computed relative to the slide's total size, needs its immediate
parent (the inner group's wrapper `<div>`) to _also_ be exactly the slide's size for that
percentage to land in the right place. That's why `renderGroup`'s wrapper is stretched with
`left/top/right/bottom: 0` (100%×100% of its own containing block) instead of being left
auto-sized — every group wrapper in a nesting chain ends up pixel-for-pixel the same size as the
slide root, so a percentage computed relative to the slide root resolves identically no matter
which wrapper it's actually nested inside.

## Key design decision: positions computed in JS via CoordinateMap, not composed via nested CSS transforms

A shape's position could instead be built from nested `<div>`s each carrying a CSS
`transform: translate(...) scale(...)` mirroring its OOXML group transform, rather than the
percentage scheme above. We didn't do that — `composeGroupMap` walks the ancestor chain in EMU
space and every leaf element gets one final `left`/`top`/`width`/`height` (as a percentage, see
above), already relative to the slide. This makes each element's box independently readable from
its own inline styles (useful for debugging and for a future formatting pass) without having to
account for inherited CSS transforms. The one thing this trades away: **rotation does not compose
across nested groups** — `computeBox` applies a shape's (or group's) own `rotation`/flip as a CSS
`transform` around its own box, uncomposed with any ancestor group's rotation. Spec-correct for
the (common) unrotated-group case; a rotated group with rotated descendants will look wrong.
Fixing this means switching the rotated subtree to the nested-CSS-transform approach above —
deliberately not done for this first pass.

## Key design decision: font inheritance chain resolved eagerly per run, not via CSS cascade

A run's effective font could instead lean on the browser's own CSS cascade — apply each level's
style to its own DOM ancestor and let `inherit`/unset properties flow down naturally. We didn't do
that: `text-style.ts`'s `resolveEffectiveRunProperties` walks the whole chain in JS for every run
and writes the fully-merged result as that run's own inline styles (mirrors
`resolveInheritedTransform`'s approach to placeholder position, and the same "every element's box
independently readable from its own inline styles" rationale as the CoordinateMap decision above).
The chain, lowest to highest priority: `Presentation.defaultTextStyle` → the slide master's
title/body/other `TextStyles` (by placeholder category — title/ctrTitle use `titleStyle`, other
placeholder types `bodyStyle`, non-placeholder shapes `otherStyle`) → the master's own matching
placeholder shape's `TextBody.listStyle` → the layout's matching placeholder shape's
`TextBody.listStyle` (mirrors `resolveInheritedTransform`'s layout→master walk, via the same
`findPlaceholderMatch`, now exported from `placeholder.ts` for this reuse) → the shape's own
`TextBody.listStyle` → the paragraph's own `defRPr` → the run's own `rPr`. Each step only supplies
defaults for the paragraph's own outline level (0-based, `TextListStyle.levels[level]`). This order
is a reasonable synthesis of how real-world OOXML renderers describe the (spec-under-specified)
inheritance, not a verified bit-exact match to PowerPoint. Theme font-scheme resolution
(`+mj-lt`/`+mn-lt`/etc. → an actual typeface) is a separate, final step (`resolveTypeface`) since
it only concerns the `typeface` field and needs the theme, not a list style — done in `text.ts`
right before the style is applied, using `context.layout?.master.theme`.

## Scope boundary — what's intentionally unmodeled (yet)

- **Run-level font colour only resolves solid fills.** `RunProperties.fill` can technically be a
  gradient/pattern/blip `Fill` (WordArt-style text); `color.ts`'s `resolveFillColor` only handles
  `SolidFill`, falling back to no colour (inherits black) for the others — vanishingly rare on
  plain text runs in real decks.
- **Colour transforms on preset colours are ignored.** `color.ts` has no RGB table for the ~140
  DrawingML preset names (they map almost 1:1 to CSS's extended colour keywords, so untransformed
  presets pass straight through as CSS keywords) — a `shade`/`tint`/etc. stacked on a preset colour
  is silently dropped. Transforms are almost always applied to scheme colours in practice, not
  presets, so this is a narrow gap.
- **`ShapeProperties.geometry` is not rendered.** Every shape/picture is still a plain rectangular
  box regardless of its preset geometry (`ellipse`, `roundRect`, `triangle`, etc.) or (unmodeled)
  custom path — `fill`/`.line` paint that rectangle, not the shape's actual outline. Fixing this
  needs either CSS `clip-path` per preset (and real path data for `custGeom`, still unmodeled in
  `packages/presentation`) or an SVG-based shape renderer; deliberately out of scope for this pass.
- **Pattern fills are only approximated**, and colour transforms on preset colours are ignored —
  see `fill.ts`'s and `color.ts`'s own doc comments for why.
- **Remaining visual formatting**: paragraph alignment (`ParagraphProperties.alignment`),
  `TextBodyProperties.anchor`/`wrap`, table cell fill, table styles. The DOM structure exists
  (`.pptx-shape`, `.pptx-paragraph`, `.pptx-run`, etc.) precisely so a later pass can add CSS
  without restructuring.
- **Placeholder inheritance is supported, but with a simplified matching rule.** See
  `placeholder.ts` above. Not modeled: the spec's type-equivalence groups (e.g. a slide's
  `ctrTitle` placeholder is allowed to match a layout's `title` placeholder) — only exact type
  matches (after the index-match attempts) are tried. Real decks reliably reuse the same
  placeholder type across slide/layout/master, so this covers the common case; a shape can still
  end up unpositioned (static flow) if nothing in the chain matches or has a transform.
- **Charts/SmartArt/OLE objects.** `GraphicPlaceholder` only carries `type`; rendered as a
  `[chart]`/`[smartArt]`/`[oleObject]` text placeholder box.
- **Connectors still render as an unstyled positioned `<div>` (no line drawn)**, even though
  `applyLine` (which could paint their `ShapeProperties.line`) now exists — the reason is no
  longer "the formatting pass doesn't exist yet" but that a connector's visible line isn't its
  bounding box's border: a straight connector's line runs along the box's diagonal, and bent/curved
  connectors need real path geometry (unmodeled, see above). Drawing this properly needs an
  SVG (or similar) line/path overlay, not `border-*`; left for a dedicated pass.
- **Object URLs are never revoked.** `renderPicture` calls `URL.createObjectURL` per picture, and
  `fill.ts`'s `applyFill` does the same for any `blip` shape/picture fill — neither ever calls
  `revokeObjectURL`. Fine for a one-shot render; calling `.render()` repeatedly on the same
  `PptxPresentationElement` (or rendering many presentations in one page) will leak blob URLs.
  Deferred — would need the created URLs plumbed back up to `PptxPresentationElement` so it can
  revoke on re-render/disconnect.

## Tests

`coordinate.test.ts`, `placeholder.test.ts`, `text-style.test.ts`, `color.test.ts` and
`background.test.ts` are pure Node (no DOM) and cover the affine math, the placeholder transform
inheritance chain (layout match, master fallback, no-match), the font-property inheritance chain
(every rung from `defaultTextStyle` down to a run's own `rPr`, plus theme font-token resolution),
DrawingML colour resolution (scheme aliasing, hsl/srgb/system conversion, transforms), and the
slide/layout/master background fallback chain, directly. Everything else that touches the DOM
(`slide.test.ts`, `table.test.ts`, `text.test.ts`, `fill.test.ts`, `shape-tree.test.ts`,
`presentation-element.test.ts`) opts into `happy-dom` per-file via a
`// @vitest-environment happy-dom` docblock — the repo's root `vitest.config.ts` stays on
`environment: 'node'` for every other package, this is the only one that needs a DOM.
`slide.test.ts` asserts percentages via a local `pct()` helper that mirrors `positionElement`'s
own formula exactly (rather than hardcoding decimal strings, which would be one floating-point
rounding difference away from a false failure), and covers: top-level percentage positioning,
nested-group percentage remapping (including that the group wrapper is stretched via
`left/top/right/bottom: 0`, not just `left/top: 0`), and a placeholder shape with no own transform
picking up its layout placeholder's position. `table.test.ts` checks column/row percentages sum
to the expected split (50/50 for two equal-width columns) rather than absolute values. `text.test.ts`
covers run→span rendering, inline style application for every character-formatting field, theme
font-token resolution against a real `Theme`, and a placeholder shape picking up the master's
title style through `renderTextBody`'s `placeholder`/`context` parameters end-to-end. `fill.test.ts`
covers `applyFill`/`applyLine`/`resolveGradientCss` directly (solid/gradient/pattern/blip fill,
dash-style→border-style mapping, noFill short-circuiting); `shape-tree.test.ts` is a thin
integration check that `renderShapeTreeNode` actually wires a shape's/picture's own `fill`/`.line`
through to the rendered element, and `slide.test.ts` has the equivalent check for
`resolveEffectiveBackground` being applied by `renderSlide`.

## Next likely steps

1. `TextBodyProperties.anchor`/`wrap` → flex/white-space, `ParagraphProperties.alignment` →
   `text-align`, table cell fill, table styles.
2. Connector line rendering (see the scope boundary above) — likely an SVG line/path overlay
   sized to the connector's own box, reusing `applyLine`'s color/width/dash resolution but not its
   `border-*` output.
3. `apps/web-demo` now renders the parsed presentation into the page (see its own source) —
   confirm this still looks right whenever this package's DOM structure changes.
