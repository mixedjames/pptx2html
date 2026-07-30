# @pptx2html/to-html5

Renders a `@pptx2html/presentation` object graph (the in-memory `.pptx` DOM, produced by
`@pptx2html/reader`) into an actual HTML5 DOM. Public API is a single synchronous function,
mirroring the reader's `readPresentation`:

```ts
import { renderPresentation } from '@pptx2html/to-html5';
const el = renderPresentation(presentation); // <pptx-presentation>, shadow DOM inside
document.body.appendChild(el);
```

## Status: layout, plus formatting passes for fonts/alignment/lists, shape fill/line (+ style-matrix fallback), shape geometry, and slide backgrounds — all responsively scaled

Every slide and shape lands in the right place at the right size, including placeholder shapes
that inherit their position from the slide layout/master rather than declaring their own (very
common in real decks — see `placeholder.ts`). Run-level character formatting (typeface, size,
bold, italic, underline, strikethrough, text colour), paragraph alignment, and bulleted/numbered
lists all render, fully resolved through the same OOXML text-property inheritance chain —
run/paragraph → shape → placeholder layout/master → master's title/body/other style →
presentation default → (fonts and bullet glyphs only) theme font scheme (see "Key design
decision: font/alignment/bullet inheritance chain" below). Shape/picture `fill`/`.line`
(§20.1.2.2.35's spPr) also render, as CSS background/border (`fill.ts`) — solid fills, linear
gradients, and (approximated) patterns/images for fill; solid-colored, dashed/dotted/double
borders for line. When a shape/picture has **no** explicit `spPr` fill/line of its own —
PowerPoint's Shape Styles gallery writes shapes exactly this way, via a bare `p:style`
`fillRef`/`lnRef` — `style-matrix.ts` resolves that reference against the theme's format-scheme
style matrix instead, so such a shape isn't left with no fill/border at all (see "Key design
decision: style-matrix (`p:style`) resolution" below). A shape/picture's own preset outline
(§20.1.9.18, `a:prstGeom`) now shapes that fill/line too, for a common subset of presets
(`shape-geometry.ts`) — `rect`/`roundRect`/`ellipse` via CSS `border-radius` (keeping full
fill/line fidelity, including gradients/patterns/images), and nine further presets (triangle,
right triangle, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, 5-point star) via a
real SVG `<path>` outline (solid fill/stroke only — see the scope boundary). A slide's own
background renders too, falling back through layout then master (`background.ts`), reusing the
same `fill.ts` machinery. Every absolute magnitude this pass introduces — font size, border width,
list indentation — scales with the slide via CSS container query units rather than a fixed px/pt
(see "Key design decision: absolute sizes scale via `cqw`" below), consistent with how
position/size already scale via percentages. Table cell/table styling is still unrendered —
deliberately deferred to a later pass.

## Layout

- `units.ts` — `EMU_PER_PX` (9525, i.e. 96 CSS px/inch) and `emuToPx`, exported for consumers but
  no longer used internally (see the percentage-based-layout design decision). `EMU_PER_PT`
  (12700) and `fontSizeToEmu` convert a run's `FontSize` (hundredths of a point) to EMU.
  `emuToCqw(emu, slideWidth)` expresses any EMU magnitude as a percentage of the slide's own
  width, suffixed `cqw` (CSS container query width units) instead of a fixed px/pt — see the
  `cqw` design decision below.
- `coordinate.ts` — pure math, no DOM, fully unit-testable in Node: `CoordinateMap` (an affine
  translate+scale from a local EMU space to the slide's root EMU space), `IDENTITY_MAP`,
  `composeGroupMap` (extends a map with a group's own transform so its children's transforms —
  expressed against `chOff`/`chExt`, §20.1.7.6 — resolve correctly), `computeBox` (map + a
  `Transform2D` → slide-relative offset/extents, **still in EMU**, plus rotation/flip — converting
  to a CSS unit is left to the caller, see below).
- `text.ts` / `table.ts` — pure content renderers (`TextBody` → `<div><p>…</p></div>`,
  `Table` → `<table>`), no absolute positioning of their own; a paragraph is `<p>` with its
  effective `alignment` applied as `text-align` (`distributed` approximated via
  `text-align: justify` plus `text-align-last: justify`, since CSS has no "distributed" keyword)
  and, if it resolves a bullet, `padding-left`/`text-indent` (from `marginLeft`/`indent`, in `cqw`)
  plus a `<span class="pptx-bullet">` prepended before its runs (see the design decision below); a
  run or field is a `<span class="pptx-run">` (text content plus its resolved inline font styling,
  font-size in `cqw` via `emuToCqw`/`fontSizeToEmu` — see below), a break is `<br>`. `renderTable`
  sizes itself to `100%`/`100%` of its containing `graphicFrame` div, `table-layout: fixed`, with
  `<col>` widths and `<tr>` heights expressed as a percentage of the column-width/row-height
  totals (so column proportions survive scaling; row-height percentages are best-effort —
  browsers that ignore them just size rows by content instead).
- `text-style.ts` — pure logic, no DOM: `levelChain` (private) collects a paragraph's outline
  level from every rung of the list-style inheritance chain (see the design decision below) as an
  ordered array of `TextListStyleLevel`s; `resolveEffectiveRunProperties` merges each level's
  `runProperties` field-by-field into one `RunProperties`; `resolveEffectiveAlignment`,
  `resolveEffectiveBullet` and `resolveEffectiveIndent` all instead use the shared `resolveScalar`
  helper — "first (highest-priority) defined wins" rather than a field-by-field merge, since each
  is one paragraph-level value (or, for a bullet, one whole tagged-union value — a `CharBullet`
  and an `AutoNumberBullet` aren't merged together) rather than a set of independent fields.
  `resolveTypeface` resolves a `+mj-lt`/`+mn-lt`/etc. theme font token against a `FontScheme`,
  passing any literal font name straight through.
- `bullet.ts` — pure logic, no DOM: `formatAutoNumber(n, scheme)` renders a 1-based ordinal as its
  scheme's label (`"3."`, `"c)"`, `"iv."`, etc. — ten `AutoNumberScheme`s covering arabic/alpha/
  roman numerals crossed with period/right-paren suffixes). `NumberingState` tracks the running
  counter per outline level as `renderTextBody` walks a text body's paragraphs in order — this is
  the one piece of list rendering that's inherently stateful across paragraphs (unlike
  alignment/bullet-glyph resolution, which `text-style.ts` resolves independently per paragraph),
  since a number depends on how many auto-numbered siblings at the same level came before it; see
  the design decision below for its restart/continue rules.
- `color.ts` — pure logic, no DOM: `resolveColor` turns a DrawingML `Color` (srgb/scheme/system/
  preset/hsl, plus its `ColorTransform` modifiers — lumMod/lumOff/shade/tint/satMod/hueMod/alpha)
  into a CSS colour string; `resolveFillColor` is the `Fill`-level wrapper for the common
  single-colour case (only `SolidFill` resolves to a colour — gradient/pattern/blip fills need
  more than one CSS colour), used for both run text colour (`text.ts`) and a shape's `line`
  colour (`fill.ts`). Scheme colour resolution assumes the default clrMap (bg1→lt1, tx1→dk1,
  bg2→lt2, tx2→dk2), since slide/layout clrMap overrides are unmodeled in `packages/presentation`.
- `fill.ts` — DOM-touching (unlike `color.ts`): `applyFill` sets a shape/picture's CSS
  `background-color`/`background-image` from its `ShapeProperties.fill`, and `applyLine` sets
  `border-width`/`-style`/`-color` from its `ShapeProperties.line` (`border-width` in `cqw` via
  `emuToCqw`, hence the extra `slideWidth` parameter). `resolveGradientCss` (the one pure piece,
  exported and separately tested) converts a `GradientFill` to a CSS `linear-gradient(...)`,
  including translating DrawingML's shade-path angle convention (clockwise from east) to CSS's
  (clockwise from north). A `blip` fill creates an object URL exactly like `renderPicture` does
  for a picture's own image — same never-revoked caveat, see below. `applySvgFill`/`applySvgLine`
  are the SVG-`<path>` equivalents of `applyFill`/`applyLine`, used for the preset shapes
  `shape-geometry.ts` renders as a real outline instead of a rectangle (see `shape-tree.ts` below)
  — `applySvgFill` only handles a solid fill (`fill` attribute), `applySvgLine` a solid/dashed/
  dotted stroke (`stroke`/`stroke-width`/`stroke-dasharray`, width and dash lengths both in `cqw`
  so they stay a consistent screen size regardless of the path's own non-uniformly-stretched
  coordinate space — see `shape-geometry.ts`'s own doc comment); gradient/pattern/blip fills and a
  non-solid outline colour are left unset on these, the same "no faithful non-rectangular
  equivalent, don't guess a colour" reasoning `applyLine` already uses for a non-solid border
  colour.
- `shape-geometry.ts` — pure logic, no DOM: `presetShapePath` renders a `PresetGeometry`
  (§20.1.9.18, `a:prstGeom`) as an SVG path's `d` string in a fixed `0 0 100 100` coordinate space,
  for a common subset of nine presets whose outline a CSS box genuinely can't express (triangle,
  `rtTriangle`, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, `star5`) — the
  caller (`shape-tree.ts`) stretches it non-uniformly onto the shape's actual box via
  `preserveAspectRatio="none"`, mirroring how every position/size elsewhere in this package is a
  percentage of its own box rather than an absolute unit. Returns `undefined` for `rect`/
  `roundRect`/`ellipse` and any preset outside the modeled subset — the caller falls back to a
  plain rectangle for the latter (today's pre-existing behaviour) and to `nativeBorderRadius` for
  the former. `nativeBorderRadius` returns a CSS `border-radius` percentage for `roundRect` (from
  its `adj` guide, capped at 50% so `adj`'s spec max of 50000 yields a full stadium/pill) and a
  fixed `50%` for `ellipse` — deliberately **not** folded into `presetShapePath`, since a plain
  `<div>` with `border-radius` already draws both exactly, including a border that correctly
  follows the rounded outline (an SVG path buys nothing there and would cost full fill fidelity —
  see the scope boundary below), unlike the nine presets `presetShapePath` covers. Adjustment-guide
  handling reads a preset's own named guide (`ShapeGuide`, only a literal `val N` override — see
  `packages/presentation`'s own doc comment, no formula evaluator) when present, else a fixed
  per-preset default chosen to look reasonable — not transcribed from the spec's own `<gdLst>`
  formulas, so not guaranteed bit-exact; same approximation tier as this package's other
  best-effort stand-ins (pattern fills, colour-transform ordering).
- `style-matrix.ts` — pure logic, no DOM: `resolveStyleFill`/`resolveStyleLine` resolve a shape's
  `p:style/fillRef`/`lnRef` (`StyleMatrixReference`, `packages/presentation`'s `shape-style.ts`)
  against the theme's `FormatScheme.fillStyles`/`.lineStyles` (1-based `index`), substituting the
  reference's own `color` for every `phClr` ("placeholder colour") scheme-colour reference found
  anywhere inside the resolved Fill/Line — recursively, so it reaches every gradient stop and a
  pattern's fore/background colour, not just a top-level solid fill. The substitution merges the
  style entry's own local colour transforms (e.g. the theme's fillStyleLst commonly stacks
  `tint`/`lumMod` on its `phClr` entries) underneath whatever transforms the reference's own colour
  carries (rare in practice) — a flat-field merge, not a spec-accurate ordered composition of two
  transform stacks, the same approximation tier `color.ts`'s own colour-transform doc comment
  already documents elsewhere in this package. Returns `undefined` for a missing reference, missing
  theme, or an out-of-range index — the caller (`shape-tree.ts`) only reaches for this when the
  shape has no explicit `spPr` fill/line of its own, see the design decision below.
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
  `computeBox`'s EMU box to `left`/`top`/`width`/`height` **percentages of `context.slideSize`**
  and also sets `box-sizing: border-box` — OOXML sizes a shape's outline _within_ its bounding box
  (§20.1.2.2.24), not outside it, matching CSS's `border-box` rather than the default `content-box`
  (which would render a bordered shape larger than its declared width/height by the border's own
  width). `renderShape`/`renderPicture` also call `fill.ts`'s `applyFill`/`applyLine` — but not with
  `ShapeProperties.fill`/`.line` directly; both first go through `effectiveFill`/`effectiveLine`
  (the shape's own `spPr` fill/line if it has one, else its `p:style/fillRef`/`lnRef` resolved via
  `style-matrix.ts`'s `resolveStyleFill`/`resolveStyleLine` against `context.layout?.master.theme`'s
  `formatScheme` — see the design decision below). A picture's fill shows through any transparent
  pixels in the image itself, since `Picture` shares `ShapeProperties` with `Shape` — `renderConnector`
  doesn't call either yet, see the scope boundary below. Before doing that, `renderShape` first checks
  `shape-geometry.ts`'s `presetShapePath` against the shape's own `geometry`: if it returns a path
  (one of the nine non-rectangular presets it models), `renderShape` appends `renderShapeOutline`'s
  `<svg>` overlay instead — sized to exactly cover the shape's box, `viewBox="0 0 100 100"` with
  `preserveAspectRatio="none"` so the fixed path stretches non-uniformly onto whatever aspect ratio
  the box actually has, one `<path>` styled via `fill.ts`'s `applySvgFill`/`applySvgLine` — and
  skips `applyFill`/`applyLine` on the `<div>` itself (no CSS background/border to paint underneath
  an opaque SVG shape). Otherwise it checks `nativeBorderRadius` (covers `roundRect`/`ellipse`) and
  sets `border-radius` before the normal `applyFill`/`applyLine` calls, so those two presets keep
  full CSS fill fidelity (gradients, patterns, images) that the SVG path route can't offer.
  `renderPicture` only takes the `nativeBorderRadius` half of this — clipping an `<img>` to one of
  the nine SVG-path presets isn't implemented, see the scope boundary below — but that alone covers
  the common oval/rounded-corner photo crop. `group` renders as an anchor `<div>` stretched to exactly
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
  follows automatically, plus `container-type: inline-size` so descendants can size themselves in
  `cqw` against _this_ element's own rendered width (see the `cqw` design decision below). Applies
  `resolveEffectiveBackground`'s result via `fill.ts`'s `applyFill` directly on the slide div,
  before its `shapeTree` (walked from `IDENTITY_MAP` with a `RenderContext` built from
  `slide.layout`, the passed-in `slideSize`, and an optional `defaultTextStyle` — the
  presentation's `p:defaultTextStyle`, passed down from `presentation-element.ts` since
  `Presentation` itself isn't otherwise threaded this deep).
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

## Key design decision: a shape's style-matrix reference is a whole-value fallback, only reached when spPr has no fill/line at all

PowerPoint's Shape Styles gallery — the default way a shape gets a fill/border when drawn via the
UI, not just an edge case — writes a shape with a bare `p:style/fillRef`/`lnRef` and **no**
`spPr/solidFill`/`ln` of its own at all. Before this decision, `renderShape`/`renderPicture` only
ever looked at `ShapeProperties.fill`/`.line`, so any such shape rendered with no fill/border
whatsoever — this is what real decks look like far more often than an explicit `spPr` fill, since
manually setting `spPr` XML by hand (as opposed to using the gallery) is the unusual case.
`effectiveFill`/`effectiveLine` (`shape-tree.ts`) fix this: `shape.properties.fill ??
resolveStyleFill(shape.style?.fillRef, formatScheme)`, and the `.line` equivalent — the shape's own
`spPr` value wins outright when present, `style-matrix.ts`'s resolution only runs as a fallback.
This is a **whole-value** fallback, not a field-level merge (an explicit `spPr` fill entirely
replaces the style reference, never blends with it) — the same "first defined wins outright"
simplification `background.ts`'s `resolveEffectiveBackground` already uses for the slide/layout/
master background chain, chosen for the same reason: real decks essentially never partially
override a style reference (a shape either fully relies on its gallery style, or a user manually
picked its own fill/line, replacing the style reference's fill/line entirely in the UI's own
model). The resolved fill/line then flows into the same `applyFill`/`applyLine`/preset-geometry
path an explicit `spPr` fill/line would have — a style-matrix-sourced fill renders identically to
an equivalent literal one, no separate code path downstream of `effectiveFill`/`effectiveLine`.

## Key design decision: font/alignment/bullet inheritance chain resolved eagerly per run/paragraph, not via CSS cascade

A run's effective font (or a paragraph's effective alignment/bullet) could instead lean on the
browser's own CSS cascade — apply each level's style to its own DOM ancestor and let
`inherit`/unset properties flow down naturally. We didn't do that: `text-style.ts`'s `levelChain`
walks the whole chain in JS for every paragraph, and `resolveEffectiveRunProperties`/
`resolveEffectiveAlignment`/`resolveEffectiveBullet`/`resolveEffectiveIndent` write the
fully-resolved result as that run's/paragraph's own inline styles (mirrors
`resolveInheritedTransform`'s approach to placeholder position, and the same "every element's box
independently readable from its own inline styles" rationale as the CoordinateMap decision above).
The chain, lowest to highest priority: `Presentation.defaultTextStyle` → the slide master's
title/body/other `TextStyles` (by placeholder category — title/ctrTitle use `titleStyle`, other
placeholder types `bodyStyle`, non-placeholder shapes `otherStyle`) → the master's own matching
placeholder shape's `TextBody.listStyle` → the layout's matching placeholder shape's
`TextBody.listStyle` (mirrors `resolveInheritedTransform`'s layout→master walk, via the same
`findPlaceholderMatch`, now exported from `placeholder.ts` for this reuse) → the shape's own
`TextBody.listStyle`. Run properties then layer the paragraph's own `defRPr` and finally the run's
own `rPr` on top; alignment/bullet/margin/indent instead take the paragraph's own value outright if
it has one (each a single scalar — or, for bullet, one whole tagged-union value — not a set of
fields to merge). Each step only supplies defaults for the paragraph's own outline level (0-based,
`TextListStyleLevel`, indexed via `TextListStyle.levels[level]`). This order is a reasonable
synthesis of how real-world OOXML renderers describe the (spec-under-specified) inheritance, not a
verified bit-exact match to PowerPoint. Theme font-scheme resolution (`+mj-lt`/`+mn-lt`/etc. → an
actual typeface) is a separate, final step (`resolveTypeface`) since it only concerns the
`typeface` field and needs the theme, not a list style — done in `text.ts` right before the style
is applied (for both run text and, since a bullet glyph can reference the theme the same way, a
bullet's own `font`), using `context.layout?.master.theme`.

## Key design decision: bullets render as an explicit queryable `<span>`, not a CSS `::before`

A bullet/number could be painted as a `list-style` marker or a `::before` pseudo-element on the
paragraph, closer to how a browser renders a native `<ul>`/`<ol>`. We didn't do that, for the same
reason every other resolved value in this package ends up as a real inline style rather than
something CSS-cascade-derived (see the design decision above): a `::before`'s content isn't
queryable from JS (or, in this codebase's case, from a `happy-dom` test) the way a real DOM node
is, and OOXML paragraphs don't actually nest into `<li>`s the way HTML lists do — a "list" here is
just a run of sibling paragraphs that happen to share an outline level and an auto-number scheme,
which can be interrupted by a non-bulleted paragraph or a different level at any point (unlike an
HTML `<ol>`, which is a single element the whole list lives inside). So `text.ts` prepends a real
`<span class="pptx-bullet">` (the glyph or `bullet.ts`'s formatted number) plus a space, as the
first children of the paragraph's own `<p>` — styled from the bullet's own `font`/`color`/
`sizePercent` overrides, falling back to the paragraph's "ambient" run properties (an empty run
resolved through the ordinary `resolveEffectiveRunProperties` chain) for anything unset, since a
bullet with no override of its own inherits the character formatting of the text it precedes.
Numbering itself (`bullet.ts`'s `NumberingState`) is the one genuinely stateful piece — `renderTextBody`
holds one `NumberingState` per text body and walks paragraphs in order, since each auto-numbered
label depends on how many same-level, same-scheme siblings came before it. Its restart rules
(chosen to match how PowerPoint's own outline numbering behaves, not spec-mandated): a paragraph
at level _L_ resets any tracked counter at a level deeper than _L_ (so a nested numbered list
restarts the next time it's entered); a non-auto-numbered, non-empty paragraph at level _L_ also
resets _L_'s own counter (a plain or char-bulleted paragraph breaks the numbered run); a scheme
change at the same level also restarts rather than continues. The hanging indent itself
(`padding-left`/`text-indent`, from `marginLeft`/`indent`) falls back to a synthesized default —
`(level + 1) * 0.5in` margin, a `0.25in` hang — only when a bullet is present but the chain
resolves neither value at all, so a bare `<a:buChar>` with no `marL`/`indent` anywhere still
renders with a visible hang instead of running the glyph straight into the text; real decks almost
always set these explicitly (typically via the master's list style) so this rarely triggers.

**A genuinely empty paragraph (no runs at all) never gets a bullet, and never touches
`NumberingState`.** PowerPoint itself only shows a bullet/number on a blank line while it's the
one being actively edited — not in the rendered slideshow, and decks routinely end a list with a
trailing empty paragraph as a structural artifact. `renderTextBody` special-cases this: `bullet`
is forced to `undefined` for an empty paragraph regardless of what the inheritance chain would
otherwise resolve, and the blank line neither consumes a number (`numbering.next`) nor breaks a
running list (`numbering.break`) — so a numbered list interrupted by a blank paragraph still
resumes its count correctly on the far side, rather than restarting at 1.

## Key design decision: absolute sizes (font size, border width) scale via CSS container query units, not JS

Font size and border width are the first _magnitudes_ this renderer introduces that aren't
themselves a position/size on the slide (those already scale for free via percentages — see
above). A fixed `pt`/`px` value would look disproportionately large or small as `<pptx-presentation>`
is resized, since it wouldn't shrink/grow along with the slide the way percentage-based position
does. The options were: (a) reintroduce a `ResizeObserver` + JS recomputation — exactly what the
percentage-based-layout decision above deliberately avoided for position, so rejected for the same
reason; (b) `vw` (viewport width) units — wrong whenever `<pptx-presentation>` isn't the full
browser viewport width, the common embedded case; (c) CSS **container query units** (`cqw`),
relative to the nearest ancestor with `container-type` set — `slide.ts` sets
`container-type: inline-size` on `.pptx-slide` itself, so `1cqw` is exactly 1% of the _slide's own_
rendered width, matching the reference dimension `positionElement`'s percentages already use. (c)
was chosen: `units.ts`'s `emuToCqw(emu, slideWidth)` computes `(emu / slideWidth) * 100 + 'cqw'` —
literally the same formula as a position percentage, just a different unit suffix — used by
`text.ts` for `font-size` (via `fontSizeToEmu`) and `fill.ts`'s `applyLine` for `border-width`.
Trade-off: requires Container Query Unit support (Chrome 105+, Safari 16+, Firefox 110+ — fine for
a 2026 target) and, as a knock-on effect, **`happy-dom`'s CSSOM doesn't recognize `cqw` as a valid
length yet** — it silently drops any `el.style.fontSize`/`.borderWidth` assignment using it, so the
DOM-level tests that exercise these (`text.test.ts`, `fill.test.ts`) can't assert on those specific
properties' values; see their `NOTE` comments and the Tests section below.

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
- **`ShapeProperties.geometry` is rendered for a common subset of presets only.**
  `shape-geometry.ts` covers twelve presets total: `rect` (the pre-existing default), `roundRect`
  and `ellipse` (native `border-radius`), and nine more via a real SVG outline (triangle,
  `rtTriangle`, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, `star5`) — see
  `shape-geometry.ts`'s and `fill.ts`'s own doc comments above. Outside that set (the other ~170
  `ST_ShapeType` names — arrows, callouts, stars other than `star5`, flowchart shapes, etc.) a
  shape/picture still renders as a plain rectangular box, and `custGeom` (freeform/custom path
  outlines) remains entirely unmodeled — `packages/presentation`'s `CustomGeometry` carries no path
  data (see that package's own scope boundary), so `presetShapePath` returns `undefined` for it and
  the shape falls back to a rectangle. Adjustment-guide (`avLst`) handling is also approximate, not
  spec-exact — see `shape-geometry.ts`'s doc comment. On the nine SVG-path presets, only a solid
  fill/stroke is supported; a gradient/pattern/blip fill or non-solid outline colour renders
  unfilled/uncoloured rather than an approximated single colour (see `fill.ts`'s `applySvgFill`/
  `applySvgLine`). `renderPicture` (`shape-tree.ts`) only applies the native `border-radius` half
  of this — an irregularly-cropped picture (e.g. a triangle-cropped photo) isn't clipped to its
  preset's real outline, since that would need an `<svg><image>` + `<clipPath>` overlay rather than
  the plain `border-radius` an `<img>` supports directly; deferred since `rect`/`roundRect`/
  `ellipse` covers the overwhelming majority of real picture crops.
- **Style-matrix resolution (`style-matrix.ts`) only covers `fillRef`/`lnRef`.** `effectRef`
  (a shape's effect style, e.g. shadow) and `fontRef` (its default text colour when a run doesn't
  set one) are unmodeled, matching `ShapeStyle`'s own scope in `packages/presentation` — neither
  effect rendering nor a fontRef-colour text fallback exist yet. `renderConnector` doesn't call
  `effectiveFill`/`effectiveLine` at all (it doesn't call `applyFill`/`applyLine` either, see
  connectors below), so a connector's own `p:style` (parsed, since `ShapeStyle` is on
  `ConnectionShape` too) currently goes unused. `resolveStyleFill`/`resolveStyleLine`'s `phClr`
  substitution is also a flat-field transform merge, not a spec-accurate ordered composition — see
  `style-matrix.ts`'s own doc comment.
- **Pattern fills are only approximated**, and colour transforms on preset colours are ignored —
  see `fill.ts`'s and `color.ts`'s own doc comments for why. The hatch overlay's own spacing
  (`2px`/`8px` in the `repeating-linear-gradient`) is also a fixed px, not `cqw` — it's a purely
  decorative stand-in already, not a real DrawingML magnitude, so it wasn't brought into the `cqw`
  pass below.
- **Bullet size/font-size interaction is an approximation.** `bullet.sizePercent` is meant to be a
  percentage of "the paragraph's own text size" (§21.1.2.4.9); since there's no single such value
  once runs can each set their own `fontSize`, the bullet span's size is computed from the
  paragraph's _ambient_ (no-run-override) resolved font size instead — accurate whenever the
  paragraph doesn't mix run-level font sizes on the same line as the bullet, which covers the
  overwhelming majority of real decks.
- **`buSzPts` (a point-size bullet override) and a handful of double-parenthesis
  `ST_TextAutonumberScheme` variants are unmodeled** — see `packages/presentation/CLAUDE.md`'s
  scope boundary; `buSzPct` and the ten common schemes `AutoNumberScheme` models cover the
  overwhelming majority of real decks.
- **Remaining visual formatting**: `TextBodyProperties.anchor`/`wrap`, table cell fill, table
  styles. The DOM structure exists (`.pptx-shape`, `.pptx-paragraph`, `.pptx-run`, etc.) precisely
  so a later pass can add CSS without restructuring.
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

`coordinate.test.ts`, `placeholder.test.ts`, `text-style.test.ts`, `color.test.ts`,
`background.test.ts`, `units.test.ts`, `bullet.test.ts`, `shape-geometry.test.ts` and
`style-matrix.test.ts` are pure Node
(no DOM) and cover the
affine math, the placeholder transform inheritance chain (layout match, master fallback,
no-match), the shared `levelChain` walked by the font-property/alignment/bullet/indent
inheritance resolvers (every rung from `defaultTextStyle` down to a run's own `rPr`/a paragraph's
own `algn`/bullet/`marL`/`indent`, plus theme font-token
resolution), DrawingML colour resolution (scheme aliasing, hsl/srgb/system conversion,
transforms), the slide/layout/master background fallback chain, the EMU→px/pt/cqw conversion
math, `presetShapePath`/`nativeBorderRadius`'s path/radius output per preset (including the
`adj`-guide default vs. explicit-override cases and the 50%-cap edge case for `roundRect`/
parallelogram/trapezoid/hexagon/octagon), and `resolveStyleFill`/`resolveStyleLine`'s `phClr`
substitution (a top-level solid fill, every gradient stop, a pattern's fore/background colour, the
local-transform-merge case, an out-of-range index, and a missing reference/format-scheme), directly.
Everything else that touches the DOM (`slide.test.ts`, `table.test.ts`,
`text.test.ts`, `fill.test.ts`, `shape-tree.test.ts`, `presentation-element.test.ts`) opts into
`happy-dom` per-file via a
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
font-token resolution against a real `Theme`, a placeholder shape picking up the master's title
style through `renderTextBody`'s `placeholder`/`context` parameters end-to-end, and paragraph
`alignment` → `text-align` (including the `distributed` → `text-align`+`text-align-last`
approximation and the no-alignment default). `fill.test.ts`
covers `applyFill`/`applyLine`/`resolveGradientCss` directly (solid/gradient/pattern/blip fill,
dash-style→border-style mapping, noFill short-circuiting), plus `applySvgFill`/`applySvgLine`'s
SVG-`fill`/`stroke` equivalents (solid resolves, gradient/noFill/undefined leave `fill: none`, a
line with no fill of its own falls back to `stroke: currentColor`); `shape-tree.test.ts` is a thin
integration check that `renderShapeTreeNode` actually wires a shape's/picture's own `fill`/`.line`
through to the rendered element, and its "preset geometry" block covers the three-way branch in
`renderShape` end-to-end: `ellipse`/`roundRect` still go through the CSS `applyFill`/`applyLine`
path (with `border-radius` set, no `<svg>` child), `triangle` renders an `<svg><path>` instead (CSS
background/border left unset, the path's own `d`/`fill`/`stroke` asserted directly), and an
unmodeled preset (e.g. `cloud`) falls back to today's plain rectangle — plus a picture-with-`ellipse`
case for `nativeBorderRadius`'s img-crop path. Its "style-matrix (p:style) fallback" block covers
`effectiveFill`/`effectiveLine` end-to-end through `renderShapeTreeNode`: a shape/picture with only
a `style.fillRef`/`.lineRef` (no `spPr` fill/line) resolves and renders that style's colour, an
explicit `spPr` fill wins outright over a style reference when both are present, a style-sourced
fill still feeds the SVG preset-geometry path (not just the CSS `applyFill` one), and a shape with
neither an `spPr` fill/line nor a style reference stays unstyled. `slide.test.ts` has the equivalent
fill/line
integration check for `resolveEffectiveBackground` being applied by `renderSlide`. `text.test.ts`'s "bullets and
numbering" block covers char-bullet rendering (glyph, own font/colour), sequential auto-numbering
across sibling paragraphs, a nested sub-list restarting and the outer list resuming its count
after, an explicit `buNone` suppressing a bullet inherited from the shape's own list style, the
ambient-run-property colour fallback when a bullet has no colour of its own, and a trailing/
interior empty paragraph getting no bullet and neither consuming a number nor breaking a running
numbered list — `bullet.test.ts` covers `formatAutoNumber`'s ten schemes (including the alpha
scheme's z→aa wraparound and a `1994` roman-numeral case) and `NumberingState`'s
continue/restart/break rules directly, DOM-free.

**Known test-environment gap**: `text.test.ts` and `fill.test.ts` can't assert on the actual
`style.fontSize`/`style.borderWidth`/`style.paddingLeft`/`style.textIndent`/`style.strokeWidth`
values `applyRunStyle`/`applyLine`/`applyIndent`/`applySvgLine` set, because those are now `cqw`
values and `happy-dom`'s CSSOM silently drops style assignments in units it doesn't recognize yet
(confirmed real browsers with Container Query Unit support accept them fine — see the `cqw` design
decision above). Both files have a `NOTE` comment at the point this bites; the actual
`emuToCqw`/`fontSizeToEmu` conversion math is still fully covered, DOM-free, in `units.test.ts`.

## Next likely steps

1. `TextBodyProperties.anchor`/`wrap` → flex/white-space, table cell fill, table styles.
2. Connector line rendering (see the scope boundary above) — likely an SVG line/path overlay
   sized to the connector's own box, reusing `applyLine`'s color/width/dash resolution but not its
   `border-*` output. `shape-tree.ts`'s new `renderShapeOutline` (an `<svg>` sized to a shape's own
   box, see above) is a directly reusable pattern for this now that it exists.
3. Widening `shape-geometry.ts`'s preset coverage — arrows (`rightArrow`/`leftArrow`/etc.), other
   star counts (`star4`/`star6`/etc.), and flowchart shapes are all common in real decks and not
   yet modeled; each is a straightforward addition to `presetShapePath`'s switch.
4. Gradient/pattern/blip fill support on the nine SVG-path presets (currently solid-only, see the
   scope boundary above) — an SVG `<linearGradient>`/`<pattern>` `<defs>` entry referenced via
   `fill="url(#id)"` would parallel `resolveGradientCss`'s CSS-gradient output without the
   rectangular-background constraint, but needs its own angle/stop-position math since SVG
   gradients use `objectBoundingBox` coordinates, not CSS's `linear-gradient()` syntax.
5. Clipping a picture to one of `shape-geometry.ts`'s nine SVG-path presets (not just `roundRect`/
   `ellipse`'s native `border-radius`) — needs an `<svg><image>` + `<clipPath>` overlay in place of
   the plain `<img>` `renderPicture` emits today.
6. `p:style/effectRef` and `fontRef` (see the scope boundary above) — `effectRef` needs effect
   rendering to exist at all first (a bigger, separate gap); `fontRef` is a smaller, self-contained
   addition — a run/paragraph with no resolved colour of its own could fall back to its shape's
   `fontRef`'s colour before defaulting to black, in `text-style.ts`'s
   `resolveEffectiveRunProperties` chain.
7. Wiring a connector's own `p:style` (already parsed, `ShapeStyle` is on `ConnectionShape` too)
   into whatever connector line rendering eventually lands (see item 2 above) — today it's parsed
   but unused, since `renderConnector` doesn't call `applyFill`/`applyLine`/`effectiveFill`/
   `effectiveLine` at all yet.
8. `apps/web-demo` now renders the parsed presentation into the page (see its own source) —
   confirm this still looks right whenever this package's DOM structure changes.
