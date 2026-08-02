# @pptx2html/to-html5

Renders a `@pptx2html/presentation` object graph (the in-memory `.pptx` DOM, produced by
`@pptx2html/reader`) into an actual HTML5 DOM. Public API is a single synchronous function,
mirroring the reader's `readPresentation`:

```ts
import { renderPresentation } from '@pptx2html/to-html5';
const el = renderPresentation(presentation); // <pptx-presentation>, shadow DOM inside
document.body.appendChild(el);
```

## Status: layout, formatting passes (fonts/alignment/lists, shape fill/line, geometry, backgrounds) — all responsively scaled — plus slideshow navigation

Every slide and shape lands in the right place at the right size, including placeholder shapes
that inherit their position from the slide layout/master rather than declaring their own (very
common in real decks — see `@pptx2html/presentation`'s `resolve/placeholder.ts`). Run-level
character formatting (typeface, size,
bold, italic, underline, strikethrough, text colour), paragraph alignment, and bulleted/numbered
lists all render, fully resolved through the same OOXML text-property inheritance chain —
run/paragraph → shape → placeholder layout/master → master's title/body/other style →
presentation default → (fonts and bullet glyphs only) theme font scheme (see "Key design
decision: font/alignment/bullet inheritance is applied eagerly" below). Shape/picture `fill`/`.line`
(§20.1.2.2.35's spPr) also render, as CSS background/border (`fill.ts`) — solid fills, linear
gradients, and (approximated) patterns/images for fill; solid-colored, dashed/dotted/double
borders for line. When a shape/picture has **no** explicit `spPr` fill/line of its own —
PowerPoint's Shape Styles gallery writes shapes exactly this way, via a bare `p:style`
`fillRef`/`lnRef` — `@pptx2html/presentation`'s `resolve/style-matrix.ts` resolves that reference
against the theme's format-scheme
style matrix instead, so such a shape isn't left with no fill/border at all (see "Key design
decision: style-matrix (`p:style`) resolution" below). A shape/picture's own preset outline
(§20.1.9.18, `a:prstGeom`) now shapes that fill/line too, for a common subset of presets
(`shape-geometry.ts`) — `rect`/`roundRect`/`ellipse` via CSS `border-radius` (keeping full
fill/line fidelity, including gradients/patterns/images), and nine further presets (triangle,
right triangle, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, 5-point star) via a
real SVG `<path>` outline (solid fill/stroke only — see the scope boundary). A slide's own
background renders too, falling back through layout then master (`@pptx2html/presentation`'s
`resolve/background.ts`), reusing the
same `fill.ts` machinery. Every absolute magnitude this pass introduces — font size, border width,
list indentation — scales with the slide via CSS container query units rather than a fixed px/pt
(see "Key design decision: absolute sizes scale via `cqw`" below), consistent with how
position/size already scale via percentages. Table cell/table styling is still unrendered —
deliberately deferred to a later pass. `<pptx-presentation>` now renders as an actual
slideshow rather than every slide stacked one below the next — one slide visible at a time,
advanced by click/keyboard (see "Key design decision: slideshow navigation" below). Navigating
between slides now (new) plays `Slide.transition`'s own effect when it's `push` or `fade` (§19.3.1.49,
`p:transition` — see "Key design decision: push/fade slide transitions" below) — a real animation,
driven by the Web Animations API rather than a plain CSS `transition` (see that same design
decision for why), rather than the instant swap every other effect kind (and a slide with no
`transition` at all) still falls back to. `Slide.timing`'s per-element/per-build animation tree
remains entirely unconsumed here still — see root `CLAUDE.md`'s Todos.

## Layout

**Note**: the OOXML inheritance-walking logic this section used to describe here —
`coordinate.ts` (`CoordinateMap`/`composeGroupMap`/`computeBox`/`ElementBox`), `text-style.ts`
(the font/alignment/bullet/indent chain), `placeholder.ts` (`resolveInheritedTransform`),
`background.ts` (`resolveEffectiveBackground`), `bullet.ts` (`formatAutoNumber`/`NumberingState`)
and `style-matrix.ts` (`resolveStyleFill`/`resolveStyleLine`) — has moved to
`packages/presentation/src/resolve/` (same filenames), since none of it actually renders anything;
see that package's CLAUDE.md for what's there now and why. This package now only _calls_ it and
turns the result into DOM/CSS/SVG — the entries below cover what's left here.

- `units.ts` — `EMU_PER_PX` (9525, i.e. 96 CSS px/inch) and `emuToPx`, exported for consumers but
  no longer used internally (see the percentage-based-layout design decision). `EMU_PER_PT`
  (12700) and `fontSizeToEmu` convert a run's `FontSize` (hundredths of a point) to EMU.
  `emuToCqw(emu, slideWidth)` expresses any EMU magnitude as a percentage of the slide's own
  width, suffixed `cqw` (CSS container query width units) instead of a fixed px/pt — see the
  `cqw` design decision below.
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
- `color.ts` — a thin CSS-formatting wrapper (unlike its former self, see the Layout note above):
  `resolveColor`/`resolveFillColor` keep their old signatures and CSS-string return type exactly
  (nothing downstream needed to change), but internally just call `@pptx2html/presentation`'s
  `resolveColor`/`resolveFillColor` (aliased on import to `resolveColorComponents`/
  `resolveFillColorComponents` to avoid shadowing) and format the resulting `ResolvedColor` —
  `rgbToCss` (unchanged) plus a `toCss` adapter for the `{ type: 'preset' }` case, which just
  passes the preset name through as-is (DrawingML preset names match CSS's extended colour
  keywords almost 1:1).
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
- `render-context.ts` — `RenderContext`, threaded alongside `CoordinateMap` (now
  `@pptx2html/presentation`'s) through every `renderShapeTreeNode` call. It `extends`
  `@pptx2html/presentation`'s `TextStyleContext` (`layout`/`defaultTextStyle` — the two pieces
  the moved text-style/placeholder resolvers need) rather than redeclaring those fields, adding
  only `slideSize` — the one field that's genuinely this renderer's own (every element's EMU box
  gets divided by it to become a percentage). Because `RenderContext` structurally satisfies
  `TextStyleContext`, every call site that already passes its full `context` to a
  `@pptx2html/presentation` resolver keeps compiling unchanged.
- `shape-tree.ts` — `renderShapeTreeNode`, the one recursive dispatcher over `ShapeTreeNode`'s
  five kinds. `shape`/`picture`/`connector` position from their own transform if they have one,
  else from `@pptx2html/presentation`'s `resolveInheritedTransform`; `graphicFrame` always has its
  own transform (mandatory in the schema) so it skips that path. All four get positioned by
  `positionElement`, which converts `computeBox`'s EMU box to `left`/`top`/`width`/`height`
  **percentages of `context.slideSize`** and also sets `box-sizing: border-box` — OOXML sizes a
  shape's outline _within_ its bounding box (§20.1.2.2.24), not outside it, matching CSS's
  `border-box` rather than the default `content-box` (which would render a bordered shape larger
  than its declared width/height by the border's own width). `renderShape`/`renderPicture` also
  call `fill.ts`'s `applyFill`/`applyLine` — but not with `ShapeProperties.fill`/`.line` directly;
  both first go through `effectiveFill`/`effectiveLine` (the shape's own `spPr` fill/line if it has
  one, else its `p:style/fillRef`/`lnRef` resolved via `@pptx2html/presentation`'s
  `resolveStyleFill`/`resolveStyleLine` against `context.layout?.master.theme`'s `formatScheme` —
  see the design decision below). A picture's fill shows through any transparent
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
- `slide.ts` — `renderSlide`: one `.pptx-slide` div, `position: relative; overflow: hidden`,
  `width: 100%` with `aspect-ratio: <slideWidth> / <slideHeight>` (raw EMU numbers — `aspect-ratio`
  only cares about the ratio) so it fills whatever width its container gives it and its height
  follows automatically, plus `container-type: inline-size` so descendants can size themselves in
  `cqw` against _this_ element's own rendered width (see the `cqw` design decision below). Applies
  `@pptx2html/presentation`'s `resolveEffectiveBackground` result via `fill.ts`'s `applyFill`
  directly on the slide div,
  before its `shapeTree` (walked from `IDENTITY_MAP` with a `RenderContext` built from
  `slide.layout`, the passed-in `slideSize`, and an optional `defaultTextStyle` — the
  presentation's `p:defaultTextStyle`, passed down from `presentation-element.ts` since
  `Presentation` itself isn't otherwise threaded this deep).
- `presentation-element.ts` — `PptxPresentationElement`, a `<pptx-presentation>` custom element.
  Shadow DOM is attached in the constructor; `.render(presentation)` replaces its slide children.
  Renders as a **slideshow**, not a stacked list — all slides are rendered into the shadow DOM up
  front (`.render()`) but only the current one is visible (`.pptx-slide--active`, CSS
  `display: none` on the rest — see the design decision below); `definePresentationElement()`
  registers the tag (idempotent). Navigation: `.next()`/`.previous()`/`.goToSlide(index)` (all
  clamped to the slide range), `.currentSlideIndex`/`.slideCount` getters, a click anywhere on the
  element advances one slide, and (once focused — a click also focuses it, since the element sets
  `tabIndex = 0`) `ArrowRight`/`ArrowDown`/`PageDown`/space/`Enter` advance, `ArrowLeft`/`ArrowUp`/
  `PageUp`/backspace retreat, `Home`/`End` jump to the first/last slide — mirroring PowerPoint's own
  presentation-mode key bindings. `goToSlide` (new) also plays a `Slide.transition` (the
  destination's own going forward, the outgoing slide's own reversed going backward — see "Key
  design decision: push/fade slide transitions" below) when its `effect.kind` is `'push'` or
  `'fade'` — see that section for the mechanism (`#animatePush`/`#animateFade`,
  `#beginTransitionFrame`, `#awaitTransition`/`#finalizeTransition`); every other effect kind, and a
  slide with no `transition` at all, still takes the unchanged instant-swap path
  (`#updateActiveSlide`).
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
percentage scheme above. We didn't do that — `@pptx2html/presentation`'s `composeGroupMap` walks
the ancestor chain in EMU space and every leaf element gets one final `left`/`top`/`width`/`height`
(as a percentage, see above), already relative to the slide. This makes each element's box
independently readable from its own inline styles (useful for debugging and for a future
formatting pass) without having to account for inherited CSS transforms. The one thing this trades
away: **rotation does not compose across nested groups** — `computeBox` returns a shape's (or
group's) own `rotation`/flip uncomposed with any ancestor group's, in `ElementBox`; applying that
as a CSS `transform` around the element's own box is still `to-html5`'s own choice (`positionElement`
in `shape-tree.ts`). Spec-correct for the (common) unrotated-group case; a rotated group with
rotated descendants will look wrong. Fixing this means switching the rotated subtree to the
nested-CSS-transform approach above — deliberately not done for this first pass.

## Key design decision: a shape's style-matrix reference is a whole-value fallback, only reached when spPr has no fill/line at all

PowerPoint's Shape Styles gallery — the default way a shape gets a fill/border when drawn via the
UI, not just an edge case — writes a shape with a bare `p:style/fillRef`/`lnRef` and **no**
`spPr/solidFill`/`ln` of its own at all. Before this decision, `renderShape`/`renderPicture` only
ever looked at `ShapeProperties.fill`/`.line`, so any such shape rendered with no fill/border
whatsoever — this is what real decks look like far more often than an explicit `spPr` fill, since
manually setting `spPr` XML by hand (as opposed to using the gallery) is the unusual case.
`effectiveFill`/`effectiveLine` (`shape-tree.ts`) fix this: `shape.properties.fill ??
resolveStyleFill(shape.style?.fillRef, formatScheme)`, and the `.line` equivalent — the shape's own
`spPr` value wins outright when present, `@pptx2html/presentation`'s `resolveStyleFill`/
`resolveStyleLine` (the substitution logic itself, and its own doc comments, now live there — see
that package's CLAUDE.md) only runs as a fallback. This is a **whole-value** fallback, not a
field-level merge (an explicit `spPr` fill entirely
replaces the style reference, never blends with it) — the same "first defined wins outright"
simplification `@pptx2html/presentation`'s `resolveEffectiveBackground` already uses for the
slide/layout/master background chain, chosen for the same reason: real decks essentially never
partially
override a style reference (a shape either fully relies on its gallery style, or a user manually
picked its own fill/line, replacing the style reference's fill/line entirely in the UI's own
model). The resolved fill/line then flows into the same `applyFill`/`applyLine`/preset-geometry
path an explicit `spPr` fill/line would have — a style-matrix-sourced fill renders identically to
an equivalent literal one, no separate code path downstream of `effectiveFill`/`effectiveLine`.

## Key design decision: font/alignment/bullet inheritance is applied eagerly as inline styles, not via CSS cascade

The chain itself — `Presentation.defaultTextStyle` → the slide master's title/body/other
`TextStyles` → the master's/layout's matching placeholder shape's `TextBody.listStyle` → the
shape's own → the paragraph's own `defRPr`/alignment/bullet/indent → the run's own `rPr` — is
`@pptx2html/presentation`'s `resolve/text-style.ts` now, see that package's CLAUDE.md for the
exact rung-by-rung order and reasoning. What's still `to-html5`'s own choice is _how_ the result
gets applied: a run's effective font (or a paragraph's effective alignment/bullet) could instead
lean on the browser's own CSS cascade — apply each level's style to its own DOM ancestor and let
`inherit`/unset properties flow down naturally. We didn't do that — `text.ts` calls
`resolveEffectiveRunProperties`/`resolveEffectiveAlignment`/`resolveEffectiveBullet`/
`resolveEffectiveIndent` once per run/paragraph and writes the fully-resolved result as that
element's own inline styles (mirrors `resolveInheritedTransform`'s approach to placeholder
position, and the same "every element's box independently readable from its own inline styles"
rationale as the `CoordinateMap` decision above). Theme font-scheme resolution (`+mj-lt`/`+mn-lt`/
etc. → an actual typeface) is a separate, final step (`resolveTypeface`, also
`@pptx2html/presentation`'s) since it only concerns the `typeface` field and needs the theme, not a
list style — called from `text.ts` right before the style is applied (for both run text and, since
a bullet glyph can reference the theme the same way, a bullet's own `font`), using
`context.layout?.master.theme`.

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
`<span class="pptx-bullet">` (the glyph or `@pptx2html/presentation`'s `formatAutoNumber`'s
formatted number) plus a space, as the
first children of the paragraph's own `<p>` — styled from the bullet's own `font`/`color`/
`sizePercent` overrides, falling back to the paragraph's "ambient" run properties (an empty run
resolved through the ordinary `resolveEffectiveRunProperties` chain) for anything unset, since a
bullet with no override of its own inherits the character formatting of the text it precedes.
Numbering itself (`@pptx2html/presentation`'s `NumberingState`) is the one genuinely stateful
piece — `renderTextBody`
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

## Key design decision: fontRef text-colour/typeface fallback lives in `resolveEffectiveRunProperties`, not a new resolver

A run with no colour/typeface of its own relies on the shape's own `p:style/fontRef`
(§20.1.4.1.17) for a default, which is what PowerPoint's Shape Styles gallery writes by default.
Rather than a separate resolver alongside `resolveStyleFill`/`resolveStyleLine` (the `fillRef`/
`lnRef` equivalent, `style-matrix.ts`), `@pptx2html/presentation`'s `resolveEffectiveRunProperties`
now takes an optional `ShapeStyle` argument and folds `fontRef`'s colour (as a `SolidFill`) and
font collection (as the same `+mj-lt`/`+mn-lt` theme token `resolveTypeface` already resolves — no
new font-resolution mechanism needed) into `levelChain` as its own rung, ranked **above** the
master/placeholder chain (`context.defaultTextStyle`, the master's title/body/other category style,
and any matching placeholder's own list style) but **below** the shape's own list style and the
paragraph/run's own formatting.

That ranking is deliberate, not the more obvious "last resort, loses to everything" — and getting
it backwards was an actual bug caught via `apps/web-demo/src/Presentation1.pptx`: a plain autoshape
(no `p:ph`) falls back to the master's `otherStyle` category for its text defaults, and a real
Office theme's `otherStyle` level 0 almost always sets an explicit `solidFill`/`typeface` (it's the
generic "any text box" default). If `fontRef` were ranked below that — as it originally was — the
master's generic dark colour would always clobber the shape's own, deliberately-chosen `fontRef`
colour (typically a light colour meant to contrast with the shape's own fill), which is backwards
from what PowerPoint itself renders: a shape's own directly-authored quick style is more specific
than a generic template default and should win, only losing to formatting actually applied to that
shape's own text. See `packages/presentation/CLAUDE.md`'s `levelChain` doc comment for the exact
rung ordering.

`renderRun`/`renderBulletSpan`'s ambient-run-properties call both take the rendering shape's own
`style` and pass it through for this reason (`text.ts`, `shape-tree.ts`'s `renderShape`). A shape
with an explicit `spPr` fill/line but no run colour of its own still gets this fallback correctly —
`fontRef` and `fillRef`/`lnRef` are independent references on the same `p:style`, not a single
all-or-nothing switch.

## Key design decision: a shape's text body is vertically anchored via flexbox

`a:bodyPr/@anchor` (§21.1.2.1.1 — top/center/bottom/justified) was previously read by the reader
into `TextBody.properties.anchor` but never consulted here at all, so text always rendered flush at
the top of a shape's box regardless of what the deck authored — visibly wrong for anything that
isn't top-anchored, e.g. a number centered inside a small circle. `renderShape` (`shape-tree.ts`)
now resolves the effective anchor via `@pptx2html/presentation`'s new `resolveEffectiveAnchor`
(own value, else the matching placeholder's own `bodyPr` in the layout then the master, else `'t'`)
and sets `display: flex; flex-direction: column; justify-content: <mapped>` on the shape's own div
before appending `.pptx-text-body` — `justify-content` rather than `align-items`/`align-content`
since the anchor axis is the block (vertical) axis in a column-direction flex container. This is
safe to set unconditionally alongside the SVG preset-outline overlay some shapes also get: an SVG
element positioned `absolute` (as that overlay always is) is excluded from flex layout entirely, so
turning the shape div into a flex container never disturbs it. `'just'` (anchor-justified —
distributing multiple paragraphs to fill the box) has no single-block flexbox equivalent and is
approximated as top-anchored (`justifyContentForAnchor`'s own doc comment) — this renderer doesn't
distribute paragraph spacing to fill a box the way real anchor-justified text does, a much rarer
case than plain top/center/bottom anchoring.

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

## Key design decision: all slides render into the DOM up front; navigation toggles visibility, it doesn't re-render

`<pptx-presentation>` could instead render only the current slide and call `renderSlide` again on
every `.next()`/`.previous()`/`.goToSlide()`, discarding the rest. We didn't do that:
`PptxPresentationElement.render()` renders every slide once into `#slidesContainer` and keeps them
all in the shadow DOM; navigation just toggles which one carries `.pptx-slide--active` (CSS
`display: none` on the rest). Two reasons: (1) it's what slide transitions (`push`/`fade`, now
implemented — see below — and per-slide/per-element animations off `Slide.timing`, still
unconsumed, see root `CLAUDE.md`'s Todos) need anyway, since animating a transition _between_ two
slides means both their DOM subtrees have to exist simultaneously at some point, not be created on
demand mid-transition; (2) it keeps `.next()`/`.previous()` synchronous and re-render-free, matching this
package's existing "resolve once, mutate style/class thereafter" style elsewhere (e.g. inline style
application in the font/alignment/bullet pass). Trade-off: every slide's full shape tree is built
and sitting in the DOM even for slides the user never reaches — fine at typical deck sizes (tens of
slides), not investigated for decks with hundreds.

Within that, advancing was originally a same-instant CSS swap (`display: none` ↔ `block`) for every
slide, deliberately minimal for that first pass. `push`/`fade` (new) now animate instead — see "Key
design decision: push/fade slide transitions" below — but every other effect kind, and a slide with
no `transition` at all, still take exactly this same-instant path unchanged.

Keyboard bindings mirror PowerPoint's own presentation-mode keys (right/down/page-down/space/enter
advance; left/up/page-up/backspace retreat; home/end jump to first/last) rather than inventing a
new scheme, since anyone who's used PowerPoint's slideshow mode already knows them. The element
sets `tabIndex = 0` (unless the host page already set its own `tabindex`) so it's keyboard-focusable
at all (custom elements aren't by default) and a click both focuses it (native behaviour for a
`tabindex`-bearing element) and advances one slide — consistent with PowerPoint's own "click
advances" behaviour, and meaning a single click is enough to both enter the presentation and start
advancing through it, no separate focus step needed. Clicking never retreats (matching PowerPoint)
— only keyboard/`.previous()` do. This is set in `connectedCallback()`, not the constructor: the
Custom Elements spec forbids a constructor from adding attributes to the element, and `tabIndex`
reflects to the `tabindex` attribute — doing it in the constructor throws (`NotSupportedError`) in
spec-strict implementations (WebKit), aborting the whole upgrade so the element never gets its
`render`/`next`/etc. methods at all. Everything else construction-time here (attaching the shadow
root, appending the `<style>`/slides-container children _inside that shadow root_, registering
event listeners) is fine in the constructor — the restriction is specifically about the element's
own light-DOM attributes/children, not its shadow tree or listeners.

## Key design decision: push/fade slide transitions

`Slide.transition` (§19.3.1.49, `p:transition`) describes the effect played when the presentation
_arrives at that slide_ going forward — OOXML has no notion of "backward" for it to describe.
`goToSlide` (`presentation-element.ts`) therefore looks up **the destination's own transition when
advancing forward, but the _outgoing_ slide's own transition when retreating backward**, and plays
it in reverse — going back to slide N-1 undoes whichever animation originally brought slide N into
view, rather than consulting whatever (if anything) slide N-1 separately authors for its own
forward arrival. Getting this backward (pun intended) was an actual bug in an earlier version of
this code, which always read the destination's transition regardless of direction — harmless when
adjacent slides share the same effect, but visibly wrong whenever they don't (e.g. slide 2 authors
`fade` and slide 1 authors `push`: retreating from 2 to 1 must undo slide 2's fade, not play a push
just because slide 1 happens to have its own, unrelated transition). Only two of `TransitionEffect`'s
~20 `kind`s are implemented — `push` (`SideDirectionTransitionEffect`, direction `l`/`u`/`r`/`d`,
defaulting to `'l'`) and `fade` (`FadeTransitionEffect`) — every other kind (wipe, cut, dissolve,
wheel, split, ...) and a slide with no `transition` at all still take the pre-existing instant
`display: none`/`block` swap (`#updateActiveSlide`) unchanged; see the scope boundary below for why
the rest aren't done yet.

**Both slides coexist and animate simultaneously, briefly.** The pre-existing "all slides render
into the DOM up front" decision (above) already keeps every slide's subtree alive; a transition
additionally makes both the outgoing and incoming slide `position: absolute` (added via inline
`style.position`, never a CSS class — `slide.ts`'s `renderSlide` already sets `position: relative`
as an _inline_ style, which a class rule can never override) and `display: block`
(`.pptx-slide--transitioning` + `.pptx-slide--active` on both) for the transition's duration, so
each can be animated independently. `.pptx-presentation` needed two additions to support this:
`overflow: hidden` (so a mid-push slide is clipped at the container edge instead of bleeding into
the page) and an explicit `aspect-ratio` set from `presentation.slideSize` in `.render()` (mirroring
`slide.ts`'s own per-slide line) — without it, the container would collapse to zero height the
moment neither slide is contributing to normal-flow layout.

**Playback is driven by the Web Animations API (`Element.animate()`), not a plain CSS
`transition`.** `#animatePush`/`#animateFade` call `.animate(keyframes, options)` directly on each
slide with explicit from/to keyframes, rather than writing a CSS `transition-duration` and then
mutating `style.transform`/`style.opacity` and letting the browser interpolate. Two concrete wins
from this, beyond the motivating one (see `docs/scroll-driven-playback.md` — a WAAPI `Animation`'s
`currentTime` is directly seekable, unlike a CSS transition, which a future scroll-driven playback
mode will need): (1) the reflow-forcing hack a CSS-transition version needs (write the start value,
force a synchronous reflow read, then write the end value — otherwise both writes collapse into one
style recalculation and never animate) is unnecessary and gone entirely, since `.animate()` takes
both endpoints as data in one call; (2) the defensive "write this property even though it doesn't
change, in case a prior different-effect transition left a stale value" keyframe writes are also
gone — `#finalizeTransition` (below) always `.cancel()`s a finished animation, fully releasing its
property override, so there's no stale-value risk left to guard against. `#currentAnimations:
Animation[]` (non-empty _is_ the "a transition is in flight" flag) tracks both animations;
`#awaitTransition` does `Promise.all([...].map(a => a.finished)).then(() => this.#finalizeTransition(...))`,
with a `.catch(() => {})` for the expected rejection when `render()` cancels an in-flight animation
(a new presentation rendered mid-transition).

**Cancelling a finished animation at cleanup time is a correctness requirement, not just hygiene.**
The same slide elements are reused across every future transition (`render()` builds them once);
leaving a finished, `fill: 'forwards'`-holding `Animation` attached to an element would stack
ambiguously against whatever plays on it next. `#finalizeTransition` calls `.cancel()` on both
animations before clearing state — safe to do visually, since cancelling reverts each property to
its element's _base_ style, which for every keyframe this file ever animates (`transform`/
`opacity`) happens to already equal the intended resting value, because neither is ever written as a
raw inline style anywhere else in this file.

**Duration is a documented approximation, now resolved by `@pptx2html/presentation`, not private
here.** OOXML's `TransitionSpeed` (`'slow'|'med'|'fast'`) is qualitative — no spec-mandated
millisecond value — `resolveTransitionDurationMs` (`packages/presentation/src/resolve/timing.ts`)
fixes `fast: 400, med: 700, slow: 1000`, the same tier of best-effort stand-in as this package's
other underspecified-OOXML-magnitude choices (`shape-geometry.ts`'s adjustment-guide defaults,
`fill.ts`'s pattern-hatch spacing) — moved out of this package since any renderer driving a
transition needs the identical answer (see that file's own doc comment, and root `CLAUDE.md`'s
scroll-driven-playback design note for why this matters beyond just avoiding duplication).

**Testing a WAAPI-driven mechanism against `happy-dom`, which implements none of it.** This
package's `happy-dom` test environment (see Tests below) has _no_ `Element.prototype.animate` at
all — not a partial/CSS-only gap the way `transitionend` was for the previous CSS-transition
version, a complete absence. `presentation-element.test.ts`'s `describe('push/fade transitions', ...)`
block installs its own `HTMLElement.prototype.animate` mock (a `FakeAnimation` class whose
`.finished` promise resolves via a `setTimeout` matching the requested duration) rather than relying
on any real implementation — see Tests below for the full mechanism and why `vi.useFakeTimers()`
alone isn't enough here (`Animation.finished` is a real `Promise`, so its `.then()` runs on the
microtask queue, which needs `vi.advanceTimersByTimeAsync()`, not the synchronous
`vi.advanceTimersByTime()` the pre-WAAPI version used).

**Navigation is ignored outright while a transition is in flight**, rather than interrupted-and-
restarted or queued: `goToSlide`'s very first check is whether `#currentAnimations.length > 0`
(non-empty _is_ the "animating" flag, no separate boolean), returning immediately if so — a
click/keypress/method call mid-transition has no effect at all until the current one finishes. This
was a deliberate simplicity choice over the alternatives (both of which need correctly canceling
and restarting an in-flight animation, a meaningfully harder problem) for this first pass.

**A directional effect's direction reverses for backward navigation, on top of the
which-slide's-transition fix above.** A `push` authored `direction: 'l'` on slide N plays
push-left when advancing into slide N, but push-right when retreating out of slide N back to
N-1 (`REVERSE_SIDE_DIRECTION`) — the same authored transition, played frame-for-frame in
reverse, which is what "undoing" an animation means. This compounds with, not replaces,
`goToSlide`'s outgoing-vs-destination lookup: the _which-transition_ fix says whose `effect` to
read for a backward step, and this direction flip says how to play that specific effect once
you're playing it in reverse.

**`fade`'s `throughBlack: true` renders identically to a plain crossfade this round** — the real
two-stage fade-out-to-black-then-in-from-black animation `throughBlack` describes is a deferred
follow-up (see the scope boundary below), not implemented as a distinct path yet.

## Scope boundary — what's intentionally unmodeled (yet)

- **Run-level font colour only resolves solid fills.** `RunProperties.fill` can technically be a
  gradient/pattern/blip `Fill` (WordArt-style text); `@pptx2html/presentation`'s `resolveFillColor`
  only handles `SolidFill`, falling back to no colour (inherits black) for the others — vanishingly
  rare on plain text runs in real decks.
- **Colour transforms on preset colours are ignored.** `@pptx2html/presentation`'s `resolve/color.ts`
  has no RGB table for the ~140 DrawingML preset names (they map almost 1:1 to CSS's extended
  colour keywords, so untransformed presets pass straight through as an opaque name — `to-html5`'s
  own (now much thinner) `color.ts` hands it to CSS as-is) — a `shade`/`tint`/etc. stacked on a
  preset colour is silently dropped. Transforms are almost always applied to scheme colours in
  practice, not presets, so this is a narrow gap.
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
- **Style-matrix resolution (`@pptx2html/presentation`'s `resolve/style-matrix.ts`) only covers
  `fillRef`/`lnRef`; `fontRef` (its default text colour/typeface fallback) is now resolved too, but
  via a different mechanism** — see "Key design decision: fontRef text-colour/typeface fallback"
  below. `effectRef` (a shape's effect style, e.g. shadow) remains unmodeled, matching
  `ShapeStyle`'s own remaining scope gap in `packages/presentation` — no effect rendering exists
  yet. `renderConnector` doesn't call `effectiveFill`/`effectiveLine` at all (it doesn't call
  `applyFill`/`applyLine` either, see connectors below), so a connector's own `p:style` (parsed,
  since `ShapeStyle` is on `ConnectionShape` too) currently goes unused — including its `fontRef`,
  though a connector never has a `textBody` in the schema, so this is moot in practice.
  `resolveStyleFill`/`resolveStyleLine`'s `phClr` substitution is also a flat-field transform
  merge, not a spec-accurate ordered composition — see that file's own doc comment.
- **Pattern fills are only approximated**, and colour transforms on preset colours are ignored —
  see `fill.ts`'s and `@pptx2html/presentation`'s `resolve/color.ts`'s own doc comments for why. The hatch overlay's own spacing
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
- **Remaining visual formatting**: `TextBodyProperties.wrap`, table cell fill, table
  styles — `anchor` (vertical text alignment) is now handled, see "Key design decision: a shape's
  text body is vertically anchored via flexbox" below. The DOM structure exists (`.pptx-shape`,
  `.pptx-paragraph`, `.pptx-run`, etc.) precisely so a later pass can add CSS without restructuring.
- **Placeholder inheritance is supported, but with a simplified matching rule.** See
  `@pptx2html/presentation`'s `resolve/placeholder.ts`. Not modeled: the spec's type-equivalence groups (e.g. a slide's
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
- **Only `push` and `fade` play a real slide transition.** Every other `TransitionEffect.kind`
  (`wipe`, `cut`, `dissolve`, `newsflash`, `wheel`, `split`, `strips`, `zoom`, `blinds`, `checker`,
  `comb`, `randomBar`, `circle`, `diamond`, `plus`, `pull`, `cover`, `random`, `wedge`) still takes
  the plain instant swap — each is a reasonably self-contained addition to `goToSlide`'s dispatch in
  `presentation-element.ts` once picked up, following the same pattern `#animatePush`/`#animateFade`
  establish. `fade`'s `throughBlack: true` variant renders as a plain crossfade rather than its own
  real fade-to-black-then-in animation (see "Key design decision: push/fade slide transitions"
  above). `SlideTransition.advanceOnClick`/`.advanceAfter` (auto-advance timers) and
  `.sound`/`TransitionSoundAction` (playing/stopping audio during the transition) are both parsed by
  `packages/reader` already but entirely unconsumed here — a slide's transition only plays in
  response to explicit navigation (click/keyboard/`goToSlide`), never on its own after a timeout,
  and no audio ever plays.

## Tests

The inheritance-resolution tests that used to live here (`coordinate.test.ts`,
`placeholder.test.ts`, `text-style.test.ts`, `background.test.ts`, `bullet.test.ts`,
`style-matrix.test.ts`) moved to `packages/presentation/src/resolve/` alongside the source they
test — see that package's CLAUDE.md for what they cover now. `shape-geometry.test.ts` and
`units.test.ts` are the remaining pure-Node (no DOM) test files here, covering
`presetShapePath`/`nativeBorderRadius`'s path/radius output per preset (including the `adj`-guide
default vs. explicit-override cases and the 50%-cap edge case for `roundRect`/parallelogram/
trapezoid/hexagon/octagon) and the EMU→px/pt/cqw conversion math, directly.

`color.test.ts` is a `happy-dom`-free integration check of a different kind: it stayed in this
package (unlike its former `color.ts` self, see the Layout note above) because it still exercises
something genuinely `to-html5`'s own — `resolveColor`/`resolveFillColor`'s CSS-string output — and
it kept passing completely unchanged through the `color.ts` split, since that split preserved
their public signature and behaviour exactly. `packages/presentation/src/resolve/color.test.ts` is
the new, separate test for the pure `ResolvedColor`-returning resolver the CSS wrapper now calls.

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
integration check for `resolveEffectiveBackground` being applied by `renderSlide`. Its new "text body
anchor and fontRef fallback" block covers `justify-content: center`/`flex-end`/`flex-start` for
`anchor="ctr"`/`"b"`/absent respectively, and a run with no colour of its own picking up its
shape's `p:style/fontRef` colour end-to-end through `renderShapeTreeNode` — `resolveEffectiveAnchor`
and `resolveEffectiveRunProperties`'s `fontRef` ranking (above the master/placeholder chain, below
the shape's own list style) are each covered directly, DOM-free, by
`packages/presentation/src/resolve/text-style.test.ts`. `text.test.ts`'s "bullets and
numbering" block covers char-bullet rendering (glyph, own font/colour), sequential auto-numbering
across sibling paragraphs, a nested sub-list restarting and the outer list resuming its count
after, an explicit `buNone` suppressing a bullet inherited from the shape's own list style, the
ambient-run-property colour fallback when a bullet has no colour of its own, and a trailing/
interior empty paragraph getting no bullet and neither consuming a number nor breaking a running
numbered list — `formatAutoNumber`'s ten schemes and `NumberingState`'s continue/restart/break
rules are covered directly, DOM-free, by `packages/presentation/src/resolve/bullet.test.ts`.

**Known test-environment gap**: `text.test.ts` and `fill.test.ts` can't assert on the actual
`style.fontSize`/`style.borderWidth`/`style.paddingLeft`/`style.textIndent`/`style.strokeWidth`
values `applyRunStyle`/`applyLine`/`applyIndent`/`applySvgLine` set, because those are now `cqw`
values and `happy-dom`'s CSSOM silently drops style assignments in units it doesn't recognize yet
(confirmed real browsers with Container Query Unit support accept them fine — see the `cqw` design
decision above). Both files have a `NOTE` comment at the point this bites; the actual
`emuToCqw`/`fontSizeToEmu` conversion math is still fully covered, DOM-free, in `units.test.ts`.

`presentation-element.test.ts` (new coverage) exercises the slideshow behaviour end-to-end:
`.render()` starts at slide 0 with exactly one `.pptx-slide--active`; `.next()`/`.previous()`
advance/retreat and clamp at the first/last slide (a no-op past either end, not a wraparound);
`.goToSlide()` clamps an out-of-range index the same way; a synthetic `click()` on the element
advances one slide; and a dispatched `keydown` `KeyboardEvent` for `ArrowRight`/space advances,
`ArrowLeft` retreats, and `Home`/`End` jump to the first/last slide.

Its nested `describe('push/fade transitions', ...)` block scopes both `vi.useFakeTimers()`/
`vi.useRealTimers()` and a `HTMLElement.prototype.animate` mock to itself via `beforeEach`/
`afterEach`, so the file's pre-existing synchronous tests are unaffected. The mock — a local
`FakeAnimation` class plus a `vi.fn()` installed directly on the prototype (`happy-dom` has no
`animate` method at all to `vi.spyOn`) — records each call's `target`/`keyframes`/`options` into a
`recordedAnimations` array and returns a `FakeAnimation` whose `.finished` promise resolves via a
`setTimeout` matching the requested duration; a `latestAnimation(el)` helper looks up the most
recently queued recording for a given element, which is what makes the multi-transition tests (two
or three sequential navigations in one `it`) tractable without index arithmetic. Assertions on
_which_ animation was requested (keyframes, duration) read `recordedAnimations`/`latestAnimation`
synchronously right after calling `.next()`/`.previous()`/`.goToSlide()` — `.animate()` itself is
called synchronously inside `goToSlide`, only its `.finished` resolution is async — while assertions
on _post-completion_ state (`await vi.advanceTimersByTimeAsync(durationMs)`, not the synchronous
`vi.advanceTimersByTime()` a plain-timer design could use, since `Animation.finished` is a real
`Promise` whose `.then()` needs a microtask flush between simulated ticks) check classes/`position`
and that a subsequent navigation succeeds, proving `#currentAnimations` was cleared. Coverage: push
forward on both axes and with an omitted (defaulting to `'l'`) direction, push direction reversing
for backward navigation (authoring the transition on the _outgoing_ slide, matching the corrected
which-slide's-transition lookup above — see that section's own note on the bug this fixes), a
dedicated regression test with three slides (push then fade) proving backward navigation undoes the
outgoing slide's own effect kind rather than the destination's, fade forward, `fade`'s
`throughBlack: true` rendering identically to a plain fade, the no-transition and
unsupported-effect-kind cases both still taking the unchanged instant-swap path (no
`pptx-slide--transitioning` class ever appears, `recordedAnimations` stays empty), navigation being
ignored entirely across every input method (click/keydown/`.next()`/`.previous()`/`.goToSlide()`)
while a transition is in flight and succeeding again once it elapses, and the exact
`fast`/`med`/`slow` duration mapping (400/700/1000ms) as a pinned regression contract.

## Next likely steps

1. Widening slide-transition coverage beyond `push`/`fade` — each remaining `TransitionEffect.kind`
   (`wipe`, `cut`, `dissolve`, `wheel`, `split`, ...) is a reasonably self-contained addition to
   `goToSlide`'s dispatch in `presentation-element.ts`, following the `#animatePush`/`#animateFade`
   pattern (now WAAPI-based — see that design decision above); `fade`'s `throughBlack: true`
   two-stage fade-to-black-then-in animation, `advanceOnClick`/`advanceAfter` auto-advance timers,
   and `TransitionSoundAction` playback are separate, currently unstarted pieces of that same
   surface area (see the scope boundary above). Per-slide/per-element animations off `Slide.timing`
   remain a separate, bigger piece of work (root `CLAUDE.md`'s Todos). This session's two enabling
   pieces for a future scroll-driven-playback feature — a central duration-resolution API in
   `@pptx2html/presentation` (`resolve/timing.ts`, see that package's CLAUDE.md) and migrating
   slide-transition playback off plain CSS `transition` onto the Web Animations API — are both
   done; read root `CLAUDE.md`'s "Future feature: scroll-driven playback" section /
   `docs/scroll-driven-playback.md` before picking up either of the two items above, since the same
   constraints (duration logic stays in `@pptx2html/presentation`, prefer WAAPI over plain CSS
   `transition`/`@keyframes`) apply to them too.
2. `TextBodyProperties.wrap` → `white-space`, table cell fill, table styles — `anchor` is now done
   (see "Key design decision: a shape's text body is vertically anchored via flexbox" above).
3. Connector line rendering (see the scope boundary above) — likely an SVG line/path overlay
   sized to the connector's own box, reusing `applyLine`'s color/width/dash resolution but not its
   `border-*` output. `shape-tree.ts`'s new `renderShapeOutline` (an `<svg>` sized to a shape's own
   box, see above) is a directly reusable pattern for this now that it exists.
4. Widening `shape-geometry.ts`'s preset coverage — arrows (`rightArrow`/`leftArrow`/etc.), other
   star counts (`star4`/`star6`/etc.), and flowchart shapes are all common in real decks and not
   yet modeled; each is a straightforward addition to `presetShapePath`'s switch.
5. Gradient/pattern/blip fill support on the nine SVG-path presets (currently solid-only, see the
   scope boundary above) — an SVG `<linearGradient>`/`<pattern>` `<defs>` entry referenced via
   `fill="url(#id)"` would parallel `resolveGradientCss`'s CSS-gradient output without the
   rectangular-background constraint, but needs its own angle/stop-position math since SVG
   gradients use `objectBoundingBox` coordinates, not CSS's `linear-gradient()` syntax.
6. Clipping a picture to one of `shape-geometry.ts`'s nine SVG-path presets (not just `roundRect`/
   `ellipse`'s native `border-radius`) — needs an `<svg><image>` + `<clipPath>` overlay in place of
   the plain `<img>` `renderPicture` emits today.
7. `p:style/effectRef` (see the scope boundary above) needs effect rendering to exist at all
   first (a bigger, separate gap) — `fontRef` is done, see "Key design decision: fontRef
   text-colour/typeface fallback lives in `resolveEffectiveRunProperties`, not a new resolver"
   above.
8. Wiring a connector's own `p:style` (already parsed, `ShapeStyle` is on `ConnectionShape` too)
   into whatever connector line rendering eventually lands (see item 3 above) — today it's parsed
   but unused, since `renderConnector` doesn't call `applyFill`/`applyLine`/`effectiveFill`/
   `effectiveLine` at all yet.
9. `apps/web-demo` now renders the parsed presentation into the page (see its own source) —
   confirm this still looks right whenever this package's DOM structure changes, and now also
   gains a working slideshow (click/keyboard navigation) for free once it re-renders with this
   change, since `renderPresentation`'s returned element already wires it up.
