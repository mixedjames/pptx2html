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
real SVG `<path>` outline (solid fill/stroke only — see the scope boundary). A shape's freeform
`custGeom` outline (§20.1.9.8, new — a later session) renders the same way, via
`shape-geometry.ts`'s `customGeometryPath` — including a boolean Merge Shapes result (multiple
`a:path` subpaths in one outline, rendering a cutout under SVG's default nonzero fill rule), the
concrete case that motivated adding it (`apps/web-demo/src/Presentation1.pptx`'s slide 3). A slide's own
background renders too, falling back through layout then master (`@pptx2html/presentation`'s
`resolve/background.ts`), reusing the
same `fill.ts` machinery. Every absolute magnitude this pass introduces — font size, border width,
list indentation — scales with the slide via CSS container query units rather than a fixed px/pt
(see "Key design decision: absolute sizes scale via `cqw`" below), consistent with how
position/size already scale via percentages. Table cell/table styling is still unrendered —
deliberately deferred to a later pass. `<pptx-presentation>` now renders as an actual
slideshow rather than every slide stacked one below the next — one slide visible at a time,
advanced by click/keyboard (see "Key design decision: slideshow navigation" below). Navigating
between slides now (new) plays `Slide.transition`'s own effect when it's `push`, `fade`, or (new)
`morph` (§19.3.1.49, `p:transition` — see "Key design decision: push/fade slide transitions" and
"Key design decision: Morph transitions" below) — a real animation, driven by the Web Animations
API rather than a plain CSS `transition` (see those same design decisions for why), rather than the
instant swap every other effect kind (and a slide with no `transition` at all) still falls back to.
Morph is the first of these to genuinely need a shape-level correspondence between two slides,
supplied by `@pptx2html/presentation`'s `resolveMorphMatch` and reduced to plain, freestanding-
HTML-safe data before this package ever stores it — see that design decision for the full
mechanism and why that reduction is a hard requirement, not a preference. `Slide.timing`'s
per-element/per-build animation tree
(new) is now partially consumed — a `p:animEffect` whose `filter` is `"fade"` targeting a whole
shape plays automatically, via the Web Animations API, as soon as its slide becomes active (see
"Key design decision: `Slide.timing` fade animations" below) — every other behavior kind, every
other `animEffect` filter, and implicit paragraph/graphic builds are still unconsumed and reported
as unsupported. See root `CLAUDE.md`'s Todos for the remaining scope.

**Scroll-driven playback (new) is also implemented now**, as a second, separate custom element —
`<pptx-scroll-presentation>` (`scroll-presentation-element.ts`, `renderScrollPresentation`) —
alongside the click-driven `<pptx-presentation>` above, not a mode on it. Every slide transition and
build animation this package already knows how to play is positioned on one scrubbable
absolute-millisecond timeline (`scroll-timeline.ts`'s `resolveScrollTimeline`) instead of being
fired-and-forgotten in real time; scrolling a track inside the element's own shadow DOM sets each
`Animation`'s `currentTime` directly. See "Key design decision: scroll-driven playback" below for
the full mechanism, and `docs/scroll-driven-playback.md` for how this differs from that document's
original framing (notably: no "is this deck fully time-resolved" gate at all — see that design
decision for why).

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
  for a consistent screen size regardless of the slide's own size, plus `vector-effect:
non-scaling-stroke` so that stroke renders uniformly despite the path's own non-uniformly-
  stretched coordinate space — see "Key design decision: SVG stroke width uses `vector-effect:
non-scaling-stroke`" below); gradient/pattern/blip fills and a
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
  the former, or now first tries `customGeometryPath` (new — see below) for a `custGeom` shape.
  `customGeometryPath` renders `@pptx2html/presentation`'s `CustomGeometry.pathLst` (§20.1.9.8)
  into the same `0 0 100 100` space: each `a:path`'s local `w`/`h` coordinate space scales onto it
  the same way `presetShapePath`'s fixed space does, its `moveTo`/`lnTo`/`quadBezTo`/`cubicBezTo`/
  `close` commands map directly onto SVG's `M`/`L`/`Q`/`C`/`Z`, and multiple `a:path` entries
  concatenate into multiple subpaths within one `d` string — which SVG's default nonzero fill rule
  renders as a hole wherever they overlap with opposite winding, exactly what a boolean "Subtract"
  shape (an outer outline plus an inner cutout — PowerPoint's own `custGeom` output for a Merge
  Shapes operation, e.g. `apps/web-demo/src/Presentation1.pptx`'s slide 3) needs, with no separate
  fill-rule logic required. `arcTo` (§20.1.9.9) is the one command needing real conversion, not a
  1:1 syntax mapping: OOXML defines it as a portion of an ellipse (radii `wR`/`hR`) that the
  current pen position sits on at angle `stAng`, swinging `swAng` further, so the ellipse's centre
  and end point are derived from the current pen position before being handed to SVG's own
  `A rx ry 0 large-arc-flag sweep-flag x y` — not guaranteed bit-exact (the same tier as this
  file's other underspecified-magnitude approximations) but not a stand-in either, it's the actual
  ellipse math. A path whose `w`/`h` is `0` (nothing to scale against) is skipped;
  `customGeometryPath` returns `undefined` if nothing in `pathLst` was renderable, same fallback as
  an unmodeled preset. `nativeBorderRadius` returns a CSS `border-radius` percentage for `roundRect` (from
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
  `shape-geometry.ts`'s `presetShapePath` (falling back to `customGeometryPath` for a `custGeom`
  shape, new) against the shape's own `geometry`: if either returns a path (one of the nine
  non-rectangular presets `presetShapePath` models, or a `custGeom` with usable `pathLst` data),
  `renderShape` appends `renderShapeOutline`'s
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
  design decision below). `renderShapeTreeNode` (new) also stamps every rendered element with
  `dataset.pptxShapeId` — its own `nonVisual.id`, the same id an `AnimationTarget` addresses — so
  `presentation-element.ts`'s `#playSlideAnimations` can find a fade animation's target element by
  a plain `querySelector('[data-pptx-shape-id="…"]')` within the slide's own subtree.
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
  (`#updateActiveSlide`). `#updateActiveSlide` (new) also calls `#playSlideAnimations`, which plays
  the newly-active slide's fade animations (see `animation.ts` and "Key design decision:
  `Slide.timing` fade animations" below) — this runs on every path that makes a slide active: the
  initial `render()`, an instant no-transition swap, and `#finalizeTransition` once a push/fade
  slide transition settles.
- `animation.ts` — pure logic, no DOM: `collectFadeAnimations` walks a slide's `Slide.timing` tree
  (`@pptx2html/presentation`'s `SlideTiming`/`TimeNode`) collecting the one behavior this renderer
  plays — a `p:animEffect` (§19.7.9) with `filter === 'fade'` targeting a whole shape — into a flat
  `ShapeFadeAnimation[]` (`shapeId`/`direction`/`delayMs`/`durationMs`), and reports everything else
  it walks past (other filters, non-shape targets, every other behavior kind, an `'indefinite'`
  duration, an unconsumed `buildList`) via its `report` callback — see "Key design decision:
  `Slide.timing` fade animations" below for what's approximated and why.
- `transition-keyframes.ts` — pure keyframe/transform helpers factored out of
  `presentation-element.ts` (`offscreenTransform`, `IDENTITY_TRANSFORM`, `REVERSE_SIDE_DIRECTION`,
  `morphKeyframes`, `readShapeBox`, `findShapeElement`) so both it and `scroll-presentation-
element.ts` produce pixel-identical `push`/`fade`/`morph` visuals from one source. Nothing here
  mutates the DOM (`findShapeElement`/`readShapeBox` only read), which is what makes reuse from a
  scrubbed (not just fire-and-forget) player possible.
- `contain-size.ts` — `observeContainSize`, a small `ResizeObserver`-backed helper shared by
  `presentation-element.ts` and `scroll-presentation-element.ts` that keeps their respective sized
  container letterboxed/pillarboxed within whatever box their host is given, toggling between
  width-driven and height-driven sizing depending on which one actually fits — see "Key design
  decision: every element's position is a CSS percentage of the slide" below for why this needed a
  (deliberately minimal) `ResizeObserver` rather than being CSS-only like everything else in this
  package.
- `scroll-timeline.ts` — pure logic, no DOM: `resolveScrollTimeline` assembles a whole
  `Presentation` into one absolute-millisecond `ScrollTimeline` (a `ScrollSegment` per slide, each
  with an optional `transition` window and a `content` window) for `scroll-presentation-element.ts`
  to scrub through — see "Key design decision: scroll-driven playback" below for the full
  mechanism, in particular why this needed no new resolvers in `@pptx2html/presentation` at all.
- `scroll-presentation-element.ts` — `PptxScrollPresentationElement`, the `<pptx-scroll-
presentation>` custom element — see "Key design decision: scroll-driven playback" below.
- `index.ts` — barrel + `renderPresentation`/`renderScrollPresentation`, which register the
  respective element and return one instance with `.render()` already called.

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

**Bug, a later session: width-only scaling cropped the deck on a host proportionally wider than the
deck's own aspect ratio.** `.pptx-slide` at `width: 100%` with `aspect-ratio` locking height means
height is _purely_ derived from whatever width the container hands it, with no ceiling — fine for
this component's whole history, since it was only ever embedded with an auto-height host (no fixed
vertical box), so the page just grew as tall as the content needed. That assumption broke once
`<pptx-presentation>`/`<pptx-scroll-presentation>` started being given an _explicit_, fixed
viewport-height box (`apps/web-demo`/`apps/pages`'s full-screen layouts, this same session): on an
especially wide viewport, `width: 100%` still forces a wide box, `aspect-ratio` still derives a
correspondingly tall height from it — and that height can exceed the fixed box it's actually been
given, with nothing capping it, cropped by `overflow: hidden`.

A first fix attempt gave `.pptx-presentation` `max-width: 100%; max-height: 100%` and no explicit
width/height of its own, reasoning that `aspect-ratio` would fill in a size satisfying both caps —
**this doesn't work, and was a second regression (the deck stopped rendering at all) caught
immediately after shipping the first one.** `max-width`/`max-height` alone are only ever a ceiling
on whatever size the normal sizing algorithm would otherwise produce for a _non-replaced_ element
(a plain `div`, unlike `img`/`video`, which get this exact behavior for free from `object-fit:
contain`) — they don't themselves supply a preferred size to size toward. With no explicit
width/height at all, a block box falls back to shrink-to-fit sizing based on its content's
intrinsic size — and `.pptx-presentation`'s content (`.pptx-slide`, itself `width: 100%` of its
parent) is entirely percentage-based, contributing nothing to a shrink-to-fit computation, since a
percentage can't resolve without already knowing the size it's a percentage _of_. Net effect: the
container collapsed to near-zero size instead of filling the available space.

**The actual fix is a small `ResizeObserver`, `contain-size.ts`'s `observeContainSize`** — used
identically by both `presentation-element.ts` (`#slidesContainer`) and
`scroll-presentation-element.ts` (`#viewport`), the one other place with the exact same bug
(`.pptx-scroll-viewport` was `top: 0; left: 0; width: 100%`, no height ceiling, for the same
reason). Rather than trying to express "contain-fit, whichever axis is binding" as a single CSS
rule (which turned out to have no reliable non-replaced-element equivalent — see above), it
toggles between two states that are each individually simple, standard, and unambiguous: the
default `width: 100%; height: auto` (letting `aspect-ratio`, still set inline as before, derive
height — identical to the original pre-bug behavior), switching to `width: auto; height: 100%`
only when that default would actually _overflow_ the host's own available height. Measuring actual
rendered overflow, rather than comparing the host's and deck's aspect ratios numerically ahead of
time, is what makes this self-adapt to both cases without needing to know _why_ a host does or
doesn't have an independent height: a host with no independent height (the traditional
just-embedded-in-a-page case) always has its own height derived from `.pptx-presentation`'s own
content, so the two can never disagree — this never switches modes there, exactly preserving
today's behavior; only a host with a genuinely independent, smaller height (the new full-viewport
case) can ever trigger it. `:host` also switched to `display: grid; place-items: center` (from
`display: block`) so the sized-and-toggled child centers itself — letterboxed or pillarboxed as
needed — instead of sitting flush in a corner whenever it ends up smaller than the host on either
axis. See `contain-size.ts`'s own doc comment for the full reasoning (including why this needed
JS at all, given this package's general "no JS resize handling" preference — see the design
decision above) and `contain-size.test.ts` for coverage of the actual decision logic (`clientHeight`/
`getBoundingClientRect` are stubbed directly, since `happy-dom` does no real layout).

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

## Key design decision: SVG stroke width uses `vector-effect: non-scaling-stroke`

An actual bug, caught by the user right after `custGeom` support (above) shipped: a shape's SVG
outline stroke (`applySvgLine`, `fill.ts`) rendered visibly thicker on some edges than others —
e.g. thicker along vertical edges than horizontal ones — whenever the shape's own box wasn't
square. The cause is `shape-geometry.ts`'s fixed `0 0 100 100` viewBox, `preserveAspectRatio="none"`
stretched non-uniformly onto the shape's actual (usually non-square) box (see `shape-tree.ts`'s
`renderShapeOutline`): a stroke is computed by offsetting the path in that _pre-transform_ user
space, and only then does the browser apply the viewBox's own scale to get to screen pixels — so a
constant-width offset in user space ends up a different number of screen pixels depending on which
axis it's measured along, whenever that scale isn't the same in both directions. Giving
`stroke-width` a real CSS length (`cqw`, already the case before this fix) does **not** avoid this:
a CSS length used for an SVG geometry property still resolves into that same pre-transform user
space, then gets warped by the same non-uniform transform as everything else — the file's own
former doc comment claimed otherwise, which was the actual root cause of the bug shipping
unnoticed. The real fix is `vector-effect="non-scaling-stroke"` (an SVG attribute, not a CSS
property — broader support), set on the path whenever `applySvgLine` draws a stroke at all: it
computes the stroke (both its width and its dasharray pattern) _after_ the transform instead of
before, so it renders as the same physical thickness on every edge regardless of the path's own
aspect ratio — restoring the "the SVG-native equivalent of `applyLine`'s CSS `border-width`"
property the code already claimed to have, correctly this time. `fill.test.ts`'s `applySvgLine`
block covers that the attribute is set exactly when a stroke is actually drawn (not for an explicit
`noFill` line).

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

## Key design decision: `Slide.timing` fade animations

`Slide.transition` (above) is a single whole-slide effect; `Slide.timing` (§19.3.1.48, `p:timing`)
is a different, much richer thing — a whole tree of per-element/per-build animations (PowerPoint's
Animation Pane), where each leaf behavior targets one shape (or its text/background) and starts
per its own click/delay/previous-effect condition. This first pass plays exactly one leaf behavior
kind out of that tree: a `p:animEffect` (§19.7.9) whose `filter` is `"fade"` and whose target is a
whole shape (`AnimationTarget`'s `{ kind: 'shape' }` case) — the common "Fade In"/"Fade Out"
entrance/exit effect. Everything else the tree can contain — every other `animEffect` filter (wipe,
blinds, ...), every other behavior kind (`set`/`anim`/`animClr`/`animMotion`/`animRot`/
`animScale`/`cmd`/`audio`/`video`), a fade targeting text/background rather than a whole shape, and
`SlideTiming.buildList`'s implicit paragraph/graphic builds — is walked past and reported via
`reportUnsupported`, not silently dropped (`animation.ts`'s `collectFadeAnimations`).

**Deliberately not a faithful implementation of PowerPoint's click-driven build system.** A real
Animation Pane sequence nests effects inside `par`/`seq` containers whose own start conditions
compose down the tree — an "After Previous" effect three levels deep waits for everything ahead of
it, and an "On Click" effect waits for an actual click before starting. `to-html5` has no concept
of an in-slide "build step" at all — a click always advances to the _next slide_ (see "slideshow
navigation" above), never to the next animation within the current one — so faithfully modeling
click-gated sequencing isn't an option without first inventing that concept, a materially bigger
change than "start with fade" calls for. Instead, `collectFadeAnimations` treats every fade leaf
node independently: its delay is `resolveTimeNodeStartMs`'s _own local_ answer for that node alone
(`@pptx2html/presentation`'s `resolve/timing.ts`, new — see that package's CLAUDE.md), never
composed with any ancestor `par`/`seq`/`excl` container's own offset, and a node gated on
`onClick`/`onNext`/etc. with no numeric delay fallback still plays — immediately, at `delayMs: 0`
— rather than waiting for a trigger nothing here can satisfy. `animation.ts`'s own doc comment
covers the same ground; both approximations are reported once per slide
(`animation-trigger-unmodeled`), not per node, to avoid flooding the log on a deck with many
click-triggered effects.

**Playback is triggered from the one place a slide actually becomes visible.**
`PptxPresentationElement`'s `#updateActiveSlide` — already the single method every navigation path
funnels through, whether an instant swap, the initial `render()`, or `#finalizeTransition` settling
a push/fade slide transition — now also calls `#playSlideAnimations(this.#currentIndex)`. This
means a slide's fade animations replay in full every time that slide becomes active again,
including navigating back to a previously-visited slide — closer to how PowerPoint itself replays
a slide's animations on (re-)entry than a "play once ever" model would be, and it comes for free
from reusing the existing navigation chokepoint rather than adding a new one.

**Target lookup is a plain `querySelector`, keyed by `data-pptx-shape-id`.** `shape-tree.ts`'s
`renderShapeTreeNode` (see above) stamps every rendered element with its own `nonVisual.id` — the
same id space `AnimationTarget.shapeId` addresses — so `#playSlideAnimations` finds a fade's target
with `slide.querySelector('[data-pptx-shape-id="…"]')` scoped to that slide's own subtree, rather
than needing a shapeId → element map threaded down through `renderSlide`.

**Playback uses the Web Animations API (`Element.animate()`), fire-and-forget, not tracked or
cancellable.** Consistent with this package's push/fade slide-transition playback (see above) and
the same scroll-driven-playback motivation (a WAAPI `Animation`'s `currentTime` is directly
seekable — see `docs/scroll-driven-playback.md`). Unlike slide transitions, though, fade
animations are **not** added to `#currentAnimations` and never cancelled — `render()`'s existing
"cancel every in-flight animation" cleanup only ever targeted slide-transition animations
(mid-transition re-render is the scenario it guards), and a fade animation is a short, one-shot
`fill: 'forwards'` effect with nothing else competing to interrupt it. Revisiting a slide queues a
_new_ `Animation` on the same element/property each time rather than reusing or explicitly
disposing of the previous one — by WAAPI's own composite-ordering rules the newest animation wins
visually once it starts, so this doesn't look wrong, but it's the same class of "stale animation
object left attached" debt `#finalizeTransition`'s own doc comment flags for slide transitions, not
yet paid down here. Fine at the scale a single slide's animations run — not revisited for a slide
with many fade-triggering navigations in one session.

## Key design decision: Morph transitions

`goToSlide` now plays PowerPoint's Morph transition (`MorphTransitionEffect`) too, via new
`morph.ts` and `#animateMorph`, when `@pptx2html/presentation`'s `resolveMorphMatch` found a
confident-enough shape correspondence between the two slides — otherwise it degrades to a plain
crossfade (`#animateFade`). Morph is a fundamentally different kind of transition from every other
one this package renders: `push`/`fade` are a single canned animation applied to one slide, but
Morph requires a _correspondence between two slides' shapes_ — computed at `render()` time, not
navigation time, and reduced away to plain data before being retained (see the next paragraph for
why that reduction is a requirement, not a style choice).

**No dependency on the `Presentation`/`Slide` object graph survives `render()` — this was an
explicit, user-stated requirement, not an incidental design preference.** A key target for this
project's output is a freestanding HTML file; every other piece of state `PptxPresentationElement`
retains after `render()` (`#transitions`, `#animations`) was already small, flat, JSON-serializable
data extracted from the graph, never the graph itself. An earlier draft of this feature considered
storing `#slideGraphs: readonly Slide[]` and calling `resolveMorphMatch` lazily inside `goToSlide` —
rejected before being implemented, specifically because it would have been the first thing in this
element to hold the live object graph after creation. Instead, `render()` calls
`resolveSlideMorphMatch(previousSlide, slide, report)` (`morph.ts`, new) once per slide, up front
(same "log is complete even for slides never visited" principle `reportSlideLevelFeatures` already
follows for `transition`/`timing`), which immediately reduces `resolveMorphMatch`'s actual
`ShapeTreeNode` references down to a `MorphMatchSummary` — plain shape-id numbers only
(`{ matched: { outgoingShapeId, incomingShapeId }[], disappearingShapeIds, appearingShapeIds }`) —
before anything is stored in `#morphMatches`. `#animateMorph` (below) never touches `Slide`/
`ShapeTreeNode` at all, only these plain ids plus `data-pptx-shape-id` DOM lookups.

**Confidence check and crossfade fallback.** `resolveSlideMorphMatch` also owns the "is this match
good enough to actually play" decision (`MIN_MORPH_MATCH_RATIO`, currently 1/3 of the larger side's
shape count) — deliberately _not_ triggered by the mere presence of appearing/disappearing shapes,
since real Morph slides routinely add/remove a shape on purpose (that's normal authoring, not a
failure); it's triggered by an overall low match _rate_, the signal that the name-matching
heuristic likely had little to go on at all. Below that threshold — or when there's no previous
slide to morph from (Morph authored on the deck's first slide, nonsensical but not something to
crash on) — it reports `morph-match-degraded` via `UnsupportedFeatureCollector` and `goToSlide`
falls back to `#animateFade`, per the user's explicit requirement that a failed/degraded match be
just as visible in the log as any other unsupported feature, not a silent approximation.

**Box/fill values are read back off each shape's own already-rendered inline style, not
recomputed.** `readShapeBox` reads a matched shape element's own `left`/`top`/`width`/`height`/
`transform` — set once by `shape-tree.ts`'s `positionElement` at render time — directly as CSS
strings, passed straight through into WAAPI keyframes with no numeric parsing/interpolation logic
of this file's own. This mirrors `coordinate.ts`'s own "every element's box independently readable
from its own inline styles" design decision (see above) — `#animateMorph` needed no new
coordinate-resolution path at all, just a second reader of state `shape-tree.ts` already computes.
Fill/colour interpolation is **not implemented this round** — only position/size/rotation/opacity
animate; a matched shape's background colour, gradient, or SVG-outline fill jumps instantly rather
than tweening, the same "ship the dominant visual signature first, defer fill nuance" tier as this
package's other incremental passes (e.g. the nine SVG-path presets' solid-only fill).

**The arriving copy alone moves, fully opaque throughout; the departing copy is hidden instantly,
not faded — crossfading both was a real bug, caught by the user seeing text through a large solid
shape mid-transition.** Like `push`/`fade`, both the outgoing and incoming slide stay
simultaneously visible and independently animatable during a Morph transition
(`#beginTransitionFrame`). The first version of `#animateMorph` animated _both_ copies of a matched
pair between the same two boxes, fading 1→0 (departing) and 0→1 (incoming) respectively — the exact
technique `#animateFade` correctly uses for whole-slide crossfades, just applied per-shape. That's
wrong for a matched pair: alpha-compositing two _overlapping, opaque_ copies makes both visibly
translucent for the middle of the tween, wrongly revealing whatever sits behind them — obvious on a
large solid shape, which is exactly how the user caught it (a big orange shape that should only
translate was visibly showing text through it partway through). A shape that persists between two
slides should look like one continuous, fully opaque object gliding, the way real PowerPoint Morph
renders it — not two ghosts blending. The fix: only the _arriving_ copy animates
(`morphKeyframes`, deliberately with no `opacity` field at all, so it never leaves full opacity);
the _departing_ copy gets a single zero-duration `Animation`
(`departingEl.animate([{ opacity: 0 }], { duration: 0, fill: 'forwards' })`) — an instant switch,
not a fade, timed to coincide exactly with the arriving copy's first frame (`morphKeyframes`'
`from` endpoint is always the departing copy's own box), so nothing is ever visibly uncovered. A
zero-duration `Animation` is still a real one (`.finished` resolves, `.cancel()` works), so it
needs no special-casing anywhere else. `disappearingShapeIds`/`appearingShapeIds` are unaffected by
any of this — those genuinely have no "moving box" interpretation (an object being added or
removed, not relocated), so a smooth opacity fade is the _correct_ rendering for them, reusing the
same plain single-keyframe fade `#playSlideAnimations` already established for `Slide.timing`
fades.

**Departing/arriving roles are fixed to the DOM parameters, not to `match`'s own labels — a real
bug caught by a backward-navigation test before it shipped.** `resolveSlideMorphMatch` always
matches "earlier-in-deck slide" against "later-in-deck slide" (`outgoingShapeId`/`incomingShapeId`
name-space), regardless of which direction the presentation actually navigates — but
`#animateMorph`'s own `outgoing`/`incoming` parameters always mean "currently shown, departing"/
"about to show, arriving", exactly like `#animatePush`/`#animateFade`. An earlier version of this
method conflated the two, remapping _which DOM element_ played which role for backward navigation
(`outgoingSideEl = forward ? outgoing : incoming`) — which silently inverted the fade direction and
looked up disappearing/appearing shapes in the wrong slide's subtree whenever navigating backward
through a morph. The fix keeps `outgoing`'s shapes always fading 1→0 and `incoming`'s shapes always
fading 0→1 (matching every other transition in this file), and instead swaps only _which shape id_
(`departingShapeId`/`arrivingShapeId`, `fadeOutIds`/`fadeInIds`) gets looked up in each, based on
`forward`. Covered by `presentation-element.test.ts`'s "reverses departing/arriving roles correctly
for backward navigation" test — written specifically because reasoning through this by hand is
exactly the kind of thing that's easy to get backward.

**Every `Animation` `#animateMorph` produces — matched-pair tweens and plain
disappearing/appearing fades alike — feeds into the existing `#awaitTransition`/
`#finalizeTransition` bookkeeping, not a new tracking mechanism.** This matters for correctness,
not just code reuse: `#finalizeTransition`'s existing cancellation loop is what reverts every
animated shape's inline-style override back to its own base style once the transition settles.
Without that, a departing shape (whose finished, `fill: 'forwards'` animation left it stuck at
opacity `0`) would stay invisible the next time its own slide became active again — cancelling
drops the override and falls back to the element's own unmodified inline style, which is always
the correct resting state (see
`#finalizeTransition`'s own doc comment for why this is true of every property this file ever
animates). `#playSlideAnimations`'s fade animations are the one exception to "always track and
cancel," and deliberately so — see that section's own note on why a short, non-repositioning,
one-shot fade doesn't need it the way a departing morph shape's position override does.

**Known gap: a Morph transition triggered via a non-adjacent `goToSlide()` jump can animate against
the wrong DOM.** `resolveSlideMorphMatch` computes each slide's match against its immediate
predecessor in the deck (`presentation.slides[slideIndex - 1]`) — correct for `.next()`/`.previous()`
sequential navigation, which is how a real presentation is actually driven. But `goToSlide(n)`
allows jumping to _any_ index from wherever `#currentIndex` currently is (same as it always has for
`push`/`fade`, not new to Morph) — if that jump lands on a Morph-transitioned slide from somewhere
other than its immediate predecessor, `#animateMorph` will look for shape ids that were matched
against a slide that isn't actually the one being shown as `outgoing`, and simply skip any it can't
find (`findShapeElement` returning `null`) rather than animating them. Not observed to crash or
mis-animate anything it _does_ find — just silently animates fewer shapes than a sequential
approach to the same destination would. Not fixed this round: real presentations are driven
sequentially, and a general fix (matching against whichever slide is _actually_ current at
navigation time, not a fixed predecessor) would mean computing `resolveMorphMatch` at navigation
time again, reintroducing exactly the live-object-graph dependency this design deliberately avoided
— see the "no dependency on the object graph" paragraph above.

## Key design decision: scroll-driven playback

`docs/scroll-driven-playback.md` scoped this as a future feature and recorded two pieces of
enabling groundwork (the central duration-resolution API in `@pptx2html/presentation`'s
`resolve/timing.ts`, and migrating slide-transition playback onto the Web Animations API) without
building the feature itself. This section covers what actually got built, including one real
reframing of that document's own original premise.

**No "is this deck fully time-resolved" gate — scroll position replaces the click/auto-advance
trigger entirely, so a deck's authored click-gating becomes irrelevant.** The design note's
original framing assumed the feature could only work for a deck where every wait had already been
replaced by a number (no `onClick`/`onNext` anywhere) — reasoned to be the _uncommon_ case for a
real `.pptx`. That assumption turned out to be the wrong frame: reaching a point on the scroll axis
_is_ the trigger, the same way a click already is for `PptxPresentationElement` — so a slide's
`advanceOnClick`/`advanceAfter` and a `Slide.timing` node's `onClick`/`onNext` start condition are
simply never consulted by `resolveScrollTimeline` at all, for either kind of segment. What's
actually required is much narrower: every segment needs a _duration_ (to occupy scroll distance)
and a _sequence position_ (to land in the right order) — both already exist or default sensibly for
virtually every deck, via machinery this package already had:

- `resolveTransitionDurationMs` already defaults an absent `speed` to `fast` (400ms), even with
  zero authored transition data at all.
- `collectFadeAnimations` (`animation.ts`, unchanged, reused verbatim) already resolves a
  click-gated fade's start via `resolveTimeNodeStartMs` and, when that's `'indefinite'`, already
  treats it as `delayMs: 0` and plays it immediately — exactly the "sequence, not full timing"
  policy scroll mode needs, for free.
- A slide with **no transition at all**, or one whose `effect.kind` isn't among the three this
  package actually animates (`push`/`fade`/`morph`), gets a **synthetic default of `push`
  (direction `'u'`)** in scroll mode specifically — a _different_ fallback than real-time mode's
  instant `display: none`/`block` swap, chosen because an abrupt cut reads badly mid-scrub and a
  vertical push matches the scrollytelling mental model directly (scrolling down brings the next
  slide up into view). Reported via `transition-effect-approximated-for-scroll` only when an
  authored-but-unsupported effect was actually substituted — a wholly absent transition is not
  "unsupported," so it's silently defaulted, matching how real-time mode never reports a slide with
  no transition either.

Net effect: **`resolveScrollTimeline` needed zero new resolvers in `@pptx2html/presentation`.**
Everything genuinely reusable (`resolveTransitionDurationMs`, `resolveTimeNodeStartMs`,
`resolveTimeNodeDuration` via `collectFadeAnimations`, `resolveMorphMatch` via
`resolveSlideMorphMatch`) already existed and was already renderer-agnostic — this is a new
composition of existing primitives plus new orchestration/DOM, entirely inside this package, the
same "resolution logic lives with the model, rendering policy lives with the renderer" split this
file's other design decisions already follow.

**A separate custom element, not a mode on `PptxPresentationElement`.** Click-driven navigation
(`goToSlide`'s mid-transition blocking, `#currentAnimations`) and continuous scroll-scrub are
different enough state machines that folding scroll mode into the existing class would just mean
two code paths inside it. `PptxScrollPresentationElement` reuses `renderSlide` for the actual
per-slide DOM — no rendering logic is duplicated, only orchestration is new — and, like
`PptxPresentationElement`, keeps no dependency on the `Presentation`/`Slide` object graph after
`render()` returns: `resolveScrollTimeline`'s `ScrollSegment`s already only carry plain data
(numbers, `ShapeFadeAnimation`, `MorphMatchSummary`, `TransitionEffect`), the same reduction
`morph.ts`'s own doc comment explains is a hard requirement, not a preference, for this project's
freestanding-HTML-file export goal.

**Every `Animation` is created once, paused immediately, and scrubbed via `currentTime` —
`render()` builds them all up front, `seekTo(ms)` never calls `.animate()` again.** Unlike
`PptxPresentationElement`, which fires a new `Animation` per navigation and lets it play out in real
time, this element creates exactly one `Animation` per transition-segment participant (the
`push`/`fade` pair, or a Morph segment's per-shape matched/disappearing/appearing animations) and
per content-phase fade at `render()` time, via `el.animate(keyframes, { duration, fill: 'both' })`
immediately followed by `.pause()`. `seekTo(ms)` — the core, directly testable entry point, with no
DOM-measurement dependency of its own — locates the active segment (`locateSegment`, a pure module-
level function) and just writes `animation.currentTime` on whichever animations are relevant, using
`fill: 'both'` so a paused animation holds its start/end frame correctly outside `[0, duration]`
too. One consequence worth calling out: **scrubbing backward through a transition needs no separate
"reverse" logic at all** — unlike `PptxPresentationElement`'s `REVERSE_SIDE_DIRECTION`-based replay
for backward navigation, setting a _smaller_ `currentTime` on the exact same `Animation` the user
scrubbed forward through already shows it partway undone, for free, as a direct consequence of
WAAPI's own scrubbing semantics.

**Morph's departing-shape hide is a direct `opacity` style write, not a third `Animation` — and
needs its own explicit reset, unlike everything else in this element.** Real-time `#animateMorph`
hides a matched pair's departing copy via a zero-duration `Animation`, safe there because exactly
one transition is ever "in flight" and `#finalizeTransition`'s cancellation loop reverts it once
that transition settles. Scroll mode has no such moment — any segment can become active or inactive
repeatedly as the user scrubs back and forth, so a `fill: 'both'` Animation permanently pinning that
shape's opacity to 0 would never naturally un-pin itself once the user scrubs back into that slide's
_own_ content phase (where the shape is no longer "departing," just present). `seekTo` instead
writes `opacity: '0'` directly on each matched pair's departing element only while that transition's
window is the active state, tracks which elements it touched in `#hiddenMorphShapes`, and resets
them (`opacity: ''`) at the very start of the _next_ `seekTo` call before that call's own active
state re-applies whatever's needed — so a shape reset can never lag a frame behind the state that
should own it.

**No `position: sticky`/`fixed` anywhere, and — since a later session — the visible content isn't
even a descendant of the scrolling element.** The first version of this element nested
`.pptx-scroll-viewport` (the actual rendered slides) inside `.pptx-scroll-track` (the scrolling
element), compensating for scroll with a manual `transform: translateY(scrollTop)` write every
frame — a deliberate substitute for `position: sticky`/`fixed` (both known sources of
cross-browser inconsistency, and both hand control of "where does this sit on screen" to the
browser's own reflow timing, the wrong trade for a feature whose later goal is _exact_, frame-synced
"perfect" pinning). **That version visibly jittered**, caught by the user in actual browser use: the
rendered slide subtree is real, non-trivial DOM, and having it live inside the actively-scrolling
element meant the browser's own native scroll compositing of that subtree and this class's
once-per-frame JS correction were two independent, not-quite-synchronized sources of truth for the
same box. The fix went a step further than the original plan: **`.pptx-scroll-viewport` is now a
sibling of `.pptx-scroll-track`, not its descendant, and needs no positioning write on scroll at
all** — since it was never part of the scrolling subtree, it simply never moves, no transform
required to keep it visually put. `.pptx-scroll-track` (`overflow-y: auto`) is still the one real
scroll region, a child of this element's own shadow DOM, not `window`/`document` — the host page
must give `<pptx-scroll-presentation>` an explicit box (e.g. `height: 100vh`) and must not nest it
inside another scrollable ancestor, since this element assumes it owns the one scroll region
involved — but now it holds nothing but `.pptx-scroll-spacer`, a plain in-flow div whose only job is
to _be_ the scroll distance (`height: totalDurationMs * pixelsPerSecond / 1000`), no visible content
at all. Both `.pptx-scroll-track` and `.pptx-scroll-viewport` are `position: absolute` against
`:host` (now `position: relative` for exactly this), sized to cover the same box — the track sits on
top (`z-index: 1`) purely so it keeps hit-testing scroll/wheel/touch gestures anywhere over the
visible area (wheel/touch scroll delegation only walks up the ancestor chain from whatever's hit,
never sideways to a sibling, so it has to be topmost to receive them) — a known trade-off, since it
also means the track captures clicks; harmless today since nothing this package renders is
interactive, but would need revisiting if that changes. `.pptx-scroll-viewport` holds every rendered
slide permanently stacked underneath it (unlike `PptxPresentationElement`, which only goes
`position: absolute` transiently during a transition, this element has no single-slide "normal flow"
state to return to).

**RAF-throttled scroll handling — the `scroll` listener does no work itself.** `#handleScroll` just
records the latest `scrollTop`; a `requestAnimationFrame` loop (`#flushScroll`, started on first
scroll, cancelled in `disconnectedCallback`) reads it once per frame and calls `seekTo` —
coalescing however many `scroll` events the browser fires within one frame into a single update,
standard scroll-perf practice, and independently useful for a future "perfect pinning" pass (one
clock, read once per frame, not once per event) even though there's no longer a second write
(the viewport's own positioning) to keep in sync alongside it. Tested by stubbing
`requestAnimationFrame`/`cancelAnimationFrame` directly (`vi.stubGlobal`) rather than fighting
fake-timer/rAF interaction — see Tests below.

**Deliberately out of scope for this pass**, all noted in `docs/scroll-driven-playback.md`: native
CSS `animation-timeline: scroll()`/`view()` (would need re-deriving every `Animation`'s timing as
scroll-range percentages, and uneven browser support — the JS-driven approach above is what that
document's own WAAPI-migration rationale already pointed toward); scroll-snap-to-slide-boundary
(native `scroll-snap-type` could layer on later without touching `seekTo`'s own math); and any
`Slide.timing` behavior beyond the fade-on-whole-shape coverage real-time mode already has — scroll
mode inherits exactly that coverage via the unchanged `collectFadeAnimations`, not a new animation
pass.

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
- **`ShapeProperties.geometry` is rendered for a common subset of presets, plus `custGeom` with a
  usable `pathLst`.** `shape-geometry.ts` covers twelve presets total: `rect` (the pre-existing
  default), `roundRect` and `ellipse` (native `border-radius`), and nine more via a real SVG
  outline (triangle, `rtTriangle`, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon,
  `star5`) — see `shape-geometry.ts`'s and `fill.ts`'s own doc comments above. Outside that set
  (the other ~170 `ST_ShapeType` names — arrows, callouts, stars other than `star5`, flowchart
  shapes, etc.) a shape/picture still renders as a plain rectangular box. `custGeom` freeform
  outlines (new — a later session) now render too, via `customGeometryPath`, for the "every path
  fully resolved to literal coordinates" case `packages/presentation`'s `CustomGeometry.pathLst`
  models (see that package's own scope boundary) — a `custGeom` whose reader parse dropped every
  `a:path` (a `gdLst`-guide-referenced point somewhere in all of them) still falls back to a
  rectangle, same as before. Adjustment-guide (`avLst`) handling is also approximate, not
  spec-exact — see `shape-geometry.ts`'s doc comment; `customGeometryPath`'s `arcTo` conversion is
  real ellipse math rather than an approximation, but not independently verified against a real
  arc-containing deck yet (slide 3's own `custGeom` example only exercises `moveTo`/`lnTo`/`close`).
  On the nine SVG-path presets and on `custGeom`, only a solid
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
- **`push`, `fade`, and now `morph` play a real slide transition; every other
  `TransitionEffect.kind`** (`wipe`, `cut`, `dissolve`, `newsflash`, `wheel`, `split`, `strips`,
  `zoom`, `blinds`, `checker`, `comb`, `randomBar`, `circle`, `diamond`, `plus`, `pull`, `cover`,
  `random`, `wedge`) still takes the plain instant swap — each is a reasonably self-contained
  addition to `goToSlide`'s dispatch in `presentation-element.ts` following the same pattern
  `#animatePush`/`#animateFade` establish. `morph` (see "Key design decision: Morph transitions"
  above for the full mechanism) plays a real per-shape interpolation when
  `@pptx2html/presentation`'s `resolveMorphMatch` finds a confident-enough shape correspondence
  between the two slides, verified against a real Morph transition on
  `apps/web-demo/src/Presentation1.pptx`'s slide 4 (all 4 shapes matched, including one behind a
  `custGeom` freeform outline); below that confidence threshold, or when there's no previous slide
  to morph from, it falls back to a plain crossfade (`#animateFade`) and reports
  `morph-match-degraded` via `UnsupportedFeatureCollector` — the user's explicit requirement that a
  degraded match be just as visible in the log as any other unsupported feature, not a silent
  approximation. Fill/colour interpolation on a matched shape is **not** implemented — only
  position/size/rotation/opacity animate, the colour/gradient/pattern jumps instantly — and
  word/character-level text morph (`MorphOption`'s `byWord`/`byChar`, vs. the `byObject` this
  always plays regardless of the authored option) isn't modeled at all; both are documented,
  deferred gaps, not silent ones — see "Key design decision: Morph transitions" above, and its own
  note on the one known correctness gap (a non-adjacent `goToSlide()` jump landing on a Morph slide
  can silently animate fewer shapes than expected). `fade`'s `throughBlack: true` variant renders as
  a plain crossfade rather than its own
  real fade-to-black-then-in animation (see "Key design decision: push/fade slide transitions"
  above). `SlideTransition.advanceOnClick`/`.advanceAfter` (auto-advance timers) and
  `.sound`/`TransitionSoundAction` (playing/stopping audio during the transition) are both parsed by
  `packages/reader` already but entirely unconsumed here — a slide's transition only plays in
  response to explicit navigation (click/keyboard/`goToSlide`), never on its own after a timeout,
  and no audio ever plays. `SlideTransition.durationMs` (new, from the `p14:dur` extension
  attribute) _is_ consumed already, though — `resolveTransitionDurationMs(transition.speed,
transition.durationMs)` is resolved once in `goToSlide` and threaded into `#animatePush`/
  `#animateFade`, which no longer resolve duration themselves.
- **Only a plain "fade" `animEffect` on a whole shape plays from `Slide.timing`.** See "Key design
  decision: `Slide.timing` fade animations" above for the full picture; in scope-boundary terms:
  every other `animEffect` filter (wipe, blinds, wheel, ...), every other behavior kind (`set`,
  `anim`, `animClr`, `animMotion`, `animRot`, `animScale`, `cmd`, `audio`, `video`), a fade
  targeting shape text/background rather than a whole shape, and `SlideTiming.buildList`'s implicit
  paragraph/graphic builds (staged per-paragraph/per-series reveal) are all unmodeled — reported via
  `animation-effect-unmodeled`/`animation-behavior-unmodeled`/`animation-build-unmodeled`, not
  silently dropped. Real click-driven build-step sequencing (an "On Click"/"After Previous" effect
  actually waiting on its trigger, rather than every fade playing immediately at its own local
  delay) is also unmodeled — `to-html5` has no concept of an in-slide "build step" separate from
  slide-granular navigation at all yet; adding one is a materially bigger change than this pass
  attempted (`animation-trigger-unmodeled`).

## Tests

The inheritance-resolution tests that used to live here (`coordinate.test.ts`,
`placeholder.test.ts`, `text-style.test.ts`, `background.test.ts`, `bullet.test.ts`,
`style-matrix.test.ts`) moved to `packages/presentation/src/resolve/` alongside the source they
test — see that package's CLAUDE.md for what they cover now. `shape-geometry.test.ts` and
`units.test.ts` are the remaining pure-Node (no DOM) test files here, covering
`presetShapePath`/`nativeBorderRadius`'s path/radius output per preset (including the `adj`-guide
default vs. explicit-override cases and the 50%-cap edge case for `roundRect`/parallelogram/
trapezoid/hexagon/octagon) and the EMU→px/pt/cqw conversion math, directly. `shape-geometry.test.ts`'s
new `customGeometryPath` block covers: a single moveTo/lnTo/close subpath with non-uniform `w`≠`h`
scaling, multiple `a:path` entries concatenating into multiple subpaths (a boolean-subtract cutout
shape), `quadBezTo`/`cubicBezTo` mapping to `Q`/`C`, `arcTo`'s ellipse-derived endpoint/flags for
both a positive and a negative swing angle, and a zero-`w`/`h` path being dropped (falling back to
`undefined` when nothing else is renderable). `shape-tree.test.ts`'s "preset geometry" block gained
an end-to-end case rendering a two-subpath `custGeom` shape through `renderShapeTreeNode`, alongside
its existing `ellipse`/`triangle`/unmodeled-preset (`cloud`) cases.

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
`fast`/`med`/`slow` duration mapping (400/700/1000ms) as a pinned regression contract. Two more
cases: an explicit `durationMs` (`p14:dur`) overriding `speed` entirely, and a genuinely still
unmodeled effect kind (`wipe`) taking the instant-swap path and being reported via
`transition-effect-unmodeled`.

Nested inside that same block (reusing its `FakeAnimation`/`beforeEach`/`afterEach` and
`slideEls`), `describe('morph transitions', ...)` has its own `morphShape`/`buildMorphPresentation`
fixture builders (real `Shape`s with a `properties.transform`, rendered through the actual
pipeline — not hand-set inline styles — so `readShapeBox` exercises real `positionElement` output)
and covers: for a matched pair, the departing copy gets a single zero-duration `opacity: 0`
`Animation` (not a fade) while the arriving copy's keyframes tween `left`/`top`/`width`/`height`/
`transform` between the two real boxes with no `opacity` field at all (keyframe `left` values
asserted against each element's own rendered `style.left`, the same "compare against the real
computed value, not a hardcoded percentage" convention `slide.test.ts`'s `pct()` helper
established) — this is the regression test for the translucency bug documented in "Key design
decision: Morph transitions" above (an earlier crossfading-both-copies version made both visibly
translucent mid-transition); a disappearing shape fades out in place and an appearing one fades in
in place in the same transition, still a real two-keyframe opacity fade (unaffected by the above,
since neither has a "moving box" interpretation); a low-confidence match (two shapes sharing no
name at all) falls back to a plain whole-slide crossfade and reports `morph-match-degraded`; a
Morph transition authored on the deck's very first slide reports the same code (nothing to morph
from); and — the regression test that caught the departing/arriving role-swap bug, a second real
bug documented in "Key design decision: Morph transitions" above — navigating backward through an
already-played morph reverses which copy is instantly hidden vs. which one moves, asserted against
both elements' own real boxes.

`morph.test.ts` (new, pure Node, mirrors `animation.test.ts`'s style below) covers
`resolveSlideMorphMatch` directly with a `vi.fn()` `report` callback: no previous slide reports and
falls back; a fully-confident match reduces to plain shape-id pairs with nothing reported; a mixed
match (some matched, some appearing/disappearing) still reports nothing since match _rate_ stays
high; a below-threshold match rate and a zero-match slide both fall back and report
`morph-match-degraded`; and a 100%-match case right at the confidence ceiling still plays.

`animation.test.ts` (new, pure Node — mirrors `shape-geometry.test.ts`/`timing.test.ts`'s
DOM-free-logic style) covers `collectFadeAnimations` directly, with small `animEffect`/`par`/`seq`
builder helpers: a fade-in and fade-out `animEffect` targeting a shape, the `transition`-absent
default of `'in'`, walking into nested `par`/`seq`/`excl` containers to find fade behaviors several
levels deep, using a node's own numeric start delay, a click-gated node still playing at `delayMs:
0` while reporting `animation-trigger-unmodeled` once, an `'indefinite'`-duration fade being
skipped and reported, a non-`"fade"` filter and a non-shape target both being reported and not
played, every other behavior kind being reported once per distinct kind (deduplicated), and an
unconsumed `buildList` being reported. `shape-tree.test.ts`'s new "data-pptx-shape-id" block checks
a rendered shape's and group's own element carry that dataset attribute, set from `nonVisual.id`.
`presentation-element.test.ts`'s new `describe('Slide.timing fade animations', ...)` block (a
separate, simpler `HTMLElement.prototype.animate` mock than the push/fade block's `FakeAnimation` —
just a `vi.fn()` recording each call, since fade-animation playback never awaits `.finished`)
covers: a fade-in plays on the target shape (found via `data-pptx-shape-id`) as soon as its slide
is shown by `render()`, with the right keyframes/duration/delay/`fill: 'forwards'`; a fade-out
gets reversed keyframes; and navigating away then back to a slide replays its fade animations
(two recorded `.animate()` calls on the same target across the two visits), proving
`#playSlideAnimations` runs from `#updateActiveSlide` on every path that makes a slide active, not
just the first.

`scroll-timeline.test.ts` (new, pure Node, mirrors `animation.test.ts`/`morph.test.ts`'s
DOM-free-logic style) covers `resolveScrollTimeline` directly: an empty deck, the first slide never
getting a `transition` segment even when it authors one, a fully static slide's content duration
being exactly the dwell floor (default and a custom `minDwellMs` override), content duration being
driven by a slide's own fades when that exceeds the floor, an authored `push`/`fade` transition
playing as authored with `resolveTransitionDurationMs`'s own duration, an absent transition falling
back to the synthetic push-up default unreported, an unmodeled effect kind falling back to it _and_
reporting `transition-effect-approximated-for-scroll`, a confidently-matched Morph keeping its own
effect and carrying the reduced `MorphMatchSummary`, a low-confidence Morph match degrading to a
plain `fade` and reporting `morph-match-degraded` (reusing `resolveSlideMorphMatch`'s own existing
behavior, not reimplemented), multi-slide absolute-ms chaining, and a forwarded
`animation-build-unmodeled` report keyed to the right slide.

`scroll-presentation-element.test.ts` (new) reuses the `happy-dom`/`HTMLElement.prototype.animate`
mocking convention `presentation-element.test.ts` established, but a simpler `FakeAnimation` — no
`.finished` promise or `setTimeout`-driven resolution needed, since scroll-mode playback never
awaits completion, only reads/writes `currentTime` on an already-paused animation. Covers: one
element per slide rendered into the viewport; only the first slide visible at `ms: 0`; `seekTo`
during a static slide's own content phase showing only that slide; `seekTo` mid-transition showing
both participating slides and setting each animation's `currentTime` to the elapsed offset; scrubbing
backward being a smaller `currentTime` on the exact same `Animation` instance (no new one created);
every animation being `paused` immediately after creation; `seekTo` clamping to `[0,
totalDurationMs]`; a content-phase fade's `currentTime` being relative to its own slide's content
start and clamped to its own `[0, delay + duration]` domain even when the slide's own dwell is
longer; a Morph transition's departing shape getting `opacity: '0'` for the transition window and
having it reset once scrubbed back into its own slide's content phase; a low-confidence Morph match
still scrubbing correctly as a plain crossfade; `pixelsPerSecond` resizing the track; and, with
`requestAnimationFrame`/`cancelAnimationFrame` stubbed via `vi.stubGlobal` (not fake timers — real
`requestAnimationFrame`/fake-timer interaction is exactly the fragility `docs/scroll-driven-
playback.md` and this file's own design decision above call out avoiding), that multiple `scroll`
events dispatched before the stubbed frame callback fires coalesce into exactly one `seekTo` call,
using only the latest `scrollTop`.

## Next likely steps

1. Widening slide-transition coverage beyond `push`/`fade` — each remaining `TransitionEffect.kind`
   (`wipe`, `cut`, `dissolve`, `wheel`, `split`, ...) is a reasonably self-contained addition to
   `goToSlide`'s dispatch in `presentation-element.ts`, following the `#animatePush`/`#animateFade`
   pattern (now WAAPI-based — see that design decision above); `fade`'s `throughBlack: true`
   two-stage fade-to-black-then-in animation, `advanceOnClick`/`advanceAfter` auto-advance timers,
   and `TransitionSoundAction` playback are separate, currently unstarted pieces of that same
   surface area. **`morph` is done, a later session** — parsing, shape-matching
   (`@pptx2html/presentation`'s `resolve/morph.ts`'s `resolveMorphMatch`), and rendering
   (`morph.ts`'s `resolveSlideMorphMatch`, `#animateMorph`) — see "Key design decision: Morph
   transitions" above for the full mechanism and its honesty caveats (fill/colour doesn't
   interpolate, word/character-level text morph isn't modeled, a non-adjacent `goToSlide()` jump
   onto a Morph slide can silently animate fewer shapes than expected). Remaining Morph work, if
   picked up further: fill/colour interpolation on matched shapes (reading `background-color` off
   each element the same way box values already are, for the solid-fill case at least);
   `byWord`/`byChar` text-level matching (needs `text.ts` to tag individual words/characters
   instead of one `<span>` per run — a real DOM restructuring, not a small addition); and possibly
   tuning `MIN_MORPH_MATCH_RATIO` (`morph.ts`) against more real Morph-authored decks than the one
   fixture this was verified against.
   Per-slide/per-element animations off `Slide.timing`
   now have a first slice done (plain `"fade"` `animEffect` only — see "Key design decision:
   `Slide.timing` fade animations" above); widening that is its own, larger piece of remaining
   work — other `animEffect` filters (each a fairly mechanical addition once one exists, though
   several — wipe, blinds — imply a CSS/SVG technique of their own, not just a new opacity
   keyframe pair), `animMotion`/`animRot`/`animScale`/`animClr` (each a different CSS property to
   drive), and real click-driven build-step sequencing (needs `to-html5` to grow an in-slide
   "build step" concept distinct from slide-granular navigation — a materially bigger change than
   any single new behavior kind, see that design decision's own "not a faithful implementation"
   note). This session's two enabling pieces for a future scroll-driven-playback feature — a
   central duration-resolution API in `@pptx2html/presentation` (`resolve/timing.ts`, see that
   package's CLAUDE.md — now including `resolveTimeNodeStartMs`, the per-node start-delay answer
   fade-animation playback also uses) and migrating slide-transition playback off plain CSS
   `transition` onto the Web Animations API — are both done, and fade-animation playback followed
   the same WAAPI-based pattern; read root `CLAUDE.md`'s "Future feature: scroll-driven playback"
   section / `docs/scroll-driven-playback.md` before picking up any of the items above, since the same
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
5. **Done, a later session**: `custGeom` freeform-outline rendering (`shape-geometry.ts`'s
   `customGeometryPath`, `@pptx2html/presentation`'s `CustomGeometry.pathLst`) — see the Layout
   section above and the scope boundary's updated `ShapeProperties.geometry` entry. Verified
   end-to-end against a real fixture, `apps/web-demo/src/Presentation1.pptx`'s slide 3 (a boolean
   "Subtract" shape). Remaining gaps in this specific area: `arcTo` conversion is unverified
   against a real arc-containing deck (see the scope boundary), and gradient/pattern/blip fill on a
   `custGeom` outline is unmodeled, same as the nine preset SVG paths (item 6 below covers both).
6. Gradient/pattern/blip fill support on the nine SVG-path presets and `custGeom` outlines
   (currently solid-only, see the scope boundary above) — an SVG `<linearGradient>`/`<pattern>`
   `<defs>` entry referenced via
   `fill="url(#id)"` would parallel `resolveGradientCss`'s CSS-gradient output without the
   rectangular-background constraint, but needs its own angle/stop-position math since SVG
   gradients use `objectBoundingBox` coordinates, not CSS's `linear-gradient()` syntax.
7. Clipping a picture to one of `shape-geometry.ts`'s nine SVG-path presets or a `custGeom` outline
   (not just `roundRect`/`ellipse`'s native `border-radius`) — needs an `<svg><image>` +
   `<clipPath>` overlay in place of the plain `<img>` `renderPicture` emits today.
8. `p:style/effectRef` (see the scope boundary above) needs effect rendering to exist at all
   first (a bigger, separate gap) — `fontRef` is done, see "Key design decision: fontRef
   text-colour/typeface fallback lives in `resolveEffectiveRunProperties`, not a new resolver"
   above.
9. Wiring a connector's own `p:style` (already parsed, `ShapeStyle` is on `ConnectionShape` too)
   into whatever connector line rendering eventually lands (see item 3 above) — today it's parsed
   but unused, since `renderConnector` doesn't call `applyFill`/`applyLine`/`effectiveFill`/
   `effectiveLine` at all yet.
10. `apps/web-demo` now renders the parsed presentation into the page (see its own source) —
    confirm this still looks right whenever this package's DOM structure changes, and now also
    gains a working slideshow (click/keyboard navigation) for free once it re-renders with this
    change, since `renderPresentation`'s returned element already wires it up.
11. **Done, a later session**: scroll-driven playback (`scroll-timeline.ts`, `scroll-presentation-
element.ts`, `renderScrollPresentation`) — see "Key design decision: scroll-driven playback" above
    for the full mechanism, including a later fix within that same session (the viewport-jitter bug,
    same section) for why the visible content lives outside the scrolling element entirely rather
    than inside it. `apps/web-demo` now wires up `<pptx-scroll-presentation>` too, via a playback-mode
    toggle next to the file picker (`index.html`/`index.ts`) that re-renders the already-parsed
    `Presentation` in whichever mode is selected, without re-parsing. Remaining gaps, all
    deliberately deferred (see that design decision's own closing paragraph): native CSS
    `animation-timeline: scroll()`/`view()` as a possible perf optimization over the current
    JS-driven scroll listener; scroll-snap-to-slide-boundary; the "perfect pinning" goal (still
    unstarted — the sibling-viewport fix already removes the browser's native scroll compositing
    from the equation, which if anything makes a future exact/eased pin easier to build, not harder);
    and the `.pptx-scroll-track` capturing all clicks trade-off (harmless today, no interactive
    content exists yet — see the design decision's own note). Also unstarted: exposing `minDwellMs`
    (currently a fixed default in `scroll-timeline.ts`) as a public property the way
    `pixelsPerSecond` already is, if a real deck's static slides need visible tuning.
