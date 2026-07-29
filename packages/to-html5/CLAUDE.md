# @pptx2html/to-html5

Renders a `@pptx2html/presentation` object graph (the in-memory `.pptx` DOM, produced by
`@pptx2html/reader`) into an actual HTML5 DOM. Public API is a single synchronous function,
mirroring the reader's `readPresentation`:

```ts
import { renderPresentation } from '@pptx2html/to-html5';
const el = renderPresentation(presentation); // <pptx-presentation>, shadow DOM inside
document.body.appendChild(el);
```

## Status: first pass — layout only, no formatting

Every slide and shape lands in the right place at the right size, including placeholder shapes
that inherit their position from the slide layout/master rather than declaring their own (very
common in real decks — see `placeholder.ts`). Nothing about how it looks (fill, line, color,
font, text run styling, backgrounds) is rendered yet — that's deliberately deferred to a later
pass. See "Scope boundary" below for the precise line.

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
  `Table` → `<table>`), no absolute positioning of their own; a paragraph is `<p>`, a run is a
  text node, a break is `<br>`, a field renders its `cachedText`. `renderTable` sizes itself to
  `100%`/`100%` of its containing `graphicFrame` div, `table-layout: fixed`, with `<col>` widths
  and `<tr>` heights expressed as a percentage of the column-width/row-height totals (so column
  proportions survive scaling; row-height percentages are best-effort — browsers that ignore them
  just size rows by content instead).
- `placeholder.ts` — pure logic, no DOM: `resolveInheritedTransform` walks the OOXML placeholder
  inheritance chain (§19.3.1.36) — slide placeholder → matching layout placeholder (by
  type+index, falling back to index-only then type-only) → matching master placeholder, same
  rule — and returns the first transform found. `shape`/`picture`/`connector` almost always rely
  on this: real decks routinely omit `xfrm` on placeholder shapes entirely.
- `render-context.ts` — `RenderContext { slideSize, layout }`, threaded alongside `CoordinateMap`
  through every `renderShapeTreeNode` call. Unlike the map (which changes at every group nesting
  level), this is constant for the whole slide — `slideSize` is what every element's EMU box gets
  divided by to become a percentage, `layout` is for placeholder resolution.
- `shape-tree.ts` — `renderShapeTreeNode`, the one recursive dispatcher over `ShapeTreeNode`'s
  five kinds. `shape`/`picture`/`connector` position from their own transform if they have one,
  else from `resolveInheritedTransform`; `graphicFrame` always has its own transform (mandatory in
  the schema) so it skips that path. All four get positioned by `positionElement`, which converts
  `computeBox`'s EMU box to `left`/`top`/`width`/`height` **percentages of `context.slideSize`**.
  `group` renders as an anchor `<div>` stretched to exactly cover the slide (`left/top/right/bottom:
0`, i.e. 100%×100% of its own containing block) — its own transform only feeds
  `composeGroupMap` for its children, it draws nothing itself, and groups never carry a
  placeholder (no `nvPr`/`ph` in the schema for `grpSp`). Percentages need a well-defined
  containing block to resolve against; stretching every nesting level to exactly the slide's size
  is what makes a doubly-nested child's percentage position still resolve correctly (see the
  design decision below).
- `slide.ts` — `renderSlide`: one `.pptx-slide` div, `position: relative; overflow: hidden`,
  `width: 100%` with `aspect-ratio: <slideWidth> / <slideHeight>` (raw EMU numbers — `aspect-ratio`
  only cares about the ratio) so it fills whatever width its container gives it and its height
  follows automatically. Holds the slide's own `shapeTree` walked from `IDENTITY_MAP` with a
  `RenderContext` built from `slide.layout` and the passed-in `slideSize`.
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

## Scope boundary — what's intentionally unmodeled (yet)

- **All visual formatting**: `ShapeProperties.fill`/`.line`/`.geometry`, `Background`, run-level
  `RunProperties` (bold/italic/color/font/size), paragraph alignment, table cell fill, table
  styles. The DOM structure exists (`.pptx-shape`, `.pptx-paragraph`, etc.) precisely so a later
  pass can add CSS without restructuring.
- **Placeholder inheritance is supported, but with a simplified matching rule.** See
  `placeholder.ts` above. Not modeled: the spec's type-equivalence groups (e.g. a slide's
  `ctrTitle` placeholder is allowed to match a layout's `title` placeholder) — only exact type
  matches (after the index-match attempts) are tried. Real decks reliably reuse the same
  placeholder type across slide/layout/master, so this covers the common case; a shape can still
  end up unpositioned (static flow) if nothing in the chain matches or has a transform.
- **Charts/SmartArt/OLE objects.** `GraphicPlaceholder` only carries `type`; rendered as a
  `[chart]`/`[smartArt]`/`[oleObject]` text placeholder box.
- **Connectors** render as an unstyled positioned `<div>` (no line drawn) — actual line
  rendering depends on the formatting pass reading `ShapeProperties.line`.
- **Object URLs are never revoked.** `renderPicture` calls `URL.createObjectURL` per picture and
  never `revokeObjectURL`s it. Fine for a one-shot render; calling `.render()` repeatedly on the
  same `PptxPresentationElement` (or rendering many presentations in one page) will leak blob
  URLs. Deferred — would need the created URLs plumbed back up to
  `PptxPresentationElement` so it can revoke on re-render/disconnect.

## Tests

`coordinate.test.ts` and `placeholder.test.ts` are pure Node (no DOM) and cover the affine math
and the placeholder inheritance chain (layout match, master fallback, no-match) directly.
Everything else that touches the DOM (`slide.test.ts`, `table.test.ts`, `text.test.ts`,
`presentation-element.test.ts`) opts into `happy-dom` per-file via a
`// @vitest-environment happy-dom` docblock — the repo's root `vitest.config.ts` stays on
`environment: 'node'` for every other package, this is the only one that needs a DOM.
`slide.test.ts` asserts percentages via a local `pct()` helper that mirrors `positionElement`'s
own formula exactly (rather than hardcoding decimal strings, which would be one floating-point
rounding difference away from a false failure), and covers: top-level percentage positioning,
nested-group percentage remapping (including that the group wrapper is stretched via
`left/top/right/bottom: 0`, not just `left/top: 0`), and a placeholder shape with no own transform
picking up its layout placeholder's position. `table.test.ts` checks column/row percentages sum
to the expected split (50/50 for two equal-width columns) rather than absolute values.

## Next likely steps

1. Wire a formatting pass on top of this structure: `Fill`/`Color`/`Line` → CSS
   background/border, `RunProperties` → span styling, `TextBodyProperties.anchor`/`wrap` →
   flex/white-space, `ParagraphProperties.alignment` → `text-align`.
2. `apps/web-demo` now renders the parsed presentation into the page (see its own source) —
   confirm this still looks right whenever this package's DOM structure changes.
