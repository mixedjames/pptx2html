# @pptx2html/presentation

The in-memory model of a `.pptx` file: interfaces that closely mirror ECMA-376 Part 1 (the
OOXML spec), organized the same way the spec itself is split — DrawingML (§20, shapes/color/
fill/text primitives shared across all of OOXML) vs. PresentationML (§19,
presentation-specific parts) — **plus** the pure functions that compute the well-defined
values OOXML's own inheritance/derivation rules imply but don't store explicitly (an effective
transform, an effective run's character formatting, a resolved fill). No parsing (that's
`@pptx2html/reader`'s job) and no output-format-specific logic — CSS strings, SVG path syntax,
DOM nodes, or any other renderer's own representation — that's a renderer's job (currently
`@pptx2html/to-html5`, deliberately not the only one this package expects). See "Key design
decision: resolution logic lives with the model" below for why that line is drawn there and not
at "no logic at all."

## Status: skeleton plus a resolution layer, consumed by a working reader and a first-pass renderer

The type graph is complete enough that `@pptx2html/reader` parses real `.pptx`
byte streams into it end-to-end (see `packages/reader/CLAUDE.md`), and
`@pptx2html/to-html5` renders the result into a laid-out, responsively-scaled HTML5 DOM
with several formatting passes done (font inheritance including alignment — typeface/
size/bold/italic/underline/strike/color/alignment, all resolved through the full
placeholder/master/theme chain; shape/picture fill/line, including PowerPoint's Shape
Styles gallery `p:style` fillRef/lnRef fallback via `FormatScheme.fillStyles`/`.lineStyles`;
a common subset of preset shape geometry; slide background; bulleted and numbered lists,
including auto-numbering schemes and indentation). The inheritance-walking logic behind all of
that — except the shape-geometry/CSS-formatting half, which is genuinely renderer-specific —
now lives in this package's own `resolve/` directory (see Layout below), not `to-html5`'s;
`to-html5` calls it and formats the result as DOM/CSS. Table cell fill and table styles are the
main remaining rendering gap — see `packages/to-html5/CLAUDE.md`. `tsc -b`, `eslint`,
`vitest run` and `prettier --check` are all green.

`Slide` also now carries an optional `timing` (`presentationml/animation.ts`, §19.3.1.48's
`p:timing`): the element/build animation timing tree and build list `reader` parses off a slide.
`to-html5` doesn't consume the tree itself yet (per-element/build animation rendering is still
unstarted; `apps/web-demo` just logs it to the console) — but unlike when this was first written,
there's no longer "nothing to derive" here: `resolve/timing.ts` (new) computes a node's/tree's own
effective duration, since a scroll-driven-playback feature (see root `CLAUDE.md`'s design note,
`docs/scroll-driven-playback.md`) needs exactly that answer and it's renderer-agnostic, same as
every other file in `resolve/`.

`Slide` also now carries an optional `transition` (`presentationml/transition.ts`, §19.3.1.49's
`p:transition`): the single whole-slide effect (fade/wipe/push/split/wheel/zoom/... — the base
schema's `EG_SlideTransition` choice group, ten shapes covering ~20 effect names) plus speed and
advance-on-click/advance-after-N-ms settings, played when the presentation advances _into_ this
slide. Distinct from `timing`'s per-element animation tree — a transition is a single flat value,
not a tree — but its `speed` → millisecond mapping is exactly the kind of renderer-agnostic
question `resolve/timing.ts`'s `resolveTransitionDurationMs` now answers (moved here from
`to-html5`, which now consumes it for its `push`/`fade` transition playback — see that package's
CLAUDE.md). Its optional sound (`TransitionSoundAction`, §19.3.1.49/50's `p:sndAc`) reuses
`drawingml/media.ts`'s `MediaPart` rather than inventing a parallel representation, the same
relationship-resolved-to-object-graph choice every other embedded binary in this package makes;
sound playback itself remains unconsumed by `to-html5`.

## Key design decision: resolved object graph, not relationship IDs

The raw OOXML package splits a presentation across many XML parts connected by
relationship IDs (`r:id`) — a `Slide` part points at its layout only via a string
ID that has to be resolved through `_rels/*.rels`. We deliberately did **not**
mirror that indirection here. Instead every reference is a real nested object:
`Slide.layout` is an actual `SlideLayout`, `SlideLayout.master` an actual
`SlideMaster`, `SlideMaster.theme` an actual `Theme`, embedded images are an
in-memory `MediaPart` (contentType + bytes), not an `r:embed` string. This makes
the tree trivial to consume for rendering; resolving relationship IDs into object
references is entirely `@pptx2html/reader`'s job, not this package's.

One consequence: `SlideMaster.layouts` and `SlideLayout.master` are mutually
referential. Every field here is `readonly`, so building that cycle needs a
two-phase construction (build with a temporary mutable slot, then freeze) — see
`Mutable<T>` in `packages/reader/src/mutable.ts` for how the reader does it. This
package itself has no construction helpers; it only declares the shape.

## Key design decision: resolution logic lives with the model, not each renderer

`resolve/`'s functions (an effective transform, an effective run's character formatting, a
resolved fill, a resolved colour, a formatted auto-number label) started out in `to-html5` and
moved here once it became clear they were never actually rendering anything — they take objects
from this package's own graph and return other objects/values from that same graph, purely by
applying OOXML's own inheritance/fallback rules. Any future renderer (a canvas- or PDF-based one,
a second HTML renderer with different DOM/CSS choices, ...) needs the exact same answer to "what
is this placeholder's effective transform" or "what is this run's effective font" that
`@pptx2html/to-html5` does — those aren't rendering decisions, they're facts about the
presentation. Duplicating this logic per renderer would both be wasted work and a correctness
risk (two renderers silently disagreeing on, say, the placeholder-matching fallback order).

The line this package actually draws is **not** "no logic at all" — it's "no output-format-specific
logic": nothing in `resolve/` produces a CSS string, SVG path, DOM node, or any other renderer's
own representation. Concretely, `color.ts`'s `resolveColor` stops at a structured `ResolvedColor`
(`{ rgb, alpha }` or an opaque preset name) rather than a CSS colour string — the RGB/HSL/
colour-transform math is exactly as reusable as `placeholder.ts`'s transform-inheritance walk, but
formatting that result as `rgb(...)`/`rgba(...)` assumes a CSS-consuming caller, so that step
stays behind in `to-html5`'s own (now much thinner) `color.ts`. `shape-geometry.ts`'s preset-outline
logic in `to-html5` is a clearer case of the same line: it stays there entirely, since its SVG-path/
`border-radius` output is inherently tied to _how_ `to-html5` chooses to draw a shape (a
`0 0 100 100` viewBox non-uniformly stretched over a CSS box), not a fact the presentation model
itself defines.

## Layout

- `drawingml/` — shared primitives, one concept per file: `units.ts` (`Emu`,
  `Angle`, `Percentage`, `FontSize`), `color.ts` (srgb/scheme/system/preset/hsl +
  transform modifiers), `geometry.ts` (`Transform2D`, preset/custom geometry),
  `media.ts` (`MediaPart`), `fill.ts` (none/solid/gradient/pattern/blip),
  `line.ts`, `text.ts` (`TextBody`/`Paragraph`/`TextRun`/`LineBreak`/`TextField`),
  `shape-common.ts` (`NonVisualDrawingProperties` — including `Placeholder`
  type/index, used to resolve a placeholder shape's inherited position/size when
  it has none of its own — and `ShapeProperties`).
- `theme.ts` — `ColorScheme` (12 named slots), `FontScheme`, `FormatScheme` (`name` plus
  `fillStyles`/`lineStyles`, each always 3 entries — the fill/line style matrices a shape's
  `p:style/fillRef`/`lnRef` points into by 1-based index, reusing `Fill`/`Line` directly since a
  style-matrix entry is structurally identical to a shape's own `spPr` fill/line; `effectStyleLst`/
  `bgFillStyleLst` remain unmodeled).
- `presentationml/` — `shape-tree.ts` (`Shape`/`Picture`/`ConnectionShape`/
  `GraphicFrame`/`GroupShape`, the `ShapeTreeNode` union), `shape-style.ts`
  (`ShapeStyle`/`StyleMatrixReference` — a shape/picture/connector's optional `p:style`
  `fillRef`/`lnRef`, each a 1-based index into `FormatScheme.fillStyles`/`.lineStyles` plus a
  `Color` that substitutes for that style's `phClr` placeholder; resolving the substitution is
  `resolve/style-matrix.ts`'s job, below), `table.ts`,
  `common-slide-data.ts` (`CommonSlideData`/`Background`, shared by every
  slide-like part), `animation.ts` (`SlideTiming`/`TimeNode`/`BuildListEntry` — a slide's optional
  `p:timing`, consumed by `Slide.timing` in `slide.ts`; see the design-decision note above for why
  it has no `resolve/` counterpart), `transition.ts` (`SlideTransition`/`TransitionEffect`/
  `TransitionSoundAction` — a slide's optional `p:transition`, consumed by `Slide.transition` in
  `slide.ts`; same no-`resolve/`-counterpart reasoning as `animation.ts`, see above), then the part
  hierarchy: `slide-master.ts` → `slide-layout.ts` → `slide.ts` → `notes.ts`, and `presentation.ts`
  as the DOM root.
- `resolve/` — pure functions that compute a derived/effective value from the graph above; no
  DOM, no CSS, no output-format-specific logic anywhere in this directory (verified file-by-file
  when it was moved here from `to-html5`, and enforced going forward by the same absence). One
  file per resolved concept: `placeholder.ts` (`findPlaceholderMatch`, `resolveInheritedTransform`
  — the placeholder shape → layout → master transform-inheritance walk, §19.3.1.36),
  `text-style.ts` (`resolveEffectiveRunProperties`/`resolveEffectiveAlignment`/
  `resolveEffectiveBullet`/`resolveEffectiveIndent`/`resolveTypeface` — the run/paragraph → shape
  → placeholder layout/master → master category style → presentation default → theme font-scheme
  chain, §21.1.2, plus the `TextStyleContext` type a renderer's own per-slide context can extend
  rather than duplicate), `background.ts` (`resolveEffectiveBackground` — slide → layout → master,
  §19.3.1.7), `style-matrix.ts` (`resolveStyleFill`/`resolveStyleLine` — a `p:style` `fillRef`/
  `lnRef` resolved against `FormatScheme`, `phClr` substitution included, §20.1.4.2.10/2.12),
  `bullet.ts` (`formatAutoNumber`/`NumberingState` — auto-number label formatting and the
  running-counter state a numbered list needs across paragraphs, §20.1.10.51), `coordinate.ts`
  (`CoordinateMap`/`composeGroupMap`/`computeBox`/`ElementBox` — the affine math that turns a
  shape's own `Transform2D`, possibly nested inside group transforms, into a slide-root-relative
  EMU box; stays in EMU rather than any renderer's own unit, see `to-html5/CLAUDE.md`'s
  percentage-based-layout decision for how that package turns it into CSS), `color.ts`
  (`ResolvedColor`/`resolveColor`/`resolveFillColor` — DrawingML colour resolution, scheme
  aliasing, HSL conversion, colour-transform math, stopping at a structured `{ rgb, alpha }` or
  `{ preset }` result rather than a CSS string — `to-html5`'s own `color.ts` is now a thin
  CSS-formatting wrapper around this), `timing.ts` (new — `resolveTransitionDurationMs` — a
  `TransitionSpeed`'s `fast`/`med`/`slow` → millisecond mapping, moved here from `to-html5`;
  `resolveTimeNodeDuration`/`resolveSlideTimingDuration` — a `TimeNode`'s/`SlideTiming`'s own
  effective duration, or `'indefinite'` if it or any descendant can't complete without external
  input, e.g. an `onClick`/`onNext` wait with no numeric delay fallback — the "is this
  presentation fully time-resolved" question `docs/scroll-driven-playback.md` needs; see that
  file's own doc comments for exactly which composition/repeat rules are exact vs. a documented
  approximation).
- `index.ts` / `drawingml/index.ts` / `presentationml/index.ts` / `resolve/index.ts` — barrel
  re-exports only.

Every `.ts` file in `packages/reader/src` corresponds 1:1 to a file in `drawingml/`/
`presentationml/` specifically — if you add a type there, the reader will need a matching parser
file. `resolve/` has no reader counterpart by design: the reader never resolves inheritance (see
`packages/reader/CLAUDE.md`'s own scope boundary), only this package and its renderers do.

## Scope boundary — what's intentionally unmodeled

Each gap below is called out with a `§`-referenced comment at its point in the
code (search for "unmodeled for the skeleton"):

- Custom geometry path data (`custGeom` in `geometry.ts` carries no path).
- Effect and table style matrices referenced by a shape's/table's style index. `FormatScheme`'s
  fill/line style matrices (`fillStyleLst`/`lnStyleLst`) _are_ modeled and consumed (see
  `theme.ts`/`shape-style.ts` above) — but `effectStyleLst`/`bgFillStyleLst` (shadows and other
  effects, plus a separate background-fill matrix) and table style matrices (banded rows, header
  row styling, etc., referenced by a table's own style ID) remain bare/unmodeled.
- Preset shape geometry (`geometry.ts`'s `PresetGeometry`) is modeled for every preset name (it's
  just a string + optional adjustment guides) but only a common subset of ~180 `ST_ShapeType`
  presets is actually turned into a real outline by a consumer — see `to-html5/CLAUDE.md`'s
  `shape-geometry.ts`.
- Chart, SmartArt and OLE object internals — `GraphicPlaceholder` in
  `shape-tree.ts` only preserves which of the three it is.
- Text autofit in `TextBodyProperties`.
- Theme overrides at the slide/layout level (color/font map overrides) — `to-html5`'s colour
  resolution assumes the default clrMap (bg1→lt1, tx1→dk1, bg2→lt2, tx2→dk2).
- Custom shows on the root `Presentation`.
- Path gradients (only linear-angle gradients are modeled in `GradientFill`).
- Bullet/numbering (`Bullet`, `AutoNumberScheme` in `drawingml/text.ts`) covers `buNone`/
  `buChar`/`buAutoNum` plus their `buFont`/`buClr`/`buSzPct` overrides, and `marL`/`indent`
  paragraph indentation — but not `buSzPts` (a point-size bullet override, vs. the modeled
  `buSzPct` percentage one) or the handful of double-parenthesis `ST_TextAutonumberScheme`
  variants `AutoNumberScheme` doesn't include (the ten common ones are). Other paragraph-level
  properties a real `a:pPr`/`a:lvlNpPr` can carry (tab stops, line spacing, space before/after)
  remain fully unmodeled.
- Non-bullet/indent/alignment/font paragraph properties in a `TextListStyle` level
  (`TextListStyleLevel` models `algn`/bullet/`marL`/`indent`/`defRPr` — everything
  `to-html5`'s paragraph-level inheritance pass needs — but not tab stops, line spacing, etc.).
  `TextListStyle` (`drawingml/text.ts`) backs `TextBody.listStyle`, `SlideMaster.textStyles`
  (title/body/other) and `Presentation.defaultTextStyle` alike — see `to-html5/CLAUDE.md`'s
  font-inheritance design decision for how a consumer walks all three.
- `animation.ts`'s timing tree covers the common par/seq/excl containers and set/anim/animEffect/
  animClr/animMotion/animRot/animScale/cmd/audio/video behaviors, but not: a relative (`p:by`)
  colour shift on `animClr` or a colour value on an `anim` keyframe (`from`/`to` absolute colours
  only); an `inkTgt` animation target (ink annotations); a sequence's own advance conditions
  (`p:seq`'s `prevCondLst`/`nextCondLst`, used by interactive sequences); a `p:iterate` (by-letter/
  by-word text animation); or a paragraph build's `"cust"` build type (a custom build fully
  described by the real timing tree, with no implicit shorthand to represent).
- `transition.ts`'s `TransitionEffect` covers every effect in the base schema's
  `EG_SlideTransition` choice group (§19.3.1.49 — blinds/checker/circle/comb/cover/cut/diamond/
  dissolve/fade/newsflash/plus/pull/push/random/randomBar/split/strips/wedge/wheel/wipe/zoom), but
  not PowerPoint's newer "fancy" transitions (Morph, Reveal, Ripple, Honeycomb, Vortex, Shred,
  Switch, Airplane, Cube, Doors, Gallery, Prism, Origami, Pan, Ferris, Fracture, Crush, Curtains,
  Windows, Warp, Glitter, Flythrough, ...), which PowerPoint authors as `p14:`/`p15:`/
  `p159:` extension elements inside `p:transition`'s `p:extLst` rather than through that group —
  parsing `p:extLst` at all is a separate, currently-unstarted piece of surface area.

**Rule for extending this**: if the reader needs to surface one of these, add the
type here first, then implement the parser in `packages/reader`. Don't let the
reader parse-then-discard data this package can't yet represent — that's exactly
the pattern `packages/reader/CLAUDE.md`'s "Scope boundary" section commits to
avoiding.

## Tests

`resolve/`'s seven files each have a `*.test.ts` sitting next to them, all pure Node (no DOM —
nothing here needs one): `placeholder.test.ts` covers `resolveInheritedTransform`'s exact-match/
type-only-fallback/master-fallback/no-match cases; `text-style.test.ts` covers the full
`levelChain` walk (every rung from `defaultTextStyle` down to a run's own `rPr`/a paragraph's own
`algn`/bullet/`marL`/`indent`, plus theme font-token resolution) via
`TextStyleContext`-constructed fixtures rather than a renderer's own richer context type;
`background.test.ts` covers the slide→layout→master fallback chain; `style-matrix.test.ts` covers
`phClr` substitution (including through gradient stops) and the out-of-range/missing-reference
cases; `bullet.test.ts` covers `formatAutoNumber`'s ten schemes and `NumberingState`'s
continue/restart/break rules; `coordinate.test.ts` covers the affine math directly (including
composing across nested groups); `color.test.ts` asserts on `ResolvedColor` structures (unrounded
RGB tuples, a `{ type: 'preset' }` case) rather than CSS strings — `to-html5`'s own `color.test.ts`
still separately covers the CSS-formatting step end-to-end, and kept passing unchanged through
this split since `resolveColor`/`resolveFillColor`'s public signature there didn't change.
`timing.test.ts` (new) covers `resolveTransitionDurationMs`'s fast/med/slow/absent-defaults-to-fast
mapping and `resolveTimeNodeDuration`'s full composition rules directly, with small `leaf`/`par`/
`seq`/`excl` fixture builders (mirroring `text-style.test.ts`'s builder-helper convention rather
than hand-writing nested object literals per test): explicit numeric/`'indefinite'` duration on a
leaf, `par`/concurrent-`seq` (max) vs. non-concurrent-`seq`/`excl` (sum) composition, a child's
numeric start-condition delay adding into its contribution, a click-gated child (no numeric delay)
propagating `'indefinite'` up the whole tree, an OR'd condition list resolving via its one numeric
member rather than going indefinite, `repeatCount`/`repeatCount: 'indefinite'`/`repeatDuration`/
`autoReverse` each exercised on a leaf, a container's own explicit `duration` overriding its
children-derived computation, and one nested-container composition case.

## Next likely steps

1. Font/alignment inheritance, shape fill/line (including the `p:style` fillRef/lnRef fallback),
   preset shape geometry (a common subset), slide background, and bulleted/numbered lists are all
   done. Table cell fill is the next thing `@pptx2html/to-html5` needs to render — `TableCell.fill`
   already exists here and `reader` already parses it, `to-html5`'s `table.ts` just doesn't apply
   it yet (see that package's CLAUDE.md). Table _style matrices_ referenced by a table's style ID
   (banded rows, header row styling, etc.) are a separate, unmodeled gap — see the scope boundary
   above.
2. Custom geometry path data is the remaining layout (not just formatting) gap most likely to
   visibly matter next.
3. Widening `to-html5`'s preset-geometry coverage beyond its current common subset (see that
   package's CLAUDE.md) doesn't need anything new here — `PresetGeometry` already carries every
   preset name and any literal `val N` adjustment guide, a consumer just needs to turn more of
   them into real outlines.
4. **Done, this session**: `resolve/timing.ts` — `resolveTimeNodeDuration` computes a `TimeNode`'s
   own effective duration (or `'indefinite'` if it or any descendant is gated on a click/other
   external event with no numeric delay fallback — exactly the "is this presentation fully
   time-resolved" question a future scroll-driven-playback feature needs, see root `CLAUDE.md`'s
   design note / `docs/scroll-driven-playback.md`), and `resolveTransitionDurationMs` (moved from
   `to-html5`, which now consumes it) answers the smaller `TransitionSpeed` → ms question. Honest
   remaining gap, clearly documented in `timing.ts` itself rather than silently assumed: container
   scheduling (`excl`'s real "at most one active" semantics, an interactive sequence's own advance
   conditions) and the `repeatCount`/`autoReverse` interaction are approximated, not spec-exact —
   nothing currently consumes exact values, so this hasn't mattered yet. Nothing in `to-html5`
   duplicates this logic; see that package's CLAUDE.md for how it now uses
   `resolveTransitionDurationMs`.
