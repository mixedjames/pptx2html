# @pptx2html/reader

Parses a `.pptx` file's bytes into the in-memory object tree defined by
`@pptx2html/presentation` (Presentation → SlideMaster/SlideLayout/Slide/Notes →
CommonSlideData → ShapeTreeNode → drawingml primitives). Public API is a single
synchronous function:

```ts
import { readPresentation } from '@pptx2html/reader';
const presentation = readPresentation(uint8ArrayOfPptxBytes);
```

## Status: complete and green

Full read path is implemented end-to-end — themes, slide masters/layouts/slides,
notes, tables, and the whole shape tree (shapes/pictures/connectors/graphic
frames/groups) with all drawingml primitives (color, fill, line, geometry, text),
including the full text-property inheritance surface: a paragraph's own `algn`/bullet
(`buNone`/`buChar`/`buAutoNum`)/`marL`/`indent`/`defRPr`, a shape's `txBody/lstStyle`,
a slide master's `txStyles` (title/body/other), and the presentation's own
`defaultTextStyle` are all parsed by `drawingml/text.ts`'s `parseTextListStyle`
(shared across all four, since they're structurally identical — §21.1.2.4.12's
`lvl1pPr`..`lvl9pPr`) and its own `parseSharedParagraphProperties` helper (shared in turn
with `a:pPr` parsing, since a `lvlNpPr` is itself structurally a full paragraph-properties
element). A theme's `fmtScheme` fill/line/background-fill style matrices (`fillStyleLst`/
`lnStyleLst`/`bgFillStyleLst`) are parsed too (`theme.ts`, reusing `drawingml/fill.ts`'s/
`line.ts`'s own parsers — `bgFillStyleLst` is structurally just another fill list, same parser as
`fillStyleLst`), along with a shape/picture/connector's `p:style` `fillRef`/`lnRef`
(`presentationml/shape-tree.ts`'s `parseShapeStyle`) — together these resolve the fill/line
PowerPoint's Shape Styles gallery writes by default (a bare style reference, no explicit `spPr`
fill/line at all). **A slide/layout/master's own `p:bg/p:bgRef` (§19.3.1.6, new — a later
session) is now parsed too**, into `CommonSlideData.backgroundRef` (`presentationml/
common-slide-data.ts`'s `parseBackground`, reusing `shape-tree.ts`'s `parseStyleMatrixReference`
— exported from there for exactly this, since it's the same `CT_StyleMatrixReference` shape
`fillRef`/`lnRef` already parse) — previously only `p:bgPr` (a literal fill) was read at all,
silently dropping `p:bgRef` (a reference to the theme's own default background) as "unmodeled for
the skeleton." A real deck caught the gap: any slide/layout/master relying on the theme's own
background — PowerPoint's own default whenever nothing's been explicitly picked, not a rare
authoring choice — parsed as having no background anywhere in its own chain at all, which
`@pptx2html/to-html5` then rendered as no CSS background whatsoever, letting whatever sits behind
the slide show through instead. See `packages/presentation/CLAUDE.md`'s own note on this (`resolve/
background.ts`'s `resolveEffectiveBackground`, `resolve/style-matrix.ts`'s new
`resolveBackgroundStyleFill`) for the resolution side. A shape's
`a:custGeom` freeform outline (§20.1.9.8) is now parsed too — `drawingml/geometry.ts`'s `parsePath`/
`parsePathCommand`, new — into `CustomGeometry.pathLst`, for the "every point is a literal
coordinate" case (see the scope boundary below). A blip's
`r:embed` can also live nested inside `<a:blip>/<a:extLst>` rather than directly on `<a:blip>`
itself — PowerPoint's Icons gallery graphics are saved this way (`asvg:svgBlip`, an SVG-only
blip with no `r:embed` on `<a:blip>` at all) — `drawingml/fill.ts`'s `blipEmbedId` (new) checks
any extension child for its own `r:embed` as a fallback, used by both `parseFill`'s `blipFill`
case and `presentationml/shape-tree.ts`'s `parsePicture`; without this, a real, common class of
picture (any built-in PowerPoint icon) silently vanished from the shape tree entirely, with
no error — `parsePicture` returning `undefined` for an unresolvable image looks identical to a
missing/broken relationship, which is exactly what this was mistaken for until traced back to
the actual slide XML. A slide's
`p:timing` (§19.3.1.48 — element/build animation) is also parsed, by `presentationml/animation.ts`'s
`parseSlideTiming`, into `Slide.timing`; see `packages/presentation/CLAUDE.md`'s own note on this
for why (unlike everything else in this list) `to-html5` doesn't consume it yet. A slide's
`p:transition` (§19.3.1.49 — the whole-slide effect played when the presentation advances into it)
is parsed the same way by `presentationml/transition.ts`'s `parseSlideTransition`, into
`Slide.transition`, including resolving its optional `p:sndAc` sound via the same `resolveMedia`
already threaded through `readSlide` for the slide's shape tree/background — `to-html5` doesn't
consume this either, deliberately deferred alongside `timing`.

**Morph transitions are now recognized correctly (new), not silently misread as their
`mc:AlternateContent` fallback.** PowerPoint authors Morph as the _entire_ `<p:transition>` element
wrapped in `mc:AlternateContent` — `mc:Choice`'s branch carries the real `<p:transition>` with
`<p159:morph option="byObject"/>` inside it (plus, in practice, a `p14:dur` explicit-duration
attribute alongside `spd`), `mc:Fallback`'s branch carries a schema-legal plain `<p:transition>`
(typically `<p:fade/>`) for older PowerPoint versions — confirmed against a real Morph transition
authored between slides 3 and 4 of `apps/web-demo/src/Presentation1.pptx`, not just documentation.
`xml/query.ts`'s pre-existing `children()` always resolves `mc:AlternateContent` to the `Fallback`
branch, so `readSlide` used to silently parse every Morph-authored slide as a plain fade instead —
a real, previously-undetected gap this session's real fixture caught. Fixed via two additions:
`xml/query.ts`'s new `findAlternateContentChild` surfaces _both_ branches (`choice`/`resolved`)
instead of collapsing to one, and `transition.ts`'s new `pickTransitionNode` decides which whole
`<p:transition>` to actually parse — the `Choice` branch, but only when its own effect is one this
reader recognizes (currently just `p159:morph`), falling back to `resolved` (the Fallback, or the
plain unwrapped element if there was no `mc:AlternateContent` at all) otherwise, so an extension
effect this reader doesn't understand yet still gets _some_ real transition rather than none.
`SlideTransition.durationMs` (new, from `p14:dur`) is parsed alongside `spd`/`advClick`/`advTm` in
`parseSlideTransition`, unconditionally (not morph-specific — the attribute isn't schema-restricted
to it). `tsc -b`, `eslint`, `vitest run` (whole repo) and `prettier --check` all pass as of this
writing.

## Layout (mirrors `packages/presentation/src`'s own file layout 1:1)

- `opc/` — `package.ts` (wraps `fflate.unzipSync`, part lookup, content-type +
  relationship caching), `content-types.ts` ([Content_Types].xml, Override beats
  Default-by-extension), `relationships.ts` (`.rels` parsing + relative target
  resolution per OPC §9.3).
- `xml/` — `parse.ts` wraps `fast-xml-parser` with `preserveOrder: true` (required:
  the default mode regroups children by tag name, which would destroy z-order and
  run/break/field interleaving). `query.ts` has the node-walking helpers; it matches
  **elements** by local name (namespace-prefix-agnostic) but matches **attributes**
  by their exact raw name, because e.g. `p:sldMasterId` carries both a plain `id`
  and an `r:id` that collide if you strip prefixes. `query.ts`'s `children()` also
  transparently unwraps `mc:AlternateContent` to its `mc:Fallback` (or first
  `mc:Choice` if no fallback) wherever nodes are walked. `findAlternateContentChild` (new) is the
  escape hatch from that unwrapping, for a caller (so far just `presentationml/transition.ts`'s
  `pickTransitionNode`, for Morph) that wants to inspect the `Choice` branch's own content itself
  rather than always accepting whatever `children()` already collapsed to.
- `drawingml/`, `theme.ts`, `presentationml/` — one parser file per model file in
  `packages/presentation/src`. `drawingml/shape-common.ts`'s
  `parseNonVisualDrawingProperties` also parses a shape's `nvPr/ph` (§19.3.1.36,
  placeholder type/idx) when present — the reader only extracts this identity, it
  does **not** resolve placeholder inheritance (matching a slide placeholder to
  its layout/master counterpart to find an inherited transform); that's
  `@pptx2html/to-html5`'s job, since it needs the sibling layout/master shapes to
  do it (see that package's `placeholder.ts`). `theme.ts`'s `parseFormatScheme`
  parses `fmtScheme`'s `fillStyleLst`/`lnStyleLst` by reusing `drawingml/fill.ts`'s
  `parseFill`/`drawingml/line.ts`'s `parseLine` directly (a style-matrix entry is
  structurally identical to a shape's own `spPr` fill/line) — against a no-op
  `MediaResolver`, since a theme's style matrix referencing an image fill is
  vanishingly rare and not worth plumbing the theme part's own relationships through
  for. `presentationml/shape-tree.ts`'s `parseShapeStyle` parses a shape/picture/
  connector's own `p:style` `fillRef`/`lnRef`/`fontRef` (§19.3.1.44) the same way it resolves any
  other child colour, except `fontRef`'s own `idx` (§20.1.4.1.17 — `"major"`/`"minor"`/`"none"`,
  not a numeric style-matrix index the way `fillRef`/`lnRef`'s is) — `effectRef` is skipped,
  unparsed, matching `ShapeStyle`'s own scope in `packages/presentation`.
- `reader-context.ts` — `ReaderContext`, threaded through every parser: the
  `OpcPackage` plus part-name-keyed caches (`themes`, `slideMasters`,
  `slideLayouts`, `notesMasters`, `slides`, `media`) so a part referenced from
  multiple places is parsed once and the same object instance is reused.
- `mutable.ts` — `Mutable<T>` helper, used exactly once, for the
  `SlideMaster.layouts` ↔ `SlideLayout.master` cyclic reference (see below).
- `read-presentation.ts` → `presentationml/presentation.ts` — entry point:
  resolves the root `_rels/.rels` to find `ppt/presentation.xml`, then **must**
  read all slide masters (and the layouts they own) before reading slides, since
  slides resolve their layout relationship against `context.slideLayouts`, which
  `slide-master.ts` populates as a side effect of building each master.

### The SlideMaster ↔ SlideLayout cycle

`SlideMaster.layouts` and `SlideLayout.master` are mutually referential and every
field involved is `readonly`. `slide-master.ts` builds the master with a temporary
empty `layouts` (via `Mutable<SlideMaster>`), parses each layout against that same
master instance (so `layout.master === master` by identity), then assigns the
finished array and `Object.freeze`s the result. If you touch either file, keep this
two-phase order intact.

## Scope boundary

Anywhere `packages/presentation` says "unmodeled" (chart/smartArt/
oleObject internals, table style matrices, `effectStyleLst`, `p:style`'s
`effectRef`, theme overrides, custom shows, path gradients), the reader simply
never reads that XML — it doesn't parse-then-discard. `custGeom` path data (`drawingml/geometry.ts`'s
`parsePath`/`parsePathCommand`, new — a later session) is now parsed into `CustomGeometryPath`/
`PathCommand`, but only the "literal coordinates only" slice `packages/presentation`'s own type
models: a command whose point references a `gdLst` guide name instead of a literal number can't be
represented, so its whole enclosing `a:path` is dropped rather than parsed into a corrupted
outline (same "drop the unrepresentable unit, don't half-parse it" rule `parseAdjustValues`
already follows for `avLst`). Don't add handling for the genuinely-unmodeled items above
without first updating the corresponding type in `packages/presentation`.

## Tests

Unit tests sit next to each parser (`*.test.ts`). `read-presentation.test.ts` is the
full integration test: it builds a minimal but complete `.pptx` **in memory** via
`fflate.zipSync` from hand-written XML strings (no binary fixture checked in), then
asserts on the resulting graph — including that `slide.layout === master.layouts[0]`
and `notesSlide.slide === presentation.slides[0]` (object-identity reuse, not
re-parsing).

## Open TODOs / known gaps

- **No real (PowerPoint/LibreOffice-produced) `.pptx` fixture in the _automated_ tests.** Only the
  synthetic in-memory one exists there. `apps/web-demo/src/Presentation1.pptx` is real (now
  including a genuine PowerPoint-authored Morph transition, added this session), and it already
  paid for itself once — the original reasoning for wanting a real fixture (see below) was that
  real output exercises `mc:AlternateContent`/prefix-binding quirks hand-written XML won't
  naturally produce, and that's exactly what happened: manually inspecting slide 4's real
  `p:transition` is what caught that `children()`'s Fallback-preferring unwrap was silently
  swallowing Morph. It's still not wired into any automated test, though (the fix above was
  verified via a one-off script, then covered by hand-written XML fixtures shaped to match — see
  `transition.test.ts`'s "end-to-end" case). The original plan called a real _fixture_ out as a
  nice-to-have via `python-pptx`, but that's not installed in this environment and installing it
  needs a `pip install` — deliberately not done without asking first — and python-pptx can't author
  Morph specifically regardless, since it has no support for PowerPoint's newer extension
  transitions. See `/Users/james/.claude/plans/sleepy-snuggling-oasis.md` for the original
  reasoning.
- **`apps/web-demo` now renders into the page, not just `console.log`.** It depends
  on both `@pptx2html/reader` and `@pptx2html/to-html5`; `src/index.ts` calls
  `readPresentation` on the picked file, then `renderPresentation` and appends the
  result to `#output` in `src/index.html`. There's a real fixture at
  `apps/web-demo/src/Presentation1.pptx` for manual testing. Build/lint/format are
  clean; browser verification is on the user to do themselves (per their
  preference, not something to launch/kill the dev server for unprompted).
- **A sample-`.pptx`-generator script was drafted but never run**, at
  `/private/tmp/claude-501/-Users-james-Documents-Projects-Dev-pptx2html/5ecec0b5-082a-4a69-b5b0-ac3621203780/scratchpad/make-sample-pptx.mjs`
  (session-scoped scratch path, not part of the repo). It builds the same kind of
  minimal deck as the integration test's fixture, for manually testing the
  web-demo's file picker. It needs to be run from the repo root (so Node resolves
  `fflate` from the workspace `node_modules`) — that run was intentionally paused
  at the user's request. If it'd be useful longer-term, consider moving an
  equivalent script into the repo (e.g. `apps/web-demo/scripts/`) instead of
  leaving it in scratch.
