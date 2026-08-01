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
element). A theme's `fmtScheme` fill/line style matrices (`fillStyleLst`/`lnStyleLst`) are
parsed too (`theme.ts`, reusing `drawingml/fill.ts`'s/`line.ts`'s own parsers), along with a
shape/picture/connector's `p:style` `fillRef`/`lnRef` (`presentationml/shape-tree.ts`'s
`parseShapeStyle`) — together these resolve the fill/line PowerPoint's Shape Styles gallery
writes by default (a bare style reference, no explicit `spPr` fill/line at all). A slide's
`p:timing` (§19.3.1.48 — element/build animation) is also parsed, by `presentationml/animation.ts`'s
`parseSlideTiming`, into `Slide.timing`; see `packages/presentation/CLAUDE.md`'s own note on this
for why (unlike everything else in this list) `to-html5` doesn't consume it yet. `tsc -b`,
`eslint`, `vitest run` (whole repo) and `prettier --check` all pass as of this writing.

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
  `mc:Choice` if no fallback) wherever nodes are walked.
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
  connector's own `p:style` `fillRef`/`lnRef` (§19.3.1.44) the same way it resolves any
  other child colour — `effectRef`/`fontRef` are skipped, unparsed, matching
  `ShapeStyle`'s own scope in `packages/presentation`.
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

Anywhere `packages/presentation` says "unmodeled" (custGeom path data, chart/smartArt/
oleObject internals, table style matrices, `effectStyleLst`/`bgFillStyleLst`, `p:style`'s
`effectRef`/`fontRef`, theme overrides, custom shows, path gradients), the reader simply
never reads that XML — it doesn't parse-then-discard. Don't add handling for these
without first updating the corresponding type in `packages/presentation`.

## Tests

Unit tests sit next to each parser (`*.test.ts`). `read-presentation.test.ts` is the
full integration test: it builds a minimal but complete `.pptx` **in memory** via
`fflate.zipSync` from hand-written XML strings (no binary fixture checked in), then
asserts on the resulting graph — including that `slide.layout === master.layouts[0]`
and `notesSlide.slide === presentation.slides[0]` (object-identity reuse, not
re-parsing).

## Open TODOs / known gaps

- **No real (PowerPoint/LibreOffice-produced) `.pptx` fixture.** Only the synthetic
  in-memory one exists. The original plan called this out as a nice-to-have via
  `python-pptx`, but that's not installed in this environment and installing it
  needs a `pip install` — deliberately not done without asking first. If you want
  this, it's the one deferred item from the plan; see
  `/Users/james/.claude/plans/sleepy-snuggling-oasis.md` for the original
  reasoning (real output exercises `mc:AlternateContent`/prefix-binding quirks that
  hand-written XML won't naturally produce).
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
