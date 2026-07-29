# @pptx2html/presentation

Type-only in-memory DOM for a `.pptx` file. No parsing, no rendering, no runtime
logic — just interfaces that closely mirror ECMA-376 Part 1 (the OOXML spec),
organized the same way the spec itself is split: DrawingML (§20, shapes/color/fill/
text primitives shared across all of OOXML) vs. PresentationML (§19,
presentation-specific parts).

## Status: skeleton, consumed by a working reader

The type graph is complete enough that `@pptx2html/reader` parses real `.pptx`
byte streams into it end-to-end (see `packages/reader/CLAUDE.md`). Nothing renders
yet — there is no HTML/DOM output stage. `tsc -b`, `eslint`, `vitest run` and
`prettier --check` are all green.

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

## Layout

- `drawingml/` — shared primitives, one concept per file: `units.ts` (`Emu`,
  `Angle`, `Percentage`, `FontSize`), `color.ts` (srgb/scheme/system/preset/hsl +
  transform modifiers), `geometry.ts` (`Transform2D`, preset/custom geometry),
  `media.ts` (`MediaPart`), `fill.ts` (none/solid/gradient/pattern/blip),
  `line.ts`, `text.ts` (`TextBody`/`Paragraph`/`TextRun`/`LineBreak`/`TextField`),
  `shape-common.ts` (`NonVisualDrawingProperties`, `ShapeProperties`).
- `theme.ts` — `ColorScheme` (12 named slots), `FontScheme`, `FormatScheme`.
- `presentationml/` — `shape-tree.ts` (`Shape`/`Picture`/`ConnectionShape`/
  `GraphicFrame`/`GroupShape`, the `ShapeTreeNode` union), `table.ts`,
  `common-slide-data.ts` (`CommonSlideData`/`Background`, shared by every
  slide-like part), then the part hierarchy: `slide-master.ts` → `slide-layout.ts`
  → `slide.ts` → `notes.ts`, and `presentation.ts` as the DOM root.
- `index.ts` / `drawingml/index.ts` / `presentationml/index.ts` — barrel
  re-exports only.

Every `.ts` file in `packages/reader/src` corresponds 1:1 to a file here — if you
add a type here, the reader will need a matching parser file.

## Scope boundary — what's intentionally unmodeled

Each gap below is called out with a `§`-referenced comment at its point in the
code (search for "unmodeled for the skeleton"):

- Custom geometry path data (`custGeom` in `geometry.ts` carries no path).
- Table/fill/effect/line style matrices referenced by a shape's style index
  (`FormatScheme` is a bare name).
- Chart, SmartArt and OLE object internals — `GraphicPlaceholder` in
  `shape-tree.ts` only preserves which of the three it is.
- Bullet/numbering in paragraph properties.
- Text autofit in `TextBodyProperties`.
- Theme overrides at the slide/layout level (color/font map overrides).
- Custom shows and default text styles on the root `Presentation`.
- Path gradients (only linear-angle gradients are modeled in `GradientFill`).

**Rule for extending this**: if the reader needs to surface one of these, add the
type here first, then implement the parser in `packages/reader`. Don't let the
reader parse-then-discard data this package can't yet represent — that's exactly
the pattern `packages/reader/CLAUDE.md`'s "Scope boundary" section commits to
avoiding.

## Next likely steps

1. Pick one item off the unmodeled list above based on what a first renderer
   milestone actually needs (custom geometry and bullet/numbering are the two
   most likely to visibly matter first).
2. Start a rendering package (e.g. `packages/renderer` or directly in
   `apps/web-demo`) that walks this tree and produces HTML/CSS — nothing consumes
   the resolved object graph for display yet, only `console.log`s it
   (`apps/web-demo/src/index.ts`, per `packages/reader/CLAUDE.md`'s open TODOs).
