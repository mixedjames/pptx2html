# @pptx2html/presentation

Type-only in-memory DOM for a `.pptx` file. No parsing, no rendering, no runtime
logic — just interfaces that closely mirror ECMA-376 Part 1 (the OOXML spec),
organized the same way the spec itself is split: DrawingML (§20, shapes/color/fill/
text primitives shared across all of OOXML) vs. PresentationML (§19,
presentation-specific parts).

## Status: skeleton, consumed by a working reader and a first-pass renderer

The type graph is complete enough that `@pptx2html/reader` parses real `.pptx`
byte streams into it end-to-end (see `packages/reader/CLAUDE.md`), and
`@pptx2html/to-html5` renders the result into a laid-out, responsively-scaled HTML5 DOM
with several formatting passes done (font inheritance including alignment — typeface/
size/bold/italic/underline/strike/color/alignment, all resolved through the full
placeholder/master/theme chain; shape/picture fill/line; slide background). Table cell
fill and table styles are the main remaining gap — see `packages/to-html5/CLAUDE.md`.
`tsc -b`, `eslint`, `vitest run` and `prettier --check` are all green.

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
  `shape-common.ts` (`NonVisualDrawingProperties` — including `Placeholder`
  type/index, used to resolve a placeholder shape's inherited position/size when
  it has none of its own — and `ShapeProperties`).
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
- Theme overrides at the slide/layout level (color/font map overrides) — `to-html5`'s colour
  resolution assumes the default clrMap (bg1→lt1, tx1→dk1, bg2→lt2, tx2→dk2).
- Custom shows on the root `Presentation`.
- Path gradients (only linear-angle gradients are modeled in `GradientFill`).
- Indent and bullet/numbering in a `TextListStyle` level — each level (`TextListStyleLevel`)
  models `algn` (as `alignment`) and `defRPr` (as `runProperties`), since that's what
  `to-html5`'s font/alignment inheritance pass needs, but not the other paragraph properties a
  real `a:lvlNpPr` can carry. `TextListStyle` (`drawingml/text.ts`) backs `TextBody.listStyle`,
  `SlideMaster.textStyles` (title/body/other) and `Presentation.defaultTextStyle` alike — see
  `to-html5/CLAUDE.md`'s font-inheritance design decision for how a consumer walks all three.

**Rule for extending this**: if the reader needs to surface one of these, add the
type here first, then implement the parser in `packages/reader`. Don't let the
reader parse-then-discard data this package can't yet represent — that's exactly
the pattern `packages/reader/CLAUDE.md`'s "Scope boundary" section commits to
avoiding.

## Next likely steps

1. Font/alignment inheritance, shape fill/line, and slide background are all done. Table cell
   fill is the next thing `@pptx2html/to-html5` needs to render — `TableCell.fill` already
   exists here and `reader` already parses it, `to-html5`'s `table.ts` just doesn't apply it yet
   (see that package's CLAUDE.md). Table _style matrices_ referenced by a table's style ID
   (banded rows, header row styling, etc.) are a separate, unmodeled gap — see `FormatScheme`
   above.
2. Custom geometry path data and bullet/numbering are the two remaining layout
   (not just formatting) gaps most likely to visibly matter.
