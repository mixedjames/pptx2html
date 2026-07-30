# pptx2html

Converts `.pptx` files to HTML5, as three packages chained together plus a browser demo:

```
.pptx bytes
  → @pptx2html/reader        readPresentation()      → Presentation object graph
  → @pptx2html/to-html5      renderPresentation()     → <pptx-presentation> (shadow DOM)
```

`@pptx2html/presentation` isn't a pipeline stage — it's the type-only object-graph definition
(ECMA-376/OOXML-shaped) that `reader` produces and `to-html5` consumes. Each package has its own
`CLAUDE.md` with real detail (file-by-file layout, key design decisions, exact scope boundaries);
this file is the map of how they fit together and what's still open, not a substitute for reading
them.

## Status

- **`packages/presentation`** — type skeleton, complete enough to be a real target. No parsing,
  no rendering, no runtime logic.
- **`packages/reader`** — complete, parses real `.pptx` byte streams end-to-end into the
  `presentation` graph, including a theme's `fmtScheme` fill/line style matrices and a
  shape/picture/connector's own `p:style` `fillRef`/`lnRef`. Tested only against a synthetic
  in-memory fixture (built via `fflate.zipSync` from hand-written XML in
  `read-presentation.test.ts`) — see Todos below.
- **`packages/to-html5`** — layout is done: every slide and shape lands in the right place at the
  right size, including placeholder shapes that inherit position from their layout/master
  (`placeholder.ts`) and responsive scale-to-container-width via CSS percentages + `aspect-ratio`
  (no JS resize handling). Formatting is under way step by step: run-level font formatting,
  paragraph alignment, and bulleted/numbered lists (typeface/size/bold/italic/underline/
  strikethrough/color/alignment/bullets), fully resolved through the same OOXML text-property
  inheritance chain (run/paragraph → shape → placeholder layout/master → master category style →
  presentation default → theme), are all done (`text-style.ts`, `bullet.ts`, `color.ts`);
  shape/picture `fill`/`.line` → CSS background/border is also done (`fill.ts`), and so is slide
  background (`background.ts`, falling back through layout/master, reusing `fill.ts`). A shape with
  no explicit `spPr` fill/line of its own — how PowerPoint's Shape Styles gallery writes shapes by
  default — now falls back to resolving its `p:style` `fillRef`/`lnRef` against the theme's
  format-scheme style matrix instead of rendering with no fill/border at all (`style-matrix.ts`,
  wired into `shape-tree.ts`'s `effectiveFill`/`effectiveLine`). A shape's own preset outline now
  shapes that fill/line too, for a common subset of twelve presets
  (`shape-geometry.ts`) — `rect`/`roundRect`/`ellipse` via CSS `border-radius`, plus nine more
  (triangle, right triangle, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, 5-point
  star) via a real SVG `<path>` outline; every other preset and `custGeom` still render as a plain
  rectangle. Font size, border width, and list indentation — the non-position magnitudes this work
  introduces — scale with the slide via CSS container query units (`cqw`) rather than a fixed
  px/pt, the same no-JS philosophy as position/size (`units.ts`'s `emuToCqw`; see
  `packages/to-html5/CLAUDE.md`). Table styles and connector line rendering are still unstyled by
  design. This is the actively-developed package right now.
- **`apps/web-demo`** — wired to both `reader` and `to-html5`: picking a `.pptx` file renders it
  into the page. `apps/web-demo/src/Presentation1.pptx` is a real (non-synthetic) fixture for
  manual browser testing. Verifying changes here in an actual browser is on the user, by
  preference — don't launch/kill the dev server unprompted.
- **`packages/core`** — an unused scaffold left over from initial repo setup (`greet()`, one
  test). Nothing depends on it. Not part of the real pipeline — see Todos below.

## Significant todos

1. **Formatting pass in `to-html5`** — the big remaining piece, being done step by step.
   Run-level font formatting, paragraph alignment, and bulleted/numbered lists (typeface/size/
   bold/italic/underline/strikethrough/color/alignment/bullets) are all done, including full
   template/master inheritance (`SlideMaster.textStyles`, `Presentation.defaultTextStyle`, theme
   font scheme — alignment/bullet/indent all walk the exact same chain as character formatting,
   see `TextListStyleLevel` in `packages/presentation`) and auto-numbering (`to-html5`'s
   `bullet.ts`, ten `AutoNumberScheme`s). Shape/picture `ShapeProperties.fill`/`.line` → CSS
   background/border is also done (`fill.ts`), including gradients and an approximated take on
   pattern fills, and so is slide background (`background.ts`, falling back slide → layout →
   master). Font size, border width, and list indentation scale with the slide via `cqw`
   (container query width units) rather than a fixed px/pt, so they resize along with
   position/size instead of looking disproportionate at other container widths. Shape geometry
   (`ShapeProperties.geometry`, previously every shape/picture rendering as a plain rectangle
   regardless of preset) is now also done for a common subset of twelve presets (`shape-geometry.ts`
   — `rect`/`roundRect`/`ellipse` via CSS `border-radius`, nine more via an SVG `<path>` overlay).
   A shape's `p:style` `fillRef`/`lnRef` — PowerPoint's Shape Styles gallery writes shapes this way
   by default, with no explicit `spPr` fill/line at all — now resolves against the theme's
   `FormatScheme.fillStyles`/`.lineStyles` as a fallback (`style-matrix.ts`); without this a very
   common class of real-world shape rendered with no fill/border whatsoever, regardless of geometry.
   Still remaining: table cell fill/styles, the ~170 presets outside that subset, `custGeom`
   (freeform path data, unmodeled in `packages/presentation`), gradient/pattern/blip fill on the
   nine SVG-path presets (solid-only today), clipping a _picture_ to one of those nine (only
   `roundRect`/`ellipse` picture crops work so far), and `p:style`'s `effectRef`/`fontRef` (effect
   styling and a fontRef text-colour fallback, both unmodeled). The DOM structure (`.pptx-shape`,
   `.pptx-paragraph`, `.pptx-run`, etc.) already exists so most of this should be additive CSS/SVG,
   not a restructure. See `packages/to-html5/CLAUDE.md`'s scope boundary for the full list and what's
   deliberately not modeled yet.
2. **Known `to-html5` limitations**, in rough order of how often they'll bite: placeholder
   matching doesn't model the spec's type-equivalence groups (e.g. slide `ctrTitle` matching
   layout `title`); rotation doesn't compose across nested groups; connectors still render as an
   unstyled empty box (no line drawn — now blocked on real geometry/an SVG line overlay, not the
   formatting pass itself, since a connector's line isn't its bounding box's border); object URLs
   (from `renderPicture` and now `fill.ts`'s blip fills) are never revoked, so calling `.render()`
   repeatedly on the same element leaks blob URLs.
3. **`presentation`'s unmodeled-for-the-skeleton list** — custom geometry path data is the item
   most likely to visibly matter next once a real deck exercises it; full list in
   `packages/presentation/CLAUDE.md`.
4. **No real-`.pptx` fixture in `reader`'s automated tests** — only the synthetic in-memory one.
   `apps/web-demo/src/Presentation1.pptx` is real but only exercised manually in the browser, not
   wired into any test. Would need `python-pptx` or similar to generate one; deliberately not
   installed without asking first.
5. **`packages/core` is dead weight** — decide whether to delete it or repurpose it; right now it
   does nothing and nothing references it.
6. **This session's work is uncommitted.** The font/alignment/list-formatting pass —
   `TextListStyle`/`TextListStyleLevel`/`Bullet`/`AutoNumberScheme`/`defaultRunProperties`/
   `listStyle`/`textStyles`/`defaultTextStyle` in `presentation` and their parsers in `reader`,
   plus `to-html5`'s `text-style.ts` (`levelChain`, `resolveEffectiveRunProperties`,
   `resolveEffectiveAlignment`, `resolveEffectiveBullet`, `resolveEffectiveIndent`), `bullet.ts`
   (`formatAutoNumber`, `NumberingState`) and `color.ts` — the shape fill/line pass — `to-html5`'s
   `fill.ts`, wired into `shape-tree.ts`'s `renderShape`/`renderPicture` — the slide background
   pass — `to-html5`'s `background.ts`, wired into `slide.ts`'s `renderSlide` — and the
   `cqw`-based responsive sizing pass — `units.ts`'s `emuToCqw`/`fontSizeToEmu`, `slide.ts`'s
   `container-type: inline-size`, and their use in `text.ts`/`fill.ts` — plus the shape-geometry
   pass added since — `to-html5`'s new `shape-geometry.ts` (`presetShapePath`,
   `nativeBorderRadius`), `fill.ts`'s new `applySvgFill`/`applySvgLine`, and their wiring into
   `shape-tree.ts`'s `renderShape`/`renderPicture` (`renderShapeOutline`) — plus the style-matrix
   pass added since that — `FormatScheme.fillStyles`/`.lineStyles` and `ShapeStyle`/
   `StyleMatrixReference` (new `presentationml/shape-style.ts`) in `presentation`, their parsers in
   `reader` (`theme.ts`'s `parseFormatScheme`, `presentationml/shape-tree.ts`'s `parseShapeStyle`),
   and `to-html5`'s new `style-matrix.ts` (`resolveStyleFill`/`resolveStyleLine`) wired into
   `shape-tree.ts`'s new `effectiveFill`/`effectiveLine` — are all working-tree changes on top of
   the `to-html5` commit; nothing since has been committed.

## Where to look

- `packages/presentation/CLAUDE.md` — the object-graph shape, what's intentionally unmodeled.
- `packages/reader/CLAUDE.md` — parsing details, the SlideMaster↔SlideLayout cycle, open gaps.
- `packages/to-html5/CLAUDE.md` — rendering design decisions (coordinate math, percentage-based
  responsive layout, placeholder inheritance), scope boundary, test layout.
