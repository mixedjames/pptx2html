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
  `presentation` graph. Tested only against a synthetic in-memory fixture (built via
  `fflate.zipSync` from hand-written XML in `read-presentation.test.ts`) — see Todos below.
- **`packages/to-html5`** — first pass: every slide and shape lands in the right place at the
  right size, including placeholder shapes that inherit position from their layout/master
  (`placeholder.ts`) and responsive scale-to-container-width via CSS percentages + `aspect-ratio`
  (no JS resize handling). No visual formatting yet — fill/line/color/font/text-run styling,
  backgrounds, table styles are all still unstyled by design. This is the actively-developed
  package right now.
- **`apps/web-demo`** — wired to both `reader` and `to-html5`: picking a `.pptx` file renders it
  into the page. `apps/web-demo/src/Presentation1.pptx` is a real (non-synthetic) fixture for
  manual browser testing. Verifying changes here in an actual browser is on the user, by
  preference — don't launch/kill the dev server unprompted.
- **`packages/core`** — an unused scaffold left over from initial repo setup (`greet()`, one
  test). Nothing depends on it. Not part of the real pipeline — see Todos below.

## Significant todos

1. **Formatting pass in `to-html5`** — the big remaining piece. `ShapeProperties.fill`/`.line`,
   run-level bold/italic/color/font, paragraph alignment, table cell fill/styles, slide
   backgrounds. The DOM structure (`.pptx-shape`, `.pptx-paragraph`, etc.) already exists so this
   should be additive CSS, not a restructure. See `packages/to-html5/CLAUDE.md`'s scope boundary
   for the full list and what's deliberately not modeled yet.
2. **Known `to-html5` limitations**, in rough order of how often they'll bite: placeholder
   matching doesn't model the spec's type-equivalence groups (e.g. slide `ctrTitle` matching
   layout `title`); rotation doesn't compose across nested groups; connectors render as an
   unstyled empty box (no line drawn — waits on the formatting pass); `renderPicture`'s object
   URLs are never revoked, so calling `.render()` repeatedly on the same element leaks blob URLs.
3. **`presentation`'s unmodeled-for-the-skeleton list** — custom geometry path data and
   bullet/numbering are the two most likely to visibly matter next once a real deck exercises
   them; full list in `packages/presentation/CLAUDE.md`.
4. **No real-`.pptx` fixture in `reader`'s automated tests** — only the synthetic in-memory one.
   `apps/web-demo/src/Presentation1.pptx` is real but only exercised manually in the browser, not
   wired into any test. Would need `python-pptx` or similar to generate one; deliberately not
   installed without asking first.
5. **`packages/core` is dead weight** — decide whether to delete it or repurpose it; right now it
   does nothing and nothing references it.
6. **This session's work is uncommitted.** Placeholder inheritance (`presentation`/`reader`
   parsing + `to-html5` resolution) and the responsive percentage-based scaling in `to-html5`,
   plus `apps/web-demo`'s wiring, are all working-tree changes on top of the `Added reader`
   commit — nothing since has been committed.

## Where to look

- `packages/presentation/CLAUDE.md` — the object-graph shape, what's intentionally unmodeled.
- `packages/reader/CLAUDE.md` — parsing details, the SlideMaster↔SlideLayout cycle, open gaps.
- `packages/to-html5/CLAUDE.md` — rendering design decisions (coordinate math, percentage-based
  responsive layout, placeholder inheritance), scope boundary, test layout.
