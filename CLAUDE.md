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

- **`packages/presentation`** — the type skeleton, complete enough to be a real target, plus (new)
  a `resolve/` directory of pure functions that compute OOXML's derived/effective values from that
  graph — effective transform, effective run formatting, resolved fill/colour, auto-number
  formatting/state — moved here from `to-html5` since none of it actually renders anything and any
  future renderer needs the identical answer. Still no parsing and no output-format-specific logic
  (CSS/SVG/DOM) — that stays a renderer's job. Also `presentationml/animation.ts`: a `Slide`'s
  optional `timing` (§19.3.1.48, `p:timing`) — the element/build animation timing tree (`par`/
  `seq`/`excl` containers plus `set`/`anim`/`animEffect`/`animClr`/`animMotion`/`animRot`/
  `animScale`/`cmd`/`audio`/`video` leaf behaviors, each with its own timing/conditions) and build
  list (per-paragraph and per-diagram/chart/graphic implicit builds). Also `presentationml/
transition.ts`: a `Slide`'s optional `transition` (§19.3.1.49, `p:transition`) — the single
  whole-slide effect (fade/wipe/push/split/wheel/zoom/..., the base schema's ten
  `EG_SlideTransition` shapes covering ~20 effect names) played when the presentation advances into
  that slide, plus its speed/advance-on-click/advance-after-N-ms settings and an optional sound.
  Neither `timing` nor `transition` has an inheritance chain to walk (unlike text/placeholder/fill,
  the parsed data already **is** the answer), but both now have a `resolve/` counterpart anyway:
  (new) `resolve/timing.ts`'s `resolveTimeNodeDuration`/`resolveSlideTimingDuration` compute a
  timing tree's own effective duration (or `'indefinite'` if any node is gated on a click/other
  external event with no numeric delay fallback), and `resolveTransitionDurationMs` answers the
  smaller `TransitionSpeed` → ms question (moved here from `to-html5`, which now consumes it) — see
  "Future feature: scroll-driven playback" below for why this groundwork exists.
- **`packages/reader`** — complete, parses real `.pptx` byte streams end-to-end into the
  `presentation` graph, including a theme's `fmtScheme` fill/line style matrices, a
  shape/picture/connector's own `p:style` `fillRef`/`lnRef`, a slide's `p:timing`
  (`presentationml/animation.ts`'s `parseSlideTiming`), and (new) a slide's `p:transition`
  (`presentationml/transition.ts`'s `parseSlideTransition`). Tested only against a synthetic
  in-memory fixture (built via `fflate.zipSync` from hand-written XML in
  `read-presentation.test.ts`) — see Todos below.
- **`packages/to-html5`** — layout is done: every slide and shape lands in the right place at the
  right size, including placeholder shapes that inherit position from their layout/master
  (`@pptx2html/presentation`'s `resolve/placeholder.ts` now, see above) and responsive
  scale-to-container-width via CSS percentages + `aspect-ratio`
  (no JS resize handling). Formatting is under way step by step: run-level font formatting,
  paragraph alignment, and bulleted/numbered lists (typeface/size/bold/italic/underline/
  strikethrough/color/alignment/bullets), fully resolved through the same OOXML text-property
  inheritance chain (run/paragraph → shape → placeholder layout/master → master category style →
  presentation default → theme), are all done — the inheritance-walking itself now lives in
  `presentation`'s `resolve/text-style.ts`/`bullet.ts`/`color.ts`, `to-html5` calls it and turns
  the result into DOM/CSS; shape/picture `fill`/`.line` → CSS background/border is also done
  (`fill.ts`), and so is slide background (`presentation`'s `resolve/background.ts`, falling back
  through layout/master, `to-html5` reusing `fill.ts` to paint it). A shape with
  no explicit `spPr` fill/line of its own — how PowerPoint's Shape Styles gallery writes shapes by
  default — now falls back to resolving its `p:style` `fillRef`/`lnRef` against the theme's
  format-scheme style matrix instead of rendering with no fill/border at all (`presentation`'s
  `resolve/style-matrix.ts`,
  wired into `to-html5`'s `shape-tree.ts`'s `effectiveFill`/`effectiveLine`). A shape's own preset outline now
  shapes that fill/line too, for a common subset of twelve presets
  (`shape-geometry.ts`) — `rect`/`roundRect`/`ellipse` via CSS `border-radius`, plus nine more
  (triangle, right triangle, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, 5-point
  star) via a real SVG `<path>` outline; every other preset and `custGeom` still render as a plain
  rectangle. Font size, border width, and list indentation — the non-position magnitudes this work
  introduces — scale with the slide via CSS container query units (`cqw`) rather than a fixed
  px/pt, the same no-JS philosophy as position/size (`units.ts`'s `emuToCqw`; see
  `packages/to-html5/CLAUDE.md`). Table styles and connector line rendering are still unstyled by
  design. Navigating between slides (new) now plays a `Slide.transition` as a real animation, via
  the Web Animations API rather than a plain CSS `transition`, when its effect is `push` or `fade`
  — the destination's own going forward, the _outgoing_ slide's own reversed going backward
  (undoing whichever animation brought it into view) — every other effect kind, and a slide with no
  `transition` at all, still take the
  pre-existing instant `display: none`/`block` swap (`presentation-element.ts`'s `#animatePush`/
  `#animateFade` vs. `#updateActiveSlide`; see that package's CLAUDE.md for the full design). This
  is the actively-developed package right now.
- **`apps/web-demo`** — wired to both `reader` and `to-html5`: picking a `.pptx` file renders it
  into the page, and (new) also `console.log`s each slide's parsed `timing` when present, since
  `to-html5` doesn't render animations yet (see Todos below) — the object model is the deliverable
  for now, inspectable via devtools. `apps/web-demo/src/Presentation1.pptx` is a real
  (non-synthetic) fixture for manual browser testing. Verifying changes here in an actual browser
  is on the user, by preference — don't launch/kill the dev server unprompted.
- **`packages/core`** — an unused scaffold left over from initial repo setup (`greet()`, one
  test). Nothing depends on it. Not part of the real pipeline — see Todos below.

## Significant todos

1. **Formatting pass in `to-html5`** — the big remaining piece, being done step by step.
   Run-level font formatting, paragraph alignment, and bulleted/numbered lists (typeface/size/
   bold/italic/underline/strikethrough/color/alignment/bullets) are all done, including full
   template/master inheritance (`SlideMaster.textStyles`, `Presentation.defaultTextStyle`, theme
   font scheme — alignment/bullet/indent all walk the exact same chain as character formatting,
   see `TextListStyleLevel` in `packages/presentation`) and auto-numbering (now
   `packages/presentation`'s `resolve/bullet.ts`, ten `AutoNumberScheme`s — see below). Shape/picture
   `ShapeProperties.fill`/`.line` → CSS
   background/border is also done (`to-html5`'s `fill.ts`), including gradients and an approximated take on
   pattern fills, and so is slide background (`presentation`'s `resolve/background.ts`, falling back slide → layout →
   master). Font size, border width, and list indentation scale with the slide via `cqw`
   (container query width units) rather than a fixed px/pt, so they resize along with
   position/size instead of looking disproportionate at other container widths. Shape geometry
   (`ShapeProperties.geometry`, previously every shape/picture rendering as a plain rectangle
   regardless of preset) is now also done for a common subset of twelve presets (`to-html5`'s `shape-geometry.ts`
   — `rect`/`roundRect`/`ellipse` via CSS `border-radius`, nine more via an SVG `<path>` overlay).
   A shape's `p:style` `fillRef`/`lnRef` — PowerPoint's Shape Styles gallery writes shapes this way
   by default, with no explicit `spPr` fill/line at all — now resolves against the theme's
   `FormatScheme.fillStyles`/`.lineStyles` as a fallback (`presentation`'s `resolve/style-matrix.ts`); without this a very
   common class of real-world shape rendered with no fill/border whatsoever, regardless of geometry.
   All of this inheritance-resolution logic (font/alignment/bullet chain, placeholder transform,
   background fallback, style-matrix, coordinate math, colour resolution) moved from `to-html5`
   into `packages/presentation/src/resolve/` this session, once it became clear none of it was
   actually renderer-specific — see that package's CLAUDE.md for the full rationale and file list.
   Still remaining: table cell fill/styles, the ~170 presets outside that subset, `custGeom`
   (freeform path data, unmodeled in `packages/presentation`), gradient/pattern/blip fill on the
   nine SVG-path presets (solid-only today), clipping a _picture_ to one of those nine (only
   `roundRect`/`ellipse` picture crops work so far), and `p:style`'s `effectRef` (effect styling,
   unmodeled — needs effect rendering to exist at all first). A shape's `p:style/fontRef` (its
   default run colour/typeface fallback when nothing else in the chain sets one) and a text body's
   effective vertical anchor (`a:bodyPr/@anchor`, rendered via flexbox) are both done now too — see
   `packages/to-html5/CLAUDE.md`'s two newest "Key design decision" sections. The DOM structure
   (`.pptx-shape`, `.pptx-paragraph`, `.pptx-run`, etc.) already exists so most of what's left
   should be additive CSS/SVG, not a restructure. See `packages/to-html5/CLAUDE.md`'s scope
   boundary for the full list and what's deliberately not modeled yet.
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
6. **Animations aren't rendered yet.** `presentation`'s new `SlideTiming`/`TimeNode`/
   `BuildListEntry` model the full timing tree and build list (see Status above), and `reader`
   parses them, but `to-html5` doesn't consume `Slide.timing` at all — per explicit direction this
   session, that's deliberately deferred; `apps/web-demo` just `console.log`s it for now. When
   picked up: driving CSS animations/transitions from this tree, converting `animMotion`'s raw path
   string to an SVG/CSS motion path, and turning a `bldP`'s implicit per-paragraph build into
   staged reveal are all renderer-side work with nothing else needed from `presentation`/`reader`
   first. A relative (`p:by`) colour shift on `animClr` and a colour value on an `anim`/`p:tav`
   keyframe are both unmodeled (absolute `from`/`to` colours only) — see `animation.ts`'s own doc
   comments in `packages/presentation`. Read "Future feature: scroll-driven playback" below (and
   `docs/scroll-driven-playback.md`) before implementing this — it constrains _how_ this should be
   driven, not just what it should render.
7. **Slide transitions: `push`/`fade` are rendered, every other effect kind still isn't.**
   `presentation`'s `SlideTransition`/`TransitionEffect`/`TransitionSoundAction`
   (`presentationml/transition.ts`) model a slide's `p:transition`, `reader` parses it into
   `Slide.transition`, and `to-html5`'s `PptxPresentationElement.goToSlide` (new) now plays it as a
   real animation — via the Web Animations API rather than a plain CSS `transition`, so playback is
   seekable (`docs/scroll-driven-playback.md`'s design note) — when `effect.kind` is `'push'` or
   `'fade'` — forward navigation plays the destination slide's own transition, backward navigation
   instead undoes the _outgoing_ slide's own transition (reversed direction for `push`) since
   that's the animation being undone, not whatever the destination separately authors (a bug fixed
   this session — see `packages/to-html5/CLAUDE.md` for the failure mode). Duration comes from a
   documented `fast`/`med`/`slow` → ms approximation (OOXML doesn't specify one, and this mapping
   now lives in `@pptx2html/presentation`'s `resolve/timing.ts`, not privately in `to-html5`), and
   navigation is ignored entirely while a transition is in flight.
   See `packages/to-html5/CLAUDE.md`'s "Key design decision: push/fade slide transitions" for the
   full mechanism. Still unrendered: every other `TransitionEffect.kind` (wipe, cut, dissolve,
   wheel, split, ...) — each a reasonably self-contained addition following the same pattern —
   `fade`'s `throughBlack: true` two-stage fade-to-black variant (renders as a plain crossfade for
   now), `advanceOnClick`/`advanceAfter` auto-advance timers, and `TransitionSoundAction` playback.
   PowerPoint's newer "fancy" transitions (Morph, Ripple, Honeycomb, ...), authored via `p14:`/
   `p15:`/`p159:` extensions rather than the base schema's `EG_SlideTransition` group, remain
   unmodeled in `presentation`/`reader` entirely — see `packages/presentation/CLAUDE.md`'s scope
   boundary. See also "Future feature: scroll-driven playback" below — `advanceOnClick`/
   `advanceAfter` in particular are exactly the fields a future scroll-time total-duration
   computation needs.
8. **This session's work is uncommitted.** The font/alignment/list-formatting pass —
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
   `shape-tree.ts`'s new `effectiveFill`/`effectiveLine` — plus the most recent change: all six of
   those inheritance-resolution files (`text-style.ts`, `placeholder.ts`, `background.ts`,
   `bullet.ts`, `coordinate.ts`, `style-matrix.ts`) relocated from `to-html5` into
   `packages/presentation/src/resolve/`, and `color.ts` split between a new pure resolver there and
   a thinned CSS-formatting wrapper left in `to-html5` — plus this session's addition of the
   animation/timing model — `presentationml/animation.ts` in both `presentation` and `reader`
   (`SlideTiming`/`TimeNode`/`BuildListEntry` and `parseSlideTiming` respectively), `Slide.timing`
   in `presentation`, its wiring in `reader`'s `presentationml/slide.ts`, and `apps/web-demo`'s new
   per-slide `console.log` of it — plus this session's addition of the slide transition model —
   `presentationml/transition.ts` in both `presentation` and `reader` (`SlideTransition`/
   `TransitionEffect`/`TransitionSoundAction` and `parseSlideTransition` respectively),
   `Slide.transition` in `presentation`, and its wiring in `reader`'s `presentationml/slide.ts` (not
   wired into `apps/web-demo`'s console logging at that point, per that round's explicit scope) —
   plus this session's `to-html5`-only addition of `push`/`fade` slide-transition rendering —
   `presentation-element.ts`'s `#transitions` state, `#animatePush`/`#animateFade`/
   `#beginTransitionFrame`, the `REVERSE_SIDE_DIRECTION` table, and the `.pptx-slide--transitioning`
   CSS class — plus a same-session bug fix to which slide's transition backward navigation reads
   (see `packages/to-html5/CLAUDE.md`'s design decision) — plus this session's central timing API
   and Web Animations API migration: `packages/presentation/src/resolve/timing.ts`
   (`resolveTransitionDurationMs`, moved from `to-html5`'s former `TRANSITION_DURATION_MS`;
   `resolveTimeNodeDuration`/`resolveSlideTimingDuration`, new) and its test file, plus
   `presentation-element.ts`'s `#currentAnimations` state (replacing `#transitionTimeoutId`) and
   `#awaitTransition`/`#finalizeTransition` (replacing `#scheduleFinalize`) now driving playback via
   `Element.animate()` instead of a plain CSS `transition` + `setTimeout`, and
   `presentation-element.test.ts`'s new `FakeAnimation`/`HTMLElement.prototype.animate` mock (since
   `happy-dom` implements no Web Animations API at all) — plus this session's fix for two shape-text
   rendering gaps found via `apps/web-demo/src/Presentation1.pptx`'s slide 1 circle: `FontReference`/
   `FontCollectionIndex` (new, `presentationml/shape-style.ts`) and `ShapeStyle.fontRef`, parsed by
   `reader`'s `parseShapeStyle`; `presentation`'s `resolveEffectiveRunProperties` taking an optional
   `ShapeStyle` and folding `fontRef` into `levelChain` above the master/placeholder chain but below
   the shape's own list style (initially shipped as the _lowest_-priority source instead, then
   corrected the same session once real content showed the master's `otherStyle` category clobbering
   a shape's own `fontRef` colour — see `packages/to-html5/CLAUDE.md`'s design decision), and
   its new `resolveEffectiveAnchor` (own value, else layout/master placeholder inheritance, else
   `'t'`) for `a:bodyPr/@anchor`; and `to-html5`'s `renderShape` now applying both — a flexbox
   `justify-content` for vertical anchoring, and threading `shape.style` through to
   `renderTextBody`/`renderRun`/`renderBulletSpan` for the `fontRef` fallback — are all working-tree
   changes on top of the `to-html5` commit; nothing since has been committed.

## Future feature: scroll-driven playback

**Not started.** Full design note: [`docs/scroll-driven-playback.md`](docs/scroll-driven-playback.md).
In short — a future feature will let a **fully time-resolved** deck (every advance driven by a
numeric delay, nothing waiting on a click) be scrubbed by scroll position instead of played back in
real time. Two constraints from that note apply to _all_ animation/transition work from here on,
not just a future scroll feature, so they're repeated here:

1. **Duration/timing computation belongs in `packages/presentation/resolve/`, as a pure function,
   not inside a renderer.** Whether a timing tree is fully time-resolved (no `onClick`/`onNext`/etc.
   wait anywhere) and how long any node/transition takes are renderer-agnostic questions — see
   `resolve/timing.ts`.
2. **Prefer the Web Animations API over plain CSS `transition`/`@keyframes` for anything new.** A
   plain CSS transition can only be started and left to run; a WAAPI `Animation`'s `currentTime` is
   directly readable/settable, which is exactly the capability a future scroll-time player needs.
   `to-html5`'s slide-transition playback already migrated to this — see
   `packages/to-html5/CLAUDE.md`.

Don't let `PptxPresentationElement.goToSlide`'s click-navigation policies (slide-granular,
non-interruptible mid-transition) be assumed as the only reasonable navigation shape elsewhere — a
future scroll-seek API is additive, not a replacement.

## Where to look

- `packages/presentation/CLAUDE.md` — the object-graph shape, what's intentionally unmodeled.
- `packages/reader/CLAUDE.md` — parsing details, the SlideMaster↔SlideLayout cycle, open gaps.
- `packages/to-html5/CLAUDE.md` — rendering design decisions (coordinate math, percentage-based
  responsive layout, placeholder inheritance), scope boundary, test layout.
- `docs/scroll-driven-playback.md` — design note for a future scroll-driven playback feature.
