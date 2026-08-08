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
  `TransitionEffect` (new) also includes `MorphTransitionEffect` — PowerPoint's Morph (`p159:morph`,
  an `option` of `byObject`/`byWord`/`byChar`) — even though it's authored as a `p14:`/`p15:`/
  `p159:` extension rather than through the base schema's `EG_SlideTransition` group, so a consumer
  can dispatch on `effect.kind` uniformly regardless of source; `SlideTransition` also gained
  `durationMs`, an explicit millisecond override (`p14:dur`) that beats `speed`'s qualitative
  fast/med/slow mapping when present. Every _other_ fancy transition (Ripple, Honeycomb, ...)
  remains unmodeled — see this package's CLAUDE.md scope boundary. (new) `resolve/morph.ts`'s
  `resolveMorphMatch` is the shape-matching/diff step Morph itself needs — given two `Slide`s, it
  matches shapes by name (with `id`/positional tiebreaks for duplicates) and returns matched pairs
  plus appearing/disappearing lists; `to-html5` doesn't consume it yet — see Todo 7 below.
  Neither `timing` nor `transition` has an inheritance chain to walk (unlike text/placeholder/fill,
  the parsed data already **is** the answer), but both now have a `resolve/` counterpart anyway:
  (new) `resolve/timing.ts`'s `resolveTimeNodeDuration`/`resolveSlideTimingDuration` compute a
  timing tree's own effective duration (or `'indefinite'` if any node is gated on a click/other
  external event with no numeric delay fallback), and `resolveTransitionDurationMs` answers the
  smaller `TransitionSpeed` → ms question (moved here from `to-html5`, which now consumes it) — see
  "Scroll-driven playback" below for why this groundwork exists (now consumed by that feature too).
- **`packages/reader`** — complete, parses real `.pptx` byte streams end-to-end into the
  `presentation` graph, including a theme's `fmtScheme` fill/line style matrices, a
  shape/picture/connector's own `p:style` `fillRef`/`lnRef`, a slide's `p:timing`
  (`presentationml/animation.ts`'s `parseSlideTiming`), and a slide's `p:transition`
  (`presentationml/transition.ts`'s `parseSlideTransition`), including (new) correctly recognizing
  a Morph transition rather than silently reading its `mc:AlternateContent` fallback: `xml/
query.ts`'s `findAlternateContentChild` (new) surfaces both the `mc:Choice` and `mc:Fallback`
  branches of a wrapped element instead of always collapsing to the Fallback the way plain
  `children()`/`findChild()` do, and `transition.ts`'s `pickTransitionNode` (new) picks the Choice
  branch whenever its own effect is one this reader recognizes (currently just `p159:morph`),
  falling back to the Fallback's schema-legal effect otherwise — verified against a real Morph
  transition authored between slides 3 and 4 of `apps/web-demo/src/Presentation1.pptx`. Tested
  only against a synthetic in-memory fixture otherwise (built via `fflate.zipSync` from
  hand-written XML in `read-presentation.test.ts`) — see Todos below.
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
  star) via a real SVG `<path>` outline; a shape's freeform `custGeom` outline (new — a later
  session) now also renders this way, via `presentation`'s `CustomGeometry.pathLst` and
  `to-html5`'s `customGeometryPath` — including a boolean Merge Shapes result (multiple `a:path`
  subpaths rendering as a cutout under SVG's default fill rule, the case that motivated it, see
  `apps/web-demo/src/Presentation1.pptx`'s slide 3); every other preset outside the modeled subset
  still renders as a plain
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
  `#animateFade` vs. `#updateActiveSlide`; see that package's CLAUDE.md for the full design).
  Element/build animations off `Slide.timing` (new) also have a first slice rendered now: a
  `p:animEffect` whose `filter` is `"fade"` targeting a whole shape plays automatically, via the
  Web Animations API, as soon as its slide becomes active (`animation.ts`'s `collectFadeAnimations`
  plus `presentation-element.ts`'s `#playSlideAnimations`) — every other behavior kind, every other
  `animEffect` filter, and implicit paragraph/graphic builds are still unplayed and reported as
  unsupported, and real click-driven build-step sequencing isn't modeled at all (every fade plays
  at its own local delay, ignoring ancestor container timing/click-gating) — see that package's
  CLAUDE.md for the full design and its "not a faithful implementation" caveat. This is the
  actively-developed package right now.
- **`apps/web-demo`** — wired to both `reader` and `to-html5`: picking a `.pptx` file renders it
  into the page; `to-html5` now plays a slide's fade animations directly rather than needing the
  `console.log` of parsed `timing` this used to rely on for visibility (still logged, since most of
  `Slide.timing` remains unplayed — see Todos below) — the object model is otherwise still the
  deliverable, inspectable via devtools. `apps/web-demo/src/Presentation1.pptx` is a real
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
   freeform path data — now modeled and rendered, see below — for the "literal coordinates only"
   case, gradient/pattern/blip fill on the nine SVG-path presets and `custGeom` outlines
   (solid-only today), clipping a _picture_ to one of those nine or to a `custGeom` outline (only
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
3. **`presentation`'s unmodeled-for-the-skeleton list** — custom geometry path data is now modeled
   (`CustomGeometry.pathLst`, literal coordinates only) and rendered by `to-html5`; remaining items
   (effect/table style matrices, chart/SmartArt/OLE internals, text autofit, theme overrides at
   the slide/layout level, custom shows, path gradients) are the full list in
   `packages/presentation/CLAUDE.md`.
4. **No real-`.pptx` fixture in `reader`'s automated tests** — only the synthetic in-memory one.
   `apps/web-demo/src/Presentation1.pptx` is real but only exercised manually in the browser, not
   wired into any test. Would need `python-pptx` or similar to generate one; deliberately not
   installed without asking first.
5. **`packages/core` is dead weight** — decide whether to delete it or repurpose it; right now it
   does nothing and nothing references it.
6. **Animations: only a plain "fade" entrance/exit effect is rendered so far.** `presentation`'s
   `SlideTiming`/`TimeNode`/`BuildListEntry` model the full timing tree and build list (see Status
   above), `reader` parses them, and `to-html5` now plays one leaf behavior kind out of that tree —
   a `p:animEffect` whose `filter` is `"fade"` targeting a whole shape (`animation.ts`'s
   `collectFadeAnimations`, wired into `presentation-element.ts`) — via the Web Animations API, as
   soon as its slide becomes active. Everything else remains unplayed and is reported via the
   unsupported-feature log instead of silently dropped: every other `animEffect` filter (wipe,
   blinds, wheel, ...), `set`/`anim`/`animClr`/`animMotion`/`animRot`/`animScale`/`cmd`/`audio`/
   `video`, a fade targeting shape text/background rather than a whole shape, and `bldP`'s implicit
   per-paragraph build (or `bldDgm`/`bldChart`/`bldGraphic`). Real click-driven build-step
   sequencing is also unmodeled — `to-html5` has no in-slide "build step" concept distinct from
   slide-granular navigation, so every playable fade animates immediately at its own local start
   delay rather than composing with its ancestor `par`/`seq` containers' own offsets or actually
   waiting on an `onClick` trigger (see `packages/to-html5/CLAUDE.md`'s "Key design decision:
   `Slide.timing` fade animations" for the full reasoning). When picked up further: each additional
   `animEffect` filter is a fairly self-contained addition (some — wipe, blinds — imply a CSS/SVG
   clip-path technique of their own, not just a new opacity tween);
   `animMotion`/`animRot`/`animScale`/`animClr` each drive a different CSS property; converting
   `animMotion`'s raw path string to an SVG/CSS motion path is unstarted; turning a `bldP` into a
   staged reveal needs the in-slide build-step concept mentioned above, likely the biggest single
   piece left. A relative (`p:by`) colour shift on `animClr` and a colour value on an `anim`/`p:tav`
   keyframe are both unmodeled at the `presentation` layer (absolute `from`/`to` colours only) —
   see `animation.ts`'s own doc comments in `packages/presentation`. Read "Scroll-driven playback"
   below (and `docs/scroll-driven-playback.md`) before extending this
   further — it constrains _how_ this should be driven, not just what it should render; this
   session's fade pass already followed those constraints (WAAPI, duration/start-delay resolution
   in `@pptx2html/presentation`'s `resolve/timing.ts`) but is a known, documented approximation on
   the click-sequencing front, not the eventual scroll-driven answer.
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
   **Morph is now fully implemented — parsing, matching, and rendering.** `presentation`/`reader`
   recognize `p159:morph` (`MorphTransitionEffect`, its `option`, and the `p14:dur` explicit-
   duration extension attribute that often accompanies it), `@pptx2html/presentation`'s
   `resolve/morph.ts`'s `resolveMorphMatch` matches shapes between two `Slide`s' shape trees (by
   name, with `id`/positional tiebreaks for duplicates — PowerPoint's own algorithm is
   undocumented, this is a best-effort approximation), and `to-html5`'s `morph.ts`/
   `presentation-element.ts` play it as a real per-shape animation — see that package's CLAUDE.md's
   "Key design decision: Morph transitions" for the full mechanism. Verified end-to-end against the
   real Morph transition on `apps/web-demo/src/Presentation1.pptx`'s slide 4 (all 4 shapes matched,
   one of them a `custGeom` freeform outline). A key architectural constraint shaped this design:
   **no dependency on the `Presentation`/`Slide` object graph may survive `render()`**, an explicit
   requirement (this project's eventual output includes a freestanding HTML file, so the emitted
   markup shouldn't need the object model to keep working after generation) that ruled out an
   earlier draft storing the live `Slide` graphs in `PptxPresentationElement` — the match is instead
   reduced to plain shape-id data at `render()` time and that's all that's retained afterward. A
   below-confidence-threshold or first-slide match falls back to a plain crossfade and reports
   `morph-match-degraded` via `UnsupportedFeatureCollector` — the explicit requirement that a
   failed/degraded match be just as visible in the log as any other unsupported feature, not a
   silent approximation. Remaining Morph gaps, both documented rather than silent: fill/colour
   doesn't interpolate on a matched shape (position/size/rotation/opacity do), and `byWord`/`byChar`
   text-level matching isn't modeled (always plays as `byObject`, whole-shape matching) — see
   `packages/to-html5/CLAUDE.md`'s own "Next likely steps" for what picking either up would need.
   PowerPoint's other "fancy" transitions (Ripple, Honeycomb, Vortex, ...), also authored
   via `p14:`/`p15:`/`p159:` extensions, remain fully unmodeled — see
   `packages/presentation/CLAUDE.md`'s scope boundary. See also "Future feature: scroll-driven
   playback" below — `advanceOnClick`/`advanceAfter` in particular are exactly the fields a future
   scroll-time total-duration computation needs.
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
   changes on top of the `to-html5` commit; nothing since has been committed. (Much of the work
   this item describes has since landed in commits — `ecc5703`/`d10e1be`/`89a1ccc`/`2384279`/
   `fc47eb3` per `git log` — this item hasn't been re-synced to match; treat `git log`/`git diff`
   as authoritative over this paragraph for what's actually committed.) Most recently, still
   uncommitted: this session's `to-html5`-only addition of `Slide.timing` fade-animation playback —
   `to-html5`'s new `animation.ts` (`collectFadeAnimations`), `shape-tree.ts`'s `data-pptx-shape-id`
   dataset tagging, `presentation-element.ts`'s `#animations` state and `#playSlideAnimations`
   (called from `#updateActiveSlide`), and `@pptx2html/presentation`'s `resolve/timing.ts`'s new
   `resolveTimeNodeStartMs` — see `packages/to-html5/CLAUDE.md`'s "Key design decision:
   `Slide.timing` fade animations" for the design and its approximations. Also uncommitted, this
   session: `custGeom` freeform-outline support — `packages/presentation`'s `geometry.ts`
   (`CustomGeometry.pathLst`, `CustomGeometryPath`, `PathCommand`), `packages/reader`'s
   `drawingml/geometry.ts` (`parsePath`/`parsePathCommand`, literal coordinates only), and
   `to-html5`'s `shape-geometry.ts` (`customGeometryPath`) wired into `shape-tree.ts`'s
   `renderShape` alongside `presetShapePath` — motivated by and verified against
   `apps/web-demo/src/Presentation1.pptx`'s slide 3, a boolean "Subtract" Merge Shapes result.
   Also uncommitted, this session: Morph transition **parsing** (not rendering, see Todo 7 above) —
   `presentation`'s `transition.ts` (`MorphTransitionEffect`, `MorphOption`,
   `SlideTransition.durationMs`) and `resolve/timing.ts`'s `resolveTransitionDurationMs` (now takes
   an optional explicit-duration override); `reader`'s `xml/query.ts` (`findAlternateContentChild`,
   new) and `transition.ts` (`pickTransitionNode`, new, plus `p14:dur`/`p159:morph` parsing in
   `parseSlideTransition`/`parseTransitionEffect`), wired into `slide.ts`'s call site; and
   `to-html5`'s `presentation-element.ts`, updated to resolve `push`/`fade` duration once via the
   new two-argument `resolveTransitionDurationMs(speed, durationMs)` rather than each animate
   method recomputing it from `speed` alone. Verified end-to-end against
   `apps/web-demo/src/Presentation1.pptx`'s real slide 3→4 Morph transition (added to that fixture
   this session), which parses as `{ speed: 'slow', durationMs: 2000, effect: { kind: 'morph',
option: 'byObject' } }` — previously this reader silently read the `mc:Fallback`'s plain `fade`
   instead, since `children()`'s existing `mc:AlternateContent`-unwrapping always prefers Fallback
   over Choice. Also uncommitted: the Morph shape-matching resolver — `@pptx2html/presentation`'s
   new `resolve/morph.ts` (`resolveMorphMatch`, `MorphLeafShape`, `MorphShapeMatch`,
   `SlideMorphMatch`), matching two `Slide`s' shape trees by name with `id`/positional tiebreaks for
   duplicates, recursing into (but never matching) `GroupShape` containers — verified with
   `morph.test.ts`'s end-to-end case mirroring `apps/web-demo/src/Presentation1.pptx`'s real slide
   3/4 shape ids/names. Most recently, still uncommitted: Morph **rendering** — `to-html5`'s new
   `morph.ts` (`resolveSlideMorphMatch`, `MorphMatchSummary`, the confidence-threshold/crossfade-
   fallback logic) and `presentation-element.ts`'s new `#morphMatches` state, `#animateMorph`, and
   `goToSlide` dispatch branch — reducing `resolveMorphMatch`'s object-graph result to plain shape-id
   data before storing anything, per the no-object-model-after-`render()` requirement (see Todo 7
   above and Todo 9 below). Verified end-to-end against `Presentation1.pptx`'s real slide 3→4 morph
   (all 4 shapes matched, confirmed via a throwaway happy-dom script, not just unit tests) — that
   verification also caught and fixed a real bug (departing/arriving roles inverted for backward
   navigation) before it shipped. A second real bug was caught after that, this time by the user in
   actual browser use: a matched pair originally crossfaded both copies' opacity (the technique
   `#animateFade` correctly uses for whole-slide crossfades), which made both copies visibly
   translucent for the middle of the transition whenever they overlapped — obvious on a large solid
   shape, which is how it was caught (a big orange shape that should only translate was visibly
   showing text through it partway through). Fixed by animating only the arriving copy (fully
   opaque throughout) and hiding the departing copy instantly via a zero-duration `Animation`
   rather than fading it — see `packages/to-html5/CLAUDE.md`'s own note on both fixes.
9. **Freestanding-HTML-file export isn't fully possible yet — the remaining blocker is image
   embedding, not Morph.** Surfaced while scoping Morph's rendering design (see Todo 7 above): a
   real target for this project's output is a single freestanding `.html` file, and the Morph work
   was deliberately designed so nothing in `PptxPresentationElement` depends on the
   `Presentation`/`Slide` object graph after `render()` returns. But `fill.ts`'s and
   `shape-tree.ts`'s picture/blip-fill rendering still embeds images via `URL.createObjectURL` —
   blob URLs are only valid within the live browser tab/session that created them, so every image
   would break if this DOM were serialized to a static `.html` file and reopened later or
   elsewhere. Already tracked as a "never revoked, leaks memory" gap in `packages/to-html5/
CLAUDE.md`'s scope boundary, but the freestanding-file angle is a sharper reason to actually fix
   it: it needs to become a base64 `data:` URI instead. Deliberately kept out of the Morph work
   itself (a separate, pre-existing gap affecting every picture/fill on every slide, not specific to
   Morph) — a `font-family` referencing an uninstalled system font is a related, softer gap in the
   same direction, worth knowing about but not the priority.

## Scroll-driven playback

**Implemented.** Full design note (including what changed from the original plan):
[`docs/scroll-driven-playback.md`](docs/scroll-driven-playback.md). In short — `@pptx2html/to-html5`
now has a second custom element, `<pptx-scroll-presentation>` (`renderScrollPresentation`,
alongside the pre-existing click-driven `<pptx-presentation>`/`renderPresentation`), that positions
every slide transition and build animation on one scrubbable scroll timeline instead of playing them
back in real time. The original design note assumed this would only work for a **fully
time-resolved** deck (every advance driven by a numeric delay, nothing waiting on a click) — that
assumption turned out to be unnecessary: scroll position replaces the click/auto-advance trigger
entirely rather than requiring one to already be absent, so there's no such gate in the actual
implementation. See `packages/to-html5/CLAUDE.md`'s "Key design decision: scroll-driven playback"
for the full mechanism. Two constraints from the original design note guided this work and still
apply to _all_ future animation/transition work, not just this feature:

1. **Duration/timing computation belongs in `packages/presentation/resolve/`, as a pure function,
   not inside a renderer.** How long any node/transition takes is a renderer-agnostic question —
   see `resolve/timing.ts`. Scroll-driven playback needed zero new resolvers there, a validation of
   this rule rather than an exception to it.
2. **Prefer the Web Animations API over plain CSS `transition`/`@keyframes` for anything new.** A
   plain CSS transition can only be started and left to run; a WAAPI `Animation`'s `currentTime` is
   directly readable/settable, which is exactly the capability scroll-driven playback needs, and
   which `to-html5`'s slide-transition playback (both elements) is built on.

`PptxPresentationElement.goToSlide`'s click-navigation policies (slide-granular,
non-interruptible mid-transition) remain exactly that — a separate element,
`PptxScrollPresentationElement`, owns the scroll-driven navigation model instead, additive rather
than a replacement, per the original design note's own guidance.

## Where to look

- `packages/presentation/CLAUDE.md` — the object-graph shape, what's intentionally unmodeled.
- `packages/reader/CLAUDE.md` — parsing details, the SlideMaster↔SlideLayout cycle, open gaps.
- `packages/to-html5/CLAUDE.md` — rendering design decisions (coordinate math, percentage-based
  responsive layout, placeholder inheritance), scope boundary, test layout.
- `docs/scroll-driven-playback.md` — design note for the scroll-driven playback feature (implemented).
- `apps/pages/e2e/` — a minimal real-browser (Playwright/Chromium) regression suite, added after
  three real bugs in a row turned out to be invisible to `packages/to-html5`'s `happy-dom`-based
  unit tests (see that package's CLAUDE.md's "Key design decision: three behavioral scroll-mode
  bugs"). Scoped to real layout/containment and real Web Animations composite behavior specifically
  — not broad visual regression — and deliberately not wired into `npm test`/CI yet; run via
  `npm run test:e2e`.
