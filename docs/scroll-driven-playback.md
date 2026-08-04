# Scroll-driven playback (design note)

**Status: implemented (v1).** This document originally scoped a future feature and recorded
enabling groundwork done ahead of it; that groundwork is now built on top of, and the feature itself
exists — `@pptx2html/to-html5`'s `<pptx-scroll-presentation>` (`scroll-presentation-element.ts`,
`renderScrollPresentation`). See that package's CLAUDE.md's "Key design decision: scroll-driven
playback" for the actual mechanism; this document keeps the original motivation plus a record of
what changed from the original plan and what's still open. See root [`CLAUDE.md`](../CLAUDE.md) for
how this fits into the rest of the project.

## The feature

A `.pptx` deck can be authored so every advance — slide-to-slide and element-to-element — is driven
by a timer rather than a click: a slide auto-advances after N seconds
(`SlideTransition.advanceAfter`), and each animation effect starts after a fixed delay instead of
waiting for the presenter to click. PowerPoint calls this "Rehearse Timings" or a kiosk/self-running
show; the whole deck effectively becomes a fixed-length video with a known total duration (e.g. 30
seconds).

For decks like that, we want to offer an alternative to playing them back in real time: let the
user's **scroll position** stand in for the clock. Scrolling down moves forward through the show;
scrolling up moves backward; scrolling partway through a slide's entrance animation shows it
partway animated in. The deck becomes something you scrub through like a long web page, rather than
something you sit and watch — a "scrollytelling" presentation.

## The premise that changed: scroll position _is_ the trigger, not a substitute for one

The original version of this document reasoned that the feature could only work for a deck where
every wait had already been replaced by a number — no animation step anywhere gated on a click (or
mouseover, or waiting for a video to finish) — and called that the _uncommon_ case for a real
`.pptx`, a real limitation to design around.

That framing turned out to be the wrong one, surfaced while actually scoping the implementation.
**Scroll position doesn't need a deck to already be timer-driven — it replaces the click/auto-advance
trigger itself, the same way a click already does for this package's existing click-driven
element.** Reaching a point on the scroll axis _is_ the trigger. A deck's authored
`advanceOnClick`/`advanceAfter` and any `Slide.timing` node's `onClick`/`onNext` gating are simply
never consulted by scroll mode's timeline assembly at all — they're facts about how a _live
presenter_ advances the deck, irrelevant once scrolling has taken over that job entirely.

What's actually required turned out to be much narrower than "no click anywhere": every segment
scroll mode animates just needs a **duration** (to occupy scroll distance) and a **sequence
position** (to land in the right order) — and both already existed or defaulted sensibly for
virtually every deck, via machinery this package already had (`resolveTransitionDurationMs`'s
existing `speed`-absent default, `collectFadeAnimations`'s existing "play a click-gated fade
immediately at `delayMs: 0`" policy). The one genuine gap — a slide with no transition at all, or an
effect kind nothing here animates — got a new, scroll-specific synthetic default (a vertical push)
rather than blocking on it. See `packages/to-html5/CLAUDE.md`'s design decision for the full
reasoning and why this meant **no new resolvers were needed in `@pptx2html/presentation` at all** —
a genuine validation of that package's `resolve/` groundwork, not just a place it happened to be
reused.

Net result: there is **no "is this deck scroll-eligible" gate anywhere in the implementation** —
every deck renders in scroll mode, with the same play-what-you-can/report-what-you-can't philosophy
this package's real-time renderer already uses for everything else.

## What's been built

- **The central timing API in `@pptx2html/presentation`** (`resolve/timing.ts`) — unchanged from
  what this document originally described, and it turned out to need no additions at all for scroll
  mode: `resolveTimeNodeStartMs`/`resolveTimeNodeDuration` already answer exactly the per-node
  questions `collectFadeAnimations` (and, through it, scroll mode) needs, and
  `resolveTransitionDurationMs` already answers the transition-duration question with the right
  defaults built in.
- **Slide-transition (`push`/`fade`/`morph`) playback via the Web Animations API**, in both
  `<pptx-presentation>` (real-time, fire-and-forget) and now `<pptx-scroll-presentation>` (every
  `Animation` created once, paused, and scrubbed via `currentTime`) — the seekability this document
  originally called out as the motivating reason to prefer WAAPI over plain CSS `transition` is
  exactly what scroll mode's `seekTo(ms)` relies on.
- **The scroll↔time mapping itself** — `scroll-timeline.ts`'s `resolveScrollTimeline` (pure,
  assembles one absolute-millisecond timeline from a `Presentation`) plus
  `scroll-presentation-element.ts`'s `PptxScrollPresentationElement` (the DOM/scroll-listener side:
  a real scrollable track inside the element's own shadow DOM, RAF-throttled, `seekTo` writing
  `Animation.currentTime` directly). See `packages/to-html5/CLAUDE.md` for the full mechanism,
  including why nothing here uses `position: sticky`/`fixed` (a decision made specifically to leave
  room for "perfect pinning," below).

## Decisions resolved from "Explicitly undecided" (below)

- **Scroll↔time mapping: JS-driven (a real scroll listener + `Animation.currentTime`), not native
  CSS `animation-timeline: scroll()`/`view()`.** Chosen because it's what this document's own WAAPI-
  migration rationale already pointed toward, and because native scroll-timelines would mean
  re-deriving every `Animation`'s timing as scroll-range percentages against still-uneven browser
  support — noted as a possible future perf optimization, not pursued in v1.
- **Snap points: not implemented.** Scroll mode maps linearly; native CSS `scroll-snap-type` could
  layer on top later purely as a CSS addition, without touching `seekTo`'s own math at all.
- **Which package the feature lives in: entirely `@pptx2html/to-html5`.** No new code was needed in
  `@pptx2html/presentation` — see "The premise that changed" above.

## Still open

- **"Perfect pinning."** The visible content (`.pptx-scroll-viewport`) is a sibling of the scrolling
  element, not a descendant of it, and needs no positioning write on scroll at all — an earlier
  version nested it inside the scroller and compensated with a manual `transform:
translateY(scrollTop)` write per frame, which visibly jittered (the browser's own native scroll
  compositing of that real, non-trivial DOM subtree and this class's once-per-frame JS correction
  were two independent, not-quite-synchronized sources of truth for the same box — caught by the
  user in actual browser use). Moving the content outside the scrolling subtree entirely fixed the
  jitter and, if anything, makes a future exact/eased "perfect pinning" pass easier to build, not
  harder — there's no longer any browser-native scroll-linked positioning to work around at all, just
  plain JS state. See `packages/to-html5/CLAUDE.md`'s design decision for the full mechanism. Not
  started.
- **`minDwellMs` (the floor on a static slide's own scroll-timeline length) isn't configurable from
  the element yet**, unlike `pixelsPerSecond` — a fixed default in `scroll-timeline.ts`.
- Real click-driven build-step sequencing (an "On Click"/"After Previous" effect composing with its
  ancestor container's own offset) remains unmodeled in _both_ renderers, same as before this
  feature — scroll mode doesn't need it (see "The premise that changed"), and nothing about scroll
  mode changes that gap for the real-time element.

## Constraints for future animation work

Both still apply, and scroll mode's own implementation followed them:

- **Keep duration/timing computation separate from DOM-driving code**, and keep it in
  `@pptx2html/presentation`. Don't let a renderer grow its own private notion of "how long does
  this effect take."
- **Prefer the Web Animations API over plain CSS `transition`/`@keyframes`** for anything new — an
  `Animation`'s `currentTime` is directly readable/settable, which any scrubbable playback needs.
