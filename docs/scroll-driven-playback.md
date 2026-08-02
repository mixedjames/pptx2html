# Scroll-driven playback (design note)

**Status: the feature itself is not started; the enabling groundwork (below) is done.** This
document exists to scope a future feature and to make sure the animation/transition work happening
now doesn't quietly make decisions that would block it later. See root
[`CLAUDE.md`](../CLAUDE.md) for how this fits into the rest of the project.

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

## Why this is hard: not every deck has a total duration

A deck authored for a live presenter, advanced by clicking, has animation steps that wait
indefinitely for that click. There is no way to know in advance how long a real audience will take
to click through it — so there's no "total duration" to map scroll position onto. This isn't an
edge case to work around; it's the normal way most `.pptx` files are authored.

The feature can only work for decks where every wait has been replaced by a number: no animation
step anywhere is gated on a click (or a mouseover, or waiting for a video to finish, etc.) —
everything has an explicit delay/duration instead. Whether a given deck qualifies is something we
can and should detect automatically, rather than assuming.

## What's been built so far: enabling groundwork, not the feature itself

Two pieces, both scoped to _enabling_ this feature later without building the feature itself yet —
both done now:

1. **A central timing API in the object model** (`@pptx2html/presentation`'s `resolve/timing.ts`).
   Given a slide's animation timing tree or its transition, `resolveTimeNodeDuration`/
   `resolveSlideTimingDuration` answer "how long does this take, or is it 'indefinite' because
   something's gated on a click/other external event with no numeric delay fallback" — exactly the
   "is this fully time-resolved" question above. Any future renderer — the existing real-time one,
   or a future scroll-driven one — needs the identical answer, so this lives with the object model,
   not duplicated inside a renderer. `resolveTransitionDurationMs` replaces the small piece of logic
   (`fast`/`med`/`slow` → milliseconds for slide transitions) that used to be private to the HTML
   renderer. The container-scheduling/repeat-interaction math is a documented approximation, not
   spec-exact — see that file's own doc comments — since nothing consumes exact values yet.
2. **Slide-transition (`push`/`fade`) playback in `@pptx2html/to-html5` now uses the Web Animations
   API** instead of plain CSS transitions. A plain CSS `transition` can only be started and left to
   run — there's no way to ask it "what does this look like 40% of the way through," which is
   exactly the capability a scroll-driven player needs (map a scroll position to a point in an
   animation, including scrubbing backward). The Web Animations API's `Animation` object _is_
   seekable — its `currentTime` can be read and written directly — so building on it now, even
   though nothing sets `currentTime` from scroll position yet, means the rendering mechanism itself
   won't need to be replaced when that lands later. (`happy-dom`, this repo's DOM test environment,
   implements no Web Animations API at all, so `to-html5`'s tests install their own
   `HTMLElement.prototype.animate` mock — see `packages/to-html5/CLAUDE.md`'s Tests section.)

Scroll position → time mapping — the part that actually makes this a _scroll-driven_ feature — is
still explicitly **not** built. That's future work now that the above groundwork exists.

## Constraints for this and future animation work

- **Keep duration/timing computation separate from DOM-driving code**, and keep it in
  `@pptx2html/presentation`. Don't let a renderer grow its own private notion of "how long does
  this effect take."
- **Don't drive new animation work with plain CSS `transition`/`@keyframes`** the way slide
  transitions used to be — prefer the Web Animations API, or otherwise keep the "what should this
  look like, and when" question answerable independently of how it's currently being played back.
- **Keep navigation APIs slide/click-shaped separate from any future seek/scrub API.** Today's
  `PptxPresentationElement.goToSlide()`/`.next()`/`.previous()` are slide-granular and intentionally
  ignore new input while a transition plays — a fine policy for click-driven use, wrong for
  continuous scrubbing. A future `seekTo`/`setProgress`-style API should be additive, not a special
  case bolted onto `goToSlide`.

## Explicitly undecided

- How scroll position maps onto duration (linear? per-slide snap points? something else).
- What actually drives scroll-linked rendering — native CSS scroll-driven animations
  (`animation-timeline: scroll()`/`view()`) vs. a `requestAnimationFrame`/scroll-listener approach
  that sets `Animation.currentTime` directly.
- Which package the feature itself lives in.

None of the above needs to be decided to do the enabling work above — they're listed here so they
aren't accidentally decided by omission either.
