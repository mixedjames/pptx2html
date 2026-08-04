import type { MediaPart } from '../drawingml/media.js';

/** §19.3.1.49, ST_TransitionSpeed. Absent means the spec default, `"fast"`. */
export type TransitionSpeed = 'slow' | 'med' | 'fast';

/** ST_Direction's horz/vert subset, shared by several transition effects below. */
export type TransitionOrientation = 'horz' | 'vert';

/** ST_TransitionEightDirectionType — the eight compass directions cover/pull travel in. */
export type EightDirection = 'l' | 'u' | 'r' | 'd' | 'lu' | 'ru' | 'ld' | 'rd';

/** ST_TransitionSideDirectionType — the four cardinal directions push/wipe travel in. */
export type SideDirection = 'l' | 'u' | 'r' | 'd';

/** ST_TransitionCornerDirectionType — the four diagonal directions strips travels in. */
export type CornerDirection = 'lu' | 'ru' | 'ld' | 'rd';

/** Shared by split/zoom — whether the effect opens outward or closes inward. */
export type InOutDirection = 'in' | 'out';

/**
 * Blinds/checker/comb/randomBar — a single effect axis (§19.3.1.49, CT_OrientationTransition).
 * Absent `orientation` means the spec default, `"horz"`.
 */
export interface OrientationTransitionEffect {
  readonly kind: 'blinds' | 'checker' | 'comb' | 'randomBar';
  readonly orientation?: TransitionOrientation;
}

/** Circle/diamond/dissolve/newsflash/plus/random/wedge — no parameters at all (§19.3.1.49, CT_Empty). */
export interface EmptyTransitionEffect {
  readonly kind: 'circle' | 'diamond' | 'dissolve' | 'newsflash' | 'plus' | 'random' | 'wedge';
}

/**
 * Cover/pull (§19.3.1.49, CT_EightDirectionTransition). Absent `direction` means the spec
 * default, `"l"`.
 */
export interface EightDirectionTransitionEffect {
  readonly kind: 'cover' | 'pull';
  readonly direction?: EightDirection;
}

/** Cut, optionally through black first (§19.3.1.49, CT_OptionalBlackTransition). */
export interface CutTransitionEffect {
  readonly kind: 'cut';
  readonly throughBlack?: boolean;
}

/** Fade, optionally through black first (§19.3.1.49, CT_FadeTransition). */
export interface FadeTransitionEffect {
  readonly kind: 'fade';
  readonly throughBlack?: boolean;
}

/**
 * Push/wipe (§19.3.1.49, CT_SideDirectionTransition). Absent `direction` means the spec default,
 * `"l"`.
 */
export interface SideDirectionTransitionEffect {
  readonly kind: 'push' | 'wipe';
  readonly direction?: SideDirection;
}

/**
 * Split — an axis plus whether the halves open outward or close inward (§19.3.1.49,
 * CT_SplitTransition). Absent `orientation`/`direction` mean the spec defaults, `"horz"`/`"out"`.
 */
export interface SplitTransitionEffect {
  readonly kind: 'split';
  readonly orientation?: TransitionOrientation;
  readonly direction?: InOutDirection;
}

/**
 * Strips (§19.3.1.49, CT_CornerDirectionTransition). Absent `direction` means the spec default,
 * `"lu"`.
 */
export interface CornerDirectionTransitionEffect {
  readonly kind: 'strips';
  readonly direction?: CornerDirection;
}

/** Wheel — spoke count around a circle (§19.3.1.49, CT_WheelTransition). Absent means 4. */
export interface WheelTransitionEffect {
  readonly kind: 'wheel';
  readonly spokes?: number;
}

/**
 * Zoom, in toward or out from the center (§19.3.1.49, CT_InOutTransition). Absent `direction`
 * means the spec default, `"in"`.
 */
export interface ZoomTransitionEffect {
  readonly kind: 'zoom';
  readonly direction?: InOutDirection;
}

/**
 * §19.3.1.49 has no spec reference for this at all — PowerPoint's Morph extension isn't in the
 * base schema (see `MorphTransitionEffect`'s doc comment). `"byObject"` (whole shapes only, no
 * word/character-level text matching) is PowerPoint's own default when the attribute is absent.
 */
export type MorphOption = 'byObject' | 'byWord' | 'byChar';

/**
 * PowerPoint's Morph transition (`p159:morph`, wrapped in `mc:AlternateContent` alongside a
 * schema-legal fallback effect — `@pptx2html/reader`'s job to pick the right branch, see
 * `packages/reader/CLAUDE.md`). Unlike every other `TransitionEffect` member, this isn't a single
 * canned animation applied to one slide — it matches shapes between the *outgoing* and *incoming*
 * slide (by name/id; PowerPoint's own matching algorithm is undocumented, not an OOXML-specified
 * one) and interpolates position/size/rotation/fill for each matched pair, crossfading anything
 * unmatched. `option` controls whether that matching (and interpolation) also descends to
 * individual words/characters within text, or stops at whole shapes. Still included directly in
 * this union, despite being extension-sourced rather than part of the base schema's
 * `EG_SlideTransition` choice group, so a consumer can dispatch on `effect.kind` uniformly
 * regardless of where an effect came from.
 */
export interface MorphTransitionEffect {
  readonly kind: 'morph';
  readonly option?: MorphOption;
}

/**
 * One of the slide transition effects defined by the base OOXML schema (§19.3.1.49's
 * `EG_SlideTransition` choice group), plus `MorphTransitionEffect` (see its own doc comment for
 * why an extension-sourced effect lives in this union too). PowerPoint's other newer "fancy"
 * transitions (Reveal, Ripple, Honeycomb, Vortex, Shred, Switch, Airplane, ...) remain unmodeled —
 * see this package's CLAUDE.md scope boundary.
 */
export type TransitionEffect =
  | OrientationTransitionEffect
  | EmptyTransitionEffect
  | EightDirectionTransitionEffect
  | CutTransitionEffect
  | FadeTransitionEffect
  | SideDirectionTransitionEffect
  | SplitTransitionEffect
  | CornerDirectionTransitionEffect
  | WheelTransitionEffect
  | ZoomTransitionEffect
  | MorphTransitionEffect;

/**
 * A sound to play during the transition, or an instruction to stop the previous slide's
 * already-playing sound (§19.3.1.49/50, p:sndAc/CT_TransitionSoundAction).
 */
export type TransitionSoundAction =
  | { readonly kind: 'play'; readonly sound: MediaPart; readonly loop?: boolean }
  | { readonly kind: 'stop' };

/**
 * How a slide transitions in when it's shown (§19.3.1.49, p:transition/CT_SlideTransition).
 * Unlike `SlideTiming`, this isn't a tree — it's a single effect applied to the whole slide, plus
 * how the presentation advances into it from the previous one.
 */
export interface SlideTransition {
  readonly speed?: TransitionSpeed;
  /**
   * An explicit millisecond duration (`p14:dur`, a PowerPoint extension attribute) that overrides
   * `speed`'s qualitative fast/med/slow mapping when present — `spd`'s three tiers aren't
   * fine-grained enough to represent PowerPoint's Transitions-ribbon duration slider. Seen in
   * practice alongside `p159:morph` but not schema-restricted to it.
   */
  readonly durationMs?: number;
  /** Whether a mouse click advances to this slide. Absent means the spec default, `true`. */
  readonly advanceOnClick?: boolean;
  /** Milliseconds after which the presentation auto-advances into this slide, if set. */
  readonly advanceAfter?: number;
  readonly effect?: TransitionEffect;
  readonly sound?: TransitionSoundAction;
}
