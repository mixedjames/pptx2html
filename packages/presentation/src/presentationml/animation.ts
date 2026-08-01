import type { Color } from '../drawingml/color.js';
import type { Angle, Percentage } from '../drawingml/units.js';

/** A duration in milliseconds, or the literal "indefinite" (§19.7.32, ST_TLTime). */
export type TimeNodeDuration = number | 'indefinite';

/**
 * A time node's role within its parent sequence — in particular, `'clickEffect'`/`'withEffect'`/
 * `'afterEffect'` are what distinguish "on click", "with previous" and "after previous" in
 * PowerPoint's Animation Pane (§19.7.34, ST_TLTimeNodeType).
 */
export type TimeNodeRole =
  | 'tmRoot'
  | 'mainSeq'
  | 'interactiveSeq'
  | 'clickEffect'
  | 'withEffect'
  | 'afterEffect'
  | 'clickPar'
  | 'withGroup'
  | 'afterGroup';

/** PowerPoint's Animation Pane effect category (§19.7.30, ST_TLTimeNodePresetClassType). */
export type EffectPresetClass = 'entr' | 'exit' | 'emph' | 'path' | 'verb' | 'mediacall';

/** §19.7.33, ST_TLTimeNodeRestartType. */
export type TimeNodeRestart = 'always' | 'whenNotActive' | 'never';

/** §19.7.31, ST_TLTimeNodeFillType — what a node's properties do once its own timing ends. */
export type TimeNodeFill = 'remove' | 'freeze' | 'hold' | 'transition';

/** Event a `TimeCondition` fires on or listens for (§19.7.24, ST_TLTriggerEvent). */
export type TriggerEvent =
  | 'onBegin'
  | 'onEnd'
  | 'begin'
  | 'end'
  | 'onClick'
  | 'onDblClick'
  | 'onMouseOver'
  | 'onMouseOut'
  | 'onNext'
  | 'onPrev'
  | 'onStopAudio'
  | 'onBookMark';

/**
 * What a condition or behavior targets (§19.7.13/7.23, CT_TLTimeTargetElement /
 * CT_TLTextTargetElement). `shapeText`'s `paragraphRange`/`characterRange` are mutually exclusive
 * in the schema (`p:pRg` vs. `p:charRg`) and both 0-based/inclusive; neither present means the
 * whole text body.
 */
export type AnimationTarget =
  | { readonly kind: 'slide' }
  | { readonly kind: 'sound' }
  | { readonly kind: 'shape'; readonly shapeId: number }
  | { readonly kind: 'shapeBackground'; readonly shapeId: number }
  | {
      readonly kind: 'shapeText';
      readonly shapeId: number;
      readonly paragraphRange?: readonly [number, number];
      readonly characterRange?: readonly [number, number];
    };

/** A start/end/next/prev condition gating when a time node fires (§19.7.4, CT_TLTimeCondition). */
export interface TimeCondition {
  readonly delay?: number | 'indefinite';
  readonly event?: TriggerEvent;
  /** The time node whose `event` this condition listens for, when it's not the node's own parent. */
  readonly triggerTimeNodeId?: number;
  readonly target?: AnimationTarget;
}

/**
 * Timing/lifecycle data shared by every node in the tree, containers and behaviors alike
 * (§19.7.5, CT_TLCommonTimeNodeData / p:cTn). `presetClass`/`presetId`/`presetSubtype` together
 * identify one of PowerPoint's catalog effects (e.g. presetClass `"entr"` + presetId `1` is
 * "Appear") — the catalog itself is an application convention, not part of the OOXML schema, so
 * it isn't reproduced here; consumers that want a human name need their own lookup table.
 */
export interface CommonTimeNodeData {
  readonly id: number;
  readonly duration?: TimeNodeDuration;
  /** In the same 1000ths units as `Percentage` — e.g. 5000 means the node repeats 5 times. */
  readonly repeatCount?: number | 'indefinite';
  readonly repeatDuration?: TimeNodeDuration;
  readonly speed?: Percentage;
  readonly accelerate?: Percentage;
  readonly decelerate?: Percentage;
  readonly autoReverse?: boolean;
  readonly restart?: TimeNodeRestart;
  readonly fill?: TimeNodeFill;
  readonly role?: TimeNodeRole;
  readonly presetClass?: EffectPresetClass;
  readonly presetId?: number;
  readonly presetSubtype?: number;
  readonly startConditions?: readonly TimeCondition[];
  readonly endConditions?: readonly TimeCondition[];
}

/** Fields shared by every animation behavior (as opposed to container) node. */
interface AnimationBehaviorCommon {
  readonly common: CommonTimeNodeData;
  readonly target?: AnimationTarget;
  /** Property path(s) this behavior drives, e.g. `"ppt_x"`, `"style.opacity"` (§19.7.6, p:attrNameLst). */
  readonly attributeNames?: readonly string[];
}

/** A parallel-timed group of children, all starting per their own conditions (§19.7.18, p:par). */
export interface ParallelTimeNode {
  readonly kind: 'par';
  readonly common: CommonTimeNodeData;
  readonly children: readonly TimeNode[];
}

/** A sequence of children, each (unless `concurrent`) gated on the previous one (§19.7.21, p:seq). */
export interface SequenceTimeNode {
  readonly kind: 'seq';
  readonly common: CommonTimeNodeData;
  readonly concurrent?: boolean;
  readonly children: readonly TimeNode[];
}

/** At most one child active at a time (§19.7.10, p:excl) — rare; PowerPoint mostly emits `par`/`seq`. */
export interface ExclusiveTimeNode {
  readonly kind: 'excl';
  readonly common: CommonTimeNodeData;
  readonly children: readonly TimeNode[];
}

/** Sets a property to a literal value at the start of its interval, no interpolation (§19.7.22, p:set). */
export interface SetBehavior extends AnimationBehaviorCommon {
  readonly kind: 'set';
  readonly to?: string | Color;
}

/** One entry of an `AnimateBehavior`'s keyframe list (§19.7.2, p:tav). */
export interface AnimateKeyframe {
  /** Position within the behavior's own duration, 0-100000 representing 0%-100%. */
  readonly time: number;
  readonly value?: string;
  readonly formula?: string;
}

/** Interpolates an arbitrary named property between values (§19.7.1, p:anim). */
export interface AnimateBehavior extends AnimationBehaviorCommon {
  readonly kind: 'anim';
  readonly calcMode?: 'discrete' | 'lin' | 'fmla';
  readonly valueType?: 'str' | 'clr' | 'num';
  readonly from?: string;
  readonly to?: string;
  readonly by?: string;
  readonly keyframes?: readonly AnimateKeyframe[];
}

/**
 * Interpolates a colour property (§19.7.7, p:animClr). A relative `p:by` colour shift (delta
 * r/g/b or h/s/l rather than an absolute colour) is unmodeled — only the absolute `from`/`to`
 * form, which covers the overwhelming majority of authored colour animations, is represented.
 */
export interface AnimateColorBehavior extends AnimationBehaviorCommon {
  readonly kind: 'animClr';
  readonly colorSpace?: 'rgb' | 'hsl';
  readonly direction?: 'cw' | 'ccw';
  readonly from?: Color;
  readonly to?: Color;
}

/**
 * One of PowerPoint's built-in entrance/exit/emphasis effects — fade, wipe, blinds, etc.
 * (§19.7.9, p:animEffect). `filter` is the raw filter string (e.g. `"fade"`, `"wipe(right)"`,
 * `"blinds(horizontal)"`); the catalogue of valid filter names isn't modeled as its own type.
 */
export interface AnimateEffectBehavior extends AnimationBehaviorCommon {
  readonly kind: 'animEffect';
  readonly transition?: 'in' | 'out';
  readonly filter?: string;
}

/**
 * A relative x/y pair used by motion/scale behaviors, each a `Percentage` of the shape's own or
 * the slide's dimension (context-dependent — see each behavior) — distinct from `geometry.ts`'s
 * EMU-based `Point2D`, which is an absolute position.
 */
export interface AnimationPoint {
  readonly x: Percentage;
  readonly y: Percentage;
}

/**
 * Moves a shape along a path (§19.7.12, p:animMotion). `path` is the raw path string in
 * shape-fraction coordinates (e.g. `"M 0 0 L 0.5 0.5 E"`) — turning it into an SVG/CSS path is a
 * renderer's job, not this package's.
 */
export interface AnimateMotionBehavior extends AnimationBehaviorCommon {
  readonly kind: 'animMotion';
  readonly origin?: 'parent' | 'layout';
  readonly path?: string;
  readonly from?: AnimationPoint;
  readonly to?: AnimationPoint;
  readonly by?: AnimationPoint;
  /** Additional rotation applied along the path, in the same 60,000ths-of-a-degree units as `Angle` (§19.7.12, p:animMotion/@rAng). */
  readonly pathRotation?: Angle;
}

/** Rotates a shape (§19.7.19, p:animRot). */
export interface AnimateRotationBehavior extends AnimationBehaviorCommon {
  readonly kind: 'animRot';
  readonly from?: Angle;
  readonly to?: Angle;
  readonly by?: Angle;
}

/** Scales a shape (§19.7.20, p:animScale). */
export interface AnimateScaleBehavior extends AnimationBehaviorCommon {
  readonly kind: 'animScale';
  readonly from?: AnimationPoint;
  readonly to?: AnimationPoint;
  readonly by?: AnimationPoint;
}

/** Issues a media/OLE command — play, pause, an OLE verb (§19.7.8, p:cmd). */
export interface CommandBehavior extends AnimationBehaviorCommon {
  readonly kind: 'cmd';
  readonly type?: 'call' | 'evt' | 'verb';
  readonly command?: string;
}

/** Starts/stops an audio or video media clip as a timed node (§19.7.16/7.28, p:audio/p:video). */
export interface MediaNode extends AnimationBehaviorCommon {
  readonly kind: 'audio' | 'video';
}

/** A node in a slide's animation timing tree — a container or a leaf behavior (§19.7's `CT_TLTimeNodeList` members). */
export type TimeNode =
  | ParallelTimeNode
  | SequenceTimeNode
  | ExclusiveTimeNode
  | SetBehavior
  | AnimateBehavior
  | AnimateColorBehavior
  | AnimateEffectBehavior
  | AnimateMotionBehavior
  | AnimateRotationBehavior
  | AnimateScaleBehavior
  | CommandBehavior
  | MediaNode;

/** §19.7.29 ST_TLParaBuildType — how a text placeholder's paragraphs build in, absent real timing nodes for it. */
export type ParagraphBuildType = 'whole' | 'allAtOnce' | 'byParagraph';

/** An implicit per-paragraph build for a text shape (§19.5.8, p:bldP). */
export interface ParagraphBuild {
  readonly kind: 'paragraph';
  readonly shapeId: number;
  readonly buildType: ParagraphBuildType;
  /** Outline level paragraphs are grouped by, e.g. 1 to build by top-level paragraph. */
  readonly buildLevel?: number;
  readonly animateBackground?: boolean;
  readonly autoUpdateAnimBg?: boolean;
  readonly reverse?: boolean;
}

/**
 * An implicit build for a diagram, chart or other graphic frame (§19.5.2/5.3/5.9, p:bldDgm/
 * p:bldChart/p:bldGraphic). Each has its own richer `ST_TLDiagramBuildType`/`ST_TLChartBuildType`/
 * `ST_TLBuildType` enumeration of how the graphic's parts build in (by series, by category, by
 * node, ...); rather than modeling all three separately, `buildType` preserves the raw attribute
 * value verbatim.
 */
export interface GraphicBuild {
  readonly kind: 'diagram' | 'chart' | 'graphic';
  readonly shapeId: number;
  readonly buildType?: string;
}

export type BuildListEntry = ParagraphBuild | GraphicBuild;

/** A slide's element-level animation timing (§19.3.1.48, p:timing). */
export interface SlideTiming {
  /** The root of the timing tree — in practice always a `par` with `role: 'tmRoot'` when present. */
  readonly timeNodeTree?: TimeNode;
  readonly buildList?: readonly BuildListEntry[];
}
