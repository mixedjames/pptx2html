import type {
  CornerDirection,
  EightDirection,
  InOutDirection,
  MorphOption,
  SideDirection,
  SlideTransition,
  TransitionEffect,
  TransitionOrientation,
  TransitionSoundAction,
  TransitionSpeed,
} from '@pptx2html/presentation';

import type { MediaResolver } from '../drawingml/fill.js';
import { parseBoolean, parseIntAttr } from '../drawingml/units.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, children, findChild, localName } from '../xml/query.js';

function parseEnum<T extends string>(
  value: string | undefined,
  valid: ReadonlySet<string>,
): T | undefined {
  return value !== undefined && valid.has(value) ? (value as T) : undefined;
}

const SPEEDS = new Set<TransitionSpeed>(['slow', 'med', 'fast']);
const ORIENTATIONS = new Set<TransitionOrientation>(['horz', 'vert']);
const EIGHT_DIRECTIONS = new Set<EightDirection>(['l', 'u', 'r', 'd', 'lu', 'ru', 'ld', 'rd']);
const SIDE_DIRECTIONS = new Set<SideDirection>(['l', 'u', 'r', 'd']);
const CORNER_DIRECTIONS = new Set<CornerDirection>(['lu', 'ru', 'ld', 'rd']);
const IN_OUT_DIRECTIONS = new Set<InOutDirection>(['in', 'out']);

const ORIENTATION_KINDS = new Set(['blinds', 'checker', 'comb', 'randomBar']);
const EMPTY_KINDS = new Set([
  'circle',
  'diamond',
  'dissolve',
  'newsflash',
  'plus',
  'random',
  'wedge',
]);
const EIGHT_DIRECTION_KINDS = new Set(['cover', 'pull']);
const SIDE_DIRECTION_KINDS = new Set(['push', 'wipe']);
const MORPH_OPTIONS = new Set<MorphOption>(['byObject', 'byWord', 'byChar']);

/**
 * Parses one member of §19.3.1.49's `EG_SlideTransition` choice group — the effect element that,
 * if present, is `p:transition`'s first child (before `p:sndAc`). PowerPoint's `p14`/`p15`/`p159`
 * "fancy" transition extensions live elsewhere (`p:extLst`) and aren't reached from here — see
 * `TransitionEffect`'s own doc comment.
 */
function parseTransitionEffect(node: XmlNode | undefined): TransitionEffect | undefined {
  if (!node) return undefined;
  const kind = localName(node);
  if (!kind) return undefined;

  if (ORIENTATION_KINDS.has(kind)) {
    const orientation = parseEnum<TransitionOrientation>(attr(node, 'dir'), ORIENTATIONS);
    return {
      kind: kind as 'blinds' | 'checker' | 'comb' | 'randomBar',
      ...(orientation ? { orientation } : {}),
    };
  }
  if (EMPTY_KINDS.has(kind)) {
    return {
      kind: kind as 'circle' | 'diamond' | 'dissolve' | 'newsflash' | 'plus' | 'random' | 'wedge',
    };
  }
  if (EIGHT_DIRECTION_KINDS.has(kind)) {
    const direction = parseEnum<EightDirection>(attr(node, 'dir'), EIGHT_DIRECTIONS);
    return { kind: kind as 'cover' | 'pull', ...(direction ? { direction } : {}) };
  }
  if (kind === 'cut') {
    const throughBlack = parseBoolean(attr(node, 'thruBlk'));
    return { kind: 'cut', ...(throughBlack !== undefined ? { throughBlack } : {}) };
  }
  if (kind === 'fade') {
    const throughBlack = parseBoolean(attr(node, 'thruBlk'));
    return { kind: 'fade', ...(throughBlack !== undefined ? { throughBlack } : {}) };
  }
  if (SIDE_DIRECTION_KINDS.has(kind)) {
    const direction = parseEnum<SideDirection>(attr(node, 'dir'), SIDE_DIRECTIONS);
    return { kind: kind as 'push' | 'wipe', ...(direction ? { direction } : {}) };
  }
  if (kind === 'split') {
    const orientation = parseEnum<TransitionOrientation>(attr(node, 'orient'), ORIENTATIONS);
    const direction = parseEnum<InOutDirection>(attr(node, 'dir'), IN_OUT_DIRECTIONS);
    return {
      kind: 'split',
      ...(orientation ? { orientation } : {}),
      ...(direction ? { direction } : {}),
    };
  }
  if (kind === 'strips') {
    const direction = parseEnum<CornerDirection>(attr(node, 'dir'), CORNER_DIRECTIONS);
    return { kind: 'strips', ...(direction ? { direction } : {}) };
  }
  if (kind === 'wheel') {
    const spokes = parseIntAttr(attr(node, 'spokes'));
    return { kind: 'wheel', ...(spokes !== undefined ? { spokes } : {}) };
  }
  if (kind === 'zoom') {
    const direction = parseEnum<InOutDirection>(attr(node, 'dir'), IN_OUT_DIRECTIONS);
    return { kind: 'zoom', ...(direction ? { direction } : {}) };
  }
  if (kind === 'morph') {
    const option = parseEnum<MorphOption>(attr(node, 'option'), MORPH_OPTIONS);
    return { kind: 'morph', ...(option ? { option } : {}) };
  }
  return undefined;
}

/** Parses p:sndAc (§19.3.1.49/50): a p:stSnd (play, with its own embedded p:snd) or p:endSnd (stop). */
function parseSoundAction(
  sndAcNode: XmlNode | undefined,
  resolveMedia: MediaResolver,
): TransitionSoundAction | undefined {
  if (!sndAcNode) return undefined;

  const stSnd = findChild(sndAcNode, 'stSnd');
  if (stSnd) {
    const sndNode = findChild(stSnd, 'snd');
    const embed = sndNode ? attr(sndNode, 'r:embed') : undefined;
    const sound = embed ? resolveMedia(embed) : undefined;
    if (!sound) return undefined;
    const loop = parseBoolean(attr(stSnd, 'loop'));
    return { kind: 'play', sound, ...(loop !== undefined ? { loop } : {}) };
  }

  return findChild(sndAcNode, 'endSnd') ? { kind: 'stop' } : undefined;
}

/** §19.3.1.49's `EG_SlideTransition` effect is p:transition's only child besides p:sndAc, if present. */
function findEffectNode(transitionNode: XmlNode): XmlNode | undefined {
  return children(transitionNode).find((child) => localName(child) !== 'sndAc');
}

/**
 * Given both branches an `mc:AlternateContent`-wrapped `p:transition` might offer (see
 * `xml/query.ts`'s `findAlternateContentChild`), decides which whole `<p:transition>` element to
 * actually parse: the `mc:Choice` branch's own version, but *only* when its own effect content is
 * one this reader recognizes (currently just `p159:morph`) — an extension effect we don't
 * understand yet falls back to the `mc:Fallback` branch's schema-legal effect (usually a plain
 * fade or push) instead of silently losing the transition altogether, since PowerPoint always
 * writes a compatible fallback for exactly this reason. A `p:transition` with no
 * `mc:AlternateContent` wrapper at all only ever has `resolved` set, so this always returns that.
 */
export function pickTransitionNode(lookup: {
  readonly choice?: XmlNode;
  readonly resolved?: XmlNode;
}): XmlNode | undefined {
  if (lookup.choice && parseTransitionEffect(findEffectNode(lookup.choice))) {
    return lookup.choice;
  }
  return lookup.resolved;
}

/** Parses a slide's p:transition (§19.3.1.49): how it's shown when the presentation advances to it. */
export function parseSlideTransition(
  transitionNode: XmlNode | undefined,
  resolveMedia: MediaResolver,
): SlideTransition | undefined {
  if (!transitionNode) return undefined;

  const speed = parseEnum<TransitionSpeed>(attr(transitionNode, 'spd'), SPEEDS);
  const durationMs = parseIntAttr(attr(transitionNode, 'p14:dur'));
  const advanceOnClick = parseBoolean(attr(transitionNode, 'advClick'));
  const advanceAfter = parseIntAttr(attr(transitionNode, 'advTm'));
  const effect = parseTransitionEffect(findEffectNode(transitionNode));
  const sound = parseSoundAction(findChild(transitionNode, 'sndAc'), resolveMedia);

  if (
    speed === undefined &&
    durationMs === undefined &&
    advanceOnClick === undefined &&
    advanceAfter === undefined &&
    effect === undefined &&
    sound === undefined
  ) {
    return undefined;
  }

  return {
    ...(speed ? { speed } : {}),
    ...(durationMs !== undefined ? { durationMs } : {}),
    ...(advanceOnClick !== undefined ? { advanceOnClick } : {}),
    ...(advanceAfter !== undefined ? { advanceAfter } : {}),
    ...(effect ? { effect } : {}),
    ...(sound ? { sound } : {}),
  };
}
