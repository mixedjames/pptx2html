import type {
  AnimateBehavior,
  AnimateColorBehavior,
  AnimateEffectBehavior,
  AnimateKeyframe,
  AnimateMotionBehavior,
  AnimateRotationBehavior,
  AnimateScaleBehavior,
  AnimationPoint,
  AnimationTarget,
  BuildListEntry,
  Color,
  CommandBehavior,
  CommonTimeNodeData,
  EffectPresetClass,
  GraphicBuild,
  MediaNode,
  ParagraphBuild,
  ParagraphBuildType,
  SequenceTimeNode,
  SetBehavior,
  SlideTiming,
  TimeCondition,
  TimeNode,
  TimeNodeDuration,
  TimeNodeFill,
  TimeNodeRestart,
  TimeNodeRole,
  TriggerEvent,
} from '@pptx2html/presentation';

import { parseChildColor } from '../drawingml/color.js';
import { parseBoolean, parseIntAttr, parsePercentage } from '../drawingml/units.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, children, findChild, localName, textOf } from '../xml/query.js';

function parseEnum<T extends string>(
  value: string | undefined,
  valid: ReadonlySet<string>,
): T | undefined {
  return value !== undefined && valid.has(value) ? (value as T) : undefined;
}

const RESTART_TYPES = new Set<TimeNodeRestart>(['always', 'whenNotActive', 'never']);
const FILL_TYPES = new Set<TimeNodeFill>(['remove', 'freeze', 'hold', 'transition']);
const NODE_ROLES = new Set<TimeNodeRole>([
  'tmRoot',
  'mainSeq',
  'interactiveSeq',
  'clickEffect',
  'withEffect',
  'afterEffect',
  'clickPar',
  'withGroup',
  'afterGroup',
]);
const PRESET_CLASSES = new Set<EffectPresetClass>([
  'entr',
  'exit',
  'emph',
  'path',
  'verb',
  'mediacall',
]);
const TRIGGER_EVENTS = new Set<TriggerEvent>([
  'onBegin',
  'onEnd',
  'begin',
  'end',
  'onClick',
  'onDblClick',
  'onMouseOver',
  'onMouseOut',
  'onNext',
  'onPrev',
  'onStopAudio',
  'onBookMark',
]);

function parseTimeValue(value: string | undefined): TimeNodeDuration | undefined {
  if (value === undefined) return undefined;
  if (value === 'indefinite') return 'indefinite';
  return parseIntAttr(value);
}

/** Parses a p:tgtEl (§19.7.13/7.23) — sldTgt/sndTgt/spTgt(+bg/txEl)/inkTgt; inkTgt is unmodeled. */
function parseAnimationTarget(tgtElNode: XmlNode | undefined): AnimationTarget | undefined {
  if (!tgtElNode) return undefined;
  const [child] = children(tgtElNode);
  if (!child) return undefined;

  switch (localName(child)) {
    case 'sldTgt':
      return { kind: 'slide' };
    case 'sndTgt':
      return { kind: 'sound' };
    case 'spTgt': {
      const shapeId = parseIntAttr(attr(child, 'spid'));
      if (shapeId === undefined) return undefined;
      if (findChild(child, 'bg')) return { kind: 'shapeBackground', shapeId };

      const txEl = findChild(child, 'txEl');
      if (txEl) {
        const pRg = findChild(txEl, 'pRg');
        const charRg = findChild(txEl, 'charRg');
        const paragraphRange = pRg ? parseRange(pRg) : undefined;
        const characterRange = charRg ? parseRange(charRg) : undefined;
        return {
          kind: 'shapeText',
          shapeId,
          ...(paragraphRange ? { paragraphRange } : {}),
          ...(characterRange ? { characterRange } : {}),
        };
      }
      return { kind: 'shape', shapeId };
    }
    default:
      return undefined;
  }
}

function parseRange(node: XmlNode): readonly [number, number] | undefined {
  const start = parseIntAttr(attr(node, 'st'));
  const end = parseIntAttr(attr(node, 'end'));
  return start !== undefined && end !== undefined ? [start, end] : undefined;
}

/** Parses a p:cond (§19.7.4). */
function parseTimeCondition(condNode: XmlNode): TimeCondition {
  const delay = parseTimeValue(attr(condNode, 'delay'));
  const event = parseEnum<TriggerEvent>(attr(condNode, 'evt'), TRIGGER_EVENTS);
  const triggerTimeNodeId = parseIntAttr(attr(condNode, 'tn'));
  const target = parseAnimationTarget(findChild(condNode, 'tgtEl'));

  return {
    ...(delay !== undefined ? { delay } : {}),
    ...(event ? { event } : {}),
    ...(triggerTimeNodeId !== undefined ? { triggerTimeNodeId } : {}),
    ...(target ? { target } : {}),
  };
}

function parseConditionList(cTn: XmlNode, listName: string): readonly TimeCondition[] | undefined {
  const listNode = findChild(cTn, listName);
  if (!listNode) return undefined;
  const conditions = children(listNode)
    .filter((child) => localName(child) === 'cond')
    .map(parseTimeCondition);
  return conditions.length > 0 ? conditions : undefined;
}

/** Parses a p:cTn's own timing/lifecycle attributes and conditions (§19.7.5) — not its childTnLst. */
function parseCommonTimeNodeData(cTn: XmlNode): CommonTimeNodeData {
  const id = parseIntAttr(attr(cTn, 'id')) ?? 0;
  const duration = parseTimeValue(attr(cTn, 'dur'));
  const repeatCount = parseTimeValue(attr(cTn, 'repeatCount'));
  const repeatDuration = parseTimeValue(attr(cTn, 'repeatDur'));
  const speed = parsePercentage(attr(cTn, 'spd'));
  const accelerate = parsePercentage(attr(cTn, 'accel'));
  const decelerate = parsePercentage(attr(cTn, 'decel'));
  const autoReverse = parseBoolean(attr(cTn, 'autoRev'));
  const restart = parseEnum<TimeNodeRestart>(attr(cTn, 'restart'), RESTART_TYPES);
  const fill = parseEnum<TimeNodeFill>(attr(cTn, 'fill'), FILL_TYPES);
  const role = parseEnum<TimeNodeRole>(attr(cTn, 'nodeType'), NODE_ROLES);
  const presetClass = parseEnum<EffectPresetClass>(attr(cTn, 'presetClass'), PRESET_CLASSES);
  const presetId = parseIntAttr(attr(cTn, 'presetID'));
  const presetSubtype = parseIntAttr(attr(cTn, 'presetSubtype'));
  const startConditions = parseConditionList(cTn, 'stCondLst');
  const endConditions = parseConditionList(cTn, 'endCondLst');

  return {
    id,
    ...(duration !== undefined ? { duration } : {}),
    ...(repeatCount !== undefined ? { repeatCount } : {}),
    ...(repeatDuration !== undefined ? { repeatDuration } : {}),
    ...(speed !== undefined ? { speed } : {}),
    ...(accelerate !== undefined ? { accelerate } : {}),
    ...(decelerate !== undefined ? { decelerate } : {}),
    ...(autoReverse !== undefined ? { autoReverse } : {}),
    ...(restart ? { restart } : {}),
    ...(fill ? { fill } : {}),
    ...(role ? { role } : {}),
    ...(presetClass ? { presetClass } : {}),
    ...(presetId !== undefined ? { presetId } : {}),
    ...(presetSubtype !== undefined ? { presetSubtype } : {}),
    ...(startConditions ? { startConditions } : {}),
    ...(endConditions ? { endConditions } : {}),
  };
}

function parseChildTimeNodes(cTn: XmlNode): readonly TimeNode[] {
  const childTnLst = findChild(cTn, 'childTnLst');
  if (!childTnLst) return [];
  return children(childTnLst)
    .map(parseTimeNode)
    .filter((node): node is TimeNode => node !== undefined);
}

interface CommonBehaviorData {
  readonly common: CommonTimeNodeData;
  readonly target?: AnimationTarget;
  readonly attributeNames?: readonly string[];
}

/** Parses a behavior's shared p:cBhvr wrapper (§19.7.6): its own p:cTn, p:tgtEl and p:attrNameLst. */
function parseCommonBehaviorData(node: XmlNode, wrapperName: string): CommonBehaviorData {
  const wrapper = findChild(node, wrapperName);
  const cTn = wrapper ? findChild(wrapper, 'cTn') : undefined;
  const common = cTn ? parseCommonTimeNodeData(cTn) : { id: 0 };
  const target = wrapper ? parseAnimationTarget(findChild(wrapper, 'tgtEl')) : undefined;

  const attrNameLst = wrapper ? findChild(wrapper, 'attrNameLst') : undefined;
  const attributeNames = attrNameLst
    ? children(attrNameLst)
        .filter((child) => localName(child) === 'attrName')
        .map(textOf)
    : undefined;

  return {
    common,
    ...(target ? { target } : {}),
    ...(attributeNames && attributeNames.length > 0 ? { attributeNames } : {}),
  };
}

function parseSetValue(toNode: XmlNode | undefined): string | Color | undefined {
  if (!toNode) return undefined;
  const [child] = children(toNode);
  if (!child) return undefined;
  if (localName(child) === 'clrVal') return parseChildColor(child);
  return attr(child, 'val');
}

function parseSetBehavior(node: XmlNode): SetBehavior {
  const to = parseSetValue(findChild(node, 'to'));
  return {
    kind: 'set',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(to !== undefined ? { to } : {}),
  };
}

/** A keyframe's value is a CT_TLAnimVariant (§19.7.25) — str/int/flt/boolVal all carry a `val`
 * attribute; a colour keyframe (`clrVal`) is unmodeled, matching `AnimateColorBehavior`'s own
 * from/to-only scope. */
function parseKeyframeValue(valNode: XmlNode | undefined): string | undefined {
  if (!valNode) return undefined;
  const [child] = children(valNode);
  if (!child || localName(child) === 'clrVal') return undefined;
  return attr(child, 'val');
}

function parseKeyframe(tavNode: XmlNode): AnimateKeyframe | undefined {
  const time = parseIntAttr(attr(tavNode, 'tm'));
  if (time === undefined) return undefined;
  const value = parseKeyframeValue(findChild(tavNode, 'val'));
  const fmla = attr(tavNode, 'fmla');
  return { time, ...(value !== undefined ? { value } : {}), ...(fmla ? { formula: fmla } : {}) };
}

function parseAnimateBehavior(node: XmlNode): AnimateBehavior {
  const calcMode = parseEnum<'discrete' | 'lin' | 'fmla'>(
    attr(node, 'calcmode'),
    new Set(['discrete', 'lin', 'fmla']),
  );
  const valueType = parseEnum<'str' | 'clr' | 'num'>(
    attr(node, 'valueType'),
    new Set(['str', 'clr', 'num']),
  );
  const tavLst = findChild(node, 'tavLst');
  const keyframes = tavLst
    ? children(tavLst)
        .filter((child) => localName(child) === 'tav')
        .map(parseKeyframe)
        .filter((kf): kf is AnimateKeyframe => kf !== undefined)
    : undefined;

  const from = attr(node, 'from');
  const to = attr(node, 'to');
  const by = attr(node, 'by');

  return {
    kind: 'anim',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(calcMode ? { calcMode } : {}),
    ...(valueType ? { valueType } : {}),
    ...(from ? { from } : {}),
    ...(to ? { to } : {}),
    ...(by ? { by } : {}),
    ...(keyframes && keyframes.length > 0 ? { keyframes } : {}),
  };
}

function parseAnimateColorBehavior(node: XmlNode): AnimateColorBehavior {
  const colorSpace = parseEnum<'rgb' | 'hsl'>(attr(node, 'clrSpc'), new Set(['rgb', 'hsl']));
  const direction = parseEnum<'cw' | 'ccw'>(attr(node, 'dir'), new Set(['cw', 'ccw']));
  const fromNode = findChild(node, 'from');
  const toNode = findChild(node, 'to');
  const from = fromNode ? parseChildColor(fromNode) : undefined;
  const to = toNode ? parseChildColor(toNode) : undefined;

  return {
    kind: 'animClr',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(colorSpace ? { colorSpace } : {}),
    ...(direction ? { direction } : {}),
    ...(from ? { from } : {}),
    ...(to ? { to } : {}),
  };
}

function parseAnimateEffectBehavior(node: XmlNode): AnimateEffectBehavior {
  const transition = parseEnum<'in' | 'out'>(attr(node, 'transition'), new Set(['in', 'out']));
  const filter = attr(node, 'filter');
  return {
    kind: 'animEffect',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(transition ? { transition } : {}),
    ...(filter ? { filter } : {}),
  };
}

function parsePoint(node: XmlNode | undefined): AnimationPoint | undefined {
  if (!node) return undefined;
  const x = parsePercentage(attr(node, 'x'));
  const y = parsePercentage(attr(node, 'y'));
  return x !== undefined && y !== undefined ? { x, y } : undefined;
}

function parseAnimateMotionBehavior(node: XmlNode): AnimateMotionBehavior {
  const origin = parseEnum<'parent' | 'layout'>(
    attr(node, 'origin'),
    new Set(['parent', 'layout']),
  );
  const path = attr(node, 'path');
  const pathRotation = parseIntAttr(attr(node, 'rAng'));
  const from = parsePoint(findChild(node, 'from'));
  const to = parsePoint(findChild(node, 'to'));
  const by = parsePoint(findChild(node, 'by'));

  return {
    kind: 'animMotion',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(origin ? { origin } : {}),
    ...(path ? { path } : {}),
    ...(from ? { from } : {}),
    ...(to ? { to } : {}),
    ...(by ? { by } : {}),
    ...(pathRotation !== undefined ? { pathRotation } : {}),
  };
}

function parseAnimateRotationBehavior(node: XmlNode): AnimateRotationBehavior {
  const from = parseIntAttr(attr(node, 'from'));
  const to = parseIntAttr(attr(node, 'to'));
  const by = parseIntAttr(attr(node, 'by'));
  return {
    kind: 'animRot',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(from !== undefined ? { from } : {}),
    ...(to !== undefined ? { to } : {}),
    ...(by !== undefined ? { by } : {}),
  };
}

function parseAnimateScaleBehavior(node: XmlNode): AnimateScaleBehavior {
  const from = parsePoint(findChild(node, 'from'));
  const to = parsePoint(findChild(node, 'to'));
  const by = parsePoint(findChild(node, 'by'));
  return {
    kind: 'animScale',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(from ? { from } : {}),
    ...(to ? { to } : {}),
    ...(by ? { by } : {}),
  };
}

function parseCommandBehavior(node: XmlNode): CommandBehavior {
  const type = parseEnum<'call' | 'evt' | 'verb'>(
    attr(node, 'type'),
    new Set(['call', 'evt', 'verb']),
  );
  const command = attr(node, 'cmd');
  return {
    kind: 'cmd',
    ...parseCommonBehaviorData(node, 'cBhvr'),
    ...(type ? { type } : {}),
    ...(command ? { command } : {}),
  };
}

function parseMediaNode(kind: 'audio' | 'video', node: XmlNode): MediaNode {
  return { kind, ...parseCommonBehaviorData(node, 'cMediaNode') };
}

function parseTimeNode(node: XmlNode): TimeNode | undefined {
  const name = localName(node);
  switch (name) {
    case 'par':
    case 'excl': {
      const cTn = findChild(node, 'cTn');
      if (!cTn) return undefined;
      return {
        kind: name,
        common: parseCommonTimeNodeData(cTn),
        children: parseChildTimeNodes(cTn),
      };
    }
    case 'seq': {
      const cTn = findChild(node, 'cTn');
      if (!cTn) return undefined;
      const concurrent = parseBoolean(attr(node, 'concurrent'));
      const seq: SequenceTimeNode = {
        kind: 'seq',
        common: parseCommonTimeNodeData(cTn),
        ...(concurrent !== undefined ? { concurrent } : {}),
        children: parseChildTimeNodes(cTn),
      };
      return seq;
    }
    case 'set':
      return parseSetBehavior(node);
    case 'anim':
      return parseAnimateBehavior(node);
    case 'animClr':
      return parseAnimateColorBehavior(node);
    case 'animEffect':
      return parseAnimateEffectBehavior(node);
    case 'animMotion':
      return parseAnimateMotionBehavior(node);
    case 'animRot':
      return parseAnimateRotationBehavior(node);
    case 'animScale':
      return parseAnimateScaleBehavior(node);
    case 'cmd':
      return parseCommandBehavior(node);
    case 'audio':
    case 'video':
      return parseMediaNode(name, node);
    default:
      return undefined;
  }
}

/**
 * §19.7.29, ST_TLParaBuildType's wire values ("p" for by-paragraph) don't match our more readable
 * `ParagraphBuildType` names 1:1, so this maps between them. `"cust"` (a custom build fully
 * described by the real timing tree rather than this implicit shorthand) has no equivalent here
 * and is treated as absent.
 */
const PARAGRAPH_BUILD_TYPE_BY_ATTR: Readonly<Record<string, ParagraphBuildType>> = {
  whole: 'whole',
  allAtOnce: 'allAtOnce',
  p: 'byParagraph',
};

function parseParagraphBuild(node: XmlNode): ParagraphBuild | undefined {
  const shapeId = parseIntAttr(attr(node, 'spid'));
  if (shapeId === undefined) return undefined;
  const buildAttr = attr(node, 'build');
  const buildType = (buildAttr && PARAGRAPH_BUILD_TYPE_BY_ATTR[buildAttr]) || 'whole';
  const buildLevel = parseIntAttr(attr(node, 'bldLvl'));
  const animateBackground = parseBoolean(attr(node, 'animBg'));
  const autoUpdateAnimBg = parseBoolean(attr(node, 'autoUpdateAnimBg'));
  const reverse = parseBoolean(attr(node, 'rev'));

  return {
    kind: 'paragraph',
    shapeId,
    buildType,
    ...(buildLevel !== undefined ? { buildLevel } : {}),
    ...(animateBackground !== undefined ? { animateBackground } : {}),
    ...(autoUpdateAnimBg !== undefined ? { autoUpdateAnimBg } : {}),
    ...(reverse !== undefined ? { reverse } : {}),
  };
}

const GRAPHIC_BUILD_KIND: Readonly<Record<string, GraphicBuild['kind']>> = {
  bldDgm: 'diagram',
  bldChart: 'chart',
  bldGraphic: 'graphic',
};

function parseGraphicBuild(node: XmlNode): GraphicBuild | undefined {
  const kind = GRAPHIC_BUILD_KIND[localName(node) ?? ''];
  const shapeId = parseIntAttr(attr(node, 'spid'));
  if (!kind || shapeId === undefined) return undefined;
  const buildType = attr(node, 'build');
  return { kind, shapeId, ...(buildType ? { buildType } : {}) };
}

function parseBuildListEntry(node: XmlNode): BuildListEntry | undefined {
  return localName(node) === 'bldP' ? parseParagraphBuild(node) : parseGraphicBuild(node);
}

/** Parses a slide's p:timing (§19.3.1.48): its timing tree (p:tnLst) and build list (p:bldLst). */
export function parseSlideTiming(timingNode: XmlNode | undefined): SlideTiming | undefined {
  if (!timingNode) return undefined;

  const tnLst = findChild(timingNode, 'tnLst');
  const timeNodeTree = tnLst
    ? children(tnLst)
        .map(parseTimeNode)
        .find((node) => node !== undefined)
    : undefined;

  const bldLst = findChild(timingNode, 'bldLst');
  const buildList = bldLst
    ? children(bldLst)
        .map(parseBuildListEntry)
        .filter((entry): entry is BuildListEntry => entry !== undefined)
    : undefined;

  if (!timeNodeTree && !buildList) return undefined;
  return {
    ...(timeNodeTree ? { timeNodeTree } : {}),
    ...(buildList && buildList.length > 0 ? { buildList } : {}),
  };
}
