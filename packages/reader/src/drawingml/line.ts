import type {
  DashStyle,
  Line,
  LineCap,
  LineCompound,
  LineEnd,
  LineEndType,
} from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, findChild } from '../xml/query.js';
import { parseChildFill, type MediaResolver } from './fill.js';
import { parseEmu } from './units.js';

const CAP_MAP: Record<string, LineCap> = { rnd: 'round', sq: 'square', flat: 'flat' };

const COMPOUND_MAP: Record<string, LineCompound> = {
  sng: 'single',
  dbl: 'double',
  thickThin: 'thickThin',
  thinThick: 'thinThick',
  tri: 'triple',
};

const DASH_STYLES: ReadonlySet<string> = new Set<DashStyle>([
  'solid',
  'dot',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
]);

const LINE_END_TYPES: ReadonlySet<string> = new Set<LineEndType>([
  'none',
  'triangle',
  'stealth',
  'diamond',
  'oval',
  'arrow',
]);

function isEndSize(value: string | undefined): value is 'sm' | 'med' | 'lg' {
  return value === 'sm' || value === 'med' || value === 'lg';
}

function parseLineEnd(node: XmlNode | undefined): LineEnd | undefined {
  if (!node) return undefined;
  const type = attr(node, 'type');
  if (!type || !LINE_END_TYPES.has(type)) return undefined;
  const width = attr(node, 'w');
  const length = attr(node, 'len');
  return {
    type: type as LineEndType,
    ...(isEndSize(width) ? { width } : {}),
    ...(isEndSize(length) ? { length } : {}),
  };
}

/** Parses a:ln (§20.1.2.2.24). */
export function parseLine(node: XmlNode, resolveMedia: MediaResolver): Line {
  const width = parseEmu(attr(node, 'w'));
  const cap = CAP_MAP[attr(node, 'cap') ?? ''];
  const compound = COMPOUND_MAP[attr(node, 'cmpd') ?? ''];
  const fill = parseChildFill(node, resolveMedia);

  const prstDash = findChild(node, 'prstDash');
  const dashStyleValue = prstDash ? attr(prstDash, 'val') : undefined;
  const dashStyle =
    dashStyleValue && DASH_STYLES.has(dashStyleValue) ? (dashStyleValue as DashStyle) : undefined;

  const headEnd = parseLineEnd(findChild(node, 'headEnd'));
  const tailEnd = parseLineEnd(findChild(node, 'tailEnd'));

  return {
    ...(width !== undefined ? { width } : {}),
    ...(fill ? { fill } : {}),
    ...(cap ? { cap } : {}),
    ...(compound ? { compound } : {}),
    ...(dashStyle ? { dashStyle } : {}),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}
