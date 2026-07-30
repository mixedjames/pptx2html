import type {
  Color,
  ColorScheme,
  Fill,
  FontCollection,
  FontScheme,
  FormatScheme,
  Line,
  Theme,
} from '@pptx2html/presentation';

import { parseChildColor } from './drawingml/color.js';
import { parseFill, type MediaResolver } from './drawingml/fill.js';
import { parseLine } from './drawingml/line.js';
import type { XmlNode } from './xml/parse.js';
import { parseXml } from './xml/parse.js';
import { attr, children, findChild, findRoot, localName } from './xml/query.js';

const COLOR_SCHEME_SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

const FALLBACK_COLOR: Color = { type: 'srgb', value: '000000' };

function parseColorScheme(node: XmlNode | undefined): ColorScheme {
  const name = (node && attr(node, 'name')) || '';
  const slots: Record<string, Color> = {};
  for (const slot of COLOR_SCHEME_SLOTS) {
    const slotNode = node ? findChild(node, slot) : undefined;
    slots[slot] = (slotNode && parseChildColor(slotNode)) ?? FALLBACK_COLOR;
  }
  return { name, ...slots } as ColorScheme;
}

function parseFontCollection(node: XmlNode | undefined): FontCollection {
  const latinNode = node ? findChild(node, 'latin') : undefined;
  const eaNode = node ? findChild(node, 'ea') : undefined;
  const csNode = node ? findChild(node, 'cs') : undefined;
  const eastAsian = eaNode ? attr(eaNode, 'typeface') : undefined;
  const complexScript = csNode ? attr(csNode, 'typeface') : undefined;

  return {
    latin: (latinNode && attr(latinNode, 'typeface')) || '',
    ...(eastAsian ? { eastAsian } : {}),
    ...(complexScript ? { complexScript } : {}),
  };
}

function parseFontScheme(node: XmlNode | undefined): FontScheme {
  return {
    name: (node && attr(node, 'name')) || '',
    majorFont: parseFontCollection(node && findChild(node, 'majorFont')),
    minorFont: parseFontCollection(node && findChild(node, 'minorFont')),
  };
}

/** A theme's fillStyleLst/lnStyleLst entries never reference media (§20.1.4.1.19) — blipFill is
 * legal there in principle but vanishingly rare, so a no-op resolver is used rather than plumbing
 * the theme part's own relationships through just for this. */
const NO_MEDIA: MediaResolver = () => undefined;

function parseFillStyleList(node: XmlNode | undefined): readonly Fill[] {
  if (!node) return [];
  const fills: Fill[] = [];
  for (const child of children(node)) {
    const fill = parseFill(child, NO_MEDIA);
    if (fill) fills.push(fill);
  }
  return fills;
}

function parseLineStyleList(node: XmlNode | undefined): readonly Line[] {
  if (!node) return [];
  return children(node)
    .filter((child) => localName(child) === 'ln')
    .map((child) => parseLine(child, NO_MEDIA));
}

function parseFormatScheme(node: XmlNode | undefined): FormatScheme {
  return {
    name: (node && attr(node, 'name')) || '',
    fillStyles: parseFillStyleList(node && findChild(node, 'fillStyleLst')),
    lineStyles: parseLineStyleList(node && findChild(node, 'lnStyleLst')),
  };
}

/** Parses a theme part, e.g. ppt/theme/theme1.xml (§14.2.7, a:theme). */
export function parseTheme(xmlText: string): Theme {
  const root = findRoot(parseXml(xmlText), 'theme');
  const themeElements = root && findChild(root, 'themeElements');

  return {
    name: (root && attr(root, 'name')) || '',
    colorScheme: parseColorScheme(themeElements && findChild(themeElements, 'clrScheme')),
    fontScheme: parseFontScheme(themeElements && findChild(themeElements, 'fontScheme')),
    formatScheme: parseFormatScheme(themeElements && findChild(themeElements, 'fmtScheme')),
  };
}

/** A structurally valid but empty Theme, used when a slide master has no theme relationship. */
export function emptyTheme(): Theme {
  return parseTheme('<a:theme/>');
}
