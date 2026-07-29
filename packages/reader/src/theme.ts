import type {
  Color,
  ColorScheme,
  FontCollection,
  FontScheme,
  FormatScheme,
  Theme,
} from '@pptx2html/presentation';

import { parseChildColor } from './drawingml/color.js';
import type { XmlNode } from './xml/parse.js';
import { parseXml } from './xml/parse.js';
import { attr, findChild, findRoot } from './xml/query.js';

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

function parseFormatScheme(node: XmlNode | undefined): FormatScheme {
  return { name: (node && attr(node, 'name')) || '' };
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
