import type { Color, ColorTransform, SchemeColorName } from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, localName } from '../xml/query.js';
import { parseAngle, parsePercentage } from './units.js';

const SCHEME_COLOR_NAMES: ReadonlySet<string> = new Set<SchemeColorName>([
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
  'phClr',
  'bg1',
  'tx1',
  'bg2',
  'tx2',
]);

const TRANSFORM_NAMES = ['alpha', 'lumMod', 'lumOff', 'shade', 'tint', 'satMod', 'hueMod'] as const;

function parseColorTransforms(node: XmlNode): ColorTransform | undefined {
  const transforms: Partial<Record<(typeof TRANSFORM_NAMES)[number], number>> = {};
  for (const child of children(node)) {
    const name = localName(child);
    const match = TRANSFORM_NAMES.find((candidate) => candidate === name);
    if (!match) continue;
    const value = parsePercentage(attr(child, 'val'));
    if (value !== undefined) transforms[match] = value;
  }
  return Object.keys(transforms).length > 0 ? transforms : undefined;
}

/** Parses a single colour element (a:srgbClr / a:schemeClr / a:sysClr / a:prstClr / a:hslClr). */
export function parseColor(node: XmlNode): Color | undefined {
  const transforms = parseColorTransforms(node);
  const t = transforms ? { transforms } : {};

  switch (localName(node)) {
    case 'srgbClr': {
      const value = attr(node, 'val');
      return value ? { type: 'srgb', value, ...t } : undefined;
    }
    case 'schemeClr': {
      const value = attr(node, 'val');
      if (!value || !SCHEME_COLOR_NAMES.has(value)) return undefined;
      return { type: 'scheme', value: value as SchemeColorName, ...t };
    }
    case 'sysClr': {
      const value = attr(node, 'val');
      const lastColor = attr(node, 'lastClr');
      if (!value || !lastColor) return undefined;
      return { type: 'system', value, lastColor, ...t };
    }
    case 'prstClr': {
      const value = attr(node, 'val');
      return value ? { type: 'preset', value, ...t } : undefined;
    }
    case 'hslClr': {
      const hue = parseAngle(attr(node, 'hue'));
      const saturation = parsePercentage(attr(node, 'sat'));
      const luminance = parsePercentage(attr(node, 'lum'));
      if (hue === undefined || saturation === undefined || luminance === undefined) {
        return undefined;
      }
      return { type: 'hsl', hue, saturation, luminance, ...t };
    }
    default:
      return undefined;
  }
}

/** Finds and parses the first colour child of a node (e.g. a:solidFill, a theme colour slot). */
export function parseChildColor(node: XmlNode): Color | undefined {
  for (const child of children(node)) {
    const color = parseColor(child);
    if (color) return color;
  }
  return undefined;
}
