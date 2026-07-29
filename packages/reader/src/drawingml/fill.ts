import type { Fill, GradientStop, MediaPart } from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, findChild, localName } from '../xml/query.js';
import { parseChildColor } from './color.js';
import { parseAngle, parsePercentage } from './units.js';

/** Resolves an `r:embed`/`r:link` relationship id (scoped to the current part) to its media. */
export type MediaResolver = (relationshipId: string) => MediaPart | undefined;

const FILL_NAMES: ReadonlySet<string> = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'pattFill',
  'blipFill',
]);

function parseGradientStops(node: XmlNode): readonly GradientStop[] {
  const gsLst = findChild(node, 'gsLst');
  if (!gsLst) return [];
  const stops: GradientStop[] = [];
  for (const gs of children(gsLst)) {
    if (localName(gs) !== 'gs') continue;
    const position = parsePercentage(attr(gs, 'pos'));
    const color = parseChildColor(gs);
    if (position !== undefined && color) stops.push({ position, color });
  }
  return stops;
}

/** Parses a single fill element (a:noFill / solidFill / gradFill / pattFill / blipFill). */
export function parseFill(node: XmlNode, resolveMedia: MediaResolver): Fill | undefined {
  switch (localName(node)) {
    case 'noFill':
      return { type: 'none' };

    case 'solidFill': {
      const color = parseChildColor(node);
      return color ? { type: 'solid', color } : undefined;
    }

    case 'gradFill': {
      const lin = findChild(node, 'lin');
      const angle = lin ? parseAngle(attr(lin, 'ang')) : undefined;
      return {
        type: 'gradient',
        stops: parseGradientStops(node),
        ...(angle !== undefined ? { angle } : {}),
      };
    }

    case 'pattFill': {
      const preset = attr(node, 'prst');
      const fgClr = findChild(node, 'fgClr');
      const bgClr = findChild(node, 'bgClr');
      const foregroundColor = fgClr && parseChildColor(fgClr);
      const backgroundColor = bgClr && parseChildColor(bgClr);
      if (!preset || !foregroundColor || !backgroundColor) return undefined;
      return { type: 'pattern', preset, foregroundColor, backgroundColor };
    }

    case 'blipFill': {
      const blip = findChild(node, 'blip');
      const embed = blip && attr(blip, 'r:embed');
      const image = embed ? resolveMedia(embed) : undefined;
      if (!image) return undefined;

      const alphaModFix = blip && findChild(blip, 'alphaModFix');
      const opacity = alphaModFix ? parsePercentage(attr(alphaModFix, 'amt')) : undefined;
      const stretch = findChild(node, 'stretch') !== undefined || undefined;
      const tile = findChild(node, 'tile') !== undefined || undefined;

      return {
        type: 'blip',
        image,
        ...(opacity !== undefined ? { opacity } : {}),
        ...(stretch ? { stretch } : {}),
        ...(tile ? { tile } : {}),
      };
    }

    default:
      return undefined;
  }
}

/** Finds and parses the first fill child of a shape/run/cell properties node. */
export function parseChildFill(node: XmlNode, resolveMedia: MediaResolver): Fill | undefined {
  for (const child of children(node)) {
    if (FILL_NAMES.has(localName(child) ?? '')) {
      return parseFill(child, resolveMedia);
    }
  }
  return undefined;
}
