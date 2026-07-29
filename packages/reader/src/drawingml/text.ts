import type {
  LineBreak,
  Paragraph,
  ParagraphProperties,
  RunProperties,
  TextAlignment,
  TextAnchor,
  TextBody,
  TextBodyProperties,
  TextField,
  TextRun,
  TextRunElement,
  TextWrap,
} from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, findAllChildren, findChild, localName, textOf } from '../xml/query.js';
import { parseChildFill, type MediaResolver } from './fill.js';
import { parseFontSize, parseIntAttr } from './units.js';

const ALIGNMENT_MAP: Record<string, TextAlignment> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  dist: 'distributed',
};

function isWrap(value: string | undefined): value is TextWrap {
  return value === 'none' || value === 'square';
}

function isAnchor(value: string | undefined): value is TextAnchor {
  return value === 't' || value === 'ctr' || value === 'b' || value === 'just';
}

function parseBooleanAttr(value: string | undefined): boolean | undefined {
  return value === undefined ? undefined : value === '1' || value === 'true';
}

function parseRunProperties(
  node: XmlNode | undefined,
  resolveMedia: MediaResolver,
): RunProperties | undefined {
  if (!node) return undefined;

  const bold = parseBooleanAttr(attr(node, 'b'));
  const italic = parseBooleanAttr(attr(node, 'i'));
  const underlineValue = attr(node, 'u');
  const underline = underlineValue !== undefined ? underlineValue !== 'none' : undefined;
  const strikeValue = attr(node, 'strike');
  const strikethrough = strikeValue !== undefined ? strikeValue !== 'noStrike' : undefined;
  const fontSize = parseFontSize(attr(node, 'sz'));
  const fill = parseChildFill(node, resolveMedia);
  const latin = findChild(node, 'latin');
  const typeface = latin ? attr(latin, 'typeface') : undefined;
  const language = attr(node, 'lang');

  const properties: RunProperties = {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underline !== undefined ? { underline } : {}),
    ...(strikethrough !== undefined ? { strikethrough } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
    ...(fill ? { fill } : {}),
    ...(typeface ? { typeface } : {}),
    ...(language ? { language } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

function parseTextRun(node: XmlNode, resolveMedia: MediaResolver): TextRun {
  const t = findChild(node, 't');
  const properties = parseRunProperties(findChild(node, 'rPr'), resolveMedia);
  return {
    kind: 'run',
    text: t ? textOf(t) : '',
    ...(properties ? { properties } : {}),
  };
}

function parseLineBreak(node: XmlNode, resolveMedia: MediaResolver): LineBreak {
  const properties = parseRunProperties(findChild(node, 'rPr'), resolveMedia);
  return { kind: 'break', ...(properties ? { properties } : {}) };
}

function parseTextField(node: XmlNode, resolveMedia: MediaResolver): TextField {
  const t = findChild(node, 't');
  const properties = parseRunProperties(findChild(node, 'rPr'), resolveMedia);
  return {
    kind: 'field',
    fieldType: attr(node, 'type') ?? '',
    cachedText: t ? textOf(t) : '',
    ...(properties ? { properties } : {}),
  };
}

function parseParagraphProperties(node: XmlNode | undefined): ParagraphProperties | undefined {
  if (!node) return undefined;
  const algn = attr(node, 'algn');
  const alignment = algn ? ALIGNMENT_MAP[algn] : undefined;
  const level = parseIntAttr(attr(node, 'lvl'));

  const properties: ParagraphProperties = {
    ...(alignment ? { alignment } : {}),
    ...(level !== undefined ? { level } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

/** Parses a:p (§21.1.2.2.6), preserving document order since runs/breaks/fields interleave. */
function parseParagraph(node: XmlNode, resolveMedia: MediaResolver): Paragraph {
  const properties = parseParagraphProperties(findChild(node, 'pPr'));
  const runs: TextRunElement[] = [];
  for (const child of children(node)) {
    switch (localName(child)) {
      case 'r':
        runs.push(parseTextRun(child, resolveMedia));
        break;
      case 'br':
        runs.push(parseLineBreak(child, resolveMedia));
        break;
      case 'fld':
        runs.push(parseTextField(child, resolveMedia));
        break;
      default:
        break;
    }
  }
  return { ...(properties ? { properties } : {}), runs };
}

function parseBodyProperties(node: XmlNode | undefined): TextBodyProperties | undefined {
  if (!node) return undefined;
  const wrapValue = attr(node, 'wrap');
  const wrap = isWrap(wrapValue) ? wrapValue : undefined;
  const anchorValue = attr(node, 'anchor');
  const anchor = isAnchor(anchorValue) ? anchorValue : undefined;

  const properties: TextBodyProperties = {
    ...(wrap ? { wrap } : {}),
    ...(anchor ? { anchor } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

/** Parses a txBody (§21.1.2.1.5). */
export function parseTextBody(
  node: XmlNode | undefined,
  resolveMedia: MediaResolver,
): TextBody | undefined {
  if (!node) return undefined;
  const properties = parseBodyProperties(findChild(node, 'bodyPr'));
  const paragraphs = findAllChildren(node, 'p').map((p) => parseParagraph(p, resolveMedia));
  return { ...(properties ? { properties } : {}), paragraphs };
}
