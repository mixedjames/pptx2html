import type {
  AutoNumberScheme,
  Bullet,
  Color,
  Emu,
  LineBreak,
  Paragraph,
  ParagraphProperties,
  Percentage,
  RunProperties,
  TextAlignment,
  TextAnchor,
  TextBody,
  TextBodyProperties,
  TextField,
  TextListStyle,
  TextListStyleLevel,
  TextRun,
  TextRunElement,
  TextWrap,
} from '@pptx2html/presentation';

import type { XmlNode } from '../xml/parse.js';
import { attr, children, findAllChildren, findChild, localName, textOf } from '../xml/query.js';
import { parseChildColor } from './color.js';
import { parseChildFill, type MediaResolver } from './fill.js';
import { parseEmu, parseFontSize, parseIntAttr, parsePercentage } from './units.js';

const ALIGNMENT_MAP: Record<string, TextAlignment> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
  dist: 'distributed',
};

const AUTO_NUMBER_SCHEMES: ReadonlySet<string> = new Set<AutoNumberScheme>([
  'arabicPeriod',
  'arabicParenR',
  'alphaLcPeriod',
  'alphaUcPeriod',
  'alphaLcParenR',
  'alphaUcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'romanLcParenR',
  'romanUcParenR',
]);

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

/** Parses a:buChar/a:buAutoNum's shared glyph overrides (§21.1.2.4.4/4.6/4.9, buFont/buClr/buSzPct). */
function parseBulletStyle(node: XmlNode): {
  font?: string;
  color?: Color;
  sizePercent?: Percentage;
} {
  const buFont = findChild(node, 'buFont');
  const font = buFont ? attr(buFont, 'typeface') : undefined;
  const buClr = findChild(node, 'buClr');
  const color = buClr ? parseChildColor(buClr) : undefined;
  const buSzPct = findChild(node, 'buSzPct');
  const sizePercent = buSzPct ? parsePercentage(attr(buSzPct, 'val')) : undefined;

  return {
    ...(font ? { font } : {}),
    ...(color ? { color } : {}),
    ...(sizePercent !== undefined ? { sizePercent } : {}),
  };
}

/**
 * Parses a pPr/lvlNpPr's bullet (§21.1.2.4): `buNone` suppresses one that would otherwise be
 * inherited, `buChar`/`buAutoNum` are mutually exclusive per the schema. `buSzPts` (a point-size
 * bullet override, as opposed to `buSzPct`'s percentage) is unmodeled for the skeleton — `buSzPct`
 * is what real decks overwhelmingly use.
 */
function parseBullet(node: XmlNode): Bullet | undefined {
  if (findChild(node, 'buNone')) return { type: 'none' };

  const buChar = findChild(node, 'buChar');
  if (buChar) {
    const char = attr(buChar, 'char');
    return char ? { type: 'char', char, ...parseBulletStyle(node) } : undefined;
  }

  const buAutoNum = findChild(node, 'buAutoNum');
  if (buAutoNum) {
    const typeValue = attr(buAutoNum, 'type');
    const scheme: AutoNumberScheme =
      typeValue && AUTO_NUMBER_SCHEMES.has(typeValue)
        ? (typeValue as AutoNumberScheme)
        : 'arabicPeriod';
    const startAt = parseIntAttr(attr(buAutoNum, 'startAt'));
    return {
      type: 'autoNum',
      scheme,
      ...(startAt !== undefined ? { startAt } : {}),
      ...parseBulletStyle(node),
    };
  }

  return undefined;
}

/** Parses the paragraph-properties fields shared by a:pPr and a:lvlNpPr (§21.1.2.2.7/2.4.12). */
function parseSharedParagraphProperties(node: XmlNode): {
  alignment?: TextAlignment;
  bullet?: Bullet;
  marginLeft?: Emu;
  indent?: Emu;
} {
  const algn = attr(node, 'algn');
  const alignment = algn ? ALIGNMENT_MAP[algn] : undefined;
  const bullet = parseBullet(node);
  const marginLeft = parseEmu(attr(node, 'marL'));
  const indent = parseEmu(attr(node, 'indent'));

  return {
    ...(alignment ? { alignment } : {}),
    ...(bullet ? { bullet } : {}),
    ...(marginLeft !== undefined ? { marginLeft } : {}),
    ...(indent !== undefined ? { indent } : {}),
  };
}

function parseParagraphProperties(
  node: XmlNode | undefined,
  resolveMedia: MediaResolver,
): ParagraphProperties | undefined {
  if (!node) return undefined;
  const shared = parseSharedParagraphProperties(node);
  const level = parseIntAttr(attr(node, 'lvl'));
  const defaultRunProperties = parseRunProperties(findChild(node, 'defRPr'), resolveMedia);

  const properties: ParagraphProperties = {
    ...shared,
    ...(level !== undefined ? { level } : {}),
    ...(defaultRunProperties ? { defaultRunProperties } : {}),
  };
  return Object.keys(properties).length > 0 ? properties : undefined;
}

/** Parses a:p (§21.1.2.2.6), preserving document order since runs/breaks/fields interleave. */
function parseParagraph(node: XmlNode, resolveMedia: MediaResolver): Paragraph {
  const properties = parseParagraphProperties(findChild(node, 'pPr'), resolveMedia);
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

const LEVEL_TAGS = [
  'lvl1pPr',
  'lvl2pPr',
  'lvl3pPr',
  'lvl4pPr',
  'lvl5pPr',
  'lvl6pPr',
  'lvl7pPr',
  'lvl8pPr',
  'lvl9pPr',
] as const;

function parseTextListStyleLevel(
  levelNode: XmlNode,
  resolveMedia: MediaResolver,
): TextListStyleLevel | undefined {
  const shared = parseSharedParagraphProperties(levelNode);
  const runProperties = parseRunProperties(findChild(levelNode, 'defRPr'), resolveMedia);

  const level: TextListStyleLevel = {
    ...shared,
    ...(runProperties ? { runProperties } : {}),
  };
  return Object.keys(level).length > 0 ? level : undefined;
}

/**
 * Parses a per-level list style (§21.1.2.4.12, a:lstStyle, or the structurally identical
 * p:titleStyle/p:bodyStyle/p:otherStyle/p:defaultTextStyle) — each level's `algn`/bullet/`marL`/
 * `indent`/`defRPr`, not the other paragraph properties a level can also carry (unmodeled for the
 * skeleton).
 */
export function parseTextListStyle(
  node: XmlNode | undefined,
  resolveMedia: MediaResolver,
): TextListStyle | undefined {
  if (!node) return undefined;
  const levels = LEVEL_TAGS.map((tag) => {
    const levelNode = findChild(node, tag);
    return levelNode ? parseTextListStyleLevel(levelNode, resolveMedia) : undefined;
  });
  return levels.some((level) => level !== undefined) ? { levels } : undefined;
}

/** Parses a txBody (§21.1.2.1.5). */
export function parseTextBody(
  node: XmlNode | undefined,
  resolveMedia: MediaResolver,
): TextBody | undefined {
  if (!node) return undefined;
  const properties = parseBodyProperties(findChild(node, 'bodyPr'));
  const listStyle = parseTextListStyle(findChild(node, 'lstStyle'), resolveMedia);
  const paragraphs = findAllChildren(node, 'p').map((p) => parseParagraph(p, resolveMedia));
  return { ...(properties ? { properties } : {}), ...(listStyle ? { listStyle } : {}), paragraphs };
}
