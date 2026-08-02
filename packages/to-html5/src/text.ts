import type {
  Bullet,
  Color,
  Paragraph,
  Placeholder,
  RunProperties,
  ShapeStyle,
  TextAlignment,
  TextBody,
  TextRunElement,
} from '@pptx2html/presentation';
import {
  formatAutoNumber,
  NumberingState,
  resolveEffectiveAlignment,
  resolveEffectiveBullet,
  resolveEffectiveIndent,
  resolveEffectiveRunProperties,
  resolveTypeface,
} from '@pptx2html/presentation';

import { resolveColor, resolveFillColor } from './color.js';
import type { RenderContext } from './render-context.js';
import { emuToCqw, fontSizeToEmu } from './units.js';

/**
 * CSS has no "distributed" text-align keyword (OOXML's variant of justify that also stretches
 * the last line, not just the ones before it) — `text-align-last: justify` is the closest
 * approximation, applied alongside `text-align: justify` below.
 */
function applyAlignment(el: HTMLElement, alignment: TextAlignment | undefined): void {
  switch (alignment) {
    case undefined:
      return;
    case 'distributed':
      el.style.textAlign = 'justify';
      el.style.textAlignLast = 'justify';
      return;
    default:
      el.style.textAlign = alignment;
  }
}

/** Applies a run's resolved character formatting (see `text-style.ts`) as inline CSS. */
function applyRunStyle(el: HTMLElement, properties: RunProperties, context: RenderContext): void {
  const typeface = resolveTypeface(properties.typeface, context.layout?.master.theme.fontScheme);
  if (typeface) el.style.fontFamily = `"${typeface}"`;
  if (properties.fontSize !== undefined) {
    el.style.fontSize = emuToCqw(fontSizeToEmu(properties.fontSize), context.slideSize.width);
  }
  if (properties.bold) el.style.fontWeight = 'bold';
  if (properties.italic) el.style.fontStyle = 'italic';

  const decorations: string[] = [];
  if (properties.underline) decorations.push('underline');
  if (properties.strikethrough) decorations.push('line-through');
  if (decorations.length > 0) el.style.textDecoration = decorations.join(' ');

  if (properties.fill) {
    const color = resolveFillColor(properties.fill, context.layout?.master.theme.colorScheme);
    if (color) el.style.color = color;
  }
}

/**
 * PowerPoint's conventional per-level indent default (0.5in per level, 0.25in hang) applied only
 * when a bulleted paragraph resolves no `marginLeft`/`indent` of its own anywhere in the
 * inheritance chain — real decks almost always set these explicitly (usually via the master's
 * list style), this just keeps a bare bullet from rendering flush against its own text with no
 * hang at all.
 */
const DEFAULT_MARGIN_PER_LEVEL_EMU = 457200;
const DEFAULT_HANG_EMU = -228600;

function applyIndent(
  el: HTMLElement,
  marginLeft: number | undefined,
  indent: number | undefined,
  hasBullet: boolean,
  level: number,
  context: RenderContext,
): void {
  const resolvedMarginLeft =
    marginLeft ?? (hasBullet ? (level + 1) * DEFAULT_MARGIN_PER_LEVEL_EMU : undefined);
  const resolvedIndent = indent ?? (hasBullet ? DEFAULT_HANG_EMU : undefined);
  if (resolvedMarginLeft !== undefined) {
    el.style.paddingLeft = emuToCqw(resolvedMarginLeft, context.slideSize.width);
  }
  if (resolvedIndent !== undefined) {
    el.style.textIndent = emuToCqw(resolvedIndent, context.slideSize.width);
  }
}

/**
 * Renders a resolved bullet's glyph/label as a `<span class="pptx-bullet">`. A bullet's own
 * `font`/`color`/`sizePercent` (§21.1.2.4.4/4.6/4.9) win; anything unset falls back to the
 * paragraph's own "ambient" run formatting (an empty run resolved through the same inheritance
 * chain as any real run — see `resolveEffectiveRunProperties`), since an OOXML bullet otherwise
 * inherits the character formatting of the text it precedes.
 */
function renderBulletSpan(
  doc: Document,
  bullet: { readonly font?: string; readonly color?: Color; readonly sizePercent?: number },
  label: string,
  ambientRunProperties: RunProperties,
  context: RenderContext,
): HTMLElement {
  const span = doc.createElement('span');
  span.className = 'pptx-bullet';
  span.textContent = label;

  const typeface = resolveTypeface(
    bullet.font ?? ambientRunProperties.typeface,
    context.layout?.master.theme.fontScheme,
  );
  if (typeface) span.style.fontFamily = `"${typeface}"`;

  const scheme = context.layout?.master.theme.colorScheme;
  const color = bullet.color
    ? resolveColor(bullet.color, scheme)
    : ambientRunProperties.fill
      ? resolveFillColor(ambientRunProperties.fill, scheme)
      : undefined;
  if (color) span.style.color = color;

  if (ambientRunProperties.fontSize !== undefined) {
    const scale = bullet.sizePercent !== undefined ? bullet.sizePercent / 100000 : 1;
    span.style.fontSize = emuToCqw(
      fontSizeToEmu(ambientRunProperties.fontSize * scale),
      context.slideSize.width,
    );
  }

  return span;
}

function textOf(run: TextRunElement): string {
  switch (run.kind) {
    case 'run':
      return run.text;
    case 'field':
      return run.cachedText;
    case 'break':
      return '';
  }
}

function renderRun(
  doc: Document,
  run: TextRunElement,
  paragraph: Paragraph,
  textBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
  shapeStyle: ShapeStyle | undefined,
): Node {
  if (run.kind === 'break') return doc.createElement('br');

  const el = doc.createElement('span');
  el.className = 'pptx-run';
  el.textContent = textOf(run);
  const properties = resolveEffectiveRunProperties(
    run,
    paragraph,
    textBody,
    placeholder,
    context,
    shapeStyle,
  );
  applyRunStyle(el, properties, context);
  return el;
}

function renderParagraph(
  doc: Document,
  paragraph: Paragraph,
  textBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
  shapeStyle: ShapeStyle | undefined,
  bullet: Bullet | undefined,
  autoNumberLabel: string | undefined,
): HTMLElement {
  const p = doc.createElement('p');
  p.className = 'pptx-paragraph';
  applyAlignment(p, resolveEffectiveAlignment(paragraph, textBody, placeholder, context));

  const { marginLeft, indent } = resolveEffectiveIndent(paragraph, textBody, placeholder, context);
  const level = paragraph.properties?.level ?? 0;
  const hasGlyph = bullet !== undefined && bullet.type !== 'none';
  applyIndent(p, marginLeft, indent, hasGlyph, level, context);

  if (hasGlyph && (bullet.type === 'char' || bullet.type === 'autoNum')) {
    const label = bullet.type === 'char' ? bullet.char : (autoNumberLabel ?? '');
    const ambientRunProperties = resolveEffectiveRunProperties(
      { kind: 'run', text: '' },
      paragraph,
      textBody,
      placeholder,
      context,
      shapeStyle,
    );
    p.appendChild(renderBulletSpan(doc, bullet, label, ambientRunProperties, context));
    p.appendChild(doc.createTextNode(' '));
  }

  if (paragraph.runs.length === 0) {
    // An empty paragraph is still a blank line.
    p.appendChild(doc.createElement('br'));
  } else {
    for (const run of paragraph.runs) {
      p.appendChild(renderRun(doc, run, paragraph, textBody, placeholder, context, shapeStyle));
    }
  }
  return p;
}

export function renderTextBody(
  doc: Document,
  textBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
  shapeStyle?: ShapeStyle,
): HTMLElement {
  const container = doc.createElement('div');
  container.className = 'pptx-text-body';
  const numbering = new NumberingState();

  for (const paragraph of textBody.paragraphs) {
    const level = paragraph.properties?.level ?? 0;
    // PowerPoint doesn't show a bullet/number on a genuinely empty paragraph outside of edit
    // mode (a trailing blank line is common at the end of a list) — and a blank line shouldn't
    // consume a number or break a running numbered list either, so leave `numbering` untouched.
    const isEmpty = paragraph.runs.length === 0;
    const bullet = isEmpty
      ? undefined
      : resolveEffectiveBullet(paragraph, textBody, placeholder, context);

    let autoNumberLabel: string | undefined;
    if (bullet?.type === 'autoNum') {
      autoNumberLabel = formatAutoNumber(numbering.next(level, bullet), bullet.scheme);
    } else if (!isEmpty) {
      numbering.break(level);
    }

    container.appendChild(
      renderParagraph(
        doc,
        paragraph,
        textBody,
        placeholder,
        context,
        shapeStyle,
        bullet,
        autoNumberLabel,
      ),
    );
  }

  return container;
}
