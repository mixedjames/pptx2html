import type {
  Paragraph,
  Placeholder,
  RunProperties,
  TextBody,
  TextRunElement,
} from '@pptx2html/presentation';

import { resolveFillColor } from './color.js';
import type { RenderContext } from './render-context.js';
import { resolveEffectiveRunProperties, resolveTypeface } from './text-style.js';

/** Applies a run's resolved character formatting (see `text-style.ts`) as inline CSS. */
function applyRunStyle(el: HTMLElement, properties: RunProperties, context: RenderContext): void {
  const typeface = resolveTypeface(properties.typeface, context.layout?.master.theme.fontScheme);
  if (typeface) el.style.fontFamily = `"${typeface}"`;
  if (properties.fontSize !== undefined) el.style.fontSize = `${properties.fontSize / 100}pt`;
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
): Node {
  if (run.kind === 'break') return doc.createElement('br');

  const el = doc.createElement('span');
  el.className = 'pptx-run';
  el.textContent = textOf(run);
  const properties = resolveEffectiveRunProperties(run, paragraph, textBody, placeholder, context);
  applyRunStyle(el, properties, context);
  return el;
}

function renderParagraph(
  doc: Document,
  paragraph: Paragraph,
  textBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
): HTMLElement {
  const p = doc.createElement('p');
  p.className = 'pptx-paragraph';
  if (paragraph.runs.length === 0) {
    // An empty paragraph is still a blank line.
    p.appendChild(doc.createElement('br'));
  } else {
    for (const run of paragraph.runs) {
      p.appendChild(renderRun(doc, run, paragraph, textBody, placeholder, context));
    }
  }
  return p;
}

export function renderTextBody(
  doc: Document,
  textBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
): HTMLElement {
  const container = doc.createElement('div');
  container.className = 'pptx-text-body';
  for (const paragraph of textBody.paragraphs) {
    container.appendChild(renderParagraph(doc, paragraph, textBody, placeholder, context));
  }
  return container;
}
