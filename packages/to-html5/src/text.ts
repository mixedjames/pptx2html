import type { Paragraph, TextBody, TextRunElement } from '@pptx2html/presentation';

function renderRun(doc: Document, run: TextRunElement): Node {
  switch (run.kind) {
    case 'run':
      return doc.createTextNode(run.text);
    case 'break':
      return doc.createElement('br');
    case 'field':
      return doc.createTextNode(run.cachedText);
  }
}

function renderParagraph(doc: Document, paragraph: Paragraph): HTMLElement {
  const p = doc.createElement('p');
  p.className = 'pptx-paragraph';
  if (paragraph.runs.length === 0) {
    // An empty paragraph is still a blank line.
    p.appendChild(doc.createElement('br'));
  } else {
    for (const run of paragraph.runs) p.appendChild(renderRun(doc, run));
  }
  return p;
}

export function renderTextBody(doc: Document, textBody: TextBody): HTMLElement {
  const container = doc.createElement('div');
  container.className = 'pptx-text-body';
  for (const paragraph of textBody.paragraphs) {
    container.appendChild(renderParagraph(doc, paragraph));
  }
  return container;
}
