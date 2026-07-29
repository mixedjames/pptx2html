// @vitest-environment happy-dom
import type { TextBody } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { renderTextBody } from './text.js';

describe('renderTextBody', () => {
  it('renders paragraphs as <p>, runs as text, breaks as <br>, fields as cached text', () => {
    const textBody: TextBody = {
      paragraphs: [
        {
          runs: [
            { kind: 'run', text: 'Hello ' },
            { kind: 'break' },
            { kind: 'run', text: 'world' },
            { kind: 'field', fieldType: 'slidenum', cachedText: '1' },
          ],
        },
        { runs: [] },
      ],
    };

    const el = renderTextBody(document, textBody);
    const paragraphs = el.querySelectorAll('p');
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]?.querySelectorAll('br')).toHaveLength(1);
    expect(paragraphs[0]?.textContent).toBe('Hello world1');
    expect(paragraphs[1]?.querySelector('br')).not.toBeNull();
  });
});
