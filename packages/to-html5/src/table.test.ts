// @vitest-environment happy-dom
import type { Table } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
import type { RenderContext } from './render-context.js';
import { renderTable } from './table.js';

const CONTEXT: RenderContext = {
  slideSize: { width: 1, height: 1 },
  layout: undefined,
  defaultTextStyle: undefined,
};

describe('renderTable', () => {
  it('renders columns/rows and skips merged cells in favor of the spanning cell', () => {
    const table: Table = {
      type: 'table',
      columns: [{ width: 914400 }, { width: 914400 }],
      rows: [
        {
          height: 457200,
          cells: [
            { textBody: { paragraphs: [{ runs: [{ kind: 'run', text: 'A' }] }] }, colSpan: 2 },
            { textBody: { paragraphs: [] }, merged: true },
          ],
        },
        {
          height: 457200,
          cells: [
            { textBody: { paragraphs: [{ runs: [{ kind: 'run', text: 'B' }] }] } },
            { textBody: { paragraphs: [{ runs: [{ kind: 'run', text: 'C' }] }] } },
          ],
        },
      ],
    };

    const el = renderTable(document, table, CONTEXT);
    expect(el.style.width).toBe('100%');
    expect(el.style.tableLayout).toBe('fixed');

    const cols = el.querySelectorAll('col');
    expect(cols).toHaveLength(2);
    // Equal-width columns (914400 each of a 1828800 total) split the table 50/50.
    expect(cols[0]?.style.width).toBe('50%');
    expect(cols[1]?.style.width).toBe('50%');

    const rows = el.querySelectorAll('tr');
    expect(rows).toHaveLength(2);
    expect(rows[0]?.style.height).toBe('50%');
    expect(rows[0]?.querySelectorAll('td')).toHaveLength(1);
    expect((rows[0]?.querySelector('td') as HTMLTableCellElement).colSpan).toBe(2);
    expect(rows[1]?.querySelectorAll('td')).toHaveLength(2);
    expect(rows[1]?.textContent).toBe('BC');
  });

  it('reports unrendered cell fills once, counting merged cells out, with the graphicFrame shape ref', () => {
    const reportUnsupported = vi.fn();
    const table: Table = {
      type: 'table',
      columns: [{ width: 1 }],
      rows: [
        {
          height: 1,
          cells: [
            {
              textBody: { paragraphs: [] },
              fill: { type: 'solid', color: { type: 'srgb', value: 'FF0000' } },
            },
          ],
        },
        {
          height: 1,
          cells: [
            {
              textBody: { paragraphs: [] },
              fill: { type: 'solid', color: { type: 'srgb', value: '00FF00' } },
              merged: true,
            },
          ],
        },
      ],
    };

    renderTable(document, table, { ...CONTEXT, reportUnsupported }, { id: 1, name: 'Table 1' });

    expect(reportUnsupported).toHaveBeenCalledWith(
      'table-cell-fill-unmodeled',
      expect.stringContaining('1'),
      { id: 1, name: 'Table 1' },
    );
    expect(reportUnsupported).toHaveBeenCalledTimes(1);
  });

  it('does not report when no cell has a fill', () => {
    const reportUnsupported = vi.fn();
    const table: Table = {
      type: 'table',
      columns: [{ width: 1 }],
      rows: [{ height: 1, cells: [{ textBody: { paragraphs: [] } }] }],
    };

    renderTable(document, table, { ...CONTEXT, reportUnsupported });

    expect(reportUnsupported).not.toHaveBeenCalled();
  });
});
