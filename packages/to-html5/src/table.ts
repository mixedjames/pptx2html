import type { Table } from '@pptx2html/presentation';
import type { RenderContext } from './render-context.js';
import { renderTextBody } from './text.js';

export function renderTable(doc: Document, table: Table, context: RenderContext): HTMLElement {
  const el = doc.createElement('table');
  el.className = 'pptx-table';
  // Fills its containing graphicFrame div (itself sized as a percentage of the slide, see
  // shape-tree.ts), so the table scales along with everything else. table-layout: fixed makes
  // the percentage <col> widths below actually control column sizing instead of being
  // overridden by content.
  el.style.width = '100%';
  el.style.height = '100%';
  el.style.tableLayout = 'fixed';

  const totalWidth = table.columns.reduce((sum, column) => sum + column.width, 0);
  const colgroup = doc.createElement('colgroup');
  for (const column of table.columns) {
    const col = doc.createElement('col');
    if (totalWidth > 0) col.style.width = `${(column.width / totalWidth) * 100}%`;
    colgroup.appendChild(col);
  }
  el.appendChild(colgroup);

  const totalHeight = table.rows.reduce((sum, row) => sum + row.height, 0);
  const tbody = doc.createElement('tbody');
  for (const row of table.rows) {
    const tr = doc.createElement('tr');
    if (totalHeight > 0) tr.style.height = `${(row.height / totalHeight) * 100}%`;
    for (const cell of row.cells) {
      // A merged cell is covered by another cell's rowSpan/colSpan — HTML tables need no
      // placeholder for it, the spanning <td> already accounts for the grid space.
      if (cell.merged) continue;

      const td = doc.createElement('td');
      if (cell.rowSpan !== undefined && cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
      if (cell.colSpan !== undefined && cell.colSpan > 1) td.colSpan = cell.colSpan;
      // Table cells never carry a placeholder identity of their own (no nvPr/ph in the schema
      // for a:tc), so cell text only falls back through the master's otherStyle/presentation
      // default, not any placeholder-specific list style.
      td.appendChild(renderTextBody(doc, cell.textBody, undefined, context));
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  el.appendChild(tbody);

  return el;
}
