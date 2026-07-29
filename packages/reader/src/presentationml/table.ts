import type { Table, TableCell, TableColumn, TableRow } from '@pptx2html/presentation';

import { parseChildFill, type MediaResolver } from '../drawingml/fill.js';
import { parseTextBody } from '../drawingml/text.js';
import { parseEmu, parseIntAttr } from '../drawingml/units.js';
import type { XmlNode } from '../xml/parse.js';
import { attr, findAllChildren, findChild } from '../xml/query.js';

function parseTableColumn(node: XmlNode): TableColumn {
  return { width: parseEmu(attr(node, 'w')) ?? 0 };
}

function parseTableCell(node: XmlNode, resolveMedia: MediaResolver): TableCell {
  const txBody = findChild(node, 'txBody');
  const textBody = (txBody && parseTextBody(txBody, resolveMedia)) ?? { paragraphs: [] };
  const rowSpan = parseIntAttr(attr(node, 'rowSpan'));
  const colSpan = parseIntAttr(attr(node, 'gridSpan'));
  const merged = attr(node, 'hMerge') === '1' || attr(node, 'vMerge') === '1' ? true : undefined;
  const tcPr = findChild(node, 'tcPr');
  const fill = tcPr ? parseChildFill(tcPr, resolveMedia) : undefined;

  return {
    textBody,
    ...(rowSpan !== undefined ? { rowSpan } : {}),
    ...(colSpan !== undefined ? { colSpan } : {}),
    ...(merged ? { merged } : {}),
    ...(fill ? { fill } : {}),
  };
}

function parseTableRow(node: XmlNode, resolveMedia: MediaResolver): TableRow {
  return {
    height: parseEmu(attr(node, 'h')) ?? 0,
    cells: findAllChildren(node, 'tc').map((tc) => parseTableCell(tc, resolveMedia)),
  };
}

/** Parses a:tbl (§21.1.3.14). Table styles are unmodeled for the skeleton. */
export function parseTable(node: XmlNode, resolveMedia: MediaResolver): Table {
  const grid = findChild(node, 'tblGrid');
  const columns = grid ? findAllChildren(grid, 'gridCol').map(parseTableColumn) : [];
  const rows = findAllChildren(node, 'tr').map((tr) => parseTableRow(tr, resolveMedia));
  return { type: 'table', columns, rows };
}
