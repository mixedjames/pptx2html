import { describe, expect, it } from 'vitest';

import type { MediaResolver } from '../drawingml/fill.js';
import { parseXml } from '../xml/parse.js';
import { parseTable } from './table.js';

const noMedia: MediaResolver = () => undefined;

describe('presentationml/table', () => {
  it('parses columns, rows, cell text and merge flags', () => {
    const [node] = parseXml(
      `<a:tbl xmlns:a="a">
        <a:tblGrid>
          <a:gridCol w="1000"/>
          <a:gridCol w="2000"/>
        </a:tblGrid>
        <a:tr h="500">
          <a:tc gridSpan="2"><a:txBody><a:p><a:r><a:t>Header</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr>
        <a:tr h="500">
          <a:tc><a:txBody><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc hMerge="1"><a:txBody><a:p/></a:txBody></a:tc>
        </a:tr>
      </a:tbl>`,
    );

    const table = parseTable(node!, noMedia);
    expect(table.columns).toEqual([{ width: 1000 }, { width: 2000 }]);
    expect(table.rows).toHaveLength(2);
    expect(table.rows[0]!.cells[0]).toMatchObject({ colSpan: 2 });
    expect(table.rows[0]!.cells[0]!.textBody.paragraphs[0]!.runs[0]).toEqual({
      kind: 'run',
      text: 'Header',
    });
    expect(table.rows[1]!.cells[1]).toMatchObject({ merged: true });
  });
});
