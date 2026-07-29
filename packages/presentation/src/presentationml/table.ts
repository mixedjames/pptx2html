import type { Fill, TextBody } from '../drawingml/index.js';
import type { Emu } from '../drawingml/index.js';

/** §21.1.3.16, a:tc. */
export interface TableCell {
  readonly textBody: TextBody;
  readonly rowSpan?: number;
  readonly colSpan?: number;
  /** True for a cell covered by another cell's row/col span, mirroring hMerge/vMerge. */
  readonly merged?: boolean;
  readonly fill?: Fill;
}

/** §21.1.3.13, a:gridCol. */
export interface TableColumn {
  readonly width: Emu;
}

/** §21.1.3.17, a:tr. */
export interface TableRow {
  readonly height: Emu;
  readonly cells: readonly TableCell[];
}

/** A table graphic frame's content (§21.1.3.14, a:tbl). Table styles are unmodeled for the skeleton. */
export interface Table {
  readonly type: 'table';
  readonly columns: readonly TableColumn[];
  readonly rows: readonly TableRow[];
}
