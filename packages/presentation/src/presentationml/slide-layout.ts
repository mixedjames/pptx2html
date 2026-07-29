import type { CommonSlideData } from './common-slide-data.js';
import type { SlideMaster } from './slide-master.js';

/** The built-in layout presets defined by §19.7.11 (ST_SlideLayoutType). */
export type SlideLayoutType =
  | 'title'
  | 'tx'
  | 'twoColTx'
  | 'tbl'
  | 'txAndChart'
  | 'chartAndTx'
  | 'dgm'
  | 'chart'
  | 'txAndClipArt'
  | 'clipArtAndTx'
  | 'titleOnly'
  | 'blank'
  | 'txAndObj'
  | 'objAndTx'
  | 'objOnly'
  | 'obj'
  | 'txAndMedia'
  | 'mediaAndTx'
  | 'objOverTx'
  | 'txOverObj'
  | 'txAndTwoObj'
  | 'twoObjAndTx'
  | 'twoObjOverTx'
  | 'fourObj'
  | 'vertTx'
  | 'clipArtAndVertTx'
  | 'vertTitleAndTx'
  | 'vertTitleAndTxOverChart'
  | 'twoObj'
  | 'objAndTwoObj'
  | 'twoTxTwoObj'
  | 'secHead'
  | 'objTx'
  | 'picTx'
  | 'cust';

/** A slide layout part (§19.3.1.39, p:sldLayout): a reusable placeholder arrangement for a master. */
export interface SlideLayout {
  readonly commonSlideData: CommonSlideData;
  readonly master: SlideMaster;
  readonly type: SlideLayoutType;
  readonly matchingName?: string;
  readonly showMasterShapes?: boolean;
}
