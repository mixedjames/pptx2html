import type { Fill } from './fill.js';
import type { FontSize } from './units.js';

export type TextAlignment = 'left' | 'center' | 'right' | 'justify' | 'distributed';

export type TextAnchor = 't' | 'ctr' | 'b' | 'just';

export type TextWrap = 'none' | 'square';

/** Run-level character formatting (§21.1.2.3.9, a:rPr and inherited defRPr/endParaRPr). */
export interface RunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  readonly fontSize?: FontSize;
  readonly fill?: Fill;
  readonly typeface?: string;
  readonly language?: string;
}

export interface TextRun {
  readonly kind: 'run';
  readonly text: string;
  readonly properties?: RunProperties;
}

export interface LineBreak {
  readonly kind: 'break';
  readonly properties?: RunProperties;
}

/** A dynamic field such as a slide number or date (§21.1.2.3.5, a:fld). */
export interface TextField {
  readonly kind: 'field';
  readonly fieldType: string;
  readonly cachedText: string;
  readonly properties?: RunProperties;
}

export type TextRunElement = TextRun | LineBreak | TextField;

/** Paragraph-level formatting (§21.1.2.2.7, a:pPr). Bullet/numbering is unmodeled for the skeleton. */
export interface ParagraphProperties {
  readonly alignment?: TextAlignment;
  /** Outline/indent level, 0-based. */
  readonly level?: number;
}

export interface Paragraph {
  readonly properties?: ParagraphProperties;
  readonly runs: readonly TextRunElement[];
}

/** Text body block properties (§21.1.2.1.1, a:bodyPr). Autofit is unmodeled for the skeleton. */
export interface TextBodyProperties {
  readonly wrap?: TextWrap;
  readonly anchor?: TextAnchor;
}

/** A shape's or table cell's text content (§21.1.2.1.5, txBody). */
export interface TextBody {
  readonly properties?: TextBodyProperties;
  readonly paragraphs: readonly Paragraph[];
}
