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
  /**
   * Default run formatting for this paragraph (a:pPr's defRPr child) — falls back for any run
   * in the paragraph that doesn't specify a given field itself. Distinct from `TextListStyle`,
   * which supplies the *paragraph's own* per-level default before this one is layered on top;
   * see `to-html5`'s `text-style.ts` for the full run-property inheritance chain.
   */
  readonly defaultRunProperties?: RunProperties;
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

/**
 * One outline level's defaults within a `TextListStyle` (§21.1.2.4.12, a:lvl1pPr..lvl9pPr —
 * structurally a full paragraph-properties element, but only `algn` and `defRPr` are modeled
 * here; other paragraph-level properties a level can also carry, like indent/bullet/numbering,
 * are unmodeled for the skeleton).
 */
export interface TextListStyleLevel {
  readonly alignment?: TextAlignment;
  readonly runProperties?: RunProperties;
}

/**
 * Per-outline-level defaults (§21.1.2.4.12, a:lstStyle's lvl1pPr..lvl9pPr). Indexed 0-based, the
 * same as `ParagraphProperties.level`: `levels[0]` is level 1's defaults, etc. A shape's txBody
 * (`TextBody.listStyle`), a slide master's title/body/other styles (`SlideMaster.textStyles`),
 * and the presentation's own default (`Presentation.defaultTextStyle`) are all this same shape.
 */
export interface TextListStyle {
  readonly levels: readonly (TextListStyleLevel | undefined)[];
}

/** A shape's or table cell's text content (§21.1.2.1.5, txBody). */
export interface TextBody {
  readonly properties?: TextBodyProperties;
  readonly listStyle?: TextListStyle;
  readonly paragraphs: readonly Paragraph[];
}
