import type {
  FontScheme,
  Paragraph,
  Placeholder,
  PlaceholderType,
  RunProperties,
  TextBody,
  TextListStyle,
  TextRunElement,
} from '@pptx2html/presentation';

import { findPlaceholderMatch } from './placeholder.js';
import type { RenderContext } from './render-context.js';

function levelOf(style: TextListStyle | undefined, level: number): RunProperties | undefined {
  return style?.levels[level];
}

/**
 * Which of the master's title/body/other text styles (§19.3.1.53, p:txStyles) a placeholder
 * falls back to: title/ctrTitle use the title style, every other placeholder type behaves like
 * body text, and a non-placeholder shape (plain text box) uses the catch-all "other" style.
 */
function masterStyleCategory(
  type: PlaceholderType | undefined,
): 'titleStyle' | 'bodyStyle' | 'otherStyle' {
  if (type === 'title' || type === 'ctrTitle') return 'titleStyle';
  if (type === undefined) return 'otherStyle';
  return 'bodyStyle';
}

/** Layers each source over the previous, later sources winning per-field; missing fields fall through. */
function mergeRunProperties(...sources: readonly (RunProperties | undefined)[]): RunProperties {
  const merged: Record<string, unknown> = {};
  for (const source of sources) {
    if (!source) continue;
    for (const [key, value] of Object.entries(source)) {
      if (value !== undefined) merged[key] = value;
    }
  }
  return merged as RunProperties;
}

/**
 * Resolves a run's effective character formatting by walking the OOXML text-property inheritance
 * chain (§21.1.2, simplified — the spec's exact algorithm is under-specified at the edges; this
 * follows the order real-world renderers converge on), lowest to highest priority:
 *
 * 1. `context.defaultTextStyle` — the presentation's own default (`p:defaultTextStyle`).
 * 2. The slide master's title/body/other style for this placeholder's category.
 * 3. The master's own matching placeholder shape's list style, if any (mirrors
 *    `resolveInheritedTransform`'s layout->master walk, one level further out).
 * 4. The layout's own matching placeholder shape's list style, if any.
 * 5. This shape's own list style (`TextBody.listStyle`).
 * 6. This paragraph's own default run properties (`pPr`'s defRPr).
 * 7. The run's own properties (`rPr`).
 *
 * Each step only supplies the per-field defaults for the paragraph's own outline level
 * (`paragraph.properties.level`, 0-based) — a level a source doesn't define contributes nothing.
 * Theme font-scheme resolution (`+mj-lt` etc.) is deliberately not done here — see
 * `resolveTypeface`, since it only concerns the `typeface` field and needs the theme, not a list
 * style.
 */
export function resolveEffectiveRunProperties(
  run: TextRunElement,
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: RenderContext,
): RunProperties {
  const level = paragraph.properties?.level ?? 0;
  const layout = context.layout;
  const master = layout?.master;

  const masterCategoryLevel = levelOf(
    master?.textStyles?.[masterStyleCategory(placeholder?.type)],
    level,
  );

  const masterPlaceholderMatch =
    placeholder && master
      ? findPlaceholderMatch(placeholder, master.commonSlideData.shapeTree)
      : undefined;
  const masterPlaceholderLevel = levelOf(
    masterPlaceholderMatch?.kind === 'shape'
      ? masterPlaceholderMatch.textBody?.listStyle
      : undefined,
    level,
  );

  const layoutPlaceholderMatch =
    placeholder && layout
      ? findPlaceholderMatch(placeholder, layout.commonSlideData.shapeTree)
      : undefined;
  const layoutPlaceholderLevel = levelOf(
    layoutPlaceholderMatch?.kind === 'shape'
      ? layoutPlaceholderMatch.textBody?.listStyle
      : undefined,
    level,
  );

  return mergeRunProperties(
    levelOf(context.defaultTextStyle, level),
    masterCategoryLevel,
    masterPlaceholderLevel,
    layoutPlaceholderLevel,
    levelOf(shapeTextBody.listStyle, level),
    paragraph.properties?.defaultRunProperties,
    run.properties,
  );
}

const THEME_FONT_TOKENS: Record<
  string,
  { major: boolean; script: 'latin' | 'eastAsian' | 'complexScript' }
> = {
  '+mj-lt': { major: true, script: 'latin' },
  '+mn-lt': { major: false, script: 'latin' },
  '+mj-ea': { major: true, script: 'eastAsian' },
  '+mn-ea': { major: false, script: 'eastAsian' },
  '+mj-cs': { major: true, script: 'complexScript' },
  '+mn-cs': { major: false, script: 'complexScript' },
};

/**
 * Resolves a theme font reference (§20.1.4.1.24, e.g. `"+mj-lt"` for "major latin") against the
 * theme's font scheme. Any other value passes through unchanged as a literal font name — most
 * runs specify a real typeface directly rather than referencing the theme.
 */
export function resolveTypeface(
  typeface: string | undefined,
  fontScheme: FontScheme | undefined,
): string | undefined {
  if (!typeface) return undefined;
  const token = THEME_FONT_TOKENS[typeface];
  if (!token || !fontScheme) return typeface;
  const collection = token.major ? fontScheme.majorFont : fontScheme.minorFont;
  return collection[token.script] ?? collection.latin;
}
