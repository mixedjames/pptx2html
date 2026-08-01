import type {
  Bullet,
  Emu,
  Paragraph,
  Placeholder,
  PlaceholderType,
  RunProperties,
  TextAlignment,
  TextBody,
  TextListStyle,
  TextListStyleLevel,
  TextRunElement,
} from '../drawingml/index.js';
import type { SlideLayout } from '../presentationml/index.js';
import type { FontScheme } from '../theme.js';
import { findPlaceholderMatch } from './placeholder.js';

/**
 * The two pieces of slide-level context the list-style inheritance chain needs — `layout` (for
 * the master/layout placeholder walk) and `defaultTextStyle` (the presentation's own
 * `p:defaultTextStyle`, the chain's bottom rung). A renderer's own per-slide context (e.g.
 * `@pptx2html/to-html5`'s `RenderContext`, which also carries a `slideSize` that has nothing to do
 * with text-style resolution) can simply extend this rather than duplicate its fields.
 */
export interface TextStyleContext {
  readonly layout?: SlideLayout;
  readonly defaultTextStyle?: TextListStyle;
}

function levelOf(style: TextListStyle | undefined, level: number): TextListStyleLevel | undefined {
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

/**
 * Collects this paragraph's outline level from every rung of the list-style inheritance chain
 * (§21.1.2, simplified — the spec's exact algorithm is under-specified at the edges; this follows
 * the order real-world renderers converge on), lowest to highest priority:
 *
 * 1. `context.defaultTextStyle` — the presentation's own default (`p:defaultTextStyle`).
 * 2. The slide master's title/body/other style for this placeholder's category.
 * 3. The master's own matching placeholder shape's list style, if any (mirrors
 *    `resolveInheritedTransform`'s layout->master walk, one level further out).
 * 4. The layout's own matching placeholder shape's list style, if any.
 * 5. This shape's own list style (`TextBody.listStyle`).
 *
 * Shared by `resolveEffectiveRunProperties` (character formatting) and
 * `resolveEffectiveAlignment` (paragraph alignment) — both need this same chain, just merged
 * differently (a field-by-field merge for run properties vs. "first defined wins" for the single
 * `alignment` scalar).
 */
function levelChain(
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): readonly (TextListStyleLevel | undefined)[] {
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

  return [
    levelOf(context.defaultTextStyle, level),
    masterCategoryLevel,
    masterPlaceholderLevel,
    layoutPlaceholderLevel,
    levelOf(shapeTextBody.listStyle, level),
  ];
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
 * Resolves a run's effective character formatting by walking `levelChain`, then layering this
 * paragraph's own default run properties (`pPr`'s defRPr) and finally the run's own properties
 * (`rPr`) on top. Each step only supplies the per-field defaults for the paragraph's own outline
 * level (`paragraph.properties.level`, 0-based) — a level a source doesn't define contributes
 * nothing. Theme font-scheme resolution (`+mj-lt` etc.) is deliberately not done here — see
 * `resolveTypeface`, since it only concerns the `typeface` field and needs the theme, not a list
 * style.
 */
export function resolveEffectiveRunProperties(
  run: TextRunElement,
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): RunProperties {
  const chain = levelChain(paragraph, shapeTextBody, placeholder, context);
  return mergeRunProperties(
    ...chain.map((entry) => entry?.runProperties),
    paragraph.properties?.defaultRunProperties,
    run.properties,
  );
}

/**
 * Resolves a single scalar field (as opposed to `RunProperties`' field-by-field merge) by taking
 * `own` if defined, otherwise the closest-defined value found walking `chain` low to high
 * priority. Used for `alignment`, `bullet`, `marginLeft` and `indent` — each is one paragraph-level
 * value, not a set of independent properties, so "first (highest-priority) defined wins" is the
 * right merge, not a field-by-field combination.
 */
function resolveScalar<T>(
  own: T | undefined,
  chain: readonly (TextListStyleLevel | undefined)[],
  pick: (level: TextListStyleLevel) => T | undefined,
): T | undefined {
  if (own !== undefined) return own;
  return chain.reduce<T | undefined>(
    (inherited, entry) => (entry ? (pick(entry) ?? inherited) : inherited),
    undefined,
  );
}

/**
 * Resolves a paragraph's effective alignment: its own explicit `alignment` if it has one,
 * otherwise the closest-defined one found walking `levelChain` (same chain
 * `resolveEffectiveRunProperties` uses, but "first defined wins" rather than merged field-by-field
 * — `alignment` is a single scalar, not a set of independent properties).
 */
export function resolveEffectiveAlignment(
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): TextAlignment | undefined {
  const chain = levelChain(paragraph, shapeTextBody, placeholder, context);
  return resolveScalar(paragraph.properties?.alignment, chain, (level) => level.alignment);
}

/**
 * Resolves a paragraph's effective bullet (§21.1.2.4): its own if set (including an explicit
 * `{ type: 'none' }` to suppress an inherited one), otherwise the closest-defined one walking
 * `levelChain`. Not merged field-by-field with a lower level's bullet — a `CharBullet` and an
 * `AutoNumberBullet` are different shapes entirely, so "closest whole bullet wins" is the only
 * sensible rule.
 */
export function resolveEffectiveBullet(
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): Bullet | undefined {
  const chain = levelChain(paragraph, shapeTextBody, placeholder, context);
  return resolveScalar(paragraph.properties?.bullet, chain, (level) => level.bullet);
}

/**
 * Resolves a paragraph's effective indentation (§21.1.2.2.7's `marL`/`indent`): `marginLeft` is
 * the whole paragraph's left margin, `indent` is the first line's offset relative to it
 * (typically negative, hanging a bullet/number ahead of the first line's own text). Each resolved
 * independently via `resolveScalar`, same chain as alignment/bullet.
 */
export function resolveEffectiveIndent(
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): { marginLeft: Emu | undefined; indent: Emu | undefined } {
  const chain = levelChain(paragraph, shapeTextBody, placeholder, context);
  return {
    marginLeft: resolveScalar(paragraph.properties?.marginLeft, chain, (level) => level.marginLeft),
    indent: resolveScalar(paragraph.properties?.indent, chain, (level) => level.indent),
  };
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
