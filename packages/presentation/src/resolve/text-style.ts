import type {
  Bullet,
  Emu,
  Paragraph,
  Placeholder,
  PlaceholderType,
  RunProperties,
  TextAlignment,
  TextAnchor,
  TextBody,
  TextListStyle,
  TextListStyleLevel,
  TextRunElement,
} from '../drawingml/index.js';
import type {
  FontReference,
  ShapeStyle,
  ShapeTreeNode,
  SlideLayout,
} from '../presentationml/index.js';
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
 * A shape's `p:style/fontRef` (§20.1.4.1.17) as a `TextListStyleLevel`-shaped `RunProperties`
 * source (see `levelChain`'s rung 5, below): its colour becomes a `SolidFill`, and its font
 * collection becomes the same `+mj-lt`/`+mn-lt` theme token `resolveTypeface` already knows how to
 * resolve (so no new font-resolution mechanism is needed here) — `'none'` contributes no typeface
 * fallback at all.
 */
function fontReferenceRunProperties(fontRef: FontReference | undefined): RunProperties | undefined {
  if (!fontRef) return undefined;
  const typeface =
    fontRef.collection === 'major'
      ? '+mj-lt'
      : fontRef.collection === 'minor'
        ? '+mn-lt'
        : undefined;
  return {
    fill: { type: 'solid', color: fontRef.color },
    ...(typeface ? { typeface } : {}),
  };
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
 * 5. This shape's own `p:style/fontRef` (§20.1.4.1.17), if any — see `fontReferenceRunProperties`.
 *    Placed *above* the master/placeholder chain (1-4): PowerPoint's Shape Styles gallery writes a
 *    `fontRef` as the shape's own directly-authored quick-style choice, which is more specific than
 *    — and visibly must outrank — the master's generic template default (`otherStyle` in
 *    particular almost always sets an explicit dark colour at level 0, which would otherwise always
 *    clobber a lighter `fontRef` colour chosen to contrast with the shape's own fill).
 * 6. This shape's own list style (`TextBody.listStyle`) — still outranks `fontRef`, since anything
 *    the author explicitly formatted on the shape's own text should win over a gallery default.
 *
 * Shared by `resolveEffectiveRunProperties` (character formatting) and
 * `resolveEffectiveAlignment` (paragraph alignment) — both need this same chain, just merged
 * differently (a field-by-field merge for run properties vs. "first defined wins" for the single
 * `alignment` scalar). Only `resolveEffectiveRunProperties` passes `shapeStyle` — `fontRef` has
 * nothing to say about alignment/bullet/indent, so the other resolvers simply never populate rung 5.
 */
function levelChain(
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
  shapeStyle?: ShapeStyle,
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

  const fontReferenceLevel: TextListStyleLevel | undefined = shapeStyle?.fontRef
    ? { runProperties: fontReferenceRunProperties(shapeStyle.fontRef) }
    : undefined;

  return [
    levelOf(context.defaultTextStyle, level),
    masterCategoryLevel,
    masterPlaceholderLevel,
    layoutPlaceholderLevel,
    fontReferenceLevel,
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
 * Resolves a run's effective character formatting by walking `levelChain` (which now includes
 * `shapeStyle`'s `fontRef`, if any, as its own rung — see that function's doc comment for exactly
 * where it sits), then layering this paragraph's own default run properties (`pPr`'s defRPr) and
 * finally the run's own properties (`rPr`) on top. Each step only supplies the per-field defaults
 * for the paragraph's own outline level (`paragraph.properties.level`, 0-based) — a level a source
 * doesn't define contributes nothing. Theme font-scheme resolution (`+mj-lt` etc.) is deliberately
 * not done here — see `resolveTypeface`, since it only concerns the `typeface` field and needs the
 * theme, not a list style.
 */
export function resolveEffectiveRunProperties(
  run: TextRunElement,
  paragraph: Paragraph,
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
  shapeStyle?: ShapeStyle,
): RunProperties {
  const chain = levelChain(paragraph, shapeTextBody, placeholder, context, shapeStyle);
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

function ownTextBody(node: ShapeTreeNode): TextBody | undefined {
  return node.kind === 'shape' ? node.textBody : undefined;
}

/**
 * Resolves a text body's effective vertical anchor (§21.1.2.1.1, a:bodyPr/@anchor): its own value
 * if set, otherwise the matching placeholder shape's own value in the layout, then the master —
 * the same placeholder-inheritance chain `resolveInheritedTransform` walks (§19.3.1.36), just for
 * `bodyPr` instead of `xfrm`. Defaults to `'t'` (top), the schema's own default, when nothing in
 * the chain sets one. `'just'` (anchor-justified, distributing multiple paragraphs to fill the
 * box) has no direct flexbox equivalent for a renderer built around a single block of paragraphs
 * and is left to whatever a consumer's own top-anchored fallback does — see `to-html5`'s own
 * mapping.
 */
export function resolveEffectiveAnchor(
  shapeTextBody: TextBody,
  placeholder: Placeholder | undefined,
  context: TextStyleContext,
): TextAnchor {
  if (shapeTextBody.properties?.anchor) return shapeTextBody.properties.anchor;

  const layout = context.layout;
  if (placeholder && layout) {
    const layoutMatch = findPlaceholderMatch(placeholder, layout.commonSlideData.shapeTree);
    const layoutAnchor = layoutMatch && ownTextBody(layoutMatch)?.properties?.anchor;
    if (layoutAnchor) return layoutAnchor;

    const masterMatch = findPlaceholderMatch(placeholder, layout.master.commonSlideData.shapeTree);
    const masterAnchor = masterMatch && ownTextBody(masterMatch)?.properties?.anchor;
    if (masterAnchor) return masterAnchor;
  }

  return 't';
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
