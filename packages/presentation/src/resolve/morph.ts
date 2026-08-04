import type { GroupShape, ShapeTreeNode, Slide } from '../presentationml/index.js';

/**
 * A leaf shape-tree node — everything a `ShapeTreeNode` can be except `GroupShape`. A `GroupShape`
 * is never itself a Morph match target (see `resolveMorphMatch`'s own doc comment for why); it's
 * only ever a container this module recurses through on the way to the leaves inside it.
 */
export type MorphLeafShape = Exclude<ShapeTreeNode, GroupShape>;

/** One shape successfully matched between the outgoing and incoming slide of a Morph transition. */
export interface MorphShapeMatch {
  readonly outgoing: MorphLeafShape;
  readonly incoming: MorphLeafShape;
}

/**
 * The result of matching two slides' shape trees for a Morph transition (§19.3.1.49's
 * `p159:morph` — see `MorphTransitionEffect` in `../presentationml/transition.js`). `matched` pairs
 * should morph (interpolate position/size/rotation/fill) from the outgoing shape to the incoming
 * one; `disappearing` (outgoing shapes with no incoming counterpart) should fade out;
 * `appearing` (incoming shapes with no outgoing counterpart) should fade in.
 */
export interface SlideMorphMatch {
  readonly matched: readonly MorphShapeMatch[];
  readonly disappearing: readonly MorphLeafShape[];
  readonly appearing: readonly MorphLeafShape[];
}

/**
 * Flattens a shape tree to its leaves, recursing through (but never including) `GroupShape`
 * containers — a group's own name/id could coincidentally match one on the other slide, but
 * matching it as a unit wouldn't mean much: PowerPoint's own Morph matches the actual visible
 * objects a deck author sees and renames, not incidental grouping structure, and a group's box
 * isn't otherwise a meaningful animation target for the same reason `to-html5`'s own group
 * rendering doesn't give one a real position/size of its own (see that package's CLAUDE.md).
 * Depth-first, document order.
 */
function flattenLeafShapes(nodes: readonly ShapeTreeNode[]): MorphLeafShape[] {
  const result: MorphLeafShape[] = [];
  for (const node of nodes) {
    if (node.kind === 'group') {
      result.push(...flattenLeafShapes(node.children));
    } else {
      result.push(node);
    }
  }
  return result;
}

/**
 * Groups leaves by their own `name` (§19.3.1.12, `p:cNvPr/@name`) — PowerPoint's own documented
 * guidance to end users is that Morph recognizes an object across slides by matching name, so
 * that's this module's primary key too. A blank name (`""`, common for shapes a user never
 * renamed) is excluded outright: with no distinguishing information at all, "matching" same-named
 * blank shapes would just be matching by position/order among unrelated objects, not identity.
 */
function groupByName(leaves: readonly MorphLeafShape[]): Map<string, MorphLeafShape[]> {
  const groups = new Map<string, MorphLeafShape[]>();
  for (const leaf of leaves) {
    const name = leaf.nonVisual.name;
    if (name === '') continue;
    const existing = groups.get(name);
    if (existing) existing.push(leaf);
    else groups.set(name, [leaf]);
  }
  return groups;
}

/**
 * A leaf's own plain text content, for the name-match fallback below — `undefined` for anything
 * without a `textBody` (pictures, connectors, ...) or whose text is entirely blank, same treatment
 * as `groupByName`'s blank-name exclusion: no distinguishing information, so nothing to key on.
 * Concatenates every run/field across every paragraph with no separator — this is an identity key,
 * not a rendered string, so paragraph/line breaks don't need to round-trip.
 */
function extractPlainText(leaf: MorphLeafShape): string | undefined {
  if (leaf.kind !== 'shape' || !leaf.textBody) return undefined;
  const text = leaf.textBody.paragraphs
    .map((paragraph) =>
      paragraph.runs
        .map((run) => {
          if (run.kind === 'run') return run.text;
          if (run.kind === 'field') return run.cachedText;
          return '';
        })
        .join(''),
    )
    .join('');
  return text === '' ? undefined : text;
}

/**
 * Groups leaves by their own plain text content (see `extractPlainText`) — the fallback key
 * `resolveMorphMatch` uses for whatever a name match left unpaired. PowerPoint's own name-based
 * guidance (see `groupByName`) assumes the same shape persisted across the duplicate-slide-then-
 * tweak authoring flow; a text box authored independently on each slide (typed fresh rather than
 * copied) never gets a shared name that way, even though it's visibly "the same" object to a
 * viewer and PowerPoint's own Morph does match it. Text content is the next-best identity signal
 * available on this package's graph for that case.
 */
function groupByText(leaves: readonly MorphLeafShape[]): Map<string, MorphLeafShape[]> {
  const groups = new Map<string, MorphLeafShape[]>();
  for (const leaf of leaves) {
    const text = extractPlainText(leaf);
    if (text === undefined) continue;
    const existing = groups.get(text);
    if (existing) existing.push(leaf);
    else groups.set(text, [leaf]);
  }
  return groups;
}

/**
 * Pairs up same-key candidates from the two slides (a "key" being whatever the caller grouped
 * by — name or, failing that, text content; see `resolveMorphMatch`). `id` (§19.3.1.12,
 * `p:cNvPr/@id`) is unique within a slide's own shape tree, so an id match within a same-key group
 * is stronger evidence of "literally the same shape" than the key match alone — real decks only
 * get same-key duplicates from copy/pasting a shape without renaming the copy, so this
 * disambiguates that case whenever the *same* one (by id) still exists on both sides. Anything
 * still unpaired after that (both the id itself changed, or there were more duplicates than id
 * matches could resolve) pairs positionally, in document order, as a last-resort tiebreak —
 * PowerPoint's own duplicate disambiguation is undocumented, so this is a best-effort stand-in,
 * the same tier as this package's other approximations for underspecified behaviour.
 */
function pairMatchingCandidates(
  outgoingCandidates: readonly MorphLeafShape[],
  incomingCandidates: readonly MorphLeafShape[],
): readonly MorphShapeMatch[] {
  const remainingIncoming = [...incomingCandidates];
  const pairs: MorphShapeMatch[] = [];
  const pairedOutgoing = new Set<MorphLeafShape>();

  for (const outgoing of outgoingCandidates) {
    const index = remainingIncoming.findIndex(
      (incoming) => incoming.nonVisual.id === outgoing.nonVisual.id,
    );
    if (index === -1) continue;
    pairs.push({ outgoing, incoming: remainingIncoming[index]! });
    remainingIncoming.splice(index, 1);
    pairedOutgoing.add(outgoing);
  }

  const remainingOutgoing = outgoingCandidates.filter((outgoing) => !pairedOutgoing.has(outgoing));
  const count = Math.min(remainingOutgoing.length, remainingIncoming.length);
  for (let i = 0; i < count; i++) {
    pairs.push({ outgoing: remainingOutgoing[i]!, incoming: remainingIncoming[i]! });
  }

  return pairs;
}

/**
 * Matches shapes between the outgoing and incoming slide of a Morph transition (§19.3.1.49,
 * `p159:morph`) — the diff step Morph fundamentally needs and every other `TransitionEffect`
 * doesn't: it isn't one canned animation applied to a single slide, it's a correspondence between
 * *two* slides' worth of shapes, interpolated for whatever matches and crossfaded for whatever
 * doesn't. This lives here, not in a renderer, for the same reason every other file in `resolve/`
 * does (see this package's "resolution logic lives with the model" design decision): which shape
 * on slide N+1 corresponds to which shape on slide N is a fact about the presentation, not a
 * rendering choice, and any future renderer needs the identical answer.
 *
 * Deliberately **not** spec-exact — there is no spec for this at all, since Morph is a PowerPoint
 * feature, not part of OOXML's own §19.3.1.49 schema (see `MorphTransitionEffect`'s doc comment in
 * `../presentationml/transition.js`). PowerPoint's own matching algorithm is undocumented; this
 * follows its one publicly-documented rule (match by shape name) first, plus `id`-based and
 * positional tiebreaks for same-named duplicates (see `groupByName`/`pairMatchingCandidates`), as
 * its primary pass. Whatever's still unmatched after that falls back to a second pass keyed on
 * plain text content instead (`groupByText`) — a text box authored fresh on each slide (typed
 * independently rather than produced by duplicating an existing shape) never shares a name across
 * the two slides even though it's visibly the same object, and PowerPoint's own Morph does match
 * that case. Both passes are best-effort approximations, the same tier as this package's other
 * stand-ins for underspecified behaviour. Only leaf shapes are matchable — see
 * `flattenLeafShapes`'s own doc comment for why a `GroupShape` itself never is — and matching is
 * global across the whole flattened tree, ignoring which group (if any) a shape sits in on either
 * side, since a real deck's Morph slide often moves a shape in or out of a group between the two
 * states.
 *
 * This function only decides *correspondence* — it returns the matched `ShapeTreeNode` pairs
 * themselves, not their resolved boxes/fills/colours (that's `coordinate.ts`/`color.ts`'s job, per
 * slide, same as any other renderer-facing resolution here) and it performs no reporting of its
 * own. A caller (`@pptx2html/to-html5`, not yet wired up) needs to decide, from `disappearing`/
 * `appearing`'s size relative to the total shape count, whether this slide's match is confident
 * enough to actually play as a morph or whether it should report a degraded/failed match through
 * its own `UnsupportedFeatureCollector` instead — this package has no dependency on that mechanism
 * (it lives downstream, in `to-html5`) and so can't report anything itself.
 */
export function resolveMorphMatch(outgoing: Slide, incoming: Slide): SlideMorphMatch {
  const outgoingLeaves = flattenLeafShapes(outgoing.commonSlideData.shapeTree);
  const incomingLeaves = flattenLeafShapes(incoming.commonSlideData.shapeTree);

  const incomingByName = groupByName(incomingLeaves);
  const matched: MorphShapeMatch[] = [];
  for (const [name, outgoingCandidates] of groupByName(outgoingLeaves)) {
    const incomingCandidates = incomingByName.get(name);
    if (incomingCandidates) {
      matched.push(...pairMatchingCandidates(outgoingCandidates, incomingCandidates));
    }
  }

  const matchedOutgoingByName = new Set(matched.map((pair) => pair.outgoing));
  const matchedIncomingByName = new Set(matched.map((pair) => pair.incoming));
  const unmatchedOutgoing = outgoingLeaves.filter((leaf) => !matchedOutgoingByName.has(leaf));
  const unmatchedIncoming = incomingLeaves.filter((leaf) => !matchedIncomingByName.has(leaf));

  const incomingByText = groupByText(unmatchedIncoming);
  for (const [text, outgoingCandidates] of groupByText(unmatchedOutgoing)) {
    const incomingCandidates = incomingByText.get(text);
    if (incomingCandidates) {
      matched.push(...pairMatchingCandidates(outgoingCandidates, incomingCandidates));
    }
  }

  const matchedOutgoing = new Set(matched.map((pair) => pair.outgoing));
  const matchedIncoming = new Set(matched.map((pair) => pair.incoming));

  return {
    matched,
    disappearing: outgoingLeaves.filter((leaf) => !matchedOutgoing.has(leaf)),
    appearing: incomingLeaves.filter((leaf) => !matchedIncoming.has(leaf)),
  };
}
