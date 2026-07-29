import type { Placeholder, ShapeTreeNode, SlideLayout, Transform2D } from '@pptx2html/presentation';

/**
 * A node's own transform, for the kinds that can carry a placeholder identity and can therefore
 * omit it (`GraphicFrame.transform` is mandatory in the schema, so it never needs resolving;
 * `GroupShape` never carries a placeholder at all — groups have no `nvPr`/`ph` in the schema).
 */
function ownTransform(node: ShapeTreeNode): Transform2D | undefined {
  switch (node.kind) {
    case 'shape':
    case 'picture':
    case 'connector':
      return node.properties.transform;
    case 'graphicFrame':
      return node.transform;
    case 'group':
      return undefined;
  }
}

/**
 * Finds the candidate whose placeholder identity best matches, per the informal rule OOXML
 * producers actually rely on: an exact type+index match wins, then any index match, then any
 * type match. Equivalence groups the spec allows (e.g. title vs. ctrTitle) are not modeled —
 * real decks reliably reuse the same type across slide/layout/master for a given placeholder.
 */
function findPlaceholderMatch(
  placeholder: Placeholder,
  candidates: readonly ShapeTreeNode[],
): ShapeTreeNode | undefined {
  const withPlaceholder = candidates
    .map((node) => ({ node, ph: node.nonVisual.placeholder }))
    .filter((entry): entry is { node: ShapeTreeNode; ph: Placeholder } => entry.ph !== undefined);

  return (
    withPlaceholder.find((c) => c.ph.index === placeholder.index && c.ph.type === placeholder.type)
      ?.node ??
    withPlaceholder.find((c) => c.ph.index === placeholder.index)?.node ??
    withPlaceholder.find((c) => c.ph.type === placeholder.type)?.node
  );
}

/**
 * Resolves the transform a placeholder shape should use when it has none of its own, by walking
 * the OOXML placeholder inheritance chain (§19.3.1.36): the matching placeholder shape in the
 * slide's layout, and if that one also has no transform of its own, the matching placeholder
 * shape in the layout's master. Returns undefined if no matching placeholder with a transform is
 * found anywhere in the chain.
 */
export function resolveInheritedTransform(
  placeholder: Placeholder,
  layout: SlideLayout,
): Transform2D | undefined {
  const layoutMatch = findPlaceholderMatch(placeholder, layout.commonSlideData.shapeTree);
  const layoutTransform = layoutMatch && ownTransform(layoutMatch);
  if (layoutTransform) return layoutTransform;

  const masterMatch = findPlaceholderMatch(placeholder, layout.master.commonSlideData.shapeTree);
  return masterMatch ? ownTransform(masterMatch) : undefined;
}
