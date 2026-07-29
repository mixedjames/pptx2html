import type {
  ConnectionShape,
  GraphicFrame,
  GroupShape,
  Picture,
  Placeholder,
  Shape,
  ShapeTreeNode,
  SlideLayout,
  Transform2D,
} from '@pptx2html/presentation';
import { type CoordinateMap, composeGroupMap, computeBox } from './coordinate.js';
import { resolveInheritedTransform } from './placeholder.js';
import type { RenderContext } from './render-context.js';
import { renderTable } from './table.js';
import { renderTextBody } from './text.js';

/**
 * Positions an element as a percentage of the slide's own size rather than in absolute px, so
 * resizing the slide (see `slide.ts`'s `width: 100%; aspect-ratio`) scales every element with
 * it for free, with no JS/ResizeObserver involved.
 */
function positionElement(
  element: HTMLElement,
  map: CoordinateMap,
  transform: Transform2D,
  context: RenderContext,
): void {
  const box = computeBox(map, transform);
  element.style.position = 'absolute';
  element.style.left = `${(box.left / context.slideSize.width) * 100}%`;
  element.style.top = `${(box.top / context.slideSize.height) * 100}%`;
  element.style.width = `${(box.width / context.slideSize.width) * 100}%`;
  element.style.height = `${(box.height / context.slideSize.height) * 100}%`;

  const transforms: string[] = [];
  if (box.flipHorizontal) transforms.push('scaleX(-1)');
  if (box.flipVertical) transforms.push('scaleY(-1)');
  if (box.rotationDeg) transforms.push(`rotate(${box.rotationDeg}deg)`);
  if (transforms.length > 0) element.style.transform = transforms.join(' ');
}

/**
 * A shape/picture/connector's own transform if it has one, otherwise the transform inherited
 * from its matching placeholder in the slide's layout/master, if any (§19.3.1.36) — most
 * placeholder shapes (title, body, etc.) omit their own position/size entirely and rely on this.
 */
function effectiveTransform(
  ownTransform: Transform2D | undefined,
  placeholder: Placeholder | undefined,
  layout: SlideLayout | undefined,
): Transform2D | undefined {
  if (ownTransform) return ownTransform;
  if (placeholder && layout) return resolveInheritedTransform(placeholder, layout);
  return undefined;
}

function renderShape(
  doc: Document,
  shape: Shape,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  const el = doc.createElement('div');
  el.className = 'pptx-shape';
  const transform = effectiveTransform(
    shape.properties.transform,
    shape.nonVisual.placeholder,
    context.layout,
  );
  if (transform) positionElement(el, map, transform, context);
  if (shape.textBody) el.appendChild(renderTextBody(doc, shape.textBody));
  return el;
}

function renderPicture(
  doc: Document,
  picture: Picture,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  const el = doc.createElement('img');
  el.className = 'pptx-picture';
  const transform = effectiveTransform(
    picture.properties.transform,
    picture.nonVisual.placeholder,
    context.layout,
  );
  if (transform) positionElement(el, map, transform, context);
  const blob = new Blob([new Uint8Array(picture.image.data)], { type: picture.image.contentType });
  el.src = URL.createObjectURL(blob);
  return el;
}

function renderConnector(
  doc: Document,
  connector: ConnectionShape,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  const el = doc.createElement('div');
  el.className = 'pptx-connector';
  const transform = effectiveTransform(
    connector.properties.transform,
    connector.nonVisual.placeholder,
    context.layout,
  );
  if (transform) positionElement(el, map, transform, context);
  return el;
}

function renderGraphicFrame(
  doc: Document,
  frame: GraphicFrame,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  // GraphicFrame.transform is mandatory in the schema (unlike sp/pic/cxnSp's spPr/xfrm), so it
  // never needs placeholder-inherited positioning.
  const el = doc.createElement('div');
  el.className = 'pptx-graphic-frame';
  positionElement(el, map, frame.transform, context);

  if (frame.graphic.type === 'table') {
    el.appendChild(renderTable(doc, frame.graphic));
  } else {
    const placeholder = doc.createElement('div');
    placeholder.className = 'pptx-graphic-placeholder';
    placeholder.textContent = `[${frame.graphic.type}]`;
    el.appendChild(placeholder);
  }
  return el;
}

function renderGroup(
  doc: Document,
  group: GroupShape,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  // A group contributes no visual box of its own here — it's an anchor stretched to exactly
  // cover the slide (`inset: 0`, i.e. 100% x 100% of its own containing block, which — since
  // every ancestor group wrapper is stretched the same way — is always exactly the slide's own
  // size). That's what makes each descendant's percentage left/top/width/height, already
  // resolved to slide-root-relative fractions by the `CoordinateMap` chain, resolve correctly
  // against *this* wrapper's box however deep the nesting goes: percentages need a
  // well-defined containing block, which a bare `left:0; top:0` (auto width/height) wrapper
  // would not have.
  const el = doc.createElement('div');
  el.className = 'pptx-group';
  el.style.position = 'absolute';
  el.style.left = '0';
  el.style.top = '0';
  el.style.right = '0';
  el.style.bottom = '0';

  const childMap = composeGroupMap(map, group.transform);
  for (const child of group.children) {
    el.appendChild(renderShapeTreeNode(doc, child, childMap, context));
  }
  return el;
}

export function renderShapeTreeNode(
  doc: Document,
  node: ShapeTreeNode,
  map: CoordinateMap,
  context: RenderContext,
): HTMLElement {
  switch (node.kind) {
    case 'shape':
      return renderShape(doc, node, map, context);
    case 'picture':
      return renderPicture(doc, node, map, context);
    case 'connector':
      return renderConnector(doc, node, map, context);
    case 'graphicFrame':
      return renderGraphicFrame(doc, node, map, context);
    case 'group':
      return renderGroup(doc, node, map, context);
  }
}
