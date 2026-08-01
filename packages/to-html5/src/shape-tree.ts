import type {
  ColorScheme,
  ConnectionShape,
  CoordinateMap,
  Emu,
  Fill,
  FormatScheme,
  GraphicFrame,
  GroupShape,
  Line,
  Picture,
  Placeholder,
  Shape,
  ShapeProperties,
  ShapeStyle,
  ShapeTreeNode,
  SlideLayout,
  Transform2D,
} from '@pptx2html/presentation';
import {
  composeGroupMap,
  computeBox,
  resolveInheritedTransform,
  resolveStyleFill,
  resolveStyleLine,
} from '@pptx2html/presentation';

import { applyFill, applyLine, applySvgFill, applySvgLine } from './fill.js';
import type { RenderContext } from './render-context.js';
import { nativeBorderRadius, presetShapePath } from './shape-geometry.js';
import { renderTable } from './table.js';
import { renderTextBody } from './text.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * Renders a preset geometry's outline (see `shape-geometry.ts`) as an `<svg>` overlay stretched
 * exactly over the shape's box — `preserveAspectRatio="none"` deliberately distorts the fixed
 * `0 0 100 100` path non-uniformly onto whatever aspect ratio the shape's actual box has, the same
 * "everything is a percentage of the box" approach `positionElement` uses for position/size.
 * `overflow: visible` keeps a thick stroke from being clipped exactly at the box edge (SVG clips
 * to the viewBox by default, unlike a CSS border which is allowed to sit astride its box edge).
 */
function renderShapeOutline(
  doc: Document,
  path: string,
  fill: Fill | undefined,
  line: Line | undefined,
  scheme: ColorScheme | undefined,
  slideWidth: Emu,
): SVGSVGElement {
  const svg = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
  svg.setAttribute('viewBox', '0 0 100 100');
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.style.position = 'absolute';
  svg.style.inset = '0';
  svg.style.width = '100%';
  svg.style.height = '100%';
  svg.style.overflow = 'visible';

  const pathEl = doc.createElementNS(SVG_NS, 'path') as SVGPathElement;
  pathEl.setAttribute('d', path);
  applySvgFill(pathEl, fill, scheme);
  applySvgLine(pathEl, line, scheme, slideWidth);
  svg.appendChild(pathEl);
  return svg;
}

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
  // PowerPoint sizes a shape's outline (see fill.ts's applyLine) within its bounding box rather
  // than growing the box by the border's width — CSS's default content-box would otherwise make
  // a bordered shape render larger than its width/height say.
  element.style.boxSizing = 'border-box';
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

/**
 * A shape/picture/connector's own `spPr` fill if it has one, otherwise its `p:style/fillRef`
 * resolved against the theme's format-scheme style matrix (§20.1.4.2.10, see `style-matrix.ts`) —
 * PowerPoint's Shape Styles gallery writes shapes with a bare style reference and no explicit
 * `spPr` fill/line at all, so without this fallback such a shape renders with no fill whatsoever.
 * A whole-value fallback, not a field-level merge — the same "first defined wins outright"
 * simplification `background.ts`'s `resolveEffectiveBackground` already uses.
 */
function effectiveFill(
  properties: ShapeProperties,
  style: ShapeStyle | undefined,
  formatScheme: FormatScheme | undefined,
): Fill | undefined {
  return properties.fill ?? resolveStyleFill(style?.fillRef, formatScheme);
}

/** `effectiveFill`'s `p:style/lnRef` equivalent. */
function effectiveLine(
  properties: ShapeProperties,
  style: ShapeStyle | undefined,
  formatScheme: FormatScheme | undefined,
): Line | undefined {
  return properties.line ?? resolveStyleLine(style?.lineRef, formatScheme);
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
  const scheme = context.layout?.master.theme.colorScheme;
  const formatScheme = context.layout?.master.theme.formatScheme;
  const fill = effectiveFill(shape.properties, shape.style, formatScheme);
  const line = effectiveLine(shape.properties, shape.style, formatScheme);
  const geometry = shape.properties.geometry;
  const outlinePath = geometry && presetShapePath(geometry);
  if (outlinePath) {
    el.appendChild(
      renderShapeOutline(doc, outlinePath, fill, line, scheme, context.slideSize.width),
    );
  } else {
    const radius = geometry && nativeBorderRadius(geometry);
    if (radius) el.style.borderRadius = radius;
    applyFill(el, fill, scheme);
    applyLine(el, line, scheme, context.slideSize.width);
  }
  if (shape.textBody) {
    el.appendChild(renderTextBody(doc, shape.textBody, shape.nonVisual.placeholder, context));
  }
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
  const scheme = context.layout?.master.theme.colorScheme;
  const formatScheme = context.layout?.master.theme.formatScheme;
  const fill = effectiveFill(picture.properties, picture.style, formatScheme);
  const line = effectiveLine(picture.properties, picture.style, formatScheme);
  // Unlike renderShape, a non-rect/roundRect/ellipse preset (e.g. a triangle-cropped picture) is
  // not yet handled here — clipping an <img> to an arbitrary SVG path needs an <svg><image> +
  // <clipPath> overlay, not the plain border-radius this native subset gets for free; deliberately
  // deferred (rect/roundRect/ellipse covers the overwhelming majority of real picture crops).
  const radius = picture.properties.geometry && nativeBorderRadius(picture.properties.geometry);
  if (radius) el.style.borderRadius = radius;
  // Shows through any transparent pixels in the image itself (e.g. a transparent PNG over a
  // colored spPr fill) — same fill/line properties a shape carries, since Picture shares
  // ShapeProperties.
  applyFill(el, fill, scheme);
  applyLine(el, line, scheme, context.slideSize.width);
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
    el.appendChild(renderTable(doc, frame.graphic, context));
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
