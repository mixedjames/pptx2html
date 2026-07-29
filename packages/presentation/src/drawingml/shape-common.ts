import type { Fill } from './fill.js';
import type { Geometry, Transform2D } from './geometry.js';
import type { Line } from './line.js';

/** Placeholder kind (§19.7.9, ST_PlaceholderType). */
export type PlaceholderType =
  | 'title'
  | 'body'
  | 'ctrTitle'
  | 'subTitle'
  | 'dt'
  | 'ftr'
  | 'sldNum'
  | 'sldImg'
  | 'pic'
  | 'chart'
  | 'tbl'
  | 'clipArt'
  | 'dgm'
  | 'media'
  | 'obj';

/**
 * A shape's placeholder identity (§19.3.1.36, p:ph), present only on placeholder shapes. A
 * placeholder shape commonly omits its own `ShapeProperties.transform`/`Table` styling and
 * inherits them from the slide layout's (or, failing that, the slide master's) placeholder shape
 * with a matching `type`/`index` — resolving that inheritance is a consumer's job (e.g.
 * `@pptx2html/to-html5`), not this package's; this type only preserves the identity needed to do
 * the match.
 */
export interface Placeholder {
  readonly type: PlaceholderType;
  readonly index: number;
}

/** Non-visual identity shared by every drawing object (§20.1.2.2.8/2.2.20/2.2.29, cNvPr). */
export interface NonVisualDrawingProperties {
  readonly id: number;
  readonly name: string;
  readonly description?: string;
  readonly hidden?: boolean;
  readonly placeholder?: Placeholder;
}

/** Visual properties shared by shapes, pictures and connectors (§20.1.2.2.35, spPr). */
export interface ShapeProperties {
  readonly transform?: Transform2D;
  readonly geometry?: Geometry;
  readonly fill?: Fill;
  readonly line?: Line;
}
