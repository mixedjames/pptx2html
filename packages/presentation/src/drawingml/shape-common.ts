import type { Fill } from './fill.js';
import type { Geometry, Transform2D } from './geometry.js';
import type { Line } from './line.js';

/** Non-visual identity shared by every drawing object (§20.1.2.2.8/2.2.20/2.2.29, cNvPr). */
export interface NonVisualDrawingProperties {
  readonly id: number;
  readonly name: string;
  readonly description?: string;
  readonly hidden?: boolean;
}

/** Visual properties shared by shapes, pictures and connectors (§20.1.2.2.35, spPr). */
export interface ShapeProperties {
  readonly transform?: Transform2D;
  readonly geometry?: Geometry;
  readonly fill?: Fill;
  readonly line?: Line;
}
