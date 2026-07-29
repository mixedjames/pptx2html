import type {
  MediaPart,
  NonVisualDrawingProperties,
  ShapeProperties,
  TextBody,
  Transform2D,
} from '../drawingml/index.js';
import type { Table } from './table.js';

/** An autoshape or text box (§19.3.1.43, p:sp). */
export interface Shape {
  readonly kind: 'shape';
  readonly nonVisual: NonVisualDrawingProperties;
  readonly properties: ShapeProperties;
  readonly textBody?: TextBody;
}

/** §19.3.1.37, p:pic. */
export interface Picture {
  readonly kind: 'picture';
  readonly nonVisual: NonVisualDrawingProperties;
  readonly properties: ShapeProperties;
  readonly image: MediaPart;
}

/** Reference to a shape's connection site, used by connectors (§20.1.2.2.10, a:stCxn/a:endCxn). */
export interface ShapeConnection {
  readonly shapeId: number;
  readonly connectionSiteIndex: number;
}

/** A straight, bent or curved connector (§19.3.1.19, p:cxnSp). */
export interface ConnectionShape {
  readonly kind: 'connector';
  readonly nonVisual: NonVisualDrawingProperties;
  readonly properties: ShapeProperties;
  readonly startConnection?: ShapeConnection;
  readonly endConnection?: ShapeConnection;
}

/**
 * A hosted graphic other than a table — chart, SmartArt diagram or OLE object. These graphic
 * types are unmodeled for the skeleton and only their kind is preserved.
 */
export interface GraphicPlaceholder {
  readonly type: 'chart' | 'smartArt' | 'oleObject' | 'unknown';
}

/** A frame hosting a non-shape graphic such as a table, chart or diagram (§19.3.1.21, p:graphicFrame). */
export interface GraphicFrame {
  readonly kind: 'graphicFrame';
  readonly nonVisual: NonVisualDrawingProperties;
  readonly transform: Transform2D;
  readonly graphic: Table | GraphicPlaceholder;
}

/** §19.3.1.22/1.32, p:grpSp / p:spTree children. */
export type ShapeTreeNode = Shape | Picture | ConnectionShape | GraphicFrame | GroupShape;

/** A group of shapes positioned and transformed as a unit (§19.3.1.22, p:grpSp). */
export interface GroupShape {
  readonly kind: 'group';
  readonly nonVisual: NonVisualDrawingProperties;
  readonly transform: Transform2D;
  readonly children: readonly ShapeTreeNode[];
}
