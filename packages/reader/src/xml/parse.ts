import { XMLParser } from 'fast-xml-parser';

/**
 * A node in the `preserveOrder` tree fast-xml-parser produces: either a text leaf
 * (`{ '#text': string }`) or an element with exactly one non-attribute key — its
 * (possibly namespace-prefixed) tag name — holding its children, plus an optional
 * `:@` key holding its raw attributes (e.g. `{ '@_id': '2', '@_r:embed': 'rId3' }`).
 */
export type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

/**
 * Parses an XML part's text into its root-level nodes. `preserveOrder` is required:
 * fast-xml-parser's default mode regroups children by tag name, which would destroy the
 * run/break/field interleaving inside a paragraph and the shape/pic/grpSp/graphicFrame/cxnSp
 * interleaving inside spTree that defines z-order.
 */
export function parseXml(xml: string): readonly XmlNode[] {
  return parser.parse(xml) as XmlNode[];
}
