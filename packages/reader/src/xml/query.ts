import type { XmlNode } from './parse.js';

const ATTRS_KEY = ':@';
const TEXT_KEY = '#text';
const ATTR_NAME_PREFIX = '@_';

function tagKey(node: XmlNode): string | undefined {
  return Object.keys(node).find((key) => key !== ATTRS_KEY);
}

export function isTextNode(node: XmlNode): boolean {
  return tagKey(node) === TEXT_KEY;
}

/** The node's tag local name (namespace prefix stripped), or undefined for a text node. */
export function localName(node: XmlNode): string | undefined {
  const key = tagKey(node);
  if (key === undefined || key === TEXT_KEY) return undefined;
  const colon = key.indexOf(':');
  return colon === -1 ? key : key.slice(colon + 1);
}

/**
 * The node's raw attributes, keyed exactly as written on the wire (including any
 * namespace prefix, e.g. `"r:embed"`) — namespace prefixes are not stripped here because
 * unlike elements, plain and prefixed attribute names commonly collide on local name alone
 * (e.g. `p:sldMasterId` carries both a plain `id` and an `r:id`).
 */
export function attrs(node: XmlNode): Record<string, string> {
  const raw = (node as Record<string, unknown>)[ATTRS_KEY] as Record<string, string> | undefined;
  if (!raw) return {};
  const result: Record<string, string> = {};
  for (const [key, value] of Object.entries(raw)) {
    result[key.slice(ATTR_NAME_PREFIX.length)] = value;
  }
  return result;
}

export function attr(node: XmlNode, name: string): string | undefined {
  return attrs(node)[name];
}

function rawChildren(node: XmlNode): readonly XmlNode[] {
  const key = tagKey(node);
  if (key === undefined) return [];
  const value = (node as Record<string, unknown>)[key];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

/**
 * The node's element children, in document order, with any `mc:AlternateContent` wrapper
 * transparently replaced by its `mc:Fallback` content (the schema-compatible baseline
 * PowerPoint always writes) — or its first `mc:Choice` if no fallback is present, rather
 * than silently dropping the content. We never evaluate `mc:Choice` requirement/namespace
 * negotiation.
 */
export function children(node: XmlNode): readonly XmlNode[] {
  const result: XmlNode[] = [];
  for (const child of rawChildren(node)) {
    if (isTextNode(child)) continue;
    if (localName(child) === 'AlternateContent') {
      const fallback = findChild(child, 'Fallback');
      const chosen = fallback ?? findChild(child, 'Choice');
      if (chosen) result.push(...children(chosen));
      continue;
    }
    result.push(child);
  }
  return result;
}

export function findChild(node: XmlNode, name: string): XmlNode | undefined {
  return children(node).find((child) => localName(child) === name);
}

export function findAllChildren(node: XmlNode, name: string): readonly XmlNode[] {
  return children(node).filter((child) => localName(child) === name);
}

/** Concatenates this node's own and nested `#text` content (e.g. an `a:t`'s value). */
export function textOf(node: XmlNode): string {
  let result = '';
  for (const child of rawChildren(node)) {
    if (isTextNode(child)) {
      result += (child as Record<string, unknown>)[TEXT_KEY] as string;
    } else if (localName(child) !== 'AlternateContent') {
      result += textOf(child);
    }
  }
  return result;
}

/** Finds the first top-level element named `name` among a parsed document's root nodes. */
export function findRoot(nodes: readonly XmlNode[], name: string): XmlNode | undefined {
  return nodes.find((node) => localName(node) === name);
}
