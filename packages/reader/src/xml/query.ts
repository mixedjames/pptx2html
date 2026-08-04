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

/**
 * Like `findChild`, but for a child that might be `mc:AlternateContent`-wrapped: returns the
 * `mc:Choice` branch's own version of `name` *alongside* whatever `children()`/`findChild()`
 * would already give you (the `mc:Fallback` branch's version, or the plain unwrapped child if
 * there's no `mc:AlternateContent` at all) — `resolved` and `choice` are the same node when
 * there's a `Choice` but no `Fallback`, mirroring `children()`'s own "falls back to `mc:Choice`
 * when no `mc:Fallback` is present" behaviour. For a caller that wants to recognize specific
 * extension content in the `Choice` branch itself (e.g. `p159:morph`) and only fall back to the
 * schema-compatible baseline when it doesn't — `children()`/`findChild()` can never expose the
 * `Choice` branch at all (see their own doc comment), since they're built for callers that don't
 * care about the difference. We still never evaluate `mc:Choice`'s `Requires` attribute.
 */
export function findAlternateContentChild(
  node: XmlNode,
  name: string,
): { readonly choice?: XmlNode; readonly resolved?: XmlNode } {
  for (const child of rawChildren(node)) {
    if (isTextNode(child)) continue;
    if (localName(child) === 'AlternateContent') {
      const choiceWrapper = findChild(child, 'Choice');
      const fallbackWrapper = findChild(child, 'Fallback');
      const resolvedWrapper = fallbackWrapper ?? choiceWrapper;
      const choice = choiceWrapper ? findChild(choiceWrapper, name) : undefined;
      const resolved = resolvedWrapper ? findChild(resolvedWrapper, name) : undefined;
      if (choice ?? resolved) {
        return { ...(choice ? { choice } : {}), ...(resolved ? { resolved } : {}) };
      }
      continue;
    }
    if (localName(child) === name) return { resolved: child };
  }
  return {};
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
