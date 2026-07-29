import { parseXml } from '../xml/parse.js';
import { attr, children, findRoot } from '../xml/query.js';

export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** Normalized (no leading slash) internal part name, or the raw URI for External targets. */
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
}

export interface Relationships {
  get(id: string): Relationship | undefined;
  findByType(type: string): readonly Relationship[];
}

const EMPTY_RELATIONSHIPS: Relationships = {
  get: () => undefined,
  findByType: () => [],
};

function directoryOf(partName: string): string {
  const slash = partName.lastIndexOf('/');
  return slash === -1 ? '' : partName.slice(0, slash);
}

/** Resolves a relationship Target against its referencing part's directory (OPC §9.3). */
function resolveTarget(basePartDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);

  const resolved = basePartDir === '' ? [] : basePartDir.split('/');
  for (const segment of target.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') resolved.pop();
    else resolved.push(segment);
  }
  return resolved.join('/');
}

/** The `_rels/<basename>.rels` part name holding a given part's relationships. */
export function relationshipsPartNameFor(partName: string): string {
  const dir = directoryOf(partName);
  const base = dir === '' ? partName : partName.slice(dir.length + 1);
  return `${dir === '' ? '' : `${dir}/`}_rels/${base}.rels`;
}

export function parseRelationships(
  ownerPartName: string,
  xmlText: string | undefined,
): Relationships {
  if (xmlText === undefined) return EMPTY_RELATIONSHIPS;

  const root = findRoot(parseXml(xmlText), 'Relationships');
  if (!root) return EMPTY_RELATIONSHIPS;

  const baseDir = directoryOf(ownerPartName);
  const relationships = new Map<string, Relationship>();
  for (const relNode of children(root)) {
    const id = attr(relNode, 'Id');
    const type = attr(relNode, 'Type');
    const target = attr(relNode, 'Target');
    if (!id || !type || !target) continue;

    const targetMode = attr(relNode, 'TargetMode') === 'External' ? 'External' : 'Internal';
    relationships.set(id, {
      id,
      type,
      target: targetMode === 'External' ? target : resolveTarget(baseDir, target),
      targetMode,
    });
  }

  return {
    get: (id) => relationships.get(id),
    findByType: (type) => [...relationships.values()].filter((rel) => rel.type === type),
  };
}
