import { parseXml } from '../xml/parse.js';
import { attr, children, findRoot, localName } from '../xml/query.js';

function normalize(partName: string): string {
  return partName.startsWith('/') ? partName.slice(1) : partName;
}

/** [Content_Types].xml (§10.1.2.2): resolves a part name to its declared MIME content type. */
export interface ContentTypes {
  contentTypeFor(partName: string): string | undefined;
}

const EMPTY_CONTENT_TYPES: ContentTypes = { contentTypeFor: () => undefined };

export function parseContentTypes(xmlText: string | undefined): ContentTypes {
  if (xmlText === undefined) return EMPTY_CONTENT_TYPES;

  const typesRoot = findRoot(parseXml(xmlText), 'Types');
  if (!typesRoot) return EMPTY_CONTENT_TYPES;

  const defaultsByExtension = new Map<string, string>();
  const overridesByPartName = new Map<string, string>();

  for (const child of children(typesRoot)) {
    const name = localName(child);
    if (name === 'Default') {
      const extension = attr(child, 'Extension');
      const contentType = attr(child, 'ContentType');
      if (extension && contentType) defaultsByExtension.set(extension.toLowerCase(), contentType);
    } else if (name === 'Override') {
      const partName = attr(child, 'PartName');
      const contentType = attr(child, 'ContentType');
      if (partName && contentType) overridesByPartName.set(normalize(partName), contentType);
    }
  }

  return {
    contentTypeFor(partName) {
      const normalized = normalize(partName);
      const override = overridesByPartName.get(normalized);
      if (override) return override;
      const dot = normalized.lastIndexOf('.');
      if (dot === -1) return undefined;
      return defaultsByExtension.get(normalized.slice(dot + 1).toLowerCase());
    },
  };
}
