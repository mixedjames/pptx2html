import { unzipSync } from 'fflate';

import { parseContentTypes, type ContentTypes } from './content-types.js';
import {
  parseRelationships,
  relationshipsPartNameFor,
  type Relationships,
} from './relationships.js';

const textDecoder = new TextDecoder('utf-8');

function normalizePartName(name: string): string {
  return name.startsWith('/') ? name.slice(1) : name;
}

/** Read access to a .pptx's underlying OPC/ZIP package: parts, content types and relationships. */
export class OpcPackage {
  private readonly parts = new Map<string, Uint8Array>();
  private readonly relationshipsCache = new Map<string, Relationships>();
  private readonly contentTypes: ContentTypes;

  constructor(data: Uint8Array) {
    const entries = unzipSync(data);
    for (const [name, bytes] of Object.entries(entries)) {
      if (name.endsWith('/')) continue; // directory entry
      this.parts.set(normalizePartName(name), bytes);
    }
    this.contentTypes = parseContentTypes(this.tryReadText('[Content_Types].xml'));
  }

  private tryReadText(partName: string): string | undefined {
    const bytes = this.parts.get(normalizePartName(partName));
    return bytes && textDecoder.decode(bytes);
  }

  readBytes(partName: string): Uint8Array {
    const bytes = this.parts.get(normalizePartName(partName));
    if (!bytes) throw new Error(`Part not found in package: ${partName}`);
    return bytes;
  }

  readText(partName: string): string {
    return textDecoder.decode(this.readBytes(partName));
  }

  contentTypeFor(partName: string): string | undefined {
    return this.contentTypes.contentTypeFor(partName);
  }

  relationshipsFor(partName: string): Relationships {
    const normalized = normalizePartName(partName);
    let relationships = this.relationshipsCache.get(normalized);
    if (!relationships) {
      relationships = parseRelationships(
        normalized,
        this.tryReadText(relationshipsPartNameFor(normalized)),
      );
      this.relationshipsCache.set(normalized, relationships);
    }
    return relationships;
  }
}
