import type { MediaPart } from '@pptx2html/presentation';

import type { OpcPackage } from '../opc/package.js';
import type { MediaResolver } from './fill.js';

const DEFAULT_CONTENT_TYPE = 'application/octet-stream';

/**
 * Builds a resolver for `r:embed`/`r:link` relationship ids scoped to `partName`'s own
 * relationships. Media parts are cached in `mediaCache` (keyed by resolved part name, shared
 * across the whole read) so an image referenced from several places is read from the zip once.
 */
export function createMediaResolver(
  pkg: OpcPackage,
  partName: string,
  mediaCache: Map<string, MediaPart>,
): MediaResolver {
  const relationships = pkg.relationshipsFor(partName);

  return (relationshipId: string): MediaPart | undefined => {
    const rel = relationships.get(relationshipId);
    if (!rel || rel.targetMode === 'External') return undefined;

    const cached = mediaCache.get(rel.target);
    if (cached) return cached;

    const media: MediaPart = {
      contentType: pkg.contentTypeFor(rel.target) ?? DEFAULT_CONTENT_TYPE,
      data: pkg.readBytes(rel.target),
    };
    mediaCache.set(rel.target, media);
    return media;
  };
}
