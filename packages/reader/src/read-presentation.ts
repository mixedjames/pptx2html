import type { Presentation } from '@pptx2html/presentation';

import { OpcPackage } from './opc/package.js';
import { readPresentationPart } from './presentationml/presentation.js';
import { RELATIONSHIP_TYPES } from './presentationml/relationship-types.js';
import { createReaderContext } from './reader-context.js';

/** Parses a .pptx file's bytes into the in-memory Presentation object tree. */
export function readPresentation(data: Uint8Array): Presentation {
  const pkg = new OpcPackage(data);
  const context = createReaderContext(pkg);

  // The package root's relationships (_rels/.rels) point at the main presentation part.
  const rootRelationships = pkg.relationshipsFor('');
  const presentationRel = rootRelationships.findByType(RELATIONSHIP_TYPES.officeDocument)[0];
  if (!presentationRel) {
    throw new Error('Package is missing its main presentation part relationship');
  }

  return readPresentationPart(context, presentationRel.target);
}
