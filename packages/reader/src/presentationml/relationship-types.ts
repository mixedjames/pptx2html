const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Relationship Type URIs (§ Annex A / OPC) used to navigate between package parts. */
export const RELATIONSHIP_TYPES = {
  officeDocument: `${NS}/officeDocument`,
  slideMaster: `${NS}/slideMaster`,
  slideLayout: `${NS}/slideLayout`,
  slide: `${NS}/slide`,
  notesSlide: `${NS}/notesSlide`,
  notesMaster: `${NS}/notesMaster`,
  theme: `${NS}/theme`,
} as const;
