import { describe, expect, it } from 'vitest';

import { parseContentTypes } from './content-types.js';

const XML = `<Types xmlns="ct">
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

describe('opc/content-types', () => {
  it('resolves an Override by exact part name over any Default', () => {
    const contentTypes = parseContentTypes(XML);
    expect(contentTypes.contentTypeFor('/ppt/slides/slide1.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    );
  });

  it('falls back to a Default by extension when there is no Override', () => {
    const contentTypes = parseContentTypes(XML);
    expect(contentTypes.contentTypeFor('ppt/media/image1.png')).toBe('image/png');
  });

  it('returns undefined for an unknown extension', () => {
    const contentTypes = parseContentTypes(XML);
    expect(contentTypes.contentTypeFor('ppt/media/image1.bmp')).toBeUndefined();
  });

  it('returns undefined content types when [Content_Types].xml is absent', () => {
    const contentTypes = parseContentTypes(undefined);
    expect(contentTypes.contentTypeFor('ppt/media/image1.png')).toBeUndefined();
  });
});
