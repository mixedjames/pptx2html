import { describe, expect, it } from 'vitest';

import { parseRelationships, relationshipsPartNameFor } from './relationships.js';

describe('opc/relationships', () => {
  it('computes the _rels part name for a nested part', () => {
    expect(relationshipsPartNameFor('ppt/slides/slide1.xml')).toBe(
      'ppt/slides/_rels/slide1.xml.rels',
    );
  });

  it('computes the root _rels part name', () => {
    expect(relationshipsPartNameFor('')).toBe('_rels/.rels');
  });

  it("resolves relative targets against the owning part's directory", () => {
    const xml = `<Relationships xmlns="r">
      <Relationship Id="rId1" Type="t" Target="../media/image1.png"/>
      <Relationship Id="rId2" Type="t" Target="slideLayout1.xml"/>
      <Relationship Id="rId3" Type="t" Target="/ppt/theme/theme1.xml"/>
    </Relationships>`;
    const rels = parseRelationships('ppt/slides/slide1.xml', xml);

    expect(rels.get('rId1')?.target).toBe('ppt/media/image1.png');
    expect(rels.get('rId2')?.target).toBe('ppt/slides/slideLayout1.xml');
    expect(rels.get('rId3')?.target).toBe('ppt/theme/theme1.xml');
  });

  it('preserves external targets untouched and marks their mode', () => {
    const xml = `<Relationships xmlns="r">
      <Relationship Id="rId1" Type="hyperlink" Target="https://example.com" TargetMode="External"/>
    </Relationships>`;
    const rels = parseRelationships('ppt/slides/slide1.xml', xml);
    const rel = rels.get('rId1');
    expect(rel?.targetMode).toBe('External');
    expect(rel?.target).toBe('https://example.com');
  });

  it('returns an empty relationship set when no .rels part exists', () => {
    const rels = parseRelationships('ppt/slides/slide1.xml', undefined);
    expect(rels.get('rId1')).toBeUndefined();
    expect(rels.findByType('anything')).toEqual([]);
  });
});
