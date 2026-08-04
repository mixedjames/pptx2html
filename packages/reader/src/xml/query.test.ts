import { describe, expect, it } from 'vitest';

import { parseXml } from './parse.js';
import {
  attr,
  children,
  findAllChildren,
  findAlternateContentChild,
  findChild,
  localName,
  textOf,
} from './query.js';

describe('xml/query', () => {
  it('preserves document order across differently-named siblings', () => {
    const [root] = parseXml(
      '<a:p xmlns:a="a"><a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r></a:p>',
    );
    const tags = children(root!).map((node) => localName(node));
    expect(tags).toEqual(['r', 'br', 'r']);
  });

  it('strips namespace prefixes from element names but not from attribute names', () => {
    const [root] = parseXml('<p:sldMasterId xmlns:p="p" xmlns:r="r" id="2147483648" r:id="rId1"/>');
    expect(localName(root!)).toBe('sldMasterId');
    expect(attr(root!, 'id')).toBe('2147483648');
    expect(attr(root!, 'r:id')).toBe('rId1');
  });

  it('replaces mc:AlternateContent with its mc:Fallback content', () => {
    const [root] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:mc="mc">
        <p:sp><p:nvSpPr/></p:sp>
        <mc:AlternateContent>
          <mc:Choice Requires="future"><p:futureShape/></mc:Choice>
          <mc:Fallback><p:sp><p:nvSpPr/></p:sp></mc:Fallback>
        </mc:AlternateContent>
      </p:spTree>`,
    );
    const tags = children(root!).map((node) => localName(node));
    expect(tags).toEqual(['sp', 'sp']);
  });

  it('falls back to mc:Choice when no mc:Fallback is present', () => {
    const [root] = parseXml(
      `<p:spTree xmlns:p="p" xmlns:mc="mc">
        <mc:AlternateContent>
          <mc:Choice Requires="x"><p:sp/></mc:Choice>
        </mc:AlternateContent>
      </p:spTree>`,
    );
    expect(children(root!).map((node) => localName(node))).toEqual(['sp']);
  });

  describe('findAlternateContentChild', () => {
    it('returns both the Choice and Fallback branch versions of a wrapped child', () => {
      const [root] = parseXml(
        `<p:sld xmlns:p="p" xmlns:mc="mc" xmlns:p159="p159">
          <mc:AlternateContent>
            <mc:Choice Requires="p159"><p:transition><p159:morph/></p:transition></mc:Choice>
            <mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>
          </mc:AlternateContent>
        </p:sld>`,
      );
      const { choice, resolved } = findAlternateContentChild(root!, 'transition');
      expect(choice && findChild(choice, 'morph')).toBeDefined();
      expect(resolved && findChild(resolved, 'fade')).toBeDefined();
    });

    it('returns the same node for choice and resolved when there is no Fallback', () => {
      const [root] = parseXml(
        `<p:sld xmlns:p="p" xmlns:mc="mc" xmlns:p159="p159">
          <mc:AlternateContent>
            <mc:Choice Requires="p159"><p:transition><p159:morph/></p:transition></mc:Choice>
          </mc:AlternateContent>
        </p:sld>`,
      );
      const { choice, resolved } = findAlternateContentChild(root!, 'transition');
      expect(choice).toBe(resolved);
    });

    it('returns only resolved for a plain, unwrapped child', () => {
      const [root] = parseXml(`<p:sld xmlns:p="p"><p:transition><p:push/></p:transition></p:sld>`);
      const { choice, resolved } = findAlternateContentChild(root!, 'transition');
      expect(choice).toBeUndefined();
      expect(resolved && findChild(resolved, 'push')).toBeDefined();
    });

    it('returns an empty object when the named child is absent entirely', () => {
      const [root] = parseXml(`<p:sld xmlns:p="p"><p:cSld/></p:sld>`);
      expect(findAlternateContentChild(root!, 'transition')).toEqual({});
    });
  });

  it('finds all matching children and concatenates nested text', () => {
    const [root] = parseXml(
      '<a:p xmlns:a="a"><a:r><a:t>Hello, </a:t></a:r><a:r><a:t>world</a:t></a:r></a:p>',
    );
    const runs = findAllChildren(root!, 'r');
    expect(runs).toHaveLength(2);
    expect(textOf(runs[0]!)).toBe('Hello, ');
    expect(findChild(root!, 'r')).toBe(runs[0]);
  });
});
