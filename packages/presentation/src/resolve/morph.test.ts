import type { GroupShape, Picture, Shape, ShapeTreeNode, Slide } from '../presentationml/index.js';
import { describe, expect, it } from 'vitest';
import { resolveMorphMatch } from './morph.js';

function shape(id: number, name: string): Shape {
  return { kind: 'shape', nonVisual: { id, name }, properties: {} };
}

function textShape(id: number, name: string, text: string): Shape {
  return {
    kind: 'shape',
    nonVisual: { id, name },
    properties: {},
    textBody: { paragraphs: [{ runs: [{ kind: 'run', text }] }] },
  };
}

function picture(id: number, name: string): Picture {
  return { kind: 'picture', nonVisual: { id, name }, properties: {}, image: {} as never };
}

function group(id: number, name: string, children: readonly ShapeTreeNode[]): GroupShape {
  return { kind: 'group', nonVisual: { id, name }, transform: {} as never, children };
}

function slide(shapeTree: readonly ShapeTreeNode[]): Slide {
  return { commonSlideData: { shapeTree }, layout: {} as never };
}

describe('resolveMorphMatch', () => {
  it('matches a single shape by name across both slides', () => {
    const outgoing = slide([shape(2, 'Title 1')]);
    const incoming = slide([shape(2, 'Title 1')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([
      { outgoing: shape(2, 'Title 1'), incoming: shape(2, 'Title 1') },
    ]);
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([]);
  });

  it('reports an outgoing shape with no counterpart as disappearing', () => {
    const outgoing = slide([shape(2, 'Title 1'), shape(3, 'Subtitle 2')]);
    const incoming = slide([shape(2, 'Title 1')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toHaveLength(1);
    expect(result.disappearing).toEqual([shape(3, 'Subtitle 2')]);
    expect(result.appearing).toEqual([]);
  });

  it('reports an incoming shape with no counterpart as appearing', () => {
    const outgoing = slide([shape(2, 'Title 1')]);
    const incoming = slide([shape(2, 'Title 1'), shape(9, 'New Callout')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toHaveLength(1);
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([shape(9, 'New Callout')]);
  });

  it('never matches shapes with a blank name, even when both sides have one', () => {
    const outgoing = slide([shape(2, '')]);
    const incoming = slide([shape(2, '')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([]);
    expect(result.disappearing).toEqual([shape(2, '')]);
    expect(result.appearing).toEqual([shape(2, '')]);
  });

  it('disambiguates same-named duplicates by matching id first', () => {
    // Two shapes share the name "Icon" on each side; only ids 10/20 persist unchanged.
    const outgoing = slide([shape(10, 'Icon'), shape(20, 'Icon')]);
    const incoming = slide([shape(20, 'Icon'), shape(10, 'Icon')]); // reordered, ids still present

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toHaveLength(2);
    expect(result.matched).toContainEqual({
      outgoing: shape(10, 'Icon'),
      incoming: shape(10, 'Icon'),
    });
    expect(result.matched).toContainEqual({
      outgoing: shape(20, 'Icon'),
      incoming: shape(20, 'Icon'),
    });
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([]);
  });

  it('falls back to positional pairing for same-named duplicates when no id matches', () => {
    // Same name, but every id changed (e.g. copy/pasted rather than the same shape edited).
    const outgoing = slide([shape(10, 'Icon'), shape(11, 'Icon')]);
    const incoming = slide([shape(30, 'Icon'), shape(31, 'Icon')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([
      { outgoing: shape(10, 'Icon'), incoming: shape(30, 'Icon') },
      { outgoing: shape(11, 'Icon'), incoming: shape(31, 'Icon') },
    ]);
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([]);
  });

  it('leaves excess same-named duplicates on the larger side unmatched', () => {
    const outgoing = slide([shape(10, 'Icon')]);
    const incoming = slide([shape(30, 'Icon'), shape(31, 'Icon')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([{ outgoing: shape(10, 'Icon'), incoming: shape(30, 'Icon') }]);
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([shape(31, 'Icon')]);
  });

  it('recurses into groups on both sides without matching the group container itself', () => {
    const innerOutgoing = shape(5, 'Bullet 1');
    const innerIncoming = shape(5, 'Bullet 1');
    const outgoing = slide([group(1, 'Same Group Name', [innerOutgoing])]);
    const incoming = slide([group(1, 'Same Group Name', [innerIncoming])]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([{ outgoing: innerOutgoing, incoming: innerIncoming }]);
  });

  it('matches a shape that moved out of a group between the two slides', () => {
    const inner = shape(5, 'Bullet 1');
    const outgoing = slide([group(1, 'Group', [inner])]);
    const incoming = slide([shape(5, 'Bullet 1')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([{ outgoing: inner, incoming: shape(5, 'Bullet 1') }]);
  });

  it('matches across different node kinds sharing the same name/id', () => {
    // A picture placeholder replaced with an autoshape, keeping the same name/id.
    const outgoing = slide([picture(7, 'Placeholder 6')]);
    const incoming = slide([shape(7, 'Placeholder 6')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([
      { outgoing: picture(7, 'Placeholder 6'), incoming: shape(7, 'Placeholder 6') },
    ]);
  });

  it('matches every shape end-to-end for a real Duplicate-Slide-authored Morph transition', () => {
    // Mirrors apps/web-demo/src/Presentation1.pptx's slides 3 and 4: identical ids/names for four
    // shapes (the standard "duplicate slide, then tweak a few shapes" Morph-authoring pattern),
    // plus a fifth — a text box authored independently on each slide rather than duplicated, so its
    // id/name differ (id 6 "TextBox 5" outgoing vs. id 3 "TextBox 2" incoming) while its text
    // content is identical — matched only via the text-content fallback pass.
    const paragraphText =
      'A paragraph of text that will appear to scroll into view during the transition giving the ' +
      'illusion that the whole system is scrolling rather than being a slide deck.';
    const outgoing = slide([
      shape(2, 'Title 1'),
      shape(5, 'Content Placeholder 4'),
      picture(4, 'Graphic 3'),
      textShape(6, 'TextBox 5', paragraphText),
      shape(13, 'Freeform 12'),
    ]);
    const incoming = slide([
      shape(2, 'Title 1'),
      shape(5, 'Content Placeholder 4'),
      picture(4, 'Graphic 3'),
      textShape(3, 'TextBox 2', paragraphText),
      shape(13, 'Freeform 12'),
    ]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toHaveLength(5);
    expect(result.matched).toContainEqual({
      outgoing: textShape(6, 'TextBox 5', paragraphText),
      incoming: textShape(3, 'TextBox 2', paragraphText),
    });
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([]);
  });

  it('matches a text box by content when its name/id differ across slides', () => {
    const outgoing = slide([textShape(6, 'TextBox 5', 'Some caption')]);
    const incoming = slide([textShape(3, 'TextBox 2', 'Some caption')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([
      {
        outgoing: textShape(6, 'TextBox 5', 'Some caption'),
        incoming: textShape(3, 'TextBox 2', 'Some caption'),
      },
    ]);
    expect(result.disappearing).toEqual([]);
    expect(result.appearing).toEqual([]);
  });

  it('does not fall back to text content for shapes already matched by name', () => {
    // Name match should win outright — the text-content pass only ever looks at shapes that
    // survived the name pass unmatched, so it must never override or duplicate a name-based pair.
    const outgoing = slide([textShape(1, 'Caption', 'Hello'), textShape(2, 'Other', 'Hello')]);
    const incoming = slide([textShape(1, 'Caption', 'Hello'), textShape(2, 'Other', 'Hello')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toHaveLength(2);
    expect(result.matched).toContainEqual({
      outgoing: textShape(1, 'Caption', 'Hello'),
      incoming: textShape(1, 'Caption', 'Hello'),
    });
    expect(result.matched).toContainEqual({
      outgoing: textShape(2, 'Other', 'Hello'),
      incoming: textShape(2, 'Other', 'Hello'),
    });
  });

  it('never matches shapes with blank/empty text content via the text-content fallback', () => {
    const outgoing = slide([shape(1, 'Box A')]);
    const incoming = slide([shape(2, 'Box B')]);

    const result = resolveMorphMatch(outgoing, incoming);
    expect(result.matched).toEqual([]);
    expect(result.disappearing).toEqual([shape(1, 'Box A')]);
    expect(result.appearing).toEqual([shape(2, 'Box B')]);
  });
});
