import type { Fill, FormatScheme, Line, StyleMatrixReference } from '@pptx2html/presentation';
import { describe, expect, it } from 'vitest';
import { resolveStyleFill, resolveStyleLine } from './style-matrix.js';

const PH_CLR_SOLID: Fill = { type: 'solid', color: { type: 'scheme', value: 'phClr' } };
const PH_CLR_SOLID_WITH_TINT: Fill = {
  type: 'solid',
  color: { type: 'scheme', value: 'phClr', transforms: { tint: 67000 } },
};
const PH_CLR_LINE: Line = { width: 12700, fill: PH_CLR_SOLID };

const FORMAT_SCHEME: FormatScheme = {
  name: 'Office',
  fillStyles: [PH_CLR_SOLID, PH_CLR_SOLID_WITH_TINT, { type: 'none' }],
  lineStyles: [PH_CLR_LINE, { width: 19050 }],
};

describe('resolveStyleFill', () => {
  it('substitutes phClr with the fillRef color, for the referenced 1-based index', () => {
    const ref: StyleMatrixReference = { index: 1, color: { type: 'scheme', value: 'accent1' } };
    expect(resolveStyleFill(ref, FORMAT_SCHEME)).toEqual({
      type: 'solid',
      color: { type: 'scheme', value: 'accent1' },
    });
  });

  it("keeps the style's own local colour transform on phClr, merged under the ref's own", () => {
    const ref: StyleMatrixReference = { index: 2, color: { type: 'scheme', value: 'accent2' } };
    expect(resolveStyleFill(ref, FORMAT_SCHEME)).toEqual({
      type: 'solid',
      color: { type: 'scheme', value: 'accent2', transforms: { tint: 67000 } },
    });
  });

  it("substitutes phClr through every gradient stop's color", () => {
    const scheme: FormatScheme = {
      name: 'Office',
      fillStyles: [
        {
          type: 'gradient',
          stops: [
            { position: 0, color: { type: 'scheme', value: 'phClr' } },
            {
              position: 100000,
              color: { type: 'scheme', value: 'phClr', transforms: { tint: 50000 } },
            },
          ],
        },
      ],
      lineStyles: [],
    };
    const ref: StyleMatrixReference = { index: 1, color: { type: 'srgb', value: 'FF0000' } };
    expect(resolveStyleFill(ref, scheme)).toEqual({
      type: 'gradient',
      stops: [
        { position: 0, color: { type: 'srgb', value: 'FF0000' } },
        { position: 100000, color: { type: 'srgb', value: 'FF0000', transforms: { tint: 50000 } } },
      ],
    });
  });

  it('returns undefined for an out-of-range index', () => {
    const ref: StyleMatrixReference = { index: 9, color: { type: 'scheme', value: 'accent1' } };
    expect(resolveStyleFill(ref, FORMAT_SCHEME)).toBeUndefined();
  });

  it('returns undefined when there is no reference or no format scheme', () => {
    expect(resolveStyleFill(undefined, FORMAT_SCHEME)).toBeUndefined();
    const ref: StyleMatrixReference = { index: 1, color: { type: 'scheme', value: 'accent1' } };
    expect(resolveStyleFill(ref, undefined)).toBeUndefined();
  });
});

describe('resolveStyleLine', () => {
  it("substitutes phClr in the line's own fill, keeping the line's other fields", () => {
    const ref: StyleMatrixReference = { index: 1, color: { type: 'scheme', value: 'accent3' } };
    expect(resolveStyleLine(ref, FORMAT_SCHEME)).toEqual({
      width: 12700,
      fill: { type: 'solid', color: { type: 'scheme', value: 'accent3' } },
    });
  });

  it('passes through a line style with no fill of its own unchanged', () => {
    const ref: StyleMatrixReference = { index: 2, color: { type: 'scheme', value: 'accent1' } };
    expect(resolveStyleLine(ref, FORMAT_SCHEME)).toEqual({ width: 19050 });
  });

  it('returns undefined for an out-of-range index', () => {
    const ref: StyleMatrixReference = { index: 9, color: { type: 'scheme', value: 'accent1' } };
    expect(resolveStyleLine(ref, FORMAT_SCHEME)).toBeUndefined();
  });
});
