import type { AnimationTarget, CommonTimeNodeData, TimeNode } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
import { collectFadeAnimations } from './animation.js';

const SHAPE_TARGET: AnimationTarget = { kind: 'shape', shapeId: 7 };

function animEffect(
  overrides: Partial<CommonTimeNodeData> = {},
  rest: { filter?: string; transition?: 'in' | 'out'; target?: AnimationTarget } = {},
): TimeNode {
  return {
    kind: 'animEffect',
    common: { id: 0, ...overrides },
    target: SHAPE_TARGET,
    filter: 'fade',
    ...rest,
  };
}

function par(children: readonly TimeNode[]): TimeNode {
  return { kind: 'par', common: { id: 0 }, children };
}

function seq(children: readonly TimeNode[]): TimeNode {
  return { kind: 'seq', common: { id: 0 }, children };
}

describe('collectFadeAnimations', () => {
  it('returns nothing and reports nothing for an absent timing', () => {
    const report = vi.fn();
    expect(collectFadeAnimations(undefined, report)).toEqual([]);
    expect(report).not.toHaveBeenCalled();
  });

  it('returns nothing and reports nothing for a timing with no timeNodeTree or buildList', () => {
    const report = vi.fn();
    expect(collectFadeAnimations({}, report)).toEqual([]);
    expect(report).not.toHaveBeenCalled();
  });

  it('collects a single fade-in animEffect targeting a shape', () => {
    const report = vi.fn();
    const fades = collectFadeAnimations(
      { timeNodeTree: animEffect({ duration: 500 }, { transition: 'in' }) },
      report,
    );
    expect(fades).toEqual([{ shapeId: 7, direction: 'in', delayMs: 0, durationMs: 500 }]);
    expect(report).not.toHaveBeenCalled();
  });

  it('defaults an absent transition to "in", per §19.7.9', () => {
    const fades = collectFadeAnimations(
      { timeNodeTree: animEffect({ duration: 500 }, { transition: undefined }) },
      vi.fn(),
    );
    expect(fades[0]!.direction).toBe('in');
  });

  it('collects a fade-out animEffect', () => {
    const fades = collectFadeAnimations(
      { timeNodeTree: animEffect({ duration: 300 }, { transition: 'out' }) },
      vi.fn(),
    );
    expect(fades).toEqual([{ shapeId: 7, direction: 'out', delayMs: 0, durationMs: 300 }]);
  });

  it("uses the node's own numeric start delay", () => {
    const fades = collectFadeAnimations(
      {
        timeNodeTree: animEffect({ duration: 500, startConditions: [{ delay: 250 }] }),
      },
      vi.fn(),
    );
    expect(fades[0]!.delayMs).toBe(250);
  });

  it('walks into par/seq/excl containers to find nested fade behaviors', () => {
    const fades = collectFadeAnimations(
      {
        timeNodeTree: par([
          seq([animEffect({ duration: 100 }), animEffect({ duration: 200 })]),
          { kind: 'excl', common: { id: 0 }, children: [animEffect({ duration: 300 })] },
        ]),
      },
      vi.fn(),
    );
    expect(fades.map((f) => f.durationMs)).toEqual([100, 200, 300]);
  });

  it('plays a click-gated fade immediately (delay 0) and reports the trigger gap once', () => {
    const report = vi.fn();
    const fades = collectFadeAnimations(
      {
        timeNodeTree: animEffect({ duration: 500, startConditions: [{ event: 'onClick' }] }),
      },
      report,
    );
    expect(fades[0]!.delayMs).toBe(0);
    expect(report).toHaveBeenCalledWith(
      'animation-trigger-unmodeled',
      expect.stringContaining('1 fade animation'),
    );
  });

  it('does not play and reports a fade with an indefinite duration', () => {
    const report = vi.fn();
    const fades = collectFadeAnimations(
      { timeNodeTree: animEffect({ duration: 'indefinite' }) },
      report,
    );
    expect(fades).toEqual([]);
    expect(report).toHaveBeenCalledWith(
      'animation-effect-unmodeled',
      expect.stringContaining('1 animation effect'),
    );
  });

  it('reports (and does not play) an animEffect with a non-fade filter', () => {
    const report = vi.fn();
    const fades = collectFadeAnimations(
      { timeNodeTree: animEffect({}, { filter: 'wipe(right)' }) },
      report,
    );
    expect(fades).toEqual([]);
    expect(report).toHaveBeenCalledWith(
      'animation-effect-unmodeled',
      expect.stringContaining('1 animation effect'),
    );
  });

  it('reports (and does not play) a fade targeting shape text rather than a whole shape', () => {
    const report = vi.fn();
    const fades = collectFadeAnimations(
      {
        timeNodeTree: animEffect({}, { target: { kind: 'shapeText', shapeId: 7 } }),
      },
      report,
    );
    expect(fades).toEqual([]);
    expect(report).toHaveBeenCalledWith(
      'animation-effect-unmodeled',
      expect.stringContaining('1 animation effect'),
    );
  });

  it('reports every other behavior kind encountered, deduplicated by kind', () => {
    const report = vi.fn();
    collectFadeAnimations(
      {
        timeNodeTree: par([
          { kind: 'set', common: { id: 0 } },
          { kind: 'animClr', common: { id: 0 } },
          { kind: 'set', common: { id: 0 } },
        ]),
      },
      report,
    );
    expect(report).toHaveBeenCalledWith(
      'animation-behavior-unmodeled',
      expect.stringContaining('animClr, set'),
    );
  });

  it('reports an unconsumed buildList', () => {
    const report = vi.fn();
    collectFadeAnimations(
      { buildList: [{ kind: 'paragraph', shapeId: 1, buildType: 'byParagraph' }] },
      report,
    );
    expect(report).toHaveBeenCalledWith(
      'animation-build-unmodeled',
      expect.stringContaining('1 paragraph/graphic build'),
    );
  });
});
