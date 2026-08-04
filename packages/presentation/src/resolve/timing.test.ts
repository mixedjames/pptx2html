import { describe, expect, it } from 'vitest';
import type { CommonTimeNodeData, TimeCondition, TimeNode } from '../presentationml/animation.js';
import {
  resolveSlideTimingDuration,
  resolveTimeNodeDuration,
  resolveTimeNodeStartMs,
  resolveTransitionDurationMs,
} from './timing.js';

function leaf(common: Partial<CommonTimeNodeData> = {}): TimeNode {
  return { kind: 'set', common: { id: 0, ...common } };
}

function par(children: readonly TimeNode[], common: Partial<CommonTimeNodeData> = {}): TimeNode {
  return { kind: 'par', common: { id: 0, ...common }, children };
}

function seq(
  children: readonly TimeNode[],
  concurrent?: boolean,
  common: Partial<CommonTimeNodeData> = {},
): TimeNode {
  return {
    kind: 'seq',
    common: { id: 0, ...common },
    children,
    ...(concurrent !== undefined ? { concurrent } : {}),
  };
}

function excl(children: readonly TimeNode[], common: Partial<CommonTimeNodeData> = {}): TimeNode {
  return { kind: 'excl', common: { id: 0, ...common }, children };
}

describe('resolveTransitionDurationMs', () => {
  it('maps fast/med/slow to their documented millisecond values', () => {
    expect(resolveTransitionDurationMs('fast')).toBe(400);
    expect(resolveTransitionDurationMs('med')).toBe(700);
    expect(resolveTransitionDurationMs('slow')).toBe(1000);
  });

  it('defaults absent speed to "fast"', () => {
    expect(resolveTransitionDurationMs(undefined)).toBe(400);
  });

  it('lets an explicit durationMs override speed outright', () => {
    expect(resolveTransitionDurationMs('slow', 2000)).toBe(2000);
    expect(resolveTransitionDurationMs(undefined, 2000)).toBe(2000);
  });
});

describe('resolveTimeNodeDuration', () => {
  it('uses an explicit numeric duration on a leaf', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 500 }))).toBe(500);
  });

  it('uses an explicit "indefinite" duration on a leaf', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 'indefinite' }))).toBe('indefinite');
  });

  it('treats a leaf with no explicit duration as instantaneous (0)', () => {
    expect(resolveTimeNodeDuration(leaf())).toBe(0);
  });

  it('derives a par container as the max across its children', () => {
    const node = par([leaf({ duration: 300 }), leaf({ duration: 700 })]);
    expect(resolveTimeNodeDuration(node)).toBe(700);
  });

  it('derives a non-concurrent seq container as the sum of its children', () => {
    const node = seq([leaf({ duration: 300 }), leaf({ duration: 700 })]);
    expect(resolveTimeNodeDuration(node)).toBe(1000);
  });

  it('derives a concurrent seq container as the max across its children', () => {
    const node = seq([leaf({ duration: 300 }), leaf({ duration: 700 })], true);
    expect(resolveTimeNodeDuration(node)).toBe(700);
  });

  it('derives an excl container as the sum of its children (documented approximation)', () => {
    const node = excl([leaf({ duration: 300 }), leaf({ duration: 700 })]);
    expect(resolveTimeNodeDuration(node)).toBe(1000);
  });

  it("adds a child's numeric start-condition delay into its contribution", () => {
    const delayed: TimeCondition = { delay: 200 };
    const node = seq([
      leaf({ duration: 300, startConditions: [delayed] }), // contributes 200 + 300 = 500
      leaf({ duration: 100 }), // contributes 100
    ]);
    expect(resolveTimeNodeDuration(node)).toBe(600);
  });

  it('makes the whole tree "indefinite" when a child is gated on a click with no numeric delay', () => {
    const clickGated: TimeCondition = { event: 'onClick' };
    const node = seq([
      leaf({ duration: 300, startConditions: [clickGated] }),
      leaf({ duration: 100 }),
    ]);
    expect(resolveTimeNodeDuration(node)).toBe('indefinite');
  });

  it("resolves via the numeric condition when start conditions are OR'd with one indefinite and one numeric", () => {
    const node = leaf({
      duration: 100,
      startConditions: [{ event: 'onClick' }, { delay: 50 }],
    });
    const wrapped = seq([node]);
    expect(resolveTimeNodeDuration(wrapped)).toBe(150); // 50 (the numeric condition wins) + 100
  });

  it('multiplies duration by repeatCount (1000ths units — 3000 means 3 repeats)', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 200, repeatCount: 3000 }))).toBe(600);
  });

  it('propagates "indefinite" for repeatCount: "indefinite"', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 200, repeatCount: 'indefinite' }))).toBe(
      'indefinite',
    );
  });

  it('lets an explicit repeatDuration override the computed base outright', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 200, repeatDuration: 5000 }))).toBe(5000);
  });

  it('approximates autoReverse as doubling one cycle', () => {
    expect(resolveTimeNodeDuration(leaf({ duration: 200, autoReverse: true }))).toBe(400);
  });

  it('combines autoReverse and repeatCount', () => {
    expect(
      resolveTimeNodeDuration(leaf({ duration: 200, autoReverse: true, repeatCount: 2000 })),
    ).toBe(800); // (200 * 2) * 2
  });

  it("lets a container's own explicit duration override its children-derived computation", () => {
    const node = par([leaf({ duration: 300 }), leaf({ duration: 700 })], { duration: 50 });
    expect(resolveTimeNodeDuration(node)).toBe(50);
  });

  it('composes nested containers (a par containing a seq containing two leaves)', () => {
    const node = par([
      seq([leaf({ duration: 100 }), leaf({ duration: 150 })]),
      leaf({ duration: 200 }),
    ]);
    expect(resolveTimeNodeDuration(node)).toBe(250); // max(100+150, 200)
  });
});

describe('resolveSlideTimingDuration', () => {
  it('is 0 when the timing is undefined', () => {
    expect(resolveSlideTimingDuration(undefined)).toBe(0);
  });

  it('is 0 when the timing has no timeNodeTree', () => {
    expect(resolveSlideTimingDuration({})).toBe(0);
  });

  it("delegates to resolveTimeNodeDuration for the timing's own tree", () => {
    expect(resolveSlideTimingDuration({ timeNodeTree: leaf({ duration: 250 }) })).toBe(250);
  });
});

describe('resolveTimeNodeStartMs', () => {
  it('is 0 when the node has no start conditions', () => {
    expect(resolveTimeNodeStartMs(leaf())).toBe(0);
  });

  it('uses an explicit numeric delay', () => {
    expect(resolveTimeNodeStartMs(leaf({ startConditions: [{ delay: 300 }] }))).toBe(300);
  });

  it('is "indefinite" for an event-gated condition with no delay fallback', () => {
    expect(resolveTimeNodeStartMs(leaf({ startConditions: [{ event: 'onClick' }] }))).toBe(
      'indefinite',
    );
  });

  it('fires immediately for a condition with neither a delay nor an event', () => {
    expect(resolveTimeNodeStartMs(leaf({ startConditions: [{}] }))).toBe(0);
  });

  it("takes the earliest of multiple OR'd conditions", () => {
    expect(
      resolveTimeNodeStartMs(
        leaf({ startConditions: [{ event: 'onClick' }, { delay: 500 }, { delay: 100 }] }),
      ),
    ).toBe(100);
  });

  it("does not add in a parent container's own start offset", () => {
    // Node-local only — composing with an ancestor's offset is a renderer's own job if it needs
    // that, not this function's (see its own doc comment).
    const child = leaf({ startConditions: [{ delay: 200 }] });
    par([child], { startConditions: [{ delay: 1000 }] });
    expect(resolveTimeNodeStartMs(child)).toBe(200);
  });
});
