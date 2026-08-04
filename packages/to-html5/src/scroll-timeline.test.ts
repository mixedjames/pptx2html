import type { Presentation, Shape, Slide, SlideLayout } from '@pptx2html/presentation';
import { describe, expect, it, vi } from 'vitest';
import { resolveScrollTimeline } from './scroll-timeline.js';

const LAYOUT: SlideLayout = {
  commonSlideData: { shapeTree: [] },
  master: { commonSlideData: { shapeTree: [] }, theme: {} as never, layouts: [] },
  type: 'blank',
};

function slide(overrides: Partial<Slide> = {}): Slide {
  return { commonSlideData: { shapeTree: [] }, layout: LAYOUT, ...overrides };
}

function presentationOf(slides: readonly Slide[]): Presentation {
  return {
    slideSize: { width: 12192000, height: 6858000 },
    slideMasters: [],
    slides,
    notesSlides: [],
  };
}

function morphShape(id: number, name: string, x = 0, y = 0): Shape {
  return {
    kind: 'shape',
    nonVisual: { id, name },
    properties: { transform: { offset: { x, y }, extents: { width: 1000000, height: 1000000 } } },
  };
}

describe('resolveScrollTimeline', () => {
  it('returns an empty timeline for a deck with no slides', () => {
    const report = vi.fn();
    expect(resolveScrollTimeline(presentationOf([]), report)).toEqual({
      segments: [],
      totalDurationMs: 0,
    });
    expect(report).not.toHaveBeenCalled();
  });

  it('the first slide never gets a transition segment, even if it authors one', () => {
    const presentation = presentationOf([slide({ transition: { effect: { kind: 'fade' } } })]);
    const [segment] = resolveScrollTimeline(presentation, vi.fn()).segments;
    expect(segment?.transition).toBeUndefined();
  });

  it('a fully static slide gets exactly the minimum-dwell content duration', () => {
    const presentation = presentationOf([slide()]);
    const { segments, totalDurationMs } = resolveScrollTimeline(presentation, vi.fn());
    expect(segments[0]?.content).toEqual({ startMs: 0, endMs: 1200, fades: [] });
    expect(totalDurationMs).toBe(1200);
  });

  it('a custom minDwellMs overrides the default floor', () => {
    const presentation = presentationOf([slide()]);
    const { totalDurationMs } = resolveScrollTimeline(presentation, vi.fn(), { minDwellMs: 300 });
    expect(totalDurationMs).toBe(300);
  });

  it("content duration is the slide's own fades' latest end time when that exceeds the dwell floor", () => {
    const presentation = presentationOf([
      slide({
        timing: {
          timeNodeTree: {
            kind: 'animEffect',
            common: { id: 0, duration: 500, startConditions: [{ delay: 2000 }] },
            target: { kind: 'shape', shapeId: 1 },
            filter: 'fade',
          },
        },
      }),
    ]);
    const { segments } = resolveScrollTimeline(presentation, vi.fn());
    expect(segments[0]?.content.endMs).toBe(2500); // 2000 delay + 500 duration, not the 1200 floor
  });

  it('an authored push/fade transition plays as authored, timed with resolveTransitionDurationMs', () => {
    const presentation = presentationOf([
      slide(),
      slide({ transition: { effect: { kind: 'push', direction: 'r' }, speed: 'slow' } }),
    ]);
    const { segments } = resolveScrollTimeline(presentation, vi.fn());
    expect(segments[1]?.transition).toEqual({
      startMs: 1200, // right after slide 0's own content (dwell floor)
      endMs: 2200, // slow = 1000ms
      effect: { kind: 'push', direction: 'r' },
      morphMatch: undefined,
    });
    expect(segments[1]?.content.startMs).toBe(2200);
  });

  it('no authored transition at all falls back to a synthetic push-up, unreported', () => {
    const presentation = presentationOf([slide(), slide()]);
    const report = vi.fn();
    const { segments } = resolveScrollTimeline(presentation, report);
    expect(segments[1]?.transition?.effect).toEqual({ kind: 'push', direction: 'u' });
    expect(segments[1]?.transition?.endMs).toBe(1200 + 400); // absent speed defaults to "fast"
    expect(report).not.toHaveBeenCalled();
  });

  it('an unmodeled effect kind falls back to a synthetic push-up and reports it, keyed to that slide', () => {
    const presentation = presentationOf([
      slide(),
      slide({ transition: { effect: { kind: 'wipe', direction: 'l' } } }),
    ]);
    const report = vi.fn();
    const { segments } = resolveScrollTimeline(presentation, report);
    expect(segments[1]?.transition?.effect).toEqual({ kind: 'push', direction: 'u' });
    expect(report).toHaveBeenCalledWith(
      'transition-effect-approximated-for-scroll',
      expect.stringContaining('"wipe"'),
      1,
    );
  });

  it('a confidently-matched morph keeps its own effect and carries the match summary', () => {
    const presentation = presentationOf([
      slide({ commonSlideData: { shapeTree: [morphShape(1, 'Title 1')] } }),
      slide({
        commonSlideData: { shapeTree: [morphShape(1, 'Title 1', 1000000, 1000000)] },
        transition: { effect: { kind: 'morph' } },
      }),
    ]);
    const { segments } = resolveScrollTimeline(presentation, vi.fn());
    expect(segments[1]?.transition?.effect).toEqual({ kind: 'morph' });
    expect(segments[1]?.transition?.morphMatch?.matched).toEqual([
      { outgoingShapeId: 1, incomingShapeId: 1 },
    ]);
  });

  it('a low-confidence morph match degrades to a plain fade and reports morph-match-degraded', () => {
    const presentation = presentationOf([
      slide({ commonSlideData: { shapeTree: [morphShape(1, 'Alpha')] } }),
      slide({
        commonSlideData: { shapeTree: [morphShape(2, 'Beta')] },
        transition: { effect: { kind: 'morph' } },
      }),
    ]);
    const report = vi.fn();
    const { segments } = resolveScrollTimeline(presentation, report);
    expect(segments[1]?.transition?.effect).toEqual({ kind: 'fade' });
    expect(segments[1]?.transition?.morphMatch).toBeUndefined();
    expect(report).toHaveBeenCalledWith('morph-match-degraded', expect.any(String), 1);
  });

  it('chains multiple slides into one absolute-ms timeline in deck order', () => {
    const presentation = presentationOf([
      slide(), // content: 0 -> 1200
      slide({ transition: { effect: { kind: 'fade' } } }), // transition: 1200 -> 1600, content: 1600 -> 2800
      slide({ transition: { effect: { kind: 'push', direction: 'l' } } }), // transition: 2800 -> 3200, content: 3200 -> 4400
    ]);
    const { segments, totalDurationMs } = resolveScrollTimeline(presentation, vi.fn());
    expect(
      segments.map((s) => [s.transition?.startMs, s.content.startMs, s.content.endMs]),
    ).toEqual([
      [undefined, 0, 1200],
      [1200, 1600, 2800],
      [2800, 3200, 4400],
    ]);
    expect(totalDurationMs).toBe(4400);
  });

  it('forwards a build/behavior report from collectFadeAnimations, keyed to the right slide', () => {
    const presentation = presentationOf([
      slide(),
      slide({
        timing: { buildList: [{ kind: 'paragraph', shapeId: 1, buildType: 'byParagraph' }] },
      }),
    ]);
    const report = vi.fn();
    resolveScrollTimeline(presentation, report);
    expect(report).toHaveBeenCalledWith(
      'animation-build-unmodeled',
      expect.stringContaining('1 paragraph/graphic build'),
      1,
    );
  });
});
