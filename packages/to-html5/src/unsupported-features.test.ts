import { describe, expect, it } from 'vitest';
import { UnsupportedFeatureCollector } from './unsupported-features.js';

describe('UnsupportedFeatureCollector', () => {
  it('collects reported features in order via .all', () => {
    const collector = new UnsupportedFeatureCollector();
    collector.report({ code: 'a', message: 'first', slideIndex: 0 });
    collector.report({ code: 'b', message: 'second', slideIndex: 1 });

    expect(collector.all).toEqual([
      { code: 'a', message: 'first', slideIndex: 0 },
      { code: 'b', message: 'second', slideIndex: 1 },
    ]);
  });

  it('groups features by slideIndex via .bySlide, preserving per-slide order', () => {
    const collector = new UnsupportedFeatureCollector();
    collector.report({ code: 'a', message: 'slide 0, first', slideIndex: 0 });
    collector.report({ code: 'b', message: 'slide 1', slideIndex: 1 });
    collector.report({ code: 'c', message: 'slide 0, second', slideIndex: 0 });

    const bySlide = collector.bySlide;
    expect([...bySlide.keys()]).toEqual([0, 1]);
    expect(bySlide.get(0)?.map((f) => f.message)).toEqual(['slide 0, first', 'slide 0, second']);
    expect(bySlide.get(1)?.map((f) => f.message)).toEqual(['slide 1']);
  });

  it('excludes presentation-level features (no slideIndex) from .bySlide', () => {
    const collector = new UnsupportedFeatureCollector();
    collector.report({ code: 'a', message: 'presentation-level' });
    collector.report({ code: 'b', message: 'slide 0', slideIndex: 0 });

    expect(collector.bySlide.size).toBe(1);
    expect(collector.all).toHaveLength(2);
  });
});
