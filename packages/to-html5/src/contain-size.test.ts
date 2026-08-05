// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { observeContainSize } from './contain-size.js';

/** happy-dom does no real layout, so `clientHeight`/`getBoundingClientRect` are stubbed directly
 *  (same convention `scroll-presentation-element.test.ts` uses for `scrollTop`) rather than
 *  relying on anything actually being rendered to a real size. */
function stubHeights(host: HTMLElement, target: HTMLElement, hostHeight: number): void {
  Object.defineProperty(host, 'clientHeight', { value: hostHeight, configurable: true });
  target.getBoundingClientRect = vi.fn(
    () => ({ height: 800, width: 1200 }) as DOMRect,
  ) as typeof target.getBoundingClientRect;
}

describe('observeContainSize', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('defaults to width-driven sizing (width: 100%, height: auto)', () => {
    const host = document.createElement('div');
    const target = document.createElement('div');
    host.appendChild(target);
    stubHeights(host, target, 0); // host has no independent height of its own (clientHeight 0)

    const controller = observeContainSize(host, target);

    expect(target.style.width).toBe('100%');
    expect(target.style.height).toBe('auto');
    controller.disconnect();
  });

  it("switches to height-driven sizing when the width-driven default would overflow the host's available height", () => {
    const host = document.createElement('div');
    const target = document.createElement('div');
    host.appendChild(target);
    // Width-driven default renders target at 800px tall (stubHeights), but host only has 600px
    // available — a proportionally-wide host relative to the deck, the actual bug this fixes.
    stubHeights(host, target, 600);

    const controller = observeContainSize(host, target);

    expect(target.style.width).toBe('auto');
    expect(target.style.height).toBe('100%');
    controller.disconnect();
  });

  it("stays width-driven when the width-driven default already fits within the host's available height", () => {
    const host = document.createElement('div');
    const target = document.createElement('div');
    host.appendChild(target);
    stubHeights(host, target, 1000); // target's 800px height comfortably fits

    const controller = observeContainSize(host, target);

    expect(target.style.width).toBe('100%');
    expect(target.style.height).toBe('auto');
    controller.disconnect();
  });

  it('reapply() re-runs the same decision on demand', () => {
    const host = document.createElement('div');
    const target = document.createElement('div');
    host.appendChild(target);
    stubHeights(host, target, 0);

    const controller = observeContainSize(host, target);
    expect(target.style.height).toBe('auto');

    // Simulate the host later gaining an independent, smaller height (e.g. a page's CSS taking
    // effect, or a new render() with a differently-shaped deck) and re-deciding without waiting
    // for an incidental resize.
    stubHeights(host, target, 600);
    controller.reapply();

    expect(target.style.width).toBe('auto');
    expect(target.style.height).toBe('100%');
    controller.disconnect();
  });

  it('disconnect() stops observing without throwing', () => {
    const host = document.createElement('div');
    const target = document.createElement('div');
    host.appendChild(target);
    stubHeights(host, target, 0);

    const controller = observeContainSize(host, target);
    expect(() => controller.disconnect()).not.toThrow();
  });
});
