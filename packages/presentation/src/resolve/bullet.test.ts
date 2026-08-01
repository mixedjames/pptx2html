import { describe, expect, it } from 'vitest';
import { formatAutoNumber, NumberingState } from './bullet.js';

describe('formatAutoNumber', () => {
  it('formats arabic schemes', () => {
    expect(formatAutoNumber(1, 'arabicPeriod')).toBe('1.');
    expect(formatAutoNumber(42, 'arabicParenR')).toBe('42)');
  });

  it('formats alpha schemes, wrapping past z into aa/ab/...', () => {
    expect(formatAutoNumber(1, 'alphaLcPeriod')).toBe('a.');
    expect(formatAutoNumber(26, 'alphaLcPeriod')).toBe('z.');
    expect(formatAutoNumber(27, 'alphaLcPeriod')).toBe('aa.');
    expect(formatAutoNumber(28, 'alphaUcParenR')).toBe('AB)');
  });

  it('formats roman schemes', () => {
    expect(formatAutoNumber(1, 'romanUcPeriod')).toBe('I.');
    expect(formatAutoNumber(4, 'romanUcPeriod')).toBe('IV.');
    expect(formatAutoNumber(9, 'romanLcParenR')).toBe('ix)');
    expect(formatAutoNumber(1994, 'romanUcPeriod')).toBe('MCMXCIV.');
  });
});

describe('NumberingState', () => {
  it('increments consecutive same-level, same-scheme numbers', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(2);
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(3);
  });

  it('honors startAt on the first call for a level, then ignores it on later calls', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod', startAt: 5 })).toBe(5);
    expect(state.next(0, { scheme: 'arabicPeriod', startAt: 5 })).toBe(6);
  });

  it('restarts when the scheme changes at the same level', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(2);
    expect(state.next(0, { scheme: 'romanUcPeriod' })).toBe(1);
  });

  it('tracks independent counters per outline level', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(2);
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(2);
  });

  it('resets a deeper level whenever a shallower-or-equal paragraph is visited', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(2);
    // Back to level 0 — the level-1 sub-list should restart the next time it's entered.
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(2);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(1);
  });

  it('break() ends a running list at that level and deeper', () => {
    const state = new NumberingState();
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(1);
    state.break(0);
    expect(state.next(0, { scheme: 'arabicPeriod' })).toBe(1);
    expect(state.next(1, { scheme: 'arabicPeriod' })).toBe(1);
  });
});
