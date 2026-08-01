import type { AutoNumberScheme } from '../drawingml/index.js';

const ROMAN_NUMERALS: readonly (readonly [number, string])[] = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

function toRoman(n: number): string {
  let result = '';
  let remaining = n;
  for (const [value, symbol] of ROMAN_NUMERALS) {
    while (remaining >= value) {
      result += symbol;
      remaining -= value;
    }
  }
  return result;
}

/** 1 -> a, 26 -> z, 27 -> aa, 28 -> ab, ... — spreadsheet-column-style base-26, not base-26 with a 0 digit. */
function toAlpha(n: number): string {
  let result = '';
  let remaining = n;
  while (remaining > 0) {
    remaining -= 1;
    result = String.fromCharCode(97 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return result;
}

/** Renders a 1-based ordinal as the label its auto-number scheme calls for (§20.1.10.51). */
export function formatAutoNumber(n: number, scheme: AutoNumberScheme): string {
  switch (scheme) {
    case 'arabicPeriod':
      return `${n}.`;
    case 'arabicParenR':
      return `${n})`;
    case 'alphaLcPeriod':
      return `${toAlpha(n)}.`;
    case 'alphaUcPeriod':
      return `${toAlpha(n).toUpperCase()}.`;
    case 'alphaLcParenR':
      return `${toAlpha(n)})`;
    case 'alphaUcParenR':
      return `${toAlpha(n).toUpperCase()})`;
    case 'romanLcPeriod':
      return `${toRoman(n).toLowerCase()}.`;
    case 'romanUcPeriod':
      return `${toRoman(n)}.`;
    case 'romanLcParenR':
      return `${toRoman(n).toLowerCase()})`;
    case 'romanUcParenR':
      return `${toRoman(n)})`;
  }
}

interface LevelCount {
  readonly next: number;
  readonly scheme: AutoNumberScheme;
}

/**
 * Tracks the running counter for each outline level's auto-numbered list as `renderTextBody`
 * walks a text body's paragraphs in order — numbering isn't a per-paragraph concern (unlike
 * alignment/bullet glyph, which `text-style.ts` resolves independently per paragraph) since each
 * number depends on how many auto-numbered siblings at the same level came before it.
 */
export class NumberingState {
  readonly #counters = new Map<number, LevelCount>();

  /**
   * Returns the next number for an auto-numbered paragraph at `level`. A shallower-or-equal
   * paragraph always starts a fresh sub-list at any deeper level (so a nested numbered list
   * restarts the next time it's entered); a same-level, same-scheme predecessor continues
   * incrementing, otherwise this call starts fresh at `bullet`'s own `startAt` (or 1).
   */
  next(
    level: number,
    bullet: { readonly scheme: AutoNumberScheme; readonly startAt?: number },
  ): number {
    for (const trackedLevel of this.#counters.keys()) {
      if (trackedLevel > level) this.#counters.delete(trackedLevel);
    }
    const existing = this.#counters.get(level);
    const value =
      existing && existing.scheme === bullet.scheme ? existing.next : (bullet.startAt ?? 1);
    this.#counters.set(level, { next: value + 1, scheme: bullet.scheme });
    return value;
  }

  /** Breaks any running numbered list at `level` and deeper — call for a paragraph at `level` that isn't auto-numbered. */
  break(level: number): void {
    for (const trackedLevel of this.#counters.keys()) {
      if (trackedLevel >= level) this.#counters.delete(trackedLevel);
    }
  }
}
