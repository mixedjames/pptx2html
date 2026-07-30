import { describe, expect, it } from 'vitest';
import { EMU_PER_PT, EMU_PER_PX, emuToCqw, emuToPx, fontSizeToEmu } from './units.js';

describe('emuToPx', () => {
  it('converts EMU to px at 96 DPI', () => {
    expect(emuToPx(EMU_PER_PX)).toBe(1);
    expect(emuToPx(914400)).toBe(96);
  });
});

describe('fontSizeToEmu', () => {
  it('converts hundredths-of-a-point to EMU', () => {
    expect(fontSizeToEmu(1800)).toBe(18 * EMU_PER_PT);
    expect(fontSizeToEmu(100)).toBe(EMU_PER_PT);
  });
});

describe('emuToCqw', () => {
  it('expresses an EMU magnitude as a percentage of the slide width, suffixed cqw', () => {
    expect(emuToCqw(914400, 9144000)).toBe('10cqw');
  });

  it('scales a font size relative to a real slide width', () => {
    // An 18pt font on a standard 10in-wide (9144000 EMU) 4:3 slide.
    expect(emuToCqw(fontSizeToEmu(1800), 9144000)).toBe('2.5cqw');
  });
});
