/** EMU per CSS pixel at 96 DPI (914400 EMU/inch ÷ 96 px/inch). */
export const EMU_PER_PX = 9525;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}
