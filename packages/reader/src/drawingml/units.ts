import type { Angle, Emu, FontSize, Percentage } from '@pptx2html/presentation';

function parseInteger(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const parsed = Number.parseInt(value, 10);
  return Number.isNaN(parsed) ? undefined : parsed;
}

export const parseEmu = (value: string | undefined): Emu | undefined => parseInteger(value);
export const parseAngle = (value: string | undefined): Angle | undefined => parseInteger(value);
export const parseFontSize = (value: string | undefined): FontSize | undefined =>
  parseInteger(value);
export const parseIntAttr = (value: string | undefined): number | undefined => parseInteger(value);

/** ST_Percentage (§20.1.10.42): a plain integer in 1000ths of a percent, or e.g. `"50%"`. */
export function parsePercentage(value: string | undefined): Percentage | undefined {
  if (value === undefined) return undefined;
  if (value.endsWith('%')) {
    const parsed = Number.parseFloat(value.slice(0, -1));
    return Number.isNaN(parsed) ? undefined : Math.round(parsed * 1000);
  }
  return parseInteger(value);
}

export function parseBoolean(value: string | undefined): boolean | undefined {
  if (value === undefined) return undefined;
  return value === '1' || value === 'true';
}
