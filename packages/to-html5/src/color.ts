import type {
  Color,
  ColorScheme,
  ColorTransform,
  Fill,
  SchemeColorName,
} from '@pptx2html/presentation';

type Rgb = readonly [number, number, number];

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function hexToRgb(hex: string): Rgb {
  const n = Number.parseInt(hex, 16);
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
}

function rgbToHsl([r, g, b]: Rgb): readonly [number, number, number] {
  const rf = r / 255;
  const gf = g / 255;
  const bf = b / 255;
  const max = Math.max(rf, gf, bf);
  const min = Math.min(rf, gf, bf);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return [0, 0, l];

  const s = d / (1 - Math.abs(2 * l - 1));
  let h: number;
  if (max === rf) h = ((gf - bf) / d) % 6;
  else if (max === gf) h = (bf - rf) / d + 2;
  else h = (rf - gf) / d + 4;
  h *= 60;
  if (h < 0) h += 360;
  return [h, s, l];
}

function hslToRgb([h, s, l]: readonly [number, number, number]): Rgb {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  const [r, g, b] =
    h < 60
      ? [c, x, 0]
      : h < 120
        ? [x, c, 0]
        : h < 180
          ? [0, c, x]
          : h < 240
            ? [0, x, c]
            : h < 300
              ? [x, 0, c]
              : [c, 0, x];
  return [(r + m) * 255, (g + m) * 255, (b + m) * 255];
}

/**
 * Approximates DrawingML's colour transform algorithms (§20.1.2.3): hue/saturation/luminance
 * modifiers are applied in HSL space, then shade/tint as an RGB blend toward black/white. Applied
 * in that fixed order since `ColorTransform` (see `packages/presentation`) doesn't preserve the
 * source XML's element order — close enough for the common case of one or two transforms stacked
 * (e.g. a themed "lighter/darker" text colour via lumMod+lumOff), not a bit-exact reimplementation
 * of PowerPoint's own renderer.
 */
function applyColorTransform(rgb: Rgb, transform: ColorTransform | undefined): Rgb {
  if (!transform) return rgb;
  let result = rgb;

  const { hueMod, satMod, lumMod, lumOff, shade, tint } = transform;
  if (
    hueMod !== undefined ||
    satMod !== undefined ||
    lumMod !== undefined ||
    lumOff !== undefined
  ) {
    const [h, s, l] = rgbToHsl(result);
    const h2 = hueMod !== undefined ? (h * (hueMod / 100000)) % 360 : h;
    const s2 = satMod !== undefined ? clamp(s * (satMod / 100000), 0, 1) : s;
    let l2 = lumMod !== undefined ? l * (lumMod / 100000) : l;
    if (lumOff !== undefined) l2 += lumOff / 100000;
    result = hslToRgb([h2, s2, clamp(l2, 0, 1)]);
  }

  if (shade !== undefined) {
    const f = shade / 100000;
    result = result.map((c) => c * f) as unknown as Rgb;
  }
  if (tint !== undefined) {
    const f = tint / 100000;
    result = result.map((c) => c * f + 255 * (1 - f)) as unknown as Rgb;
  }

  return result.map((c) => clamp(c, 0, 255)) as unknown as Rgb;
}

/** Default clrMap (§19.3.1.7 defines the override; this is the identity mapping absent one — slide/layout-level clrMap overrides are unmodeled, matching `packages/presentation`'s scope boundary). */
const SCHEME_ALIASES: Partial<Record<SchemeColorName, keyof ColorScheme>> = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
};

function baseRgb(color: Color, scheme: ColorScheme | undefined): Rgb | undefined {
  switch (color.type) {
    case 'srgb':
      return hexToRgb(color.value);
    case 'system':
      return hexToRgb(color.lastColor);
    case 'hsl':
      return hslToRgb([color.hue / 60000, color.saturation / 100000, color.luminance / 100000]);
    case 'scheme': {
      if (!scheme || color.value === 'phClr') return undefined;
      const slot = SCHEME_ALIASES[color.value] ?? (color.value as keyof ColorScheme);
      const base = scheme[slot];
      return typeof base === 'object' ? baseRgb(base, scheme) : undefined;
    }
    case 'preset':
      // No RGB table for the ~140 DrawingML preset names — they match CSS's extended colour
      // keywords almost 1:1, so `resolveColor` hands the name straight to CSS instead below,
      // meaning a transform stacked on a preset colour is silently ignored. Rare in practice:
      // transforms are almost always applied to scheme colours, not presets.
      return undefined;
  }
}

function rgbToCss(rgb: Rgb, alpha: number): string {
  const [r, g, b] = rgb.map((c) => Math.round(c));
  return alpha < 1 ? `rgba(${r}, ${g}, ${b}, ${alpha})` : `rgb(${r}, ${g}, ${b})`;
}

/** Resolves a DrawingML colour (§20.1.2.3) to a CSS colour string, applying its transforms. */
export function resolveColor(color: Color, scheme: ColorScheme | undefined): string | undefined {
  const rgb = baseRgb(color, scheme);
  if (!rgb) return color.type === 'preset' ? color.value : undefined;
  const transformed = applyColorTransform(rgb, color.transforms);
  const alpha = color.transforms?.alpha !== undefined ? color.transforms.alpha / 100000 : 1;
  return rgbToCss(transformed, alpha);
}

/**
 * Resolves a Fill to a single CSS colour string, for the common case of a solid fill (e.g. a run's
 * text colour). Gradient/pattern/blip fills need more than one CSS colour to render and are
 * unmodeled here — that's shape-fill rendering's job in a later pass, not this one.
 */
export function resolveFillColor(fill: Fill, scheme: ColorScheme | undefined): string | undefined {
  return fill.type === 'solid' ? resolveColor(fill.color, scheme) : undefined;
}
