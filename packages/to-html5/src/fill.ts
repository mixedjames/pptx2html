import type { ColorScheme, Emu, Fill, GradientFill, Line } from '@pptx2html/presentation';

import { resolveColor, resolveFillColor } from './color.js';
import { emuToCqw } from './units.js';

/** DrawingML's shade-path angle (§20.1.8.41, `<a:lin ang="...">`, 60,000ths of a degree, clockwise
 * from 3 o'clock/east — 0 means "left edge color to right edge color") is measured from a
 * different zero point than CSS's `linear-gradient()` angle (clockwise from 12 o'clock/north). */
function toCssGradientAngle(ooxmlAngle: number | undefined): number {
  const degrees = (ooxmlAngle ?? 0) / 60000;
  return (degrees + 90) % 360;
}

/** Resolves a linear gradient fill to a CSS `linear-gradient(...)` value. */
export function resolveGradientCss(
  fill: GradientFill,
  scheme: ColorScheme | undefined,
): string | undefined {
  if (fill.stops.length === 0) return undefined;
  const stops = [...fill.stops]
    .sort((a, b) => a.position - b.position)
    .map((stop) => {
      const color = resolveColor(stop.color, scheme);
      return color ? `${color} ${stop.position / 1000}%` : undefined;
    });
  if (stops.some((stop) => stop === undefined)) return undefined;
  return `linear-gradient(${toCssGradientAngle(fill.angle)}deg, ${stops.join(', ')})`;
}

/**
 * Applies a shape/picture's fill (§20.1.2.2.35, spPr's fill) as a CSS background. Solid and
 * linear-gradient fills map directly; a pattern fill is only approximated (DrawingML has ~48
 * named presets — pct10, dashDnDiag, diagCross, etc. — with no CSS equivalent, so this just
 * overlays a translucent diagonal hatch on the background color as a visual hint that a pattern
 * was there, not a faithful reproduction of any specific preset). A blip (image) fill creates an
 * object URL the same way `renderPicture` does for a picture's own image, with the same
 * never-revoked caveat (see `packages/to-html5/CLAUDE.md`).
 */
export function applyFill(
  el: HTMLElement,
  fill: Fill | undefined,
  scheme: ColorScheme | undefined,
): void {
  if (!fill) return;
  switch (fill.type) {
    case 'none':
      return;
    case 'solid': {
      const color = resolveColor(fill.color, scheme);
      if (color) el.style.backgroundColor = color;
      return;
    }
    case 'gradient': {
      const css = resolveGradientCss(fill, scheme);
      if (css) el.style.backgroundImage = css;
      return;
    }
    case 'pattern': {
      const background = resolveColor(fill.backgroundColor, scheme);
      const foreground = resolveColor(fill.foregroundColor, scheme);
      if (background) el.style.backgroundColor = background;
      if (foreground) {
        el.style.backgroundImage = `repeating-linear-gradient(45deg, ${foreground}, ${foreground} 2px, transparent 2px, transparent 8px)`;
      }
      return;
    }
    case 'blip': {
      const blob = new Blob([new Uint8Array(fill.image.data)], { type: fill.image.contentType });
      el.style.backgroundImage = `url(${URL.createObjectURL(blob)})`;
      if (fill.tile) {
        el.style.backgroundRepeat = 'repeat';
      } else {
        el.style.backgroundRepeat = 'no-repeat';
        el.style.backgroundSize = fill.stretch ? '100% 100%' : 'cover';
      }
      return;
    }
  }
}

/** Approximates a DashStyle as one of CSS's three border-style keywords (no CSS equivalent for
 * DrawingML's finer-grained dash/dot spacing presets, e.g. lgDash vs sysDash). */
function borderStyleFor(line: Line): string {
  if (line.compound === 'double') return 'double';
  switch (line.dashStyle) {
    case undefined:
    case 'solid':
      return 'solid';
    case 'dot':
    case 'sysDot':
      return 'dotted';
    default:
      return 'dashed';
  }
}

/** DrawingML's default outline weight (§20.1.2.2.24, a:ln) when `w` is present but unset by the
 * shape's own style — a commonly used default, not a value the spec itself mandates. */
const DEFAULT_LINE_WIDTH_EMU = 12700;

/**
 * Applies a shape/picture's outline (§20.1.2.2.24, spPr's ln) as a CSS border. Only a solid line
 * fill resolves to an explicit `border-color` — a gradient/pattern/blip-filled outline is rare
 * enough on a plain shape border that it's left to CSS's own default (`currentColor`) rather than
 * picking an arbitrary single color out of it. `slideWidth` scales the border's own weight with
 * the slide (see `units.ts`'s `emuToCqw`), the same way font sizes do — a border expressed in a
 * fixed px would look disproportionately thick/thin as the slide is resized.
 */
export function applyLine(
  el: HTMLElement,
  line: Line | undefined,
  scheme: ColorScheme | undefined,
  slideWidth: Emu,
): void {
  if (!line || line.fill?.type === 'none') return;
  el.style.borderWidth = emuToCqw(line.width ?? DEFAULT_LINE_WIDTH_EMU, slideWidth);
  el.style.borderStyle = borderStyleFor(line);
  const color = line.fill ? resolveFillColor(line.fill, scheme) : undefined;
  if (color) el.style.borderColor = color;
}
