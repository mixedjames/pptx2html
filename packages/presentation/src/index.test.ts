import { describe, expect, it } from 'vitest';
import type { Presentation, Slide, SlideLayout, SlideMaster, Theme } from './index.js';

const theme: Theme = {
  name: 'Office Theme',
  colorScheme: {
    name: 'Office',
    dk1: { type: 'scheme', value: 'dk1' },
    lt1: { type: 'scheme', value: 'lt1' },
    dk2: { type: 'srgb', value: '1F497D' },
    lt2: { type: 'srgb', value: 'EEECE1' },
    accent1: { type: 'srgb', value: '4F81BD' },
    accent2: { type: 'srgb', value: 'C0504D' },
    accent3: { type: 'srgb', value: '9BBB59' },
    accent4: { type: 'srgb', value: '8064A2' },
    accent5: { type: 'srgb', value: '4BACC6' },
    accent6: { type: 'srgb', value: 'F79646' },
    hlink: { type: 'srgb', value: '0000FF' },
    folHlink: { type: 'srgb', value: '800080' },
  },
  fontScheme: {
    name: 'Office',
    majorFont: { latin: 'Calibri Light' },
    minorFont: { latin: 'Calibri' },
  },
  formatScheme: { name: 'Office' },
};

const master: SlideMaster = {
  commonSlideData: { shapeTree: [] },
  theme,
  layouts: [],
};

const layout: SlideLayout = {
  commonSlideData: { shapeTree: [] },
  master,
  type: 'title',
};

const slide: Slide = {
  commonSlideData: {
    shapeTree: [
      {
        kind: 'shape',
        nonVisual: { id: 2, name: 'Title 1' },
        properties: {},
        textBody: {
          paragraphs: [
            {
              runs: [{ kind: 'run', text: 'Hello, presentation!' }],
            },
          ],
        },
      },
    ],
  },
  layout,
};

const presentation: Presentation = {
  slideSize: { width: 12192000, height: 6858000 },
  slideMasters: [master],
  slides: [slide],
  notesSlides: [],
};

describe('presentation DOM skeleton', () => {
  it('assembles a minimal object graph that satisfies the public types', () => {
    expect(presentation.slides).toHaveLength(1);
    expect(presentation.slides[0]?.layout.master.theme.name).toBe('Office Theme');

    const shape = presentation.slides[0]?.commonSlideData.shapeTree[0];
    expect(shape?.kind).toBe('shape');
  });
});
