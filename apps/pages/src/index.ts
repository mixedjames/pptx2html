import { readPresentation } from '@pptx2html/reader';
import type { UnsupportedFeature, UnsupportedFeatureCollector } from '@pptx2html/to-html5';
import { renderPresentation } from '@pptx2html/to-html5';

const demoSelect = document.getElementById('demo-select');
const status = document.getElementById('status');
const output = document.getElementById('output');
const errorLog = document.getElementById('error-log');

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

/** "sliding-panels.pptx" -> "Sliding Panels" — filenames are the only source of a display name. */
function titleFromFilename(filename: string): string {
  return filename
    .replace(/\.pptx$/i, '')
    .replace(/[-_]+/g, ' ')
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function populateDemoOptions(select: HTMLSelectElement): void {
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent =
    __DEMO_FILES__.length > 0 ? 'Choose a presentation…' : 'No demo presentations found';
  placeholder.disabled = true;
  placeholder.selected = true;
  select.appendChild(placeholder);

  for (const filename of __DEMO_FILES__) {
    const option = document.createElement('option');
    option.value = filename;
    option.textContent = titleFromFilename(filename);
    select.appendChild(option);
  }
}

function featureListItem(feature: UnsupportedFeature): HTMLLIElement {
  const item = document.createElement('li');
  const shapeLabel = feature.shape
    ? ` — shape "${feature.shape.name}" (id ${feature.shape.id})`
    : '';
  item.textContent = `${feature.message}${shapeLabel}`;
  return item;
}

/** Presentation-level entries (no slideIndex) first, then one collapsible section per slide. */
function renderErrorLog(collector: UnsupportedFeatureCollector): void {
  if (!errorLog) return;
  errorLog.replaceChildren();
  if (collector.all.length === 0) {
    errorLog.hidden = true;
    return;
  }
  errorLog.hidden = false;

  const heading = document.createElement('h2');
  heading.textContent = `Error log (${collector.all.length})`;
  errorLog.appendChild(heading);

  const presentationLevel = collector.all.filter((feature) => feature.slideIndex === undefined);
  if (presentationLevel.length > 0) {
    const list = document.createElement('ul');
    for (const feature of presentationLevel) list.appendChild(featureListItem(feature));
    errorLog.appendChild(list);
  }

  const bySlide = [...collector.bySlide.entries()].sort(([a], [b]) => a - b);
  for (const [slideIndex, features] of bySlide) {
    const details = document.createElement('details');
    const summary = document.createElement('summary');
    summary.textContent = `Slide ${slideIndex + 1} (${features.length})`;
    details.appendChild(summary);
    const list = document.createElement('ul');
    for (const feature of features) list.appendChild(featureListItem(feature));
    details.appendChild(list);
    errorLog.appendChild(details);
  }
}

async function loadDemo(filename: string): Promise<void> {
  setStatus(`Loading ${filename}…`);
  if (errorLog) errorLog.hidden = true;
  output?.replaceChildren();

  try {
    const response = await fetch(`demos/${encodeURIComponent(filename)}`);
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const buffer = await response.arrayBuffer();
    const presentation = readPresentation(new Uint8Array(buffer));
    const { element, unsupportedFeatures } = renderPresentation(presentation);
    renderErrorLog(unsupportedFeatures);
    output?.replaceChildren(element);
    setStatus(
      `Loaded ${filename}: ${presentation.slides.length} slide(s). Click the presentation or use arrow keys to navigate.`,
    );
  } catch (error) {
    console.error(error);
    setStatus(`Failed to load ${filename}: ${String(error)}`);
  }
}

if (demoSelect instanceof HTMLSelectElement) {
  populateDemoOptions(demoSelect);
  demoSelect.addEventListener('change', () => {
    const filename = demoSelect.value;
    if (filename) void loadDemo(filename);
  });
}
