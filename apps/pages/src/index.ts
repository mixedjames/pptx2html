import { readPresentation } from '@pptx2html/reader';
import type { UnsupportedFeature, UnsupportedFeatureCollector } from '@pptx2html/to-html5';
import { renderPresentation } from '@pptx2html/to-html5';

type ViewName = 'chooser' | 'presentation' | 'log';

const chooserView = document.getElementById('view-chooser');
const presentationView = document.getElementById('view-presentation');
const logView = document.getElementById('view-log');
const demoList = document.getElementById('demo-list');
const presentationTitle = document.getElementById('presentation-title');
const status = document.getElementById('status');
const output = document.getElementById('output');
const errorLogBadge = document.getElementById('error-log-badge');
const errorLogContent = document.getElementById('error-log-content');
const backToChooser = document.getElementById('back-to-chooser');
const openErrorLog = document.getElementById('open-error-log');
const backToPresentation = document.getElementById('back-to-presentation');

/** Set once a presentation has been chosen — guards direct/back navigation into the other views. */
let currentFilename: string | undefined;
let currentUnsupportedFeatures: UnsupportedFeatureCollector | undefined;

/** "sliding-panels.pptx" -> "Sliding Panels" — filenames are the only source of a display name. */
function titleFromFilename(filename: string): string {
  return filename
    .replace(/\.pptx$/i, '')
    .replace(/[-_]+/g, ' ')
    .trim()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

// Three full-screen views (chooser / presentation / error log) rather than sections on one long
// page — chosen for mobile, where scrolling past a whole rendered slide deck to reach the error
// log (or the picker) is awkward. Navigation is hash-routed (#presentation / #log, chooser is the
// hash-less default) purely so the device/browser back gesture — the mobile-native way to move
// between screens — does the right thing for free.
function showView(view: ViewName): void {
  if (chooserView) chooserView.hidden = view !== 'chooser';
  if (presentationView) presentationView.hidden = view !== 'presentation';
  if (logView) logView.hidden = view !== 'log';
}

function viewFromHash(hash: string): ViewName {
  if (hash === '#presentation') return 'presentation';
  if (hash === '#log') return 'log';
  return 'chooser';
}

function navigate(view: ViewName): void {
  const hash = view === 'chooser' ? '' : `#${view}`;
  if (location.hash === hash) {
    showView(view);
  } else {
    location.hash = hash;
  }
}

/** Reachable only once a presentation is loaded — falls back to the chooser otherwise (e.g. a
 * reload landing directly on #presentation, or a stale bookmark). */
function syncViewFromHash(): void {
  let view = viewFromHash(location.hash);
  if (view !== 'chooser' && !currentFilename) {
    view = 'chooser';
    history.replaceState(null, '', location.pathname + location.search);
  }
  showView(view);
}

function populateDemoList(list: HTMLUListElement): void {
  if (__DEMO_FILES__.length === 0) {
    const empty = document.createElement('li');
    empty.className = 'empty-state';
    empty.textContent = 'No demo presentations found.';
    list.appendChild(empty);
    return;
  }

  for (const filename of __DEMO_FILES__) {
    const item = document.createElement('li');
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'demo-button';
    button.textContent = titleFromFilename(filename);
    button.addEventListener('click', () => void selectDemo(filename));
    item.appendChild(button);
    list.appendChild(item);
  }
}

function updateErrorBadge(): void {
  if (!(errorLogBadge instanceof HTMLElement)) return;
  const count = currentUnsupportedFeatures?.all.length ?? 0;
  errorLogBadge.textContent = String(count);
  errorLogBadge.hidden = count === 0;
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
function renderErrorLogContent(): void {
  if (!errorLogContent) return;
  errorLogContent.replaceChildren();

  const collector = currentUnsupportedFeatures;
  if (!collector || collector.all.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'empty-state';
    empty.textContent = 'No unsupported features encountered for this presentation.';
    errorLogContent.appendChild(empty);
    return;
  }

  const summary = document.createElement('p');
  summary.textContent = `${collector.all.length} unsupported feature(s) found.`;
  errorLogContent.appendChild(summary);

  const presentationLevel = collector.all.filter((feature) => feature.slideIndex === undefined);
  if (presentationLevel.length > 0) {
    const list = document.createElement('ul');
    for (const feature of presentationLevel) list.appendChild(featureListItem(feature));
    errorLogContent.appendChild(list);
  }

  const bySlide = [...collector.bySlide.entries()].sort(([a], [b]) => a - b);
  for (const [slideIndex, features] of bySlide) {
    const details = document.createElement('details');
    const summaryEl = document.createElement('summary');
    summaryEl.textContent = `Slide ${slideIndex + 1} (${features.length})`;
    details.appendChild(summaryEl);
    const list = document.createElement('ul');
    for (const feature of features) list.appendChild(featureListItem(feature));
    details.appendChild(list);
    errorLogContent.appendChild(details);
  }
}

async function loadDemo(filename: string): Promise<void> {
  setStatus(`Loading ${filename}…`);
  output?.replaceChildren();
  currentUnsupportedFeatures = undefined;
  updateErrorBadge();

  try {
    const response = await fetch(`demos/${encodeURIComponent(filename)}`);
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const buffer = await response.arrayBuffer();
    const presentation = readPresentation(new Uint8Array(buffer));
    const { element, unsupportedFeatures } = renderPresentation(presentation);
    currentUnsupportedFeatures = unsupportedFeatures;
    updateErrorBadge();
    renderErrorLogContent();
    output?.replaceChildren(element);
    setStatus(
      `${presentation.slides.length} slide(s). Tap the presentation or use arrow keys to navigate.`,
    );
  } catch (error) {
    console.error(error);
    setStatus(`Failed to load ${filename}: ${String(error)}`);
  }
}

async function selectDemo(filename: string): Promise<void> {
  currentFilename = filename;
  if (presentationTitle) presentationTitle.textContent = titleFromFilename(filename);
  navigate('presentation');
  await loadDemo(filename);
}

if (demoList instanceof HTMLUListElement) populateDemoList(demoList);
backToChooser?.addEventListener('click', () => navigate('chooser'));
openErrorLog?.addEventListener('click', () => navigate('log'));
backToPresentation?.addEventListener('click', () => navigate('presentation'));
window.addEventListener('hashchange', syncViewFromHash);
syncViewFromHash();
