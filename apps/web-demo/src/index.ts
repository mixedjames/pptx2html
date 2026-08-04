import { readPresentation } from '@pptx2html/reader';
import type { Presentation } from '@pptx2html/presentation';
import type { UnsupportedFeature, UnsupportedFeatureCollector } from '@pptx2html/to-html5';
import { renderPresentation, renderScrollPresentation } from '@pptx2html/to-html5';

type Mode = 'click' | 'scroll';

const fileInput = document.getElementById('file-input');
const status = document.getElementById('status');
const output = document.getElementById('output');
const unsupportedPanel = document.getElementById('unsupported-features');
const stage = document.getElementById('stage');
const fullscreenButton = document.getElementById('fullscreen-btn');
const modeInputs = [
  ...document.querySelectorAll<HTMLInputElement>('#mode-fieldset input[name="mode"]'),
];

let currentPresentation: Presentation | undefined;
let currentFileName = '';

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

function featureListItem(feature: UnsupportedFeature): HTMLLIElement {
  const item = document.createElement('li');
  const shapeLabel = feature.shape
    ? ` — shape "${feature.shape.name}" (id ${feature.shape.id})`
    : '';
  item.textContent = `${feature.message}${shapeLabel}`;
  return item;
}

/** Presentation-level features (no slideIndex) first, then one collapsible section per slide. */
function renderUnsupportedFeaturesPanel(collector: UnsupportedFeatureCollector): void {
  if (!unsupportedPanel) return;
  unsupportedPanel.replaceChildren();
  if (collector.all.length === 0) {
    unsupportedPanel.hidden = true;
    return;
  }
  unsupportedPanel.hidden = false;

  const heading = document.createElement('div');
  heading.className = 'heading';
  const headingText = document.createElement('h2');
  headingText.textContent = `Unsupported features (${collector.all.length})`;
  const closeButton = document.createElement('button');
  closeButton.type = 'button';
  closeButton.className = 'close-button';
  closeButton.setAttribute('aria-label', 'Hide unsupported features panel');
  closeButton.textContent = '×';
  closeButton.addEventListener('click', () => {
    unsupportedPanel.hidden = true;
  });
  heading.append(headingText, closeButton);
  unsupportedPanel.appendChild(heading);

  const presentationLevel = collector.all.filter((feature) => feature.slideIndex === undefined);
  if (presentationLevel.length > 0) {
    const list = document.createElement('ul');
    for (const feature of presentationLevel) list.appendChild(featureListItem(feature));
    unsupportedPanel.appendChild(list);
  }

  const bySlide = [...collector.bySlide.entries()].sort(([a], [b]) => a - b);
  for (const [slideIndex, features] of bySlide) {
    const details = document.createElement('details');
    details.open = true;
    const summary = document.createElement('summary');
    summary.textContent = `Slide ${slideIndex + 1} (${features.length})`;
    details.appendChild(summary);
    const list = document.createElement('ul');
    for (const feature of features) list.appendChild(featureListItem(feature));
    details.appendChild(list);
    unsupportedPanel.appendChild(details);
  }
}

function currentMode(): Mode {
  return modeInputs.find((input) => input.checked)?.value === 'scroll' ? 'scroll' : 'click';
}

/** (Re-)renders `currentPresentation` in whichever mode is currently selected — the one place
 *  either a freshly-parsed file or a mode toggle ends up, so switching modes never needs a
 *  re-parse. `#output` always fills `#stage` (index.html) regardless of mode, which is what
 *  `renderScrollPresentation`'s `<pptx-scroll-presentation>` needs (an explicit box — see its own
 *  doc comment) and is harmless for `renderPresentation`'s own aspect-ratio-locked sizing. */
function renderCurrent(): void {
  if (!currentPresentation) return;
  const mode = currentMode();

  const { element, unsupportedFeatures } =
    mode === 'scroll'
      ? renderScrollPresentation(currentPresentation)
      : renderPresentation(currentPresentation);
  if (unsupportedFeatures.all.length > 0) {
    console.log('Unsupported features:', unsupportedFeatures.all);
    console.log('Unsupported features by slide:', unsupportedFeatures.bySlide);
  }
  renderUnsupportedFeaturesPanel(unsupportedFeatures);
  output?.replaceChildren(element);
  setStatus(
    `Parsed ${currentFileName}: ${currentPresentation.slides.length} slide(s) (${mode}-driven).`,
  );
}

if (fileInput instanceof HTMLInputElement) {
  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (!file) return;

    setStatus(`Reading ${file.name}...`);
    if (unsupportedPanel) unsupportedPanel.hidden = true;
    file
      .arrayBuffer()
      .then((buffer) => {
        const presentation = readPresentation(new Uint8Array(buffer));
        console.log(presentation);
        presentation.slides.forEach((slide, index) => {
          if (slide.timing) console.log(`Slide ${index + 1} timing:`, slide.timing);
        });
        currentPresentation = presentation;
        currentFileName = file.name;
        renderCurrent();
      })
      .catch((error: unknown) => {
        console.error(error);
        setStatus(`Failed to parse ${file.name}: ${String(error)}`);
      });
  });
}

for (const input of modeInputs) {
  input.addEventListener('change', renderCurrent);
}

// Fullscreens #stage (the presentation plus its own chrome overlay), not just #output — the
// overlay's own controls (including this button) stay reachable while fullscreen instead of
// disappearing along with the rest of the page, which is what fullscreening #output alone would
// do (the Fullscreen API hides everything outside the fullscreened element).
if (fullscreenButton && stage) {
  fullscreenButton.addEventListener('click', () => {
    if (document.fullscreenElement) {
      void document.exitFullscreen();
    } else {
      void stage.requestFullscreen();
    }
  });
  document.addEventListener('fullscreenchange', () => {
    const isFullscreen = document.fullscreenElement === stage;
    fullscreenButton.textContent = isFullscreen ? 'Exit fullscreen' : 'Fullscreen';
  });
}
