import { readPresentation } from '@pptx2html/reader';
import type { UnsupportedFeature, UnsupportedFeatureCollector } from '@pptx2html/to-html5';
import { renderPresentation } from '@pptx2html/to-html5';

const fileInput = document.getElementById('file-input');
const status = document.getElementById('status');
const output = document.getElementById('output');
const unsupportedPanel = document.getElementById('unsupported-features');

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

  const heading = document.createElement('h2');
  heading.textContent = `Unsupported features (${collector.all.length})`;
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
        const { element, unsupportedFeatures } = renderPresentation(presentation);
        if (unsupportedFeatures.all.length > 0) {
          console.log('Unsupported features:', unsupportedFeatures.all);
          console.log('Unsupported features by slide:', unsupportedFeatures.bySlide);
        }
        renderUnsupportedFeaturesPanel(unsupportedFeatures);
        output?.replaceChildren(element);
        setStatus(`Parsed ${file.name}: ${presentation.slides.length} slide(s).`);
      })
      .catch((error: unknown) => {
        console.error(error);
        setStatus(`Failed to parse ${file.name}: ${String(error)}`);
      });
  });
}
