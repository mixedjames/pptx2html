import { readPresentation } from '@pptx2html/reader';
import { renderPresentation } from '@pptx2html/to-html5';

const fileInput = document.getElementById('file-input');
const status = document.getElementById('status');
const output = document.getElementById('output');

function setStatus(message: string): void {
  if (status) status.textContent = message;
}

if (fileInput instanceof HTMLInputElement) {
  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (!file) return;

    setStatus(`Reading ${file.name}...`);
    file
      .arrayBuffer()
      .then((buffer) => {
        const presentation = readPresentation(new Uint8Array(buffer));
        console.log(presentation);
        output?.replaceChildren(renderPresentation(presentation));
        setStatus(`Parsed ${file.name}: ${presentation.slides.length} slide(s).`);
      })
      .catch((error: unknown) => {
        console.error(error);
        setStatus(`Failed to parse ${file.name}: ${String(error)}`);
      });
  });
}
