import { readPresentation } from '@pptx2html/reader';

const fileInput = document.getElementById('file-input');
const status = document.getElementById('status');

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
        setStatus(
          `Parsed ${file.name}: ${presentation.slides.length} slide(s). See the console for the full object tree.`,
        );
      })
      .catch((error: unknown) => {
        console.error(error);
        setStatus(`Failed to parse ${file.name}: ${String(error)}`);
      });
  });
}
