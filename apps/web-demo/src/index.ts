import { greet } from '@pptx2html/core';

const root = document.getElementById('app');
if (root) {
  root.textContent = greet('pptx2html');
}
