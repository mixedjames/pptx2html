import type { Presentation } from '@pptx2html/presentation';
import { renderSlide } from './slide.js';

const STYLES = `
  :host {
    display: block;
  }
  .pptx-presentation {
    display: flex;
    flex-direction: column;
    gap: 16px;
  }
`;

/** A `<pptx-presentation>` element: renders a `Presentation` object graph into a shadow DOM. */
export class PptxPresentationElement extends HTMLElement {
  readonly #slidesContainer: HTMLElement;

  constructor() {
    super();
    const shadow = this.attachShadow({ mode: 'open' });

    const style = this.ownerDocument.createElement('style');
    style.textContent = STYLES;
    shadow.appendChild(style);

    this.#slidesContainer = this.ownerDocument.createElement('div');
    this.#slidesContainer.className = 'pptx-presentation';
    shadow.appendChild(this.#slidesContainer);
  }

  render(presentation: Presentation): void {
    this.#slidesContainer.replaceChildren(
      ...presentation.slides.map((slide) =>
        renderSlide(
          this.ownerDocument,
          slide,
          presentation.slideSize,
          presentation.defaultTextStyle,
        ),
      ),
    );
  }
}

export function definePresentationElement(tagName = 'pptx-presentation'): void {
  if (!customElements.get(tagName)) {
    customElements.define(tagName, PptxPresentationElement);
  }
}
