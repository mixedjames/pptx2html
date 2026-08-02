/**
 * A shape/picture/connector/graphicFrame's identity (§19.3.1.12's `p:cNvPr` `id`/`name`, the same
 * pair PowerPoint's own UI shows in the Selection Pane) — enough for a user to locate the specific
 * object an `UnsupportedFeature` refers to inside the source deck.
 */
export interface UnsupportedFeatureShapeRef {
  readonly id: number;
  readonly name: string;
}

/** One thing present in a `.pptx` that this renderer doesn't (fully) render. */
export interface UnsupportedFeature {
  /** Stable machine-readable id, e.g. `'shape-geometry-unmodeled'`. */
  readonly code: string;
  /** Human-readable explanation, including the specific value where relevant. */
  readonly message: string;
  /** Absent for a presentation-level feature (nothing tying it to one slide). */
  readonly slideIndex?: number;
  /** The specific shape/picture/etc. this occurred on, when there is one. */
  readonly shape?: UnsupportedFeatureShapeRef;
}

/** Collects `UnsupportedFeature`s reported while rendering one `Presentation`. */
export class UnsupportedFeatureCollector {
  readonly #features: UnsupportedFeature[] = [];

  report(feature: UnsupportedFeature): void {
    this.#features.push(feature);
  }

  get all(): readonly UnsupportedFeature[] {
    return this.#features;
  }

  /** Every reported feature that has a `slideIndex`, grouped under it. */
  get bySlide(): ReadonlyMap<number, UnsupportedFeature[]> {
    const grouped = new Map<number, UnsupportedFeature[]>();
    for (const feature of this.#features) {
      if (feature.slideIndex === undefined) continue;
      const existing = grouped.get(feature.slideIndex);
      if (existing) existing.push(feature);
      else grouped.set(feature.slideIndex, [feature]);
    }
    return grouped;
  }
}
