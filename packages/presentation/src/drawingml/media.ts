/** A binary media part (image, audio, video) embedded in the package, resolved from its relationship. */
export interface MediaPart {
  readonly contentType: string;
  readonly data: Uint8Array;
}
