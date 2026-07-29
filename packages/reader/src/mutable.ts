/**
 * Strips `readonly` for building an object with cyclic references (SlideMaster <-> SlideLayout)
 * where all fields are declared readonly: build with this, assign the back-reference, freeze.
 */
export type Mutable<T> = { -readonly [K in keyof T]: T[K] };
