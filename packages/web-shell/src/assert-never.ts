/** Exhaustiveness guard: a compile error if a discriminated-union variant is left unhandled. */
export function assertNever(x: never): never {
  throw new Error(`unexpected variant: ${JSON.stringify(x)}`);
}
