/**
 * The ONE place the Office-host simulators reach the global object model the real bridges drive.
 *
 * The bridges are the only Office.js code in the repo and they enter the host through the GLOBAL
 * `Excel.run` / `Word.run` / `PowerPoint.run` and `Office.context.*`. To drive the REAL bridge
 * unchanged against seeded data, each `installFake*` writes an in-memory fake onto `globalThis`
 * under the host's global name (`Excel`/`Word`/`PowerPoint`/`Office`) and `restore()` removes it.
 *
 * The `@types/office-js` global declarations are *much* wider than the slice the bridges touch
 * (see each simulator's header for the exact enumerated call set). Rather than implement the full
 * typings, the fakes model only that slice and are installed through the single narrow cast in
 * {@link installGlobal} — so the `unknown`-cast at the global boundary lives in exactly one place
 * and nothing downstream needs `any`.
 */

/** The Office-host global names a simulator may install. */
export type HostGlobalName = 'Excel' | 'Word' | 'PowerPoint' | 'Office';

/**
 * `globalThis` viewed as a plain mutable record. `@types/office-js` declares `Excel`/`Word`/… as
 * precise `const` namespaces, so the ONLY way to install an in-memory fake under those names is to
 * step outside those declarations. We do that with a single `unknown` cast, isolated here, so no
 * other file in the harness needs a cast or `any`.
 */
function hostGlobals(): Record<string, unknown> {
  return globalThis as unknown as Record<string, unknown>;
}

/**
 * Install `value` as the global host namespace `name`, returning a `restore()` that puts back
 * whatever (if anything) was there before. The single `unknown` cast at the Office global boundary
 * is isolated here; callers pass their richly-typed fake and never cast themselves.
 */
export function installGlobal(name: HostGlobalName, value: object): () => void {
  const g = hostGlobals();
  const had = Object.prototype.hasOwnProperty.call(g, name);
  const prev = g[name];
  g[name] = value;
  return () => {
    if (had) g[name] = prev;
    else delete g[name];
  };
}

/** Compose several `restore()` thunks into one idempotent teardown (last-installed-first). */
export function composeRestores(restores: ReadonlyArray<() => void>): () => void {
  let done = false;
  return () => {
    if (done) return;
    done = true;
    for (const r of [...restores].reverse()) r();
  };
}
