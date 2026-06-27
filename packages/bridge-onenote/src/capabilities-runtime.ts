/**
 * Runtime capability gating for Office.js requirement sets.
 *
 * The CORRECT runtime gate for an Office.js API/event is
 * `Office.context.requirements.isSetSupported(name, version)` — NOT property truthiness.
 * Members like `worksheets.onChanged` are getters on the Office.js proxy that are *always*
 * truthy even when the host's active requirement set doesn't actually support the event, so
 * `if (sheets.onChanged)` never gates anything (the bug-bash confirmed this) and `.add()` then
 * throws on an older host.
 *
 * This helper is kept pure-ish and unit-testable: it accepts an injectable `requirements` object
 * so tests can pass a fake `{ isSetSupported }` without a live Office runtime. When no object is
 * injected it falls back to the global `Office.context.requirements`, and it returns `false`
 * defensively whenever Office / the requirements bag / the method is absent (e.g. SSR, web-shell
 * harness, or a host that predates the requirements API).
 */

/** The minimal slice of `Office.context.requirements` (`RequirementSetSupport`) we depend on. */
export interface RequirementsLike {
  isSetSupported(name: string, version?: string): boolean;
}

/**
 * True iff the host supports requirement set `name` at >= `version`.
 *
 * @param name    Requirement-set name, e.g. `'ExcelApi'`.
 * @param version Minimum version string, e.g. `'1.9'`.
 * @param requirements Injectable requirements bag; defaults to `Office.context.requirements`.
 */
export function isSet(name: string, version: string, requirements?: RequirementsLike): boolean {
  const reqs = requirements ?? defaultRequirements();
  if (!reqs || typeof reqs.isSetSupported !== 'function') return false;
  try {
    return reqs.isSetSupported(name, version) === true;
  } catch {
    // A throwing host implementation must read as "unsupported", never crash the caller.
    return false;
  }
}

/** Resolve the global `Office.context.requirements`, or `undefined` when Office is absent. */
function defaultRequirements(): RequirementsLike | undefined {
  const office = (globalThis as { Office?: { context?: { requirements?: RequirementsLike } } })
    .Office;
  return office?.context?.requirements;
}
