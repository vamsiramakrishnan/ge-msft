/**
 * Offline manifest lint — a self-contained regression guard over the two add-in manifests
 * (`manifests/m365-unified.manifest.json`, the Teams/Office unified package, and
 * `manifests/onenote.manifest.xml`, the legacy OneNote package).
 *
 * Microsoft recommends validating the manifest in CI (their `office-addin-manifest validate` calls
 * a hosted schema service). We deliberately do NOT depend on that here: the manifests are deploy
 * *templates* — every tenant-specific value is a `REPLACE_*` placeholder — so a strict schema/URL
 * validator both needs network access and would reject the placeholders. Instead this lints the
 * things that actually regress offline and that the hosted validator would never catch anyway:
 *
 *  - **Well-formedness** — the JSON parses; the XML has a balanced root.
 *  - **Required structure** — the top-level keys and the Office `extensions` block we ship.
 *  - **Wiring integrity** — every ribbon / auto-run `actionId` resolves to a declared runtime
 *    action. This is the realistic regression: rename an action, silently break the button.
 *  - **Security invariants** — the Outlook on-send event stays a *reviewable* `softBlock` gate
 *    (CLAUDE.md: the on-send gate), and the OneNote package keeps its scoped permission.
 *  - **Placeholder discipline** — every `REPLACE_*` token is a clean upper-snake placeholder, so a
 *    half-edited value (`REPLACE_web_shell`) can't slip through looking intentional.
 *
 * Pure: the two `lint*` functions take the file *contents* and return findings; the conformance
 * test reads the real files and asserts zero errors (and exercises the negative cases on synthetic
 * broken manifests). Same shape as {@link "./capability-closure"} — one definition, one test.
 */

/** A single lint result. `error` fails the build; `warn` is advisory (surfaced, non-fatal). */
export interface ManifestFinding {
  level: 'error' | 'warn';
  /** Stable machine code, e.g. `unwired-action`, `onsend-not-softblock`. */
  code: string;
  message: string;
}

/** The Office host scopes the unified package must request — one per surface we ship. */
export const REQUIRED_OFFICE_SCOPES = ['mail', 'workbook', 'document', 'presentation'] as const;

/** Office for Mac rejects ribbon groups unless their controls provide all three required sizes. */
export const REQUIRED_RIBBON_ICON_SIZES = [16, 32, 80] as const;

/** A `REPLACE_*` placeholder must be clean upper-snake — catches half-edited tenant values. */
const PLACEHOLDER_RE = /REPLACE_[A-Za-z0-9_]*/g;
const CLEAN_PLACEHOLDER_RE = /^REPLACE_[A-Z0-9][A-Z0-9_]*$/;

function err(code: string, message: string): ManifestFinding {
  return { level: 'error', code, message };
}
function warn(code: string, message: string): ManifestFinding {
  return { level: 'warn', code, message };
}

/** Walk an arbitrary JSON value, collecting every string leaf (for placeholder scanning). */
function collectStrings(value: unknown, out: string[]): void {
  if (typeof value === 'string') out.push(value);
  else if (Array.isArray(value)) for (const v of value) collectStrings(v, out);
  else if (value && typeof value === 'object')
    for (const v of Object.values(value)) collectStrings(v, out);
}

/** Shallow record accessor that never throws on a non-object. */
function obj(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' ? (value as Record<string, unknown>) : {};
}
function arr(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

/**
 * Lint the unified (Teams/Office) JSON manifest. Returns `[]` when clean; any `error`-level finding
 * should fail CI. `path` is only used to label messages.
 */
export function lintUnifiedManifest(
  source: string,
  path = 'm365-unified.manifest.json',
): ManifestFinding[] {
  const findings: ManifestFinding[] = [];

  let root: Record<string, unknown>;
  try {
    root = obj(JSON.parse(source));
  } catch (e) {
    return [err('invalid-json', `${path}: not valid JSON — ${(e as Error).message}`)];
  }

  // 1. Required top-level structure.
  for (const key of ['$schema', 'manifestVersion', 'id', 'name', 'developer', 'validDomains']) {
    if (!(key in root)) findings.push(err('missing-key', `${path}: missing top-level "${key}"`));
  }
  if (!obj(root.name).short) findings.push(err('missing-key', `${path}: missing name.short`));

  // 2. The Office extension block we ship (task pane + commands + ribbon + on-send).
  const ext = obj(arr(root.extensions)[0]);
  if (arr(root.extensions).length === 0) {
    findings.push(err('missing-extension', `${path}: no Office "extensions" block`));
    return findings; // nothing more to check without it.
  }

  // 2a. Requirements: one scope per surface, capabilities carry a name + minVersion.
  const scopes = arr(obj(ext.requirements).scopes).map(String);
  for (const need of REQUIRED_OFFICE_SCOPES) {
    if (!scopes.includes(need))
      findings.push(warn('missing-scope', `${path}: requirements.scopes is missing "${need}"`));
  }
  for (const cap of arr(obj(ext.requirements).capabilities)) {
    const c = obj(cap);
    if (!c.name || !c.minVersion)
      findings.push(
        err('bad-capability', `${path}: a requirements.capabilities entry lacks name/minVersion`),
      );
  }

  // 2b. Collect every declared runtime action id — the link targets for ribbons / auto-run.
  const actionIds = new Set<string>();
  for (const rt of arr(ext.runtimes))
    for (const a of arr(obj(rt).actions)) {
      const id = obj(a).id;
      if (typeof id === 'string') actionIds.add(id);
    }
  if (actionIds.size === 0)
    findings.push(err('no-runtime-actions', `${path}: no runtime actions declared`));

  // 2c. Wiring integrity: every ribbon control actionId resolves to a runtime action.
  for (const ribbon of arr(ext.ribbons))
    for (const tab of arr(obj(ribbon).tabs))
      for (const group of arr(obj(tab).groups)) {
        const g = obj(group);
        const groupIconSizes = new Set(arr(g.icons).map((icon) => Number(obj(icon).size)));
        for (const size of REQUIRED_RIBBON_ICON_SIZES) {
          if (!groupIconSizes.has(size)) {
            findings.push(
              err(
                'missing-ribbon-group-icon',
                `${path}: ribbon group "${String(g.id)}" is missing required ${size}px icon`,
              ),
            );
          }
        }
        for (const control of arr(obj(group).controls)) {
          const c = obj(control);
          const actionId = c.actionId;
          if (typeof actionId === 'string' && !actionIds.has(actionId))
            findings.push(
              err(
                'unwired-action',
                `${path}: ribbon control "${String(obj(control).id)}" references unknown actionId "${actionId}"`,
              ),
            );
          const iconSizes = new Set(arr(c.icons).map((icon) => Number(obj(icon).size)));
          for (const size of REQUIRED_RIBBON_ICON_SIZES) {
            if (!iconSizes.has(size)) {
              findings.push(
                err(
                  'missing-ribbon-icon',
                  `${path}: ribbon control "${String(c.id)}" is missing required ${size}px icon`,
                ),
              );
            }
          }
        }
      }

  // 2d. The Outlook on-send gate: the event must resolve to an action AND stay a reviewable
  //     softBlock (CLAUDE.md — reversible/reviewable writes; never a silent send).
  for (const evGroup of arr(ext.autoRunEvents))
    for (const ev of arr(obj(evGroup).events)) {
      const e = obj(ev);
      if (typeof e.actionId === 'string' && !actionIds.has(e.actionId))
        findings.push(
          err(
            'unwired-action',
            `${path}: autoRunEvent references unknown actionId "${e.actionId}"`,
          ),
        );
      if (e.type === 'messageSending') {
        const sendMode = obj(e.options).sendMode;
        if (sendMode !== 'softBlock')
          findings.push(
            err(
              'onsend-not-softblock',
              `${path}: messageSending event sendMode is "${String(sendMode)}", must be "softBlock" (the reviewable on-send gate)`,
            ),
          );
      }
    }

  // 3. Placeholder discipline: every REPLACE_* token is clean upper-snake.
  findings.push(...lintPlaceholders(source, path, JSON.parse(source)));

  return findings;
}

/** Scan a manifest's string leaves for malformed `REPLACE_*` placeholders. */
function lintPlaceholders(source: string, path: string, parsed?: unknown): ManifestFinding[] {
  const out: ManifestFinding[] = [];
  // Prefer string leaves when we have a parsed tree; fall back to a raw scan (XML).
  const strings: string[] = [];
  if (parsed !== undefined) collectStrings(parsed, strings);
  else strings.push(source);
  const seen = new Set<string>();
  for (const s of strings)
    for (const token of s.match(PLACEHOLDER_RE) ?? []) {
      if (seen.has(token)) continue;
      seen.add(token);
      if (!CLEAN_PLACEHOLDER_RE.test(token))
        out.push(
          warn('malformed-placeholder', `${path}: placeholder "${token}" is not clean UPPER_SNAKE`),
        );
    }
  return out;
}

/**
 * Lint the OneNote XML manifest. No XML dependency: we assert well-formedness (a single balanced
 * `<OfficeApp>` root) plus the security-relevant elements by targeted match. `path` labels messages.
 */
export function lintOneNoteManifest(
  source: string,
  path = 'onenote.manifest.xml',
): ManifestFinding[] {
  const findings: ManifestFinding[] = [];
  const text = source.trim();

  if (!text) return [err('empty', `${path}: file is empty`)];

  // Well-formedness (cheap): exactly one OfficeApp open + close, declared as a task-pane app.
  const opens = (text.match(/<OfficeApp[\s>]/g) ?? []).length;
  const closes = (text.match(/<\/OfficeApp>/g) ?? []).length;
  if (opens !== 1 || closes !== 1)
    findings.push(err('bad-root', `${path}: expected exactly one <OfficeApp> root`));
  if (!/xsi:type="TaskPaneApp"/.test(text))
    findings.push(err('not-taskpane', `${path}: OfficeApp is not a TaskPaneApp`));

  // Required structure for the OneNote (Notebook) host.
  if (!/<Host\s+Name="Notebook"\s*\/?>/.test(text))
    findings.push(err('missing-host', `${path}: missing <Host Name="Notebook">`));
  if (!/<Id>[^<]+<\/Id>/.test(text)) findings.push(err('missing-id', `${path}: missing <Id>`));
  if (!/<SourceLocation\b/.test(text))
    findings.push(err('missing-source', `${path}: missing a <SourceLocation>`));

  // Security: the OneNote package writes synthesized pages, so ReadWriteDocument is expected — but
  // it must be *exactly* that scoped permission, not a broader one.
  const perm = text.match(/<Permissions>([^<]*)<\/Permissions>/)?.[1];
  if (perm !== 'ReadWriteDocument')
    findings.push(
      err(
        'bad-permission',
        `${path}: <Permissions> is "${String(perm)}", expected ReadWriteDocument`,
      ),
    );

  // Every declared AppDomain must be https (no plaintext origin in the trust list).
  for (const m of text.matchAll(/<AppDomain>([^<]*)<\/AppDomain>/g)) {
    const d = m[1] ?? '';
    if (!/^https:\/\//.test(d))
      findings.push(err('insecure-domain', `${path}: AppDomain "${d}" is not https`));
  }

  findings.push(...lintPlaceholders(text, path));
  return findings;
}
