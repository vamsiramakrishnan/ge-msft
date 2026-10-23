/**
 * Brand loader shared by the web build (packages/web-shell/brand-assets.ts) and the manifest tools
 * (tools/release). A brand is a folder, brands/<id>/, holding:
 *
 *   brand.json   the names users see and the font stylesheet (required)
 *   icons/       the six add-in icons Office and the manifests reference (required)
 *   theme.css    CSS overrides loaded after the task-pane styles (optional)
 *   assets/      extra files served at /brand/<file>, e.g. a logo theme.css points at (optional)
 *
 * GE_BRAND picks the folder, then brands/current, then brands/default/ (the upstream look). Code identifiers
 * (openGemini, Gemini.* resids, GeminiPane) are not brand text and never change.
 */
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

export const brandsRoot = new URL('../../brands', import.meta.url).pathname;

/** The icons every brand ships; the manifests reference them at the site root. */
export const BRAND_ICONS = [
  'icon-16.png',
  'icon-32.png',
  'icon-64.png',
  'icon-80.png',
  'icon-color.png',
  'icon-outline.png',
];

const TEXT_FIELDS = ['name', 'assistantName', 'fullName', 'fontStylesheet'];

/** HTML placeholder → brand field. An unknown %GE_BRAND_*% token fails the build. */
const HTML_TOKENS = {
  GE_BRAND_NAME: 'name',
  GE_BRAND_ASSISTANT: 'assistantName',
  GE_BRAND_FULL_NAME: 'fullName',
  GE_BRAND_FONT_STYLESHEET: 'fontStylesheet',
};

/**
 * The brand to build: GE_BRAND if set, else the id in brands/current (a repository can pin its own
 * brand there, so nobody has to type it), else `default`.
 */
export function selectedBrand() {
  if (process.env.GE_BRAND) return process.env.GE_BRAND;
  const pinned = join(brandsRoot, 'current');
  return existsSync(pinned) ? readFileSync(pinned, 'utf8').trim() || 'default' : 'default';
}

export function loadBrand(id = selectedBrand()) {
  if (!/^[a-z0-9][a-z0-9-]*$/.test(id)) {
    throw new Error(`GE_BRAND "${id}" must be lowercase letters, digits and dashes.`);
  }
  const dir = join(brandsRoot, id);
  const file = join(dir, 'brand.json');
  if (!existsSync(file)) throw new Error(`Brand "${id}" not found: ${file}`);
  const raw = JSON.parse(readFileSync(file, 'utf8'));
  for (const key of TEXT_FIELDS) {
    if (typeof raw[key] !== 'string' || !raw[key].trim()) {
      throw new Error(`brands/${id}/brand.json: "${key}" must be a non-empty string.`);
    }
  }
  if (raw.tagline !== undefined && (typeof raw.tagline !== 'string' || !raw.tagline.trim())) {
    throw new Error(`brands/${id}/brand.json: "tagline" must be a non-empty string when set.`);
  }
  if (
    raw.promptName !== undefined &&
    (typeof raw.promptName !== 'string' || !raw.promptName.trim())
  ) {
    throw new Error(`brands/${id}/brand.json: "promptName" must be a non-empty string when set.`);
  }
  const list = (v, ok) => v === undefined || (Array.isArray(v) && v.length > 0 && v.every(ok));
  const validRule = (rule) =>
    (typeof rule === 'string' && rule.trim()) ||
    (rule &&
      typeof rule === 'object' &&
      typeof rule.text === 'string' &&
      rule.text.trim() &&
      list(rule.surfaces, (s) => typeof s === 'string' && s.trim()) &&
      list(rule.turns, (t) => ['chat', 'planner', 'command'].includes(t)));
  if (
    raw.promptRules !== undefined &&
    (!Array.isArray(raw.promptRules) || !raw.promptRules.every(validRule))
  ) {
    throw new Error(
      `brands/${id}/brand.json: "promptRules" must list rules: text, or { "text", "surfaces"?: [...], "turns"?: ["chat" | "planner" | "command"] }.`,
    );
  }
  if (
    !raw.fontStylesheet.startsWith('https://fonts.googleapis.com/') ||
    /["'<>\s]/.test(raw.fontStylesheet)
  ) {
    throw new Error(`brands/${id}/brand.json: "fontStylesheet" must be a Google Fonts URL.`);
  }
  const iconsDir = join(dir, 'icons');
  const missing = BRAND_ICONS.filter((icon) => !existsSync(join(iconsDir, icon)));
  if (missing.length) throw new Error(`brands/${id}/icons is missing ${missing.join(', ')}.`);
  const themePath = join(dir, 'theme.css');
  const assetsDir = join(dir, 'assets');
  const accentColor = existsSync(themePath) ? primaryColor(readFileSync(themePath, 'utf8')) : null;
  return {
    id,
    dir,
    name: raw.name,
    assistantName: raw.assistantName,
    fullName: raw.fullName,
    ...(raw.tagline ? { tagline: raw.tagline } : {}),
    ...(raw.promptName ? { promptName: raw.promptName } : {}),
    ...(raw.promptRules?.length ? { promptRules: raw.promptRules } : {}),
    ...(accentColor ? { accentColor } : {}),
    fontStylesheet: raw.fontStylesheet,
    iconsDir,
    themePath: existsSync(themePath) ? themePath : null,
    assetsDir: existsSync(assetsDir) ? assetsDir : null,
  };
}

/** The part of the brand the task-pane code reads at runtime (see packages/web-shell/src/brand.ts). */
export function brandText(brand) {
  return {
    name: brand.name,
    assistantName: brand.assistantName,
    fullName: brand.fullName,
    ...(brand.tagline ? { tagline: brand.tagline } : {}),
    ...(brand.promptName ? { promptName: brand.promptName } : {}),
    ...(brand.promptRules?.length ? { promptRules: brand.promptRules } : {}),
    ...(brand.accentColor ? { accentColor: brand.accentColor } : {}),
  };
}

/**
 * The theme's `--primary` as #RRGGBB, following one `var(--x)` step (e.g. var(--brand-red) →
 * #c8102e). Documents use it where a brand colour is wanted, e.g. a PowerPoint table header.
 */
export function primaryColor(css) {
  const value = (name) => new RegExp(`--${name}\\s*:\\s*([^;]+);`).exec(css)?.[1]?.trim();
  let v = value('primary');
  const ref = v && /^var\(--([\w-]+)\)$/.exec(v);
  if (ref) v = value(ref[1]);
  return v && /^#[0-9a-fA-F]{6}$/.test(v) ? v.toUpperCase() : null;
}

/** Replace %GE_BRAND_*% placeholders in an HTML page with the brand's (HTML-escaped) values. */
export function fillBrandTokens(html, brand) {
  return html.replace(/%(GE_BRAND_[A-Z_]+)%/g, (token, key) => {
    const field = HTML_TOKENS[key];
    if (!field) throw new Error(`Unknown brand placeholder ${token}.`);
    // The stylesheet URL is validated above and goes in verbatim, so its `&`s stay as written.
    return field === 'fontStylesheet' ? brand[field] : escapeHtml(brand[field]);
  });
}

/** Files under brands/<id>/assets/, as { published path: absolute path }. */
export function brandAssets(brand) {
  if (!brand.assetsDir) return {};
  return Object.fromEntries(
    readdirSync(brand.assetsDir).map((file) => [`brand/${file}`, join(brand.assetsDir, file)]),
  );
}

/** Short content hash, used as a cache-busting query on the theme stylesheet. */
export function contentHash(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex').slice(0, 10);
}

function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
