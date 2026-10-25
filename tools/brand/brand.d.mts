/** Types for brand.mjs (plain ESM so the release tools can run it without a build step). */
export interface Brand {
  id: string;
  dir: string;
  name: string;
  assistantName: string;
  fullName: string;
  /** Optional line under the product name in the task-pane header. */
  tagline?: string;
  /** Name the model calls itself; unset keeps the engine's default persona. */
  promptName?: string;
  /** Writing rules sent with the name on every turn; `{ text, surfaces }` limits one to those apps. */
  promptRules?: Array<
    string | { text: string; surfaces?: string[]; turns?: Array<'chat' | 'planner' | 'command'> }
  >;
  /** Outlook drafts starting with one of `replace` start with `with` instead (before approval). */
  /** The theme's --primary as #RRGGBB (e.g. a PowerPoint table header). */
  accentColor?: string;
  fontStylesheet: string;
  iconsDir: string;
  themePath: string | null;
  assetsDir: string | null;
}
export interface BrandText {
  name: string;
  assistantName: string;
  fullName: string;
  tagline?: string;
  promptName?: string;
  promptRules?: Array<
    string | { text: string; surfaces?: string[]; turns?: Array<'chat' | 'planner' | 'command'> }
  >;
  /** Outlook drafts starting with one of `replace` start with `with` instead (before approval). */
  accentColor?: string;
}
export const brandsRoot: string;
export const BRAND_ICONS: readonly string[];
export function selectedBrand(): string;
export function loadBrand(id?: string): Brand;
export function brandText(brand: Brand): BrandText;
export function primaryColor(css: string): string | null;
export function fillBrandTokens(html: string, brand: Brand): string;
export function brandAssets(brand: Brand): Record<string, string>;
export function contentHash(path: string): string;
