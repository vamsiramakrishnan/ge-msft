/**
 * The product names users see in the task pane. The build fills these from
 * brands/<GE_BRAND>/brand.json (see brand-assets.ts); without a build-time brand, as in unit tests,
 * the upstream default applies. Code identifiers (openGemini, GeminiPane) are not brand text.
 */
export interface Brand {
  /** Product name, e.g. in the header and error titles. */
  readonly name: string;
  /** How the assistant is addressed, e.g. "Ask Gemini". */
  readonly assistantName: string;
  /** Long form, e.g. page titles. */
  readonly fullName: string;
  /** Optional line under the name in the header, e.g. "Contoso AI Assistant". */
  readonly tagline?: string;
  /** Name the model calls itself in answers (e.g. "ContosoAI"); unset keeps the engine's default. */
  readonly promptName?: string;
  /**
   * Writing rules sent with the name; unset sends none. A plain rule applies on every turn in every
   * app; `{ text, surfaces, turns }` only in those Office apps and on those turns (e.g. email
   * wording only in Outlook, only when the assistant writes the email).
   */
  readonly promptRules?: readonly PromptRule[];
  /**
   * Outlook drafts that START with one of `replace` start with `with` instead (e.g. "Halo", not
   * "Yth."). Applied before the approval card; only the opening words change.
   */
  /** The brand's main colour (#RRGGBB, the theme's --primary): a new PowerPoint table's header row. */
  readonly accentColor?: string;
}

/** The turns a rule can be limited to: answering, planning, or writing into the document/email. */
export type PromptTurn = 'chat' | 'planner' | 'command';

/**
 * A brand writing rule: everywhere, or only in the named Office apps (`surfaces`) and/or only on
 * the named turns (`turns`; `command` = turns where the assistant writes into the document or email).
 */
export type PromptRule =
  | string
  | {
      readonly text: string;
      readonly surfaces?: readonly string[];
      readonly turns?: readonly PromptTurn[];
    };

const ROUTE: Readonly<Record<PromptTurn, 'default' | 'planner' | 'command'>> = {
  chat: 'default',
  planner: 'planner',
  command: 'command',
};

/** A rule for the client: text on every turn, or text limited to some turn routes. */
export type ClientRule =
  | string
  | { text: string; routes: Array<'default' | 'planner' | 'command'> };

/** The rules that apply in `surface`, each limited to its turns for the client. */
export function rulesFor(rules: readonly PromptRule[] | undefined, surface: string): ClientRule[] {
  return (rules ?? []).flatMap((rule): ClientRule[] => {
    if (typeof rule === 'string') return [rule];
    if (rule.surfaces && !rule.surfaces.includes(surface)) return [];
    return rule.turns
      ? [{ text: rule.text, routes: rule.turns.map((t) => ROUTE[t]) }]
      : [rule.text];
  });
}

declare const __GE_BRAND__: Brand | undefined;

/** Mirrors brands/default/brand.json (brand.test.ts keeps the two in step). */
export const DEFAULT_BRAND: Brand = {
  name: 'Gemini Enterprise',
  assistantName: 'Gemini',
  fullName: 'Gemini Enterprise for Microsoft 365',
};

export const brand: Brand = typeof __GE_BRAND__ === 'undefined' ? DEFAULT_BRAND : __GE_BRAND__;
