/**
 * Context budgeting + mechanism selection. These are **policy defaults**, not API limits
 * (the real caps live in the tenant's quota; see docs/api/discoveryengine/files-and-limits.md).
 * They answer the practical question: for this piece of content, do we inline it, reference
 * the indexed copy, or upload it as a session file for Python code execution?
 */

/** Conservative inline budget (tokens) for a single attached item before we prefer alternatives. */
export const DEFAULT_MAX_INLINE_TOKENS = 6000;

/** Soft ceiling (tokens) for the *total* attached inline context in one turn. */
export const DEFAULT_CONTEXT_BUDGET_TOKENS = 60000;

export type AttachStrategy = 'inline' | 'reference' | 'upload-for-code-execution';

export interface StrategyInput {
  tokensEstimate: number;
  /** The source already exists in a connected data store (can be referenced). */
  indexed?: boolean;
  /** The intent needs computation/pivots/charts over the data (favours code execution). */
  analytical?: boolean;
  maxInlineTokens?: number;
}

/**
 * Choose how to attach one source. Order of preference:
 *  1. analytical data → upload for code execution (Python over the real file);
 *  2. indexed → reference (ACL-preserving, ~free);
 *  3. fits the inline budget → inline;
 *  4. otherwise → reference if possible, else upload.
 */
export function recommendStrategy(input: StrategyInput): AttachStrategy {
  const max = input.maxInlineTokens ?? DEFAULT_MAX_INLINE_TOKENS;
  if (input.analytical) return 'upload-for-code-execution';
  if (input.indexed && input.tokensEstimate > max) return 'reference';
  if (input.tokensEstimate <= max) return input.indexed ? 'reference' : 'inline';
  return input.indexed ? 'reference' : 'upload-for-code-execution';
}

/** Running budget across the attached set; tells the UI when to stop inlining. */
export class ContextBudget {
  private used = 0;
  constructor(private readonly limit: number = DEFAULT_CONTEXT_BUDGET_TOKENS) {}

  /** Try to reserve tokens for an inline item; returns false if it would blow the budget. */
  tryAdd(tokens: number): boolean {
    if (this.used + tokens > this.limit) return false;
    this.used += tokens;
    return true;
  }

  get remaining(): number {
    return Math.max(0, this.limit - this.used);
  }

  get fraction(): number {
    return this.limit === 0 ? 1 : Math.min(1, this.used / this.limit);
  }
}
