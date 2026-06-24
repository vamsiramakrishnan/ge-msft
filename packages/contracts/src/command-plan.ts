import { z } from 'zod';
import { IntentSchema, type Intent } from './intent.js';
import { SurfaceSchema, type Surface } from './context.js';

/**
 * The planner block (ADR-0004 — the `m365-command-planner` skill). The planner turns free text into
 * a confirmable, structured plan emitted inside a fenced ` ```plan ` block of keyword lines. This is
 * the faithful TS port of `skill/m365-command-planner/scripts/parse_plan.py` (the dependency-free
 * Python checker) — keep the keyword set, validation, and structural rules in lockstep with it; the
 * TS side is authoritative. The runtime applies the authoritative parse when it dispatches the plan.
 */

/** Scalar keys (last one wins). */
const SCALAR_KEYS = new Set(['intent', 'surface', 'scope', 'confidence']);
/** List keys (accumulate, in order). */
const LIST_KEYS = new Set(['ground', 'step', 'exclude', 'clarify']);
/** Optional bracket markers, ignored. */
const BRACKETS = new Set(['plan', 'end']);
const ALL_KEYS = new Set([...SCALAR_KEYS, ...LIST_KEYS, ...BRACKETS]);

const INTENTS = new Set(IntentSchema.options);
const SURFACES = new Set(SurfaceSchema.options);
const CONFIDENCE = new Set(['high', 'medium', 'low']);

export const CommandPlanSchema = z.object({
  intent: IntentSchema,
  surface: SurfaceSchema,
  scope: z.string().optional(),
  ground: z.array(z.string()).default([]),
  steps: z.array(z.string()),
  excludes: z.array(z.string()).default([]),
  clarify: z.array(z.string()).default([]),
  confidence: z.enum(['high', 'medium', 'low']).optional(),
});
export type CommandPlan = z.infer<typeof CommandPlanSchema>;

const _FENCE = /```plan[^\S\n]*\r?\n([\s\S]*?)```/i;
const _FENCE_OPEN = /```plan[^\S\n]*\r?\n([\s\S]*)$/i;

/**
 * Return the inner text of the first ```plan fence, or `null` (→ re-prompt, not an error).
 * Tolerates an unclosed fence (a frequent real-world failure mode).
 */
export function extractPlanBlock(text: string): string | null {
  const closed = text.match(_FENCE);
  if (closed) return closed[1]!.trim();
  const open = text.match(_FENCE_OPEN);
  if (open) return open[1]!.replace(/```\s*$/, '').trim();
  return null;
}

/** One parsed keyword line: a key/value pair, a bracket marker, an error, or nothing to do. */
type ParsedLine =
  | { kind: 'pair'; key: string; value: string }
  | { kind: 'bracket' }
  | { kind: 'error'; error: string }
  | null;

/** A "did you mean" hint against the known keywords (brackets excluded), or '' when none is close. */
function didYouMean(key: string): string {
  const candidates = [...ALL_KEYS].filter((k) => !BRACKETS.has(k)).sort();
  let best: string | undefined;
  let bestDist = Infinity;
  for (const candidate of candidates) {
    const d = levenshtein(key, candidate);
    if (d < bestDist) {
      bestDist = d;
      best = candidate;
    }
  }
  // Mirror difflib.get_close_matches' default cutoff (0.6 similarity ratio).
  if (best === undefined) return '';
  const ratio = 1 - bestDist / Math.max(key.length, best.length);
  return ratio >= 0.6 ? ` — did you mean '${best}'?` : '';
}

/** Parse one keyword line into a pair, a bracket, an error, or null (blank/comment). */
function parseLine(line: string): ParsedLine {
  const trimmed = line.trim();
  if (trimmed === '' || trimmed.startsWith('#')) return null;

  const ws = trimmed.search(/\s/);
  const key = (ws === -1 ? trimmed : trimmed.slice(0, ws)).toLowerCase();
  let rest = ws === -1 ? '' : trimmed.slice(ws + 1).trim();

  if (BRACKETS.has(key)) return { kind: 'bracket' };
  if (!ALL_KEYS.has(key)) {
    return { kind: 'error', error: `unknown plan keyword '${key}'${didYouMean(key)}` };
  }
  if (rest === '') return { kind: 'error', error: `'${key}' needs a value` };

  if (key === 'intent' && !INTENTS.has(rest as Intent)) {
    return {
      kind: 'error',
      error: `unknown intent '${rest}' — expected one of ${JSON.stringify([...INTENTS].sort())}`,
    };
  }
  if (key === 'surface' && !SURFACES.has(rest.toLowerCase() as Surface)) {
    return {
      kind: 'error',
      error: `unknown surface '${rest}' — expected one of ${JSON.stringify([...SURFACES].sort())}`,
    };
  }
  if (key === 'confidence' && !CONFIDENCE.has(rest.toLowerCase())) {
    return { kind: 'error', error: `confidence must be high|medium|low — got '${rest}'` };
  }
  if (key === 'ground') {
    rest = rest.trim().replace(/^"+|"+$/g, '');
  }
  return { kind: 'pair', key, value: rest };
}

/**
 * Parse the planner's reply: extract the ```plan fence, parse each keyword line, validate
 * intent/surface/confidence, accumulate the list keys, and apply the structural rules (a plan needs
 * an `intent`, a `surface`, and at least one `step` — unless it asks a `clarify` first). Mirrors
 * `parse_plan.py`'s `parse_plan`; `plan` is `null` when there is no fence or the plan is structurally
 * incomplete. `needsClarification` is true when the plan carries any `clarify` line.
 */
export function parsePlanBlock(text: string): {
  plan: CommandPlan | null;
  errors: string[];
  needsClarification: boolean;
} {
  const inner = extractPlanBlock(text);
  if (inner === null) {
    // No ```plan fence — re-prompt, not an error (matches the python "note" path).
    return { plan: null, errors: [], needsClarification: false };
  }

  const errors: string[] = [];
  let intent: string | undefined;
  let surface: string | undefined;
  let scope: string | undefined;
  let confidence: string | undefined;
  const ground: string[] = [];
  const steps: string[] = [];
  const excludes: string[] = [];
  const clarify: string[] = [];

  for (const raw of inner.split('\n')) {
    const rec = parseLine(raw);
    if (rec === null) continue;
    if (rec.kind === 'error') {
      errors.push(rec.error);
      continue;
    }
    if (rec.kind === 'bracket') continue;

    switch (rec.key) {
      case 'intent':
        intent = rec.value;
        break;
      case 'surface':
        surface = rec.value.toLowerCase();
        break;
      case 'scope':
        scope = rec.value;
        break;
      case 'confidence':
        confidence = rec.value.toLowerCase();
        break;
      case 'ground':
        ground.push(rec.value);
        break;
      case 'step':
        steps.push(rec.value);
        break;
      case 'exclude':
        excludes.push(rec.value);
        break;
      case 'clarify':
        clarify.push(rec.value);
        break;
    }
  }

  // Structural validation (mirrors the grammar's "required" column).
  if (!intent) errors.push("plan is missing 'intent'");
  if (!surface) errors.push("plan is missing 'surface'");
  if (steps.length === 0 && clarify.length === 0) {
    errors.push("plan needs at least one 'step' (or a 'clarify' to ask first)");
  }

  const needsClarification = clarify.length > 0;

  // Only assemble a structured plan when the required scalars are present; otherwise the caller
  // re-prompts off the errors (mirrors the python `plan` dict, but typed as null when incomplete).
  if (!intent || !surface) {
    return { plan: null, errors, needsClarification };
  }

  const plan: CommandPlan = {
    intent: intent as Intent,
    surface: surface as Surface,
    ...(scope !== undefined ? { scope } : {}),
    ground,
    steps,
    excludes,
    clarify,
    ...(confidence !== undefined ? { confidence: confidence as CommandPlan['confidence'] } : {}),
  };
  return { plan, errors, needsClarification };
}

/** Classic iterative Levenshtein edit distance (for the did-you-mean hint). */
function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  let curr = new Array<number>(n + 1);
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j]! + 1, curr[j - 1]! + 1, prev[j - 1]! + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[n]!;
}
