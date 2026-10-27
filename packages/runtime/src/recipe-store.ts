import { z } from 'zod';
import { isProgramSkillDef, parseProgramBlock } from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import type { RegisteredSkill, SkillRegistry } from './skill-registry.js';

/**
 * B6 — durable recipes. A `def name(p…): … end` the model (or the user, through the planner)
 * registers lives in the in-session {@link SkillRegistry}; this store persists those definitions in
 * the host document's own metadata (`DocBridge.recoveryStorage`, the same port the cell-recovery
 * journal uses — Excel settings today) so they survive a pane reload and travel with the document.
 *
 * Trust: a stored recipe is DOCUMENT CONTENT, not configuration. On load every definition is
 * re-parsed through the grammar (`parseProgramBlock`) and re-registered through the registry's own
 * checks exactly as a freshly emitted `def` would be; a recipe is a macro whose expansion still
 * flows through type-check, dry-run and the approval gate, so a definition planted in a shared
 * document can at most propose commands the user then sees and approves. Nothing here reads the
 * recipes as instructions, and nothing executes at load time.
 */
const LineSchema = z.string().min(1).max(2048);
const RecipeSchema = z
  .object({
    name: z.string().min(1).max(64),
    params: z.array(z.string().min(1).max(64)).max(16),
    body: z.array(LineSchema).min(1).max(64),
  })
  .strict();
export const MAX_RECIPES = 32;
const MAX_BYTES = 256 * 1024;
const StoreSchema = z
  .object({ version: z.literal(1), recipes: z.array(RecipeSchema).max(MAX_RECIPES) })
  .strict();

export type StoredRecipe = z.infer<typeof RecipeSchema>;

/** The `def … end` text of a recipe — what the grammar re-parses on load. */
export function recipeSource(recipe: Pick<StoredRecipe, 'name' | 'params' | 'body'>): string {
  const params = recipe.params.map((p) => `$${p}`).join(' ');
  return [`def ${recipe.name}(${params}):`, ...recipe.body, 'end'].join('\n');
}

export interface RecipeLoadOutcome {
  /** Recipes the registry accepted, by name. */
  loaded: string[];
  /** Recipes the grammar or the registry refused (name → reason). Nothing partial is kept. */
  rejected: Array<{ name: string; reason: string }>;
}

export class RecipeStore {
  constructor(
    private readonly bridge: Pick<DocBridge, 'recoveryStorage'>,
    private readonly owner: string,
  ) {}

  /** Whether the host document can persist recipes at all. */
  get durable(): boolean {
    return Boolean(this.bridge.recoveryStorage);
  }

  private get key(): string {
    return `${this.owner}:recipes`;
  }

  /**
   * Load the stored recipes into `registry`. Each one is reconstructed as `def … end` text and
   * parsed through the grammar scoped to the names loaded so far (so a recipe may call an earlier
   * one), then registered through the registry's own validation. A stored recipe that no longer
   * parses, shadows a verb, or exceeds a bound is skipped and reported; a corrupt store loads
   * nothing. Never throws for a missing or unreadable store.
   */
  async load(registry: SkillRegistry): Promise<RecipeLoadOutcome> {
    const outcome: RecipeLoadOutcome = { loaded: [], rejected: [] };
    if (!this.bridge.recoveryStorage) return outcome;
    let raw: unknown;
    try {
      raw = await this.bridge.recoveryStorage.load(this.key);
    } catch {
      return outcome;
    }
    if (raw === undefined || raw === null || (Array.isArray(raw) && raw.length === 0))
      return outcome;
    const parsed = StoreSchema.safeParse(raw);
    if (!parsed.success) return outcome;
    for (const recipe of parsed.data.recipes) {
      // The grammar parses a fenced block, exactly as it would a model reply carrying this def.
      const { entries } = parseProgramBlock(
        '```cmd\n' + recipeSource(recipe) + '\n```',
        registry.names(),
      );
      const def = entries.find(isProgramSkillDef);
      const bad = entries.find((entry) => 'error' in entry);
      if (!def || bad) {
        outcome.rejected.push({
          name: recipe.name,
          reason: bad && 'error' in bad ? String(bad.error) : 'the stored recipe does not parse',
        });
        continue;
      }
      const result = registry.register(def);
      if (result.ok) outcome.loaded.push(result.name);
      else outcome.rejected.push({ name: recipe.name, reason: result.error });
    }
    return outcome;
  }

  /**
   * Persist the registry's current definitions. Bounded: more than {@link MAX_RECIPES} recipes, or
   * a payload over the size cap, is refused (the caller reports "kept for this session only").
   */
  async save(skills: readonly RegisteredSkill[]): Promise<void> {
    if (!this.bridge.recoveryStorage) throw new Error('This document cannot store recipes.');
    const recipes = skills.map((s) =>
      RecipeSchema.parse({ name: s.name, params: s.params, body: s.body }),
    );
    if (recipes.length > MAX_RECIPES)
      throw new Error(`At most ${MAX_RECIPES} recipes can be stored in a document.`);
    const payload = StoreSchema.parse({ version: 1, recipes });
    if (new TextEncoder().encode(JSON.stringify(payload)).byteLength > MAX_BYTES)
      throw new Error('The recipes are too large to store in this document.');
    await this.bridge.recoveryStorage.save(this.key, payload);
  }
}
