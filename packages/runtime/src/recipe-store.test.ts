import { describe, expect, it } from 'vitest';
import { RecipeStore, recipeSource, MAX_RECIPES } from './recipe-store.js';
import { SkillRegistry } from './skill-registry.js';

function storage(initial?: unknown) {
  let persisted: unknown = initial;
  const saved: unknown[] = [];
  return {
    bridge: {
      recoveryStorage: {
        load: async () => structuredClone(persisted),
        save: async (_owner: string, value: unknown) => {
          persisted = structuredClone(value);
          saved.push(persisted);
        },
      },
    },
    saved,
    get persisted() {
      return persisted;
    },
  };
}

describe('RecipeStore (B6 durable recipes)', () => {
  it('round-trips the registry through document storage, re-parsing on load', async () => {
    const host = storage();
    const store = new RecipeStore(host.bridge, 'owner');
    const registry = new SkillRegistry();
    registry.register({
      kind: 'skill-def',
      name: 'writeTotal',
      params: ['src', 'dst'],
      body: ['set $dst = (read $src | sum amount)'],
    });
    await store.save([...registry.names()].map((n) => registry.get(n)!));
    expect(host.saved).toHaveLength(1);
    const reloaded = new SkillRegistry();
    const outcome = await new RecipeStore(host.bridge, 'owner').load(reloaded);
    expect(outcome).toEqual({ loaded: ['writeTotal'], rejected: [] });
    expect(reloaded.get('writeTotal')).toEqual({
      name: 'writeTotal',
      params: ['src', 'dst'],
      body: ['set $dst = (read $src | sum amount)'],
    });
  });

  it('renders a recipe back into def … end text', () => {
    expect(recipeSource({ name: 'x', params: ['a'], body: ['outline'] })).toBe(
      'def x($a):\noutline\nend',
    );
  });

  it('skips a stored recipe that no longer passes the grammar or registry and loads the rest', async () => {
    const host = storage({
      version: 1,
      recipes: [
        { name: 'set', params: [], body: ['outline'] }, // shadows a built-in verb
        { name: 'bad', params: [], body: ['def nested():'] }, // does not parse as a body
        { name: 'ok', params: ['p'], body: ['set A1 $p'] },
      ],
    });
    const registry = new SkillRegistry();
    const outcome = await new RecipeStore(host.bridge, 'owner').load(registry);
    expect(outcome.loaded).toEqual(['ok']);
    expect(outcome.rejected.map((r) => r.name).sort()).toEqual(['bad', 'set']);
    expect(registry.has('set')).toBe(false);
  });

  it('loads nothing from a corrupt or absent store and never throws', async () => {
    const registry = new SkillRegistry();
    expect(await new RecipeStore(storage({ version: 2 }).bridge, 'o').load(registry)).toEqual({
      loaded: [],
      rejected: [],
    });
    expect(await new RecipeStore(storage(undefined).bridge, 'o').load(registry)).toEqual({
      loaded: [],
      rejected: [],
    });
    const throwing = {
      recoveryStorage: {
        load: async () => {
          throw new Error('host');
        },
        save: async () => {},
      },
    };
    expect(await new RecipeStore(throwing, 'o').load(registry)).toEqual({
      loaded: [],
      rejected: [],
    });
    expect(new RecipeStore({}, 'o').durable).toBe(false);
  });

  it('refuses to store more than the bound', async () => {
    const store = new RecipeStore(storage().bridge, 'owner');
    const many = Array.from({ length: MAX_RECIPES + 1 }, (_, i) => ({
      name: `r${i}`,
      params: [],
      body: ['outline'],
    }));
    await expect(store.save(many)).rejects.toThrow('At most');
  });
});
