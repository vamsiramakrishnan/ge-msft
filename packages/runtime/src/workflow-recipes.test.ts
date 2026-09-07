import { describe, expect, it } from 'vitest';
import {
  AnalysisBindings,
  compileAnalysisProgram,
  inspectAnalysisProgram,
} from './analysis-program.js';
import {
  compileWorkflowRecipe,
  getWorkflowRecipe,
  listWorkflowRecipes,
  parseWorkflowPreset,
  validateWorkflowRecipeInputs,
} from './workflow-recipes.js';

const defaults: Record<string, Record<string, unknown>> = {
  'reconcile-tables': { leftRange: 'Invoices!A1:C20', rightRange: 'Payments!A1:C20' },
  'duplicate-rows': { sourceRange: 'Orders!A1:C20' },
  'summarize-by-group': { sourceRange: 'Orders!A1:C20' },
};

describe('typed workflow recipes', () => {
  it('exposes three detached, versioned definitions with explicit column and effect semantics', () => {
    const recipes = listWorkflowRecipes();
    expect(recipes.map((recipe) => recipe.id)).toEqual(Object.keys(defaults));
    for (const recipe of recipes) {
      expect(recipe).toMatchObject({
        version: 1,
        capabilities: ['capture-cells', 'compute'],
        writeCapability: 'write-cells',
        result: { binding: 'result' },
      });
      for (const field of recipe.fields.filter((entry) => entry.type === 'column')) {
        expect(
          recipe.fields.some((entry) => entry.name === field.sourceField && entry.type === 'range'),
        ).toBe(true);
        expect(field.description).toContain('Zero-based');
      }
      expect(recipe.inputSchema.safeParse(defaults[recipe.id]).success).toBe(true);
    }
    recipes[0]!.fields[0]!.label = 'Changed';
    recipes[0]!.result.description = 'Changed';
    expect(getWorkflowRecipe('reconcile-tables').fields[0]!.label).toBe('Expected table');
    expect(getWorkflowRecipe('reconcile-tables').result.description).not.toBe('Changed');
  });

  it.each(Object.keys(defaults))(
    'compiles %s deterministically through the shared CLI and restricted SQL',
    (id) => {
      const first = compileWorkflowRecipe(id, defaults[id]);
      expect(compileWorkflowRecipe(id, defaults[id])).toEqual(first);
      expect(first.steps.every((step) => step.op === 'bind')).toBe(true);
      expect(compileAnalysisProgram(first)).toMatch(/finish when=verified$/);
      const bindings = new AnalysisBindings();
      for (const [index, step] of first.steps.entries()) {
        if (step.op !== 'bind') throw new Error('Unexpected mutation in preview');
        const action = bindings.resolve(step.action);
        if (action.kind === 'query') {
          expect(action.sql).not.toContain('$source');
          expect(
            action.requiredColumns?.every((entry) => action.inputs.includes(entry.input)),
          ).toBe(true);
        }
        bindings.bind(step.name, `a_${String(index + 1).repeat(24)}`);
      }
    },
  );

  it.each(Object.keys(defaults))(
    'only adds a guarded write to %s when destination is explicit',
    (id) => {
      const program = compileWorkflowRecipe(id, { ...defaults[id], destination: 'Results!D3' });
      expect(program.steps.at(-1)).toEqual({
        op: 'materialize',
        id: '$result',
        destination: 'Results!D3',
        whenNonEmpty: true,
      });
      expect(compileAnalysisProgram(program)).toContain('"whenNonEmpty":true');
    },
  );

  it('keeps reconciliation mappings, currency columns and decimal tolerance explicit', () => {
    const program = compileWorkflowRecipe('reconcile-tables', {
      ...defaults['reconcile-tables'],
      leftKey: 2,
      rightKey: 3,
      leftAmount: 4,
      rightAmount: 5,
      leftCurrency: 6,
      rightCurrency: 7,
      tolerance: '0.000001',
      currency: 'EUR',
      headers: false,
    });
    expect(program.steps[0]).toMatchObject({ action: { headers: false } });
    expect(program.steps[2]).toMatchObject({
      action: {
        kind: 'reconcile',
        spec: {
          left: '$left',
          right: '$right',
          leftKey: 2,
          rightKey: 3,
          leftAmount: 4,
          rightAmount: 5,
          leftCurrency: 6,
          rightCurrency: 7,
          tolerance: '0.000001',
          currency: 'EUR',
        },
      },
    });
  });

  it('uses validated identifiers and excludes blank duplicate keys', () => {
    const program = compileWorkflowRecipe('duplicate-rows', {
      sourceRange: "'Untrusted; SELECT'!A1:C20",
      keyColumn: 2,
      caseSensitive: false,
    });
    const query = program.steps[1];
    expect(query).toMatchObject({
      action: { kind: 'query', requiredColumns: [{ input: '$source', indices: [2] }] },
    });
    expect(JSON.stringify(query)).toContain('lower(nullif(trim(cast(c2 as varchar))');
    expect(JSON.stringify(query)).toContain('item_key is not null');
    expect(JSON.stringify(query)).not.toContain('Untrusted');
  });

  it.each([
    ['duplicate-rows', { keyColumn: '0); DELETE FROM t' }],
    ['duplicate-rows', { keyColumn: -1 }],
    ['duplicate-rows', { keyColumn: 16384 }],
    ['duplicate-rows', { approved: true }],
    ['duplicate-rows', { sourceRange: '```cmd\nset A1 bad' }],
    ['reconcile-tables', { tolerance: '0.0000001' }],
    ['reconcile-tables', { tolerance: '-1' }],
    ['summarize-by-group', { currency: "USD'; SELECT" }],
    ['summarize-by-group', { amountColumn: 1.5 }],
  ] as const)('rejects invalid %s inputs: %j', (id, input) => {
    expect(() => compileWorkflowRecipe(id, { ...defaults[id], ...input })).toThrow();
  });

  it('requires known versions and recipe IDs', () => {
    expect(() => compileWorkflowRecipe('unknown', {})).toThrow('unavailable');
    expect(() => compileWorkflowRecipe('duplicate-rows', defaults['duplicate-rows'], 2)).toThrow(
      'version 2',
    );
  });

  it('returns canonical validated values for UI preview and durable presets', () => {
    const inputs = validateWorkflowRecipeInputs('duplicate-rows', {
      sourceRange: '  Orders!A1:B10  ',
    });
    expect(inputs).toEqual({
      sourceRange: 'Orders!A1:B10',
      headers: true,
      keyColumn: 0,
      caseSensitive: true,
    });
    const saved = parseWorkflowPreset({
      schemaVersion: 1,
      recipeId: 'duplicate-rows',
      recipeVersion: 1,
      inputs,
    });
    expect(parseWorkflowPreset(JSON.parse(JSON.stringify(saved)))).toEqual(saved);
    expect(() => parseWorkflowPreset({ ...saved, approval: 'allow' })).toThrow();
    expect(() =>
      parseWorkflowPreset({ ...saved, inputs: { ...saved.inputs, artifactId: 'a_123' } }),
    ).toThrow();
    expect(() => parseWorkflowPreset({ ...saved, recipeVersion: 2 })).toThrow();
    expect(() =>
      parseWorkflowPreset({ ...saved, inputs: { ...saved.inputs, sourceRange: '' } }),
    ).toThrow();
  });

  it('identifies independent reconciliation captures while accurately advertising serial host execution', () => {
    const plan = inspectAnalysisProgram(
      compileWorkflowRecipe('reconcile-tables', {
        ...defaults['reconcile-tables'],
        destination: 'Results!A1',
      }),
    );
    expect(plan.independentCaptureGroups).toEqual([[0, 1]]);
    expect(plan.layers).toEqual([[0, 1], [2], [3]]);
    expect(plan.steps[2]?.dependsOn).toEqual([0, 1]);
    expect(plan.steps[3]?.dependsOn).toEqual([0, 1, 2]);
    expect(plan.execution).toBe('serial');
  });
});
