import { execFileSync } from 'node:child_process';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { z } from 'zod';
import { parsePlanBlock } from './command-plan.js';
import { AnalysisActionSchema } from './analysis-actions.js';
import { buildPreflightMetadata, describePreflightGuard } from './preflight-metadata.js';

const scripts = resolve('skill/m365-command-planner/scripts');
function pythonPlan(text: string): {
  errors: string[];
  plan: { step: string[]; clarify: string[] };
} {
  return JSON.parse(
    execFileSync(
      'python3',
      [
        '-c',
        'import sys,json;sys.path.insert(0,sys.argv[1]);from parse_plan import parse_plan;print(json.dumps(parse_plan(sys.stdin.read())))',
        scripts,
      ],
      { input: text, encoding: 'utf8' },
    ),
  );
}

describe('generated preflight contracts', () => {
  it('refuses unsupported schema semantics instead of losing them during JSON serialization', () => {
    expect(() => describePreflightGuard(z.string().regex(/^[A-Z]+$/))).toThrow(
      /Unsupported.*regex/,
    );
    expect(() => describePreflightGuard(z.string().email())).toThrow(/Unsupported/);
    expect(() => describePreflightGuard(z.number().multipleOf(2))).toThrow(/Unsupported/);
    expect(() => describePreflightGuard(z.coerce.number())).toThrow(/Unsupported/);
    expect(() => describePreflightGuard(z.object({}).catchall(z.string()))).toThrow(/Unsupported/);
    expect(() => describePreflightGuard(z.array(z.string()).length(2))).toThrow(/Unsupported/);
  });

  it('matches Python guard validation to the original Zod schemas', () => {
    const metadata = JSON.parse(JSON.stringify(buildPreflightMetadata()));
    const cases = [
      ['materialize', 'whenNonEmpty', true],
      ['materialize', 'whenNonEmpty', false],
      ['materialize', 'whenNonEmpty', 1],
      ['materialize', 'whenNonEmpty', 'yes'],
      ['query', 'requiredColumns', [{ input: '$source', indices: [0, 1], exactDecimal: true }]],
      ['query', 'requiredColumns', [{ input: '$source', indices: [true] }]],
      ['query', 'requiredColumns', [{ input: '', indices: [0] }]],
      ['query', 'requiredColumns', [{ input: '$source', indices: [16384] }]],
      ['query', 'requiredColumns', [{ input: '$source', indices: [0], unknown: true }]],
      ['query', 'requiredColumns', []],
      ['query', 'requiredColumns', null],
    ] as const;
    const actual: boolean[] = JSON.parse(
      execFileSync(
        'python3',
        [
          '-c',
          'import sys,json;sys.path.insert(0,"skill/shared");from language_manifest import guard_errors;data=json.load(sys.stdin);print(json.dumps([not guard_errors(value,data["metadata"]["analysisGuards"][kind][field],field) for kind,field,value in data["cases"]]))',
        ],
        { input: JSON.stringify({ metadata, cases }), encoding: 'utf8' },
      ),
    );
    const expected = cases.map(([kind, field, value]) => {
      const action = AnalysisActionSchema.options.find((entry) => entry.shape.kind.value === kind)!;
      return (action.shape as Record<string, z.ZodTypeAny>)[field]!.safeParse(value).success;
    });
    expect(actual).toEqual(expected);
  });

  it('every bundled planner example passes the production parser and Python preflight', () => {
    const directory = resolve('skill/m365-command-planner/assets/example-plans');
    for (const filename of readdirSync(directory).filter((name) => name.endsWith('.md'))) {
      const markdown = readFileSync(resolve(directory, filename), 'utf8');
      for (const block of markdown.matchAll(/```plan\s*\n([\s\S]*?)\n```/g)) {
        const text = `\`\`\`plan\n${block[1]}\n\`\`\``;
        expect(parsePlanBlock(text).errors, filename).toEqual([]);
        expect(pythonPlan(text).errors, filename).toEqual([]);
      }
    }
  });

  it.each(['workflow', 'source', 'target', 'phase', 'handoff'])(
    'both parsers reject the unsupported %s extension',
    (keyword) => {
      const text = `\`\`\`plan\nintent draft\nsurface excel\n${keyword} unsupported\nstep Prepare summary\n\`\`\``;
      expect(parsePlanBlock(text).errors).toContainEqual(
        expect.stringContaining('unknown plan keyword'),
      );
      expect(pythonPlan(text).errors).toContainEqual(
        expect.stringContaining('unknown plan keyword'),
      );
    },
  );
});
