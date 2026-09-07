import { readdirSync, readFileSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const builtin = new Set(builtinModules.map((name) => name.replace(/^node:/, '')));
interface PackageMetadata {
  name: string;
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
}
interface ModuleReference {
  specifier: string;
  typeOnly: boolean;
  line: number;
}
interface Workspace {
  dir: string;
  metadata: PackageMetadata;
}

/** Direction is intentional; declaring a dependency cannot silently authorize a reverse edge. */
const allowedDependencies: Record<string, readonly string[]> = {
  '@ge/contracts': [],
  '@ge/content': ['@ge/contracts'],
  '@ge/compute': ['@ge/contracts'],
  '@ge/deck-compiler': ['@ge/contracts'],
  '@ge/gemini-client': ['@ge/contracts'],
  '@ge/triggers': ['@ge/contracts'],
  '@ge/graph-client': ['@ge/contracts', '@ge/content'],
  '@ge/runtime': [
    '@ge/contracts',
    '@ge/content',
    '@ge/compute',
    '@ge/gemini-client',
    '@ge/triggers',
  ],
  ...Object.fromEntries(
    [
      'bridge-word',
      'bridge-excel',
      'bridge-powerpoint',
      'bridge-outlook',
      'bridge-onenote',
      'teams',
    ].map((name) => [
      `@ge/${name}`,
      ['@ge/contracts', '@ge/content', '@ge/runtime', '@ge/triggers'],
    ]),
  ),
};

function filesIn(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory() ? filesIn(resolve(dir, entry.name)) : [resolve(dir, entry.name)],
  );
}
function productionSource(file: string): boolean {
  const path = file.split(sep).join('/');
  return (
    /\.(?:[cm]?[jt]sx?)$/.test(path) &&
    !/\.(?:test|spec)\.[^/]+$/.test(path) &&
    !/(?:^|\/)(?:test-harness|__tests__|__fixtures__|fixtures)(?:\/|$)/.test(path) &&
    !/(?:^|\/)(?:preview(?:[-.]|\/)|analysis-preview\.)/.test(path)
  );
}
function packageName(specifier: string): string {
  return specifier.startsWith('@')
    ? specifier.split('/').slice(0, 2).join('/')
    : specifier.split('/')[0]!;
}
function moduleReferences(text: string, file: string): ModuleReference[] {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
  const found: ModuleReference[] = source.typeReferenceDirectives.map((reference) => ({
    specifier: reference.fileName,
    typeOnly: true,
    line: source.getLineAndCharacterOfPosition(reference.pos).line + 1,
  }));
  const add = (node: ts.Node, value: ts.Node | undefined, typeOnly: boolean) => {
    if (value && ts.isStringLiteralLike(value))
      found.push({
        specifier: value.text,
        typeOnly,
        line: source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1,
      });
  };
  const visit = (node: ts.Node): void => {
    if (ts.isImportDeclaration(node)) {
      const clause = node.importClause;
      const named = clause?.namedBindings;
      const typeOnly = Boolean(
        clause?.isTypeOnly ||
        (!clause?.name &&
          named &&
          ts.isNamedImports(named) &&
          named.elements.length > 0 &&
          named.elements.every((entry) => entry.isTypeOnly)),
      );
      add(node, node.moduleSpecifier, typeOnly);
    } else if (ts.isExportDeclaration(node)) {
      const named = node.exportClause;
      const typeOnly = Boolean(
        node.isTypeOnly ||
        (named &&
          ts.isNamedExports(named) &&
          named.elements.length > 0 &&
          named.elements.every((entry) => entry.isTypeOnly)),
      );
      add(node, node.moduleSpecifier, typeOnly);
    } else if (ts.isImportTypeNode(node) && ts.isLiteralTypeNode(node.argument))
      add(node, node.argument.literal, true);
    else if (
      ts.isImportEqualsDeclaration(node) &&
      ts.isExternalModuleReference(node.moduleReference)
    )
      add(node, node.moduleReference.expression, node.isTypeOnly);
    else if (
      ts.isCallExpression(node) &&
      (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
        (ts.isIdentifier(node.expression) && node.expression.text === 'require'))
    )
      add(node, node.arguments[0], false);
    ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
}
function workspacePackages(): Workspace[] {
  return readdirSync(resolve(root, 'packages'), { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => {
      const dir = resolve(root, 'packages', entry.name);
      return {
        dir,
        metadata: JSON.parse(readFileSync(resolve(dir, 'package.json'), 'utf8')) as PackageMetadata,
      };
    });
}
function referencesFor(dir: string): string[] {
  const file = resolve(dir, 'tsconfig.json');
  const parsed = ts.readConfigFile(file, ts.sys.readFile);
  if (parsed.error)
    throw new Error(ts.flattenDiagnosticMessageText(parsed.error.messageText, '\n'));
  return ((parsed.config as { references?: Array<{ path: string }> }).references ?? []).map(
    (reference) => resolve(dir, reference.path),
  );
}
function productionDeclarations(metadata: PackageMetadata): Record<string, string> {
  return {
    ...metadata.dependencies,
    ...metadata.optionalDependencies,
    ...metadata.peerDependencies,
  };
}
function importViolation(
  workspace: Workspace,
  file: string,
  entry: ModuleReference,
  workspaces: Map<string, Workspace>,
): string | undefined {
  const prefix = `${relative(root, file)}:${entry.line}`;
  const specifier = entry.specifier;
  if (specifier.startsWith('.')) {
    const target = resolve(dirname(file), specifier);
    if (target !== workspace.dir && !target.startsWith(`${workspace.dir}${sep}`))
      return `${prefix}: cross-package relative import ${specifier}; use the package's public entry point`;
    return;
  }
  if (specifier.startsWith('node:') || builtin.has(specifier)) return;
  const name = packageName(specifier);
  const declared = {
    ...productionDeclarations(workspace.metadata),
    ...(entry.typeOnly ? workspace.metadata.devDependencies : {}),
  };
  if (!declared[name])
    return `${prefix}: ${name} must be declared in this package's ${entry.typeOnly ? 'dependencies or devDependencies' : 'production dependencies'}`;
  if (name.startsWith('@ge/')) {
    if (!workspaces.has(name)) return `${prefix}: unknown workspace ${name}`;
    // The shell is the composition root: host adapters are assembled here, never below runtime.
    if (
      workspace.metadata.name !== '@ge/web-shell' &&
      !allowedDependencies[workspace.metadata.name]?.includes(name)
    )
      return `${prefix}: forbidden dependency direction ${workspace.metadata.name} → ${name}`;
  }
}

describe('repository dependency boundaries', () => {
  it('checks syntax-bearing imports, not comments or strings, including scoped subpaths and type-only forms', () => {
    const imports = moduleReferences(
      `// import { ignored } from 'comment';
      const text = "import('string')";
      import { value, type Mixed } from '@ge/contracts';
      import { type Only } from '@azure/msal-common/browser';
      export type { Shape } from './types.js';
      const lazy = import('@ge/compute/browser');
      type Module = import('zod').ZodType;
      const legacy = require('node:fs');`,
      'example.ts',
    );
    expect(imports.map(({ specifier, typeOnly }) => ({ specifier, typeOnly }))).toEqual([
      { specifier: '@ge/contracts', typeOnly: false },
      { specifier: '@azure/msal-common/browser', typeOnly: true },
      { specifier: './types.js', typeOnly: true },
      { specifier: '@ge/compute/browser', typeOnly: false },
      { specifier: 'zod', typeOnly: true },
      { specifier: 'node:fs', typeOnly: false },
    ]);
    expect(packageName('@duckdb/duckdb-wasm/dist/module.js')).toBe('@duckdb/duckdb-wasm');
    expect(packageName('react/jsx-runtime')).toBe('react');
    expect(productionSource('/repo/src/test-harness/fake.ts')).toBe(false);
    expect(productionSource('/repo/src/taskpane/preview-interactive.ts')).toBe(false);
    expect(productionSource('/repo/src/taskpane/analysis-preview.tsx')).toBe(false);
    expect(productionSource('/repo/src/taskpane/main.tsx')).toBe(true);
  });

  it('rejects undeclared dependencies and reverse/relative edges even when hoisting hides them', () => {
    const workspace: Workspace = {
      dir: '/repo/packages/runtime',
      metadata: {
        name: '@ge/runtime',
        dependencies: { '@ge/web-shell': '*' },
        devDependencies: { zod: '*' },
      },
    };
    const packages = new Map([
      ['@ge/web-shell', { dir: '/repo/packages/web-shell', metadata: { name: '@ge/web-shell' } }],
    ]);
    const check = (specifier: string, typeOnly = false) =>
      importViolation(
        workspace,
        '/repo/packages/runtime/src/main.ts',
        { specifier, typeOnly, line: 1 },
        packages,
      );
    expect(check('zod')).toContain('production dependencies');
    expect(check('zod', true)).toBeUndefined();
    expect(check('@ge/web-shell')).toContain('forbidden dependency direction');
    expect(check('../../web-shell/src/controller.js')).toContain('cross-package relative import');
    expect(check('node:crypto')).toBeUndefined();
  });

  it('declares every direct production import and preserves the package dependency direction', () => {
    const packages = workspacePackages();
    const byName = new Map(packages.map((workspace) => [workspace.metadata.name, workspace]));
    const violations: string[] = [];
    for (const workspace of packages) {
      expect(
        workspace.metadata.name === '@ge/web-shell' ||
          workspace.metadata.name in allowedDependencies,
        `Classify ${workspace.metadata.name} in the dependency policy`,
      ).toBe(true);
      for (const dependency of Object.keys(productionDeclarations(workspace.metadata)).filter(
        (name) => name.startsWith('@ge/'),
      )) {
        if (!byName.has(dependency))
          violations.push(`${workspace.metadata.name}: unknown declared workspace ${dependency}`);
        if (
          workspace.metadata.name !== '@ge/web-shell' &&
          !allowedDependencies[workspace.metadata.name]?.includes(dependency)
        )
          violations.push(
            `${workspace.metadata.name}: forbidden declared dependency ${dependency}`,
          );
      }
      for (const file of filesIn(resolve(workspace.dir, 'src')).filter(productionSource)) {
        for (const entry of moduleReferences(readFileSync(file, 'utf8'), file)) {
          const violation = importViolation(workspace, file, entry, byName);
          if (violation) violations.push(violation);
        }
      }
    }
    expect(violations.sort(), violations.join('\n')).toEqual([]);
  });

  it('keeps workspace declarations and TypeScript project references in agreement', () => {
    const packages = workspacePackages();
    const byName = new Map(packages.map((workspace) => [workspace.metadata.name, workspace]));
    for (const workspace of packages) {
      const declared = {
        ...productionDeclarations(workspace.metadata),
        ...workspace.metadata.devDependencies,
      };
      const expected = Object.keys(declared)
        .filter((name) => name.startsWith('@ge/'))
        .map((name) => {
          const dependency = byName.get(name);
          expect(dependency, `Unknown declared workspace ${name}`).toBeDefined();
          return dependency!.dir;
        })
        .sort();
      expect(
        referencesFor(workspace.dir).sort(),
        `${workspace.metadata.name}: tsconfig references must match declared workspace dependencies`,
      ).toEqual(expected);
    }
    expect(referencesFor(root).sort()).toEqual(packages.map((workspace) => workspace.dir).sort());
  });
});
