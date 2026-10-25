#!/usr/bin/env bun
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import {
  ActuationKindSchema, ActuationParamsSchema, CAPABILITY_REGISTRY,
  HostCapabilityProfileSchema, SurfaceSchema, listCapabilityModules,
} from '../packages/contracts/src/index.js';
import { capabilitySupport } from '../packages/runtime/src/host-profile.js';
import { NativeCertificationReportSchema } from '../packages/runtime/src/native-certification.js';

const [command = 'list', ...args] = process.argv.slice(2);
async function main(): Promise<void> {
  if (command === 'list') {
    const surface = args[0] ? SurfaceSchema.parse(args[0]) : undefined;
    console.log(JSON.stringify(CAPABILITY_REGISTRY.filter((entry) => !surface || entry.surface === surface).map((entry) => ({
      surface: entry.surface, kind: entry.kind, title: entry.title, promotion: entry.status,
      requirements: entry.requirementSets, moduleVersion: listCapabilityModules().find((module) => module.surface === entry.surface && module.kind === entry.kind)?.version,
    })), null, 2));
    return;
  }
  if (command === 'doctor') {
    if (!args[0]) throw new Error('Usage: bun scripts/office-lab.ts doctor host-profile.json');
    const profile = HostCapabilityProfileSchema.parse(JSON.parse(await readFile(args[0], 'utf8')));
    console.log(JSON.stringify({ host: profile, capabilities: CAPABILITY_REGISTRY.filter((entry) => entry.surface === profile.surface && entry.status === 'implemented').map((entry) => ({ kind: entry.kind, ...capabilitySupport(profile.surface, entry.kind, profile) })) }, null, 2));
    return;
  }
  if (command === 'scaffold') {
    const surface = SurfaceSchema.parse(args[0]);
    const kind = ActuationKindSchema.parse(args[1]);
    const fields = (args[2] ?? '').split(',').filter(Boolean);
    if (!fields.length || fields.some((field) => !Object.hasOwn(ActuationParamsSchema.shape, field)))
      throw new Error('Provide comma-separated parameter fields from ActuationParamsSchema.');
    const entry = CAPABILITY_REGISTRY.find((entry) => entry.surface === surface && entry.kind === kind);
    if (!entry) throw new Error('The capability must first have a reviewed registry entry for this surface.');
    if (listCapabilityModules().some((entry) => entry.surface === surface && entry.kind === kind))
      throw new Error('This capability already has a module. Extend its existing contract.');
    const path = resolve(`packages/bridge-${surface}/src/promotions/${kind}.ts`);
    const source = `import { ActuationParamsSchema, type CapabilityModule } from '@ge/contracts';\n\n/** Review required fields, nested strictness, native budgets, postconditions and compensation before registration. */\nexport const module: CapabilityModule = {\n  surface: '${surface}', kind: '${kind}', version: 1,\n  params: ActuationParamsSchema.pick({ ${fields.map((field) => field + ': true').join(', ')} }).strict(),\n  postcondition: ${JSON.stringify(entry.preview.join('; '))},\n  compensation: 'host-specific',\n};\n\n// Register a trusted bridge handler only after its native conformance fixture and inverse policy pass.\n// This descriptive module alone does not make the operation executable.\n`;
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, source, { flag: 'wx' });
    console.log(path);
    return;
  }
  if (command === 'check-native') {
    if (!args[0]) throw new Error('Usage: bun scripts/office-lab.ts check-native reports.json [commit] [word,excel,powerpoint,outlook]');
    const commit = args[1] ?? execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
    const required = (args[2] ?? 'word,excel,powerpoint,outlook').split(',').map((surface) => SurfaceSchema.parse(surface));
    const raw: unknown = JSON.parse(await readFile(args[0], 'utf8'));
    if (!Array.isArray(raw) || raw.length > 1000) throw new Error('Expected at most 1,000 native certification reports.');
    const reports = raw.map((report) => NativeCertificationReportSchema.parse(report));
    const expected: Record<string, string> = { word: 'word-tracked-edit', excel: 'excel-grid-readback', powerpoint: 'powerpoint-slide-readback', outlook: 'outlook-review-form' };
    for (const surface of required) {
      const selected = reports.filter((report) => report.surface === surface && report.commit === commit && report.fixtureId === expected[surface]);
      if (!selected.length || selected.some((report) => report.status !== 'passed' || !report.host || report.host.platform === 'unknown' || report.host.build === 'unknown' || !report.checks.execution || !report.checks.postcondition || !report.checks.finished || Object.values(report.checks).some((value) => !value)))
        throw new Error(`Native acceptance is missing or failed for ${surface} at ${commit}.`);
    }
    console.log(JSON.stringify({ commit, certifiedSurfaces: required, reports: reports.length }, null, 2));
    return;
  }
  throw new Error('Commands: list [surface], doctor profile.json, scaffold surface kind fields, check-native reports.json [commit] [surfaces]');
}
await main().catch((error: unknown) => { console.error(error instanceof Error ? error.message : 'Office lab failed.'); process.exitCode = 1; });
