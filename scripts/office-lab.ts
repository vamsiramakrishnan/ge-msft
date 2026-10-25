#!/usr/bin/env bun
import { readFile, mkdir, writeFile, open } from 'node:fs/promises';
import { constants as fileFlags } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import {
  ActuationKindSchema, ActuationParamsSchema, CAPABILITY_REGISTRY,
  HostCapabilityProfileSchema, SurfaceSchema, listCapabilityModules,
} from '../packages/contracts/src/index.js';
import { capabilitySupport } from '../packages/runtime/src/host-profile.js';
import { NativeCertificationReportSchema } from '../packages/runtime/src/native-certification.js';
import { assessOfficeReadiness } from '../packages/runtime/src/office-readiness.js';
import { createOfficeAcceptanceBundle, assessOfficeAcceptance, collectOfficeAcceptanceEvidence, OfficeAcceptancePlanSchema } from '../packages/runtime/src/office-acceptance.js';
import { nativeFixtureExpectation } from '../packages/runtime/src/native-certification.js';
import { OFFICE_ACCEPTANCE_FIXTURES } from '../packages/runtime/src/office-acceptance.js';
import { z } from 'zod';

async function boundedReport(file: string, limit = 8 * 1024 * 1024): Promise<string> {
  const handle = await open(file, fileFlags.O_RDONLY | fileFlags.O_NONBLOCK);
  try {
    const metadata = await handle.stat();
    if (!metadata.isFile() || metadata.size > limit) throw new Error('Acceptance reports must be regular files within the eight MiB budget.');
    const buffer = Buffer.alloc(Math.min(metadata.size + 1, limit + 1));
    let offset = 0;
    while (offset < buffer.length) {
      const { bytesRead } = await handle.read(buffer, offset, buffer.length - offset, offset);
      if (!bytesRead) break;
      offset += bytesRead;
    }
    const after = await handle.stat();
    if (offset > limit || offset > metadata.size || after.size !== metadata.size || after.mtimeMs !== metadata.mtimeMs) throw new Error('Acceptance report changed while reading. Retry from a stable file.');
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer.subarray(0, offset));
  } finally { await handle.close(); }
}

const [command = 'list', ...args] = process.argv.slice(2);
async function main(): Promise<void> {
  if (command === 'campaign-init') {
    const head=args[0]??execFileSync('git',['rev-parse','HEAD'],{encoding:'utf8'}).trim();
    const selected=(args[1]??'word,excel,powerpoint,outlook').split(',').map(value=>{
      if(!['word','excel','powerpoint','outlook'].includes(value))throw new Error('Campaigns cover Word, Excel, PowerPoint and Outlook.');
      return value as 'word'|'excel'|'powerpoint'|'outlook';
    });
    console.log(JSON.stringify(createOfficeAcceptanceBundle(head,selected),null,2));
    return;
  }
  if (command === 'campaign') {
    if(!args[0])throw new Error('Usage: office:lab campaign bundle.json [commit]');
    const head=args[1]??execFileSync('git',['rev-parse','HEAD'],{encoding:'utf8'}).trim();
    const body=await boundedReport(args[0]);
    const report=assessOfficeAcceptance(JSON.parse(body),head);
    console.log(JSON.stringify(report,null,2));
    process.exitCode=report.status==='PASSED'?0:report.status==='FAIL'?1:2;
    return;
  }
  if (command === 'campaign-collect') {
    if(!args[0]||args.length<2||args.length>1001)throw new Error('Usage: office:lab campaign-collect plan.json report.json [...]');
    const plan=OfficeAcceptancePlanSchema.parse(JSON.parse(await boundedReport(args[0],256*1024)));
    const documents:unknown[]=[];let bytes=0;
    for(const file of args.slice(1)){
      const body=await boundedReport(file,8*1024*1024-bytes);bytes+=new TextEncoder().encode(body).byteLength;
      if(bytes>8*1024*1024)throw new Error('Imported acceptance reports exceed eight MiB.');
      documents.push(JSON.parse(body));
    }
    console.log(JSON.stringify(collectOfficeAcceptanceEvidence(plan,documents),null,2));
    return;
  }
  if (command === 'fixtures') {
    const selected=args[0];
    if(selected&&!Object.hasOwn(OFFICE_ACCEPTANCE_FIXTURES,selected))throw new Error('Choose word, excel, powerpoint or outlook.');
    const ids=Object.entries(OFFICE_ACCEPTANCE_FIXTURES).filter(([host])=>!selected||host===selected).flatMap(([,values])=>values);
    console.log(JSON.stringify(ids.map(id=>({id,...nativeFixtureExpectation(id)})),null,2));
    return;
  }
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
  if (command === 'readiness') {
    if (!args[0]) throw new Error('Usage: bun scripts/office-lab.ts readiness evidence.json [commit]');
    const commit = args[1] ?? execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
    const body = await readFile(args[0], 'utf8');
    if (new TextEncoder().encode(body).byteLength > 8 * 1024 * 1024) throw new Error('Release evidence exceeds eight MiB.');
    const report = assessOfficeReadiness(JSON.parse(body), commit);
    console.log(JSON.stringify(report, null, 2));
    process.exitCode = report.status === 'PASSED' ? 0 : report.status === 'FAIL' ? 1 : 2;
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
  throw new Error('Commands: list, doctor, scaffold, fixtures, check-native, readiness, campaign-init, campaign-collect, campaign');
}
await main().catch((error: unknown) => { console.error(error instanceof z.ZodError ? 'Office report does not match the supported schema.' : error instanceof SyntaxError ? 'Office report is not valid JSON.' : error instanceof Error ? error.message : 'Office lab failed.'); process.exitCode = 1; });
