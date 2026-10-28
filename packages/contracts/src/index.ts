/**
 * @ge/contracts — the authoritative boundary between the surface-agnostic core and the
 * per-surface bridges (and the shape of every Gemini Enterprise call). Implement against this
 * exactly; every cross-boundary payload is parsed with its Zod schema on receipt. See docs/CONTRACTS.md.
 */
export * from './brand.js';
export * from './intent.js';
export * from './grounding.js';
export * from './unit.js';
export * from './finding.js';
export * from './provenance.js';
export * from './request.js';
export * from './sse.js';
export * from './anchor.js';
export * from './context.js';
export * from './doc-state.js';
export * from './document-read.js';
export * from './deck.js';
export * from './estate.js';
export * from './capability.js';
export * from './actuation-outcome.js';
export * from './capability-registry.js';
export * from './plan-graph.js';
export * from './capability-closure.js';
export * from './command-grammar.js';
export * from './command-help.js';
export * from './command-plan.js';
export * from './command-palette.js';
export * from './intent-capability.js';
export * from './quick-actions.js';
export * from './expr-grammar.js';
export * from './language-manifest.js';
export * from './skill-grammar.js';
export * from './each-grammar.js';
export * from './manifest-lint.js';
export * from './release-profile.js';
export * from './docfs.js';

export * from './analysis.js';
export * from './cells.js';
export * from './analysis-actions.js';
export * from './analysis-program-schema.js';
export * from './provenance-record.js';
export * from './evidence.js';

export * from './host-profile.js';
export * from './capability-module.js';

export * from './formula-footprint.js';

export * from './formula-policy.js';

export * from './deck-template.js';

export * from './durable-program.js';
export * from './durable-inspection.js';
export * from './office-handoff.js';

export * from './excel-native-target.js';
export * from './tenant-workflow-report.js';
export * from './office-benchmark.js';
