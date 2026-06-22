/**
 * @ge/contracts — the authoritative boundary between the surface-agnostic core and the
 * per-surface bridges (and the shape of every Gemini Enterprise call). Implement against this
 * exactly; every cross-boundary payload is parsed with its Zod schema on receipt. See docs/CONTRACTS.md.
 */
export * from './brand.js';
export * from './intent.js';
export * from './unit.js';
export * from './finding.js';
export * from './provenance.js';
export * from './request.js';
export * from './sse.js';
export * from './anchor.js';
export * from './context.js';
export * from './doc-state.js';
export * from './estate.js';
export * from './capability.js';
export * from './command-grammar.js';
export * from './expr-grammar.js';
