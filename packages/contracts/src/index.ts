/**
 * @ge/contracts — the authoritative boundary between the gateway and the clients.
 * Implement against this exactly; every cross-boundary payload is parsed with its
 * Zod schema on receipt. See docs/CONTRACTS.md.
 */
export * from './intent.js';
export * from './unit.js';
export * from './finding.js';
export * from './provenance.js';
export * from './request.js';
export * from './sse.js';
export * from './anchor.js';
export * from './context.js';
export * from './capability.js';
