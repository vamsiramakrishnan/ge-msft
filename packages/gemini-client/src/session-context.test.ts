import { describe, it, expect } from 'vitest';
import type { AssistRequest, ResolvedContext } from '@ge/contracts';
import { SessionContext, contextValueToQueryPart } from './session-context.js';
import { buildStreamAssistRequest } from './stream-assist.js';

function textCtx(id: string, text: string): ResolvedContext {
  return {
    ref: { id, kind: 'selection', surface: 'word', title: id },
    value: { as: 'text', text },
  };
}

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'u@acme' };

describe('SessionContext', () => {
  it('adds, dedupes by id, removes, and clears', () => {
    const sc = new SessionContext();
    sc.add(textCtx('a', 'first'));
    sc.add(textCtx('a', 'updated')); // same id replaces
    sc.add(textCtx('b', 'second'));
    expect(sc.size).toBe(2);
    expect((sc.list()[0]!.value as { text: string }).text).toBe('updated');
    sc.remove('a');
    expect(sc.size).toBe(1);
    sc.clear();
    expect(sc.size).toBe(0);
  });

  it('maps each context value kind to the right query part', () => {
    expect(contextValueToQueryPart({ as: 'text', text: 'hi' })).toEqual({ text: 'hi' });
    expect(
      contextValueToQueryPart({ as: 'indexed-document', documentName: 'projects/x/d', title: 'D' }),
    ).toEqual({ documentReference: { documentName: 'projects/x/d', displayTitle: 'D' } });
    expect(contextValueToQueryPart({ as: 'drive-document', driveId: 'gd1' })).toEqual({
      driveDocumentReference: { driveId: 'gd1' },
    });
    expect(
      contextValueToQueryPart({ as: 'person', displayName: 'Vamsi', email: 'v@acme' }),
    ).toEqual({ personReference: { displayName: 'Vamsi', email: 'v@acme' } });
  });
});

describe('buildStreamAssistRequest with live context', () => {
  const req: AssistRequest = {
    intent: 'ask',
    query: 'Summarize the risk',
    unit: { connectors: [], surfaceContext: { kind: 'word', selection: 'ignored when parts set' } },
  };

  it('emits multi-part query: context parts then the question', () => {
    const sc = new SessionContext();
    sc.add(textCtx('sel', 'SLA is 99.5%'));
    sc.add({
      ref: { id: 'sp', kind: 'indexed-document', surface: 'word', title: 'Policy' },
      value: { as: 'indexed-document', documentName: 'projects/x/policy', title: 'Vendor Policy' },
    });
    const body = buildStreamAssistRequest(req, cfg, undefined, sc.list());
    const query = body.query as { parts: unknown[] };
    expect(query.parts).toEqual([
      { text: 'SLA is 99.5%' },
      { documentReference: { documentName: 'projects/x/policy', displayTitle: 'Vendor Policy' } },
      { text: 'Summarize the risk' },
    ]);
  });

  it('falls back to a composed single text query when nothing is attached', () => {
    const body = buildStreamAssistRequest(req, cfg);
    expect(body.query).toHaveProperty('text');
    expect(body.query).not.toHaveProperty('parts');
  });
});
