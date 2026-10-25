import { describe, it, expect } from 'vitest';
import type { AssistRequest, ResolvedContext } from '@ge/contracts';
import { SessionContext, contextValueToQueryPart } from './session-context.js';
import {
  LANGUAGE_RULE,
  buildStreamAssistRequest,
  identityText,
  paneInstructions,
} from './stream-assist.js';

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

  it('puts the branded assistant name first when one is configured', () => {
    const named = { ...cfg, assistantName: 'ContosoAI' };
    const sc = new SessionContext();
    sc.add(textCtx('sel', 'SLA is 99.5%'));
    const parts = (
      buildStreamAssistRequest(req, named, undefined, sc.list()).query as {
        parts: unknown[];
      }
    ).parts;
    expect(parts).toEqual([
      { text: paneInstructions('ContosoAI')! },
      { text: 'SLA is 99.5%' },
      { text: 'Summarize the risk' },
    ]);
    const text = (buildStreamAssistRequest(req, named).query as { text: string }).text;
    expect(text.startsWith(`${paneInstructions('ContosoAI')}\n\n`)).toBe(true);
    expect(identityText('ContosoAI')).toContain('never Gemini');
    expect(identityText('ContosoAI')).toContain('Do not introduce yourself');
  });

  it('sends no identity line without a branded name', () => {
    const text = (buildStreamAssistRequest(req, cfg).query as { text: string }).text;
    expect(text).not.toContain('Your name is');
  });

  it('sends the brand writing rules right after the name, in the same instruction block', () => {
    const rule = 'Open Bahasa emails with "Halo", never "Yth."';
    const branded = { ...cfg, assistantName: 'ContosoAI', assistantRules: [rule] };
    const text = (buildStreamAssistRequest(req, branded).query as { text: string }).text;
    expect(
      text.startsWith(
        `<pane_instructions>\n${identityText('ContosoAI')}\n${rule}\n${LANGUAGE_RULE}\n</pane_instructions>\n\n`,
      ),
    ).toBe(true);
    expect(paneInstructions(undefined, [])).toBeUndefined();
  });

  it('sends a rule limited to command turns only when the assistant writes (not on chat or planner)', () => {
    const rule = 'Open Bahasa emails with "Halo".';
    const branded = {
      ...cfg,
      assistantName: 'ContosoAI',
      assistantRules: [{ text: rule, routes: ['command' as const] }],
    };
    const queryOf = (route: 'default' | 'planner' | 'command'): string =>
      (
        buildStreamAssistRequest(req, branded, undefined, undefined, route).query as {
          text: string;
        }
      ).text;
    expect(queryOf('command')).toContain(rule);
    expect(queryOf('default')).not.toContain(rule);
    expect(queryOf('planner')).not.toContain(rule);
    expect(queryOf('default')).toContain(identityText('ContosoAI'));
  });
});
