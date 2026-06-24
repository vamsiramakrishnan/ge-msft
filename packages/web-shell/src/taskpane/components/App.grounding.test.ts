// @vitest-environment jsdom
import { describe, it, expect, vi, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  ContextRef,
  SseEvent,
} from '@ge/contracts';
import type { CommandLoopEvent, RunCommandsOptions } from '@ge/runtime';
import type { ResolvedGrounding } from '@ge/gemini-client';
import type { HostEvent } from '@ge/triggers';
import { PanelController, type AssistLike, type ContextLister } from '../../controller.js';
import { App, invocationToGrounding, mentionToSelection } from './App.js';
import type { ComposerInvocation } from './Composer.js';

/**
 * Finding #2/#B-wire: the composer's typed `@`-mentions must be CONSUMED into structured grounding
 * (mapped to `GroundingSelection[]`, resolved via the gemini-client resolver, threaded into the turn),
 * NOT discarded with only the raw text forwarded. These tests assert the resolver-call seam directly
 * and the end-to-end App → controller → session wiring.
 */

const inv = (mentions: ComposerInvocation['mentions'], raw = ''): ComposerInvocation => ({
  scope: { kind: 'selection' },
  mentions,
  instruction: raw,
  raw,
});

describe('mentionToSelection — typed @-mention → GroundingSelection', () => {
  it('reference kinds carry no id; addressable kinds require their ref', () => {
    expect(mentionToSelection({ kind: 'this' })).toEqual({ kind: 'current-context' });
    expect(mentionToSelection({ kind: 'unit' })).toEqual({ kind: 'unit' });
    expect(mentionToSelection({ kind: 'datastore', ref: 'ds-1' })).toEqual({
      kind: 'data-store',
      id: 'ds-1',
    });
    expect(mentionToSelection({ kind: 'document', ref: 'doc-1' })).toEqual({
      kind: 'document',
      id: 'doc-1',
    });
    expect(mentionToSelection({ kind: 'upload', ref: 'file-1' })).toEqual({
      kind: 'upload',
      fileId: 'file-1',
    });
    // An addressable kind with NO ref cannot be resolved — dropped, never smuggled as text.
    expect(mentionToSelection({ kind: 'datastore' })).toBeUndefined();
  });
});

describe('invocationToGrounding — @-mentions reach the structured request path', () => {
  it('@this and a data-store mention resolve to structured grounding (queryParts + dataStoreSpecs), not raw text', () => {
    const grounding = invocationToGrounding(
      inv([{ kind: 'this' }, { kind: 'datastore', ref: 'projects/p/dataStores/ds-1' }]),
      // The bridge's live context parts the resolver addresses `@this` onto.
      { contextParts: [{ text: 'the selected paragraph' }] },
    );

    // @this → the live context query parts; @data-store → a dataStoreSpec — both STRUCTURED.
    expect(grounding?.queryParts).toEqual([{ text: 'the selected paragraph' }]);
    expect(grounding?.dataStoreSpecs).toEqual([{ dataStore: 'projects/p/dataStores/ds-1' }]);
  });

  it('a mention-free invocation produces no grounding (undefined)', () => {
    expect(invocationToGrounding(inv([]))).toBeUndefined();
  });
});

/** A minimal AssistLike that records the grounding handed to each turn. */
class RecordingAssist implements AssistLike {
  context = { size: 0 };
  runGrounding: Array<ResolvedGrounding | undefined> = [];
  askGrounding: Array<ResolvedGrounding | undefined> = [];
  attachRef(): Promise<void> {
    return Promise.resolve();
  }
  detach(): void {}
  async *ask(
    _q: string,
    opts?: { signal?: AbortSignal; grounding?: ResolvedGrounding },
  ): AsyncGenerator<SseEvent> {
    this.askGrounding.push(opts?.grounding);
    yield { type: 'done' };
  }
  apply(
    kind: ActuationRequest['kind'],
    _params: ActuationParams,
    changeId: ChangeId,
  ): Promise<ActuationResult> {
    return Promise.resolve({ ok: true, changeId, kind, location: 'A1' });
  }
  async *runCommands(
    _task: string,
    opts?: RunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent> {
    this.runGrounding.push(opts?.grounding);
    yield { type: 'done', turn: 1, answer: '' };
  }
  plan(): Promise<{ plan: null; errors: string[]; needsClarification: boolean }> {
    return Promise.resolve({ plan: null, errors: [], needsClarification: false });
  }
  ingest(_e: HostEvent): Promise<void> {
    return Promise.resolve();
  }
}

const lister = (refs: ContextRef[]): ContextLister => ({
  listContext: () => Promise.resolve(refs),
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement | undefined;
let root: Root | undefined;

afterEach(() => {
  if (root) act(() => root!.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

describe('<App/> onInvoke — structured grounding reaches the controller→session turn', () => {
  it('a chat /verb with @this forwards structured grounding to send (not just raw text)', async () => {
    const assist = new RecordingAssist();
    const controller = new PanelController(assist, lister([]));
    const sendSpy = vi.spyOn(controller, 'send');

    const el = document.createElement('div');
    container = el;
    document.body.appendChild(el);
    const r = createRoot(el);
    root = r;
    act(() => {
      r.render(createElement(App, { controller, surface: 'word' }));
    });

    // Type `@this` and submit — a chat turn (no actuating verb) routes to send().
    const input = el.querySelector('input#ask') as HTMLInputElement;
    const form = el.querySelector('form.comp') as HTMLFormElement;
    // React tracks the input value via its own setter; use the native setter so onChange fires.
    const setNativeValue = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype,
      'value',
    )?.set;
    act(() => {
      setNativeValue?.call(input, 'what is @this about?');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });

    // The send received a STRUCTURED ResolvedGrounding (the @this mention), not just the raw string.
    expect(sendSpy).toHaveBeenCalledTimes(1);
    const grounding = sendSpy.mock.calls[0]?.[1];
    expect(grounding).toBeDefined();
    // @this with no live context resolves to a structured "dropped" note — still STRUCTURED, never
    // inlined as prompt text. (The raw text still rides as the model-facing task, arg 0.)
    expect(grounding?.notes?.some((n) => n.kind === 'current-context')).toBe(true);
    expect(assist.askGrounding[0]).toBe(grounding); // it reached the session turn
  });
});
