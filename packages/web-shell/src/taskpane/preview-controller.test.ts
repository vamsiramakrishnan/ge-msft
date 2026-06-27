// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { makeMockController } from './preview.js';
import { FIXTURE_STATE } from './preview-fixtures.js';
import type { PanelState } from '../controller.js';

/**
 * Behavioral tests for the preview harness's fake `PanelController`. The render smoke test exercises
 * the *view*; this exercises the fake controller's own contract: `getState` returns the supplied
 * snapshot, `subscribe` is an inert unsubscribe, `listSkills` reflects the state (and degrades to an
 * empty list), and every action method is callable, async-resolves, and logs without mutating the
 * snapshot. These guards keep the harness honest so the smoke test isn't standing on quicksand.
 */

let logSpy: ReturnType<typeof vi.fn>;

beforeEach(() => {
  logSpy = vi.spyOn(console, 'log').mockImplementation(() => undefined) as ReturnType<typeof vi.fn>;
});

afterEach(() => {
  logSpy.mockRestore();
});

describe('makeMockController', () => {
  it('returns exactly the state snapshot it was given', () => {
    const c = makeMockController(FIXTURE_STATE);
    expect(c.getState()).toBe(FIXTURE_STATE);
  });

  it('exposes a subscribe that returns an inert unsubscribe and never invokes the listener', () => {
    const c = makeMockController(FIXTURE_STATE);
    const listener = vi.fn();
    const unsub = c.subscribe(listener);
    expect(typeof unsub).toBe('function');
    // Calling the returned unsubscribe is harmless.
    expect(() => unsub()).not.toThrow();
    expect(listener).not.toHaveBeenCalled();
  });

  it('reflects the state skills through listSkills', () => {
    const c = makeMockController(FIXTURE_STATE);
    expect(c.listSkills()).toBe(FIXTURE_STATE.skills);
  });

  it('degrades listSkills to an empty array when the snapshot has no skills', () => {
    const noSkills: PanelState = { ...FIXTURE_STATE, skills: undefined };
    const c = makeMockController(noSkills);
    expect(c.listSkills()).toEqual([]);
  });

  it('logs and resolves async refreshContext without throwing', async () => {
    const c = makeMockController(FIXTURE_STATE);
    await expect(c.refreshContext()).resolves.toBeUndefined();
    expect(logSpy).toHaveBeenCalledWith('[preview] refreshContext');
  });

  it('logs attach/detach with the chip id', async () => {
    const c = makeMockController(FIXTURE_STATE);
    await c.attach('sp:contracts');
    c.detach('sp:contracts');
    expect(logSpy).toHaveBeenCalledWith('[preview] attach', 'sp:contracts');
    expect(logSpy).toHaveBeenCalledWith('[preview] detach', 'sp:contracts');
  });

  it('logs the query on send and runCommands', async () => {
    const c = makeMockController(FIXTURE_STATE);
    await c.send('hello');
    await c.runCommands('do the thing');
    expect(logSpy).toHaveBeenCalledWith('[preview] send', 'hello');
    expect(logSpy).toHaveBeenCalledWith('[preview] runCommands', 'do the thing');
  });

  it('logs the change id on applyProposal and resolves', async () => {
    const c = makeMockController(FIXTURE_STATE);
    const id = asChangeId('c-1');
    await expect(c.applyProposal(id)).resolves.toBeUndefined();
    expect(logSpy).toHaveBeenCalledWith('[preview] applyProposal', id);
  });

  it('logs the approval-gate actions for plan and write', () => {
    const c = makeMockController(FIXTURE_STATE);
    c.approvePlan();
    c.rejectPlan();
    c.approvePendingWrite();
    c.rejectPendingWrite();
    expect(logSpy).toHaveBeenCalledWith('[preview] approvePlan');
    expect(logSpy).toHaveBeenCalledWith('[preview] rejectPlan');
    expect(logSpy).toHaveBeenCalledWith('[preview] approvePendingWrite');
    expect(logSpy).toHaveBeenCalledWith('[preview] rejectPendingWrite');
  });

  it('logs cancel and dismissSuggestion', () => {
    const c = makeMockController(FIXTURE_STATE);
    c.cancel();
    c.dismissSuggestion('s-1');
    expect(logSpy).toHaveBeenCalledWith('[preview] cancel');
    expect(logSpy).toHaveBeenCalledWith('[preview] dismissSuggestion', 's-1');
  });

  it('logs the skill name and args when a skill is invoked, and resolves', async () => {
    const c = makeMockController(FIXTURE_STATE);
    await expect(
      c.invokeSkill('flag-vendor-risk', { vendor: 'Northwind Cloud' }),
    ).resolves.toBeUndefined();
    expect(logSpy).toHaveBeenCalledWith('[preview] invokeSkill', 'flag-vendor-risk', {
      vendor: 'Northwind Cloud',
    });
  });

  it('does not mutate the supplied snapshot when actions are invoked', async () => {
    const snapshot: PanelState = { ...FIXTURE_STATE };
    const c = makeMockController(snapshot);
    await c.send('q');
    c.approvePlan();
    c.detach('x');
    expect(c.getState()).toBe(snapshot);
    expect(snapshot.messages).toBe(FIXTURE_STATE.messages);
  });
});
