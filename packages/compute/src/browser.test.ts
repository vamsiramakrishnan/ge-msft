import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { BrowserCompute } from './browser.js';
const mocks = vi.hoisted(() => ({ instantiate: vi.fn(), detach: vi.fn() }));
vi.mock('@duckdb/duckdb-wasm', () => ({
  VoidLogger: class {},
  AsyncDuckDB: class {
    instantiate = mocks.instantiate;
    detach = mocks.detach;
  },
}));
let terminated = 0;
beforeEach(() => {
  vi.useFakeTimers();
  terminated = 0;
  mocks.instantiate.mockReset().mockImplementation(() => new Promise(() => {}));
  vi.stubGlobal('location', { href: 'https://addin.test/', origin: 'https://addin.test' });
  vi.stubGlobal(
    'Worker',
    class {
      terminate() {
        terminated++;
      }
    },
  );
  vi.stubGlobal(
    'fetch',
    vi.fn(async () => ({ ok: true, blob: async () => new Blob(['wasm']) })),
  );
});
afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});
describe('compute supervision', () => {
  it('terminates initialization on cancel and never allows a second simultaneous query', async () => {
    const engine = new BrowserCompute({ workerUrl: '/worker.js', wasmUrl: '/engine.wasm' });
    const controller = new AbortController();
    const pending = engine.query('SELECT 1', [], controller.signal);
    const rejected = expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    await vi.advanceTimersByTimeAsync(1);
    await expect(engine.query('SELECT 2', [])).rejects.toThrow('already running');
    controller.abort();
    await rejected;
    expect(terminated).toBe(1);
    engine.dispose();
  });
  it('reports a deadline distinctly and closes the worker', async () => {
    const engine = new BrowserCompute({ workerUrl: '/worker.js', wasmUrl: '/engine.wasm' });
    const pending = engine.query('SELECT 1', []);
    const rejected = expect(pending).rejects.toThrow('time budget');
    await vi.advanceTimersByTimeAsync(30001);
    await rejected;
    expect(terminated).toBe(1);
    engine.dispose();
  });
  it('refuses cross-origin assets and disposed workspaces', async () => {
    const engine = new BrowserCompute({
      workerUrl: 'https://other.test/worker.js',
      wasmUrl: '/engine.wasm',
    });
    await expect(engine.query('SELECT 1', [])).rejects.toThrow('application origin');
    expect(fetch).not.toHaveBeenCalled();
    engine.dispose();
    await expect(engine.query('SELECT 1', [])).rejects.toThrow('closed');
  });
});
