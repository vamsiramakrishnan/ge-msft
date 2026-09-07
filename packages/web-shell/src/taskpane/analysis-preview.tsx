import { createRoot } from 'react-dom/client';
import { App } from './components/App.js';
import {
  AssistSession,
  RuntimeHooks,
  completedEffectsExtension,
  installRuntimeExtensions,
  type DocBridge,
} from '@ge/runtime';
import { TriggerRegistry } from '@ge/triggers';
import {
  makeCellSnapshot,
  cellsMatchRequest,
  gridForRequest,
  formulasForRequest,
  intentsForManifest,
  type CellValue,
} from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import { EXCEL_CAPABILITIES } from '@ge/bridge-excel';
import { PanelController } from '../controller.js';
import './styles.css';
import './workspace.css';

// Deliberately separate dev entry. Real runtime, approval state machine and WASM; simulated Office.
const initial: Record<string, CellValue[][]> = {
  Invoices: [
    ['Invoice', 'Amount', 'Currency'],
    ['INV-01', '100.00', 'USD'],
    ['INV-02', '125.55', 'USD'],
    ['INV-03', '75.10', 'EUR'],
    ['INV-04', '20.00', 'USD'],
    ['INV-05', 'invalid', 'USD'],
  ],
  Payments: [
    ['Invoice', 'Amount', 'Currency'],
    ['INV-01', '100.00', 'USD'],
    ['INV-02', '120.55', 'USD'],
    ['INV-03', '75.10', 'EUR'],
    ['INV-99', '12.50', 'USD'],
  ],
  Results: [],
};
const load = <T,>(key: string, fallback: T): T => {
  try {
    return (JSON.parse(localStorage.getItem(key) ?? 'null') as T) ?? fallback;
  } catch {
    return fallback;
  }
};
const sheets = load('ge-analysis-demo-cells', initial);
let failCheckpoint = false;
const span = (address: string) => {
  const m = /^(Invoices|Payments|Results)!([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address);
  if (!m) throw new Error('Use an explicit range in Invoices, Payments or Results.');
  const col = (x: string): number => [...x].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(m[3]) - 1;
  const column = col(m[2]!);
  const rows = Number(m[5] ?? m[3]) - row;
  const columns = col(m[4] ?? m[2]!) - column + 1;
  if (rows < 1 || columns < 1 || rows * columns > 100000)
    throw new Error('Invalid range dimensions.');
  return { sheet: m[1]!, row, column, rows, columns };
};
const bridge: DocBridge = {
  surface: 'excel',
  getCapabilities: () => EXCEL_CAPABILITIES,
  listContext: async () =>
    ['Invoices!A1:C6', 'Payments!A1:C5'].map((range) => ({
      id: `xl:${range}`,
      kind: 'range',
      surface: 'excel',
      title: range,
    })),
  resolveContext: async () => [],
  captureCells: async (address) => {
    const s = span(address);
    return makeCellSnapshot({
      surface: 'excel',
      documentId: 'analysis-demo',
      objectId: s.sheet,
      locator: address,
      values: Array.from({ length: s.rows }, (_, r) =>
        Array.from(
          { length: s.columns },
          (_, c) => sheets[s.sheet]?.[s.row + r]?.[s.column + c] ?? '',
        ),
      ),
      formulas: Array.from({ length: s.rows }, () => Array<string>(s.columns).fill('')),
    });
  },
  recoveryStorage: {
    load: async () => load('ge-analysis-demo-recovery', []),
    save: async (_owner, records) => {
      if (failCheckpoint) {
        failCheckpoint = false;
        throw new Error('Simulated checkpoint failure.');
      }
      localStorage.setItem('ge-analysis-demo-recovery', JSON.stringify(records));
    },
  },
  actuate: async (request) => {
    const address = request.params.target?.range ?? '';
    const before = await bridge.captureCells!(address);
    for (const condition of request.preconditions ?? []) {
      const current = await bridge.captureCells!(condition.locator);
      if (condition.hash !== current.hash)
        return {
          ok: false,
          changeId: request.changeId,
          kind: request.kind,
          error: { code: 'stale_source', message: 'Source or destination changed after preview.' },
        };
    }
    if (formulasForRequest(request).some((r) => r.some(Boolean)))
      throw new Error('The simulated host supports literal cells only.');
    const s = span(address);
    const values = gridForRequest(request);
    for (let r = 0; r < s.rows; r++) {
      const sheet = sheets[s.sheet]!;
      sheet[s.row + r] ??= [];
      for (let c = 0; c < s.columns; c++) sheet[s.row + r]![s.column + c] = values[r]?.[c] ?? '';
    }
    localStorage.setItem('ge-analysis-demo-cells', JSON.stringify(sheets));
    const after = await bridge.captureCells!(address);
    return {
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: address,
      verification: {
        status: cellsMatchRequest(after, request) ? 'verified' : 'mismatch',
        beforeHash: before.hash,
        afterHash: after.hash,
      },
      inverse: { op: 'restore-cells', range: address, values: before.values },
    };
  },
};
const hooks = new RuntimeHooks();
installRuntimeExtensions([completedEffectsExtension], { hooks, triggers: new TriggerRegistry() });
const client = {
  stream: () => {
    throw new Error('This verification harness makes no model calls.');
  },
} as unknown as StreamAssistClient;
const session = new AssistSession(bridge, client, {
  unit: { connectors: [], surfaceContext: { kind: 'excel' } },
  hooks,
  recoveryOwner: 'demo-user',
  context: { docState: false, lazyRead: false },
  compute: async () => {
    const { createBrowserCompute } = await import('@ge/compute/browser');
    return createBrowserCompute({
      workerUrl: '/compute/duckdb-worker.js',
      wasmUrl: '/compute/duckdb.wasm',
    });
  },
});
const controller = new PanelController(session, bridge);
void controller.refreshContext();
function Preview(): JSX.Element {
  return (
    <div style={{ padding: 20, background: '#f2efec', minHeight: '100vh' }}>
      <div style={{ maxWidth: 820, margin: '0 auto' }}>
        <p style={{ fontSize: 13 }}>Real DuckDB · simulated Excel host · no model calls</p>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          <button
            onClick={() => {
              sheets.Invoices![1]![1] = '101.00';
              localStorage.setItem('ge-analysis-demo-cells', JSON.stringify(sheets));
            }}
          >
            Change invoice value
          </button>
          <button
            onClick={() => {
              failCheckpoint = true;
            }}
          >
            Fail next checkpoint
          </button>
          <button
            onClick={() => {
              localStorage.removeItem('ge-analysis-demo-cells');
              localStorage.removeItem('ge-analysis-demo-recovery');
              location.reload();
            }}
          >
            Reset fixture
          </button>
        </div>
        <div
          style={{
            width: 'min(100%, 390px)',
            height: 860,
            margin: '0 auto',
            background: 'white',
            border: '1px solid #ccc',
          }}
        >
          <App
            controller={controller}
            surface="excel"
            allowedIntents={intentsForManifest(EXCEL_CAPABILITIES)}
          />
        </div>
      </div>
    </div>
  );
}
createRoot(document.getElementById('root')!).render(<Preview />);
