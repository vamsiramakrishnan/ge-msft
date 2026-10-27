import { isSet } from './capabilities-runtime.js';

/**
 * Calculation readback after a formula write.
 *
 * In `Automatic` calculation mode the host already recalculated at the sync that committed the
 * write (writes queue under `suspendApiCalculationUntilNextSync`, see {@link suspendCalculation}),
 * so this only waits for `calculationState === 'Done'`. Only a non-automatic mode (`Manual`,
 * `AutomaticExceptTables`) forces an explicit `calculate('Recalculate')` first — otherwise the
 * readback would hash stale values.
 */
interface CalculationPort {
  calculationState: string;
  calculationMode?: string;
  load(properties: string): unknown;
  calculate(type: 'Recalculate'): void;
}

/** Ending suspension is not proof of completed calculation. Bound both polls and elapsed time. */
export async function settleCalculation(
  application: CalculationPort,
  sync: () => Promise<unknown>,
  maxPolls = 12,
): Promise<void> {
  const deadline = performance.now() + 5000;
  application.load('calculationMode');
  await sync();
  if (application.calculationMode !== 'Automatic') application.calculate('Recalculate');
  for (let poll = 0; poll < maxPolls; poll++) {
    application.load('calculationState');
    await sync();
    if (application.calculationState === 'Done') return;
    if (performance.now() >= deadline) break;
    await new Promise<void>((resolve) => setTimeout(resolve, 25));
  }
  throw new Error('Excel calculation did not complete within the verification budget.');
}

export async function verifyCalculationReady(context: Excel.RequestContext): Promise<void> {
  if (!isSet('ExcelApi', '1.9')) throw new Error('This host cannot report completed calculation.');
  await settleCalculation(context.application, () => context.sync());
}

/**
 * Defer calculation until the batch syncs (ExcelApi 1.6) so a multi-cell formula write triggers
 * one recalculation instead of one per queued request. A no-op below 1.6 or when the host lacks
 * the member; calculation then simply happens as it always did.
 */
export function suspendCalculation(context: Excel.RequestContext): void {
  if (!isSet('ExcelApi', '1.6')) return;
  const application = context.application as { suspendApiCalculationUntilNextSync?: () => void };
  try {
    application.suspendApiCalculationUntilNextSync?.();
  } catch {
    // Suspension is an optimisation; a host that refuses it still calculates at the sync.
  }
}
