import { isSet } from './capabilities-runtime.js';

interface CalculationPort {
  calculationState: string;
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
  application.calculate('Recalculate');
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
