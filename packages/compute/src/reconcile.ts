import { z } from 'zod';
import type { TableArtifact } from '@ge/contracts';
export const ReconciliationSpecSchema = z
  .object({
    left: z.string(),
    right: z.string(),
    leftKey: z.number().int().nonnegative(),
    rightKey: z.number().int().nonnegative(),
    leftAmount: z.number().int().nonnegative(),
    rightAmount: z.number().int().nonnegative(),
    leftCurrency: z.number().int().nonnegative().optional(),
    rightCurrency: z.number().int().nonnegative().optional(),
    currency: z
      .string()
      .regex(/^[A-Z]{3}$/)
      .optional(),
    tolerance: z
      .string()
      .regex(/^\d{1,12}(\.\d{1,6})?$/)
      .default('0.01'),
  })
  .refine(
    (s) => Boolean(s.currency) || (s.leftCurrency !== undefined && s.rightCurrency !== undefined),
    'Choose a currency or a currency column in both tables.',
  );
export type ReconciliationSpec = z.infer<typeof ReconciliationSpecSchema>;
export function reconciliationQuery(
  left: TableArtifact,
  right: TableArtifact,
  raw: ReconciliationSpec,
): string {
  const spec = ReconciliationSpecSchema.parse(raw);
  const c = (a: TableArtifact, index: number): string => {
    const col = a.columns[index];
    if (!col) throw new Error('A selected reconciliation column no longer exists.');
    return col.name;
  };
  const side = (
    a: TableArtifact,
    key: number,
    amount: number,
    currency: number | undefined,
  ): string => {
    const money = c(a, amount);
    const unit =
      currency === undefined
        ? `'${spec.currency}'`
        : `upper(trim(cast(${c(a, currency)} as varchar)))`;
    return `select nullif(trim(cast(${c(a, key)} as varchar)), '') as item_key, ${unit} as currency,
      sum(try_cast(${money} as decimal(38,6))) as total, count(*) as records,
      sum(case when try_cast(${money} as decimal(38,6)) is null then 1 else 0 end) as invalid_amounts
      from ${a.id} group by item_key, currency`;
  };
  return `with invoices as (${side(left, spec.leftKey, spec.leftAmount, spec.leftCurrency)}), payments as (${side(right, spec.rightKey, spec.rightAmount, spec.rightCurrency)})
    select coalesce(i.item_key, p.item_key) as item_key, coalesce(i.currency, p.currency) as currency,
      cast(coalesce(i.total, 0) as varchar) as invoice_total, cast(coalesce(p.total, 0) as varchar) as payment_total,
      cast(coalesce(i.total, 0) - coalesce(p.total, 0) as varchar) as variance,
      case when coalesce(i.item_key, p.item_key) is null then 'invalid'
        when coalesce(i.currency, p.currency) is null or not regexp_full_match(coalesce(i.currency, p.currency), '[A-Z]{3}') then 'invalid'
        when coalesce(i.invalid_amounts, 0) + coalesce(p.invalid_amounts, 0) > 0 then 'invalid'
        when i.item_key is null then 'unallocated' when p.item_key is null then 'unpaid'
        when abs(i.total - p.total) > cast('${spec.tolerance}' as decimal(38,6)) then 'variance'
        else 'matched' end as status,
      coalesce(i.records, 0) as invoice_rows, coalesce(p.records, 0) as payment_rows
    from invoices i full outer join payments p on i.item_key = p.item_key and i.currency = p.currency
    order by status, currency, item_key`;
}
