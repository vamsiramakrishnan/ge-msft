import {
  ReconciliationSpecSchema,
  type ReconciliationSpec,
  type TableArtifact,
} from '@ge/contracts';
import { assertExactDecimalColumn, exactDecimalColumnSql } from './exact-decimal.js';
export { ReconciliationSpecSchema, type ReconciliationSpec } from '@ge/contracts';
export function reconciliationQuery(
  left: TableArtifact,
  right: TableArtifact,
  raw: ReconciliationSpec,
): string {
  const spec = ReconciliationSpecSchema.parse(raw);
  const c = (a: TableArtifact, index: number): string => {
    const col = a.columns[index];
    if (!col)
      throw new Error(
        `Column ${index + 1} (index ${index}) does not exist in ${a.title}; the captured range has ${a.columns.length} columns. Expand the range or choose another reconciliation column.`,
      );
    return col.name;
  };
  const side = (
    a: TableArtifact,
    key: number,
    amount: number,
    currency: number | undefined,
  ): string => {
    assertExactDecimalColumn(a, amount);
    const money = exactDecimalColumnSql(c(a, amount));
    const unit =
      currency === undefined
        ? `'${spec.currency}'`
        : `upper(trim(cast(${c(a, currency)} as varchar)))`;
    return `select nullif(trim(cast(${c(a, key)} as varchar)), '') as item_key, ${unit} as currency,
      sum(${money}) as total, count(*) as records,
      sum(case when ${money} is null then 1 else 0 end) as invalid_amounts
      from ${a.id} group by item_key, currency`;
  };
  const valid = `coalesce(i.item_key, p.item_key) is not null and coalesce(i.currency, p.currency) is not null and regexp_full_match(coalesce(i.currency, p.currency), '[A-Z]{3}') and coalesce(i.invalid_amounts, 0) + coalesce(p.invalid_amounts, 0) = 0`;
  return `with invoices as (${side(left, spec.leftKey, spec.leftAmount, spec.leftCurrency)}), payments as (${side(right, spec.rightKey, spec.rightAmount, spec.rightCurrency)})
    select coalesce(i.item_key, p.item_key) as item_key, coalesce(i.currency, p.currency) as currency,
      case when ${valid} then cast(coalesce(i.total, 0) as varchar) else null end as invoice_total,
      case when ${valid} then cast(coalesce(p.total, 0) as varchar) else null end as payment_total,
      case when ${valid} then cast(coalesce(i.total, 0) - coalesce(p.total, 0) as varchar) else null end as variance,
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
