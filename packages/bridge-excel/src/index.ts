/**
 * @ge/bridge-excel — the Excel DocBridge. Native context capture (selected range / used
 * range → table block) + address-anchored `write-cells` writes. Implements @ge/runtime's
 * DocBridge, so it plugs into the shared AssistSession loop unchanged.
 */
export { ExcelBridge, parseAddress } from './excel-bridge.js';
export { EXCEL_CAPABILITIES } from './capabilities.js';
export { rangeToContext, selectionValuesToContext, splitHeaderRows } from './capture.js';
export {
  deriveOrigin,
  selectionChanged,
  documentChanged,
  commentAdded,
  type ExcelEventSourceLike,
} from './events.js';
export {
  planWriteCells,
  planFormatCells,
  planAddComment,
  splitFormulaGrid,
  formatSourceComment,
  type WriteCellsPlan,
  type FormatCellsPlan,
  type AddCommentPlan,
  type FormulaGrid,
} from './actuate-plan.js';
export { provenanceRecord, provenanceKey, type ProvenanceRecord } from './provenance-record.js';
