import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Excel can read and write. Static today; `ExcelBridge.getCapabilities` will narrow it
 * at runtime against `Office.context.requirements.isSetSupported('ExcelApi', …)` so
 * unsupported actions (e.g. linked-entity cards on a thin platform) are never offered.
 */
export const EXCEL_CAPABILITIES: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['selection', 'range', 'sheet', 'table'],
  actuations: [
    {
      kind: 'write-cells',
      surface: 'excel',
      title: 'Write cells',
      description: 'Write values/formulas into a worksheet range.',
      reversible: true,
      appliesTo: ['range'],
    },
    {
      kind: 'format-cells',
      surface: 'excel',
      title: 'Format cells',
      description: 'Apply bold/italic/fill/number-format to a worksheet range.',
      reversible: true,
      appliesTo: ['range'],
    },
    {
      kind: 'add-comment',
      surface: 'excel',
      title: 'Add comment',
      description: 'Attach a new comment to a cell.',
      reversible: true,
      appliesTo: ['range'],
    },
    {
      kind: 'comment-reply',
      surface: 'excel',
      title: 'Reply & resolve comment',
      reversible: true,
      appliesTo: ['comment'],
    },
  ],
};
