/**
 * Office host simulators + full-stack UI integration harness.
 *
 * A clean, extractable module that fakes the GLOBAL Office object model (`Excel.run` / `Word.run` /
 * `PowerPoint.run` / `Office.context`) with in-memory, seeded data so the REAL per-surface bridges
 * run unchanged, letting an integration test drive the whole stack
 * (bridge → runtime `AssistSession` → `PanelController` → React `<App/>`) against simulated host
 * data — including reversible, gated, provenanced WRITES that mutate the fake host.
 *
 * See `packages/web-shell/test-harness/README.md` for the seed shapes, the simulator APIs, how to
 * add a surface, and how to write a new integration test.
 */

export { installFakeExcel } from './fake-excel.js';
export type {
  ExcelSimulator,
  ExcelSeed,
  SheetSeed,
  NamedRangeSeed,
  CommentSeed,
  RangeFormatSeed,
  ExcelSnapshot,
  ExcelEvents,
} from './fake-excel.js';
export { excelSeed, defaultExcelSeed } from './fake-excel.js';

export { installFakeWord } from './fake-word.js';
export type {
  WordSimulator,
  WordSeed,
  WordParagraphSeed,
  WordCommentSeed,
  WordSnapshot,
} from './fake-word.js';
export { wordSeed, defaultWordSeed } from './fake-word.js';

export { installFakePowerPoint } from './fake-powerpoint.js';
export type {
  PowerPointSimulator,
  PowerPointSeed,
  SlideSeed,
  ShapeSeed,
  PowerPointSnapshot,
} from './fake-powerpoint.js';
export { powerPointSeed, defaultPowerPointSeed } from './fake-powerpoint.js';

export { scriptedClient } from './scripted-client.js';
export type { ScriptedClient, ScriptedTurn } from './scripted-client.js';

export { makeOfficeSeed } from './fake-office.js';
export type { OfficeSeed, RequirementSets, OfficeHandlerRegistry } from './fake-office.js';

export { mountStack } from './mount-stack.js';
export type { MountedStack, MountStackOptions } from './mount-stack.js';
