/**
 * The cross-surface `Office` fake — the slice of `Office.context` the bridges drive, independent of
 * which editor is open. Built fresh per `installFake*` so each surface gets its own settings bag /
 * handler registry / requirement set, and asserted via {@link OfficeSeed} after a run.
 *
 * Enumerated host calls modelled (the fidelity boundary for `Office`):
 *   - `Office.context.requirements.isSetSupported(name, version)` — the capability gate every
 *     bridge uses (`capabilities-runtime.ts`). Backed by a seeded requirement-set map.
 *   - `Office.context.document.settings.set/get/saveAsync(...)` — Excel durable provenance
 *     (`ExcelBridge.persistProvenance`). Writes record into {@link OfficeSeed.settings}.
 *   - `Office.context.document.customXmlParts.add(xml)` — Word durable provenance
 *     (`OfficeWordHost.persistProvenance`). Appends into {@link OfficeSeed.customXmlParts}.
 *   - `Office.context.document.addHandlerAsync/removeHandlerAsync(eventType, handler)` +
 *     `Office.EventType.{DocumentSelectionChanged,ActiveViewChanged}` — host-event registration
 *     (Word selection + PowerPoint selection/view `watch()`). Handlers are retained so a test can
 *     fire them via {@link OfficeHandlerRegistry.fire}.
 */

/** A requirement-set support map: `"ExcelApi"` → the highest minor it supports (e.g. `12` for 1.12). */
export type RequirementSets = Readonly<Record<string, number>>;

/** The mutable Office-level seed a simulator threads through and a test asserts against. */
export interface OfficeSeed {
  /** `name@minor` requirement-set support (e.g. `{ ExcelApi: 12 }` ⇒ supports ExcelApi ≤ 1.12). */
  requirements: RequirementSets;
  /** Durable settings bag (Excel provenance): key → JSON-ish value last `set()`. */
  settings: Map<string, unknown>;
  /** Whether `settings.saveAsync()` was invoked at least once (Excel persists then saves). */
  settingsSaved: boolean;
  /** Custom XML parts added (Word provenance): each is the raw XML string passed to `add()`. */
  customXmlParts: string[];
}

/** Build an Office seed with sensible modern defaults; callers override per surface. */
export function makeOfficeSeed(requirements: RequirementSets): OfficeSeed {
  return {
    requirements,
    settings: new Map<string, unknown>(),
    settingsSaved: false,
    customXmlParts: [],
  };
}

/** Retained host-event handlers, so a test can simulate the host firing selection/view events. */
export interface OfficeHandlerRegistry {
  /** Fire every handler registered for `eventType` (no-op when none). */
  fire(eventType: string): void;
  /** How many handlers are currently registered for `eventType` (for teardown assertions). */
  count(eventType: string): number;
}

interface FakeRequirements {
  isSetSupported(name: string, version?: string): boolean;
}

interface FakeSettings {
  set(name: string, value: unknown): void;
  get(name: string): unknown;
  remove(name: string): void;
  saveAsync(callback?: (result: { status: string }) => void): void;
}

interface FakeCustomXmlParts {
  add(xml: string): { id: string };
}

interface FakeDocument {
  settings: FakeSettings;
  customXmlParts: FakeCustomXmlParts;
  addHandlerAsync(
    eventType: string,
    handler: (args?: unknown) => void,
    callback?: (result: { status: string }) => void,
  ): void;
  removeHandlerAsync(
    eventType: string,
    options?: { handler?: (args?: unknown) => void },
    callback?: (result: { status: string }) => void,
  ): void;
}

/** The fake `Office` namespace object installed onto `globalThis.Office`. */
export interface FakeOffice {
  context: {
    requirements: FakeRequirements;
    document: FakeDocument;
  };
  EventType: Readonly<Record<string, string>>;
}

/** The Office event names the bridges reference (mirrors `Office.EventType`). */
const EVENT_TYPE: Readonly<Record<string, string>> = {
  DocumentSelectionChanged: 'documentSelectionChanged',
  ActiveViewChanged: 'activeViewChanged',
};

/** True iff `have@haveMinor` covers the requested `name@version` (parsed `major.minor`). */
function supports(sets: RequirementSets, name: string, version?: string): boolean {
  const haveMinor = sets[name];
  if (haveMinor === undefined) return false;
  if (version === undefined) return true;
  const [, minorStr = '0'] = version.split('.');
  const wantMinor = Number.parseInt(minorStr, 10);
  return Number.isFinite(wantMinor) && haveMinor >= wantMinor;
}

/**
 * Build the `Office` fake over a shared {@link OfficeSeed}. Returns the namespace object to install
 * plus a {@link OfficeHandlerRegistry} so a test can fire host events the bridges' `watch()` wired.
 */
export function makeFakeOffice(seed: OfficeSeed): {
  office: FakeOffice;
  handlers: OfficeHandlerRegistry;
} {
  const handlerMap = new Map<string, Array<(args?: unknown) => void>>();

  const settings: FakeSettings = {
    set(name, value) {
      seed.settings.set(name, value);
    },
    get(name) {
      return seed.settings.get(name);
    },
    remove(name) {
      seed.settings.delete(name);
    },
    saveAsync(callback) {
      seed.settingsSaved = true;
      callback?.({ status: 'succeeded' });
    },
  };

  const customXmlParts: FakeCustomXmlParts = {
    add(xml) {
      seed.customXmlParts.push(xml);
      return { id: `xml-${seed.customXmlParts.length}` };
    },
  };

  const document: FakeDocument = {
    settings,
    customXmlParts,
    addHandlerAsync(eventType, handler, callback) {
      const list = handlerMap.get(eventType) ?? [];
      list.push(handler);
      handlerMap.set(eventType, list);
      callback?.({ status: 'succeeded' });
    },
    removeHandlerAsync(eventType, options, callback) {
      const list = handlerMap.get(eventType) ?? [];
      const target = options?.handler;
      handlerMap.set(eventType, target ? list.filter((h) => h !== target) : []);
      callback?.({ status: 'succeeded' });
    },
  };

  const office: FakeOffice = {
    context: {
      requirements: {
        isSetSupported: (name, version) => supports(seed.requirements, name, version),
      },
      document,
    },
    EventType: EVENT_TYPE,
  };

  const handlers: OfficeHandlerRegistry = {
    fire(eventType) {
      for (const h of handlerMap.get(eventType) ?? []) h();
    },
    count(eventType) {
      return (handlerMap.get(eventType) ?? []).length;
    },
  };

  return { office, handlers };
}
