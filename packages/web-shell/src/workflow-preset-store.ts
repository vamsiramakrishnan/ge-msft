import { parseWorkflowPreset } from '@ge/runtime';

type WorkflowPreset = ReturnType<typeof parseWorkflowPreset>;
type PresetStorage = Pick<Storage, 'getItem' | 'setItem'>;
const STORAGE_KEY = 'ge.workflow-settings.v1';
const MAX_BYTES = 32_768;
const MAX_PRESETS = 12;

/** Explicitly saved settings only. Results, credentials and approval receipts never enter this store. */
export class WorkflowPresetStore {
  constructor(private readonly storage: PresetStorage) {}

  list(): WorkflowPreset[] {
    const raw = this.storage.getItem(STORAGE_KEY);
    if (raw === null) return [];
    if (new TextEncoder().encode(raw).length > MAX_BYTES)
      throw new Error('Saved settings are too large to load. Clear saved settings to start again.');
    const value: unknown = JSON.parse(raw);
    if (!Array.isArray(value) || value.length > MAX_PRESETS)
      throw new Error('Saved settings could not be read. Clear saved settings to start again.');
    const presets = value.map((item: unknown) => parseWorkflowPreset(item));
    if (new Set(presets.map((item) => item.recipeId)).size !== presets.length)
      throw new Error(
        'Saved settings contain duplicate workflows. Clear saved settings to start again.',
      );
    return presets;
  }

  save(raw: unknown): void {
    const preset = parseWorkflowPreset(raw);
    const presets = this.list().filter((item) => item.recipeId !== preset.recipeId);
    presets.push(preset);
    if (presets.length > MAX_PRESETS) throw new Error('The saved-settings limit has been reached.');
    const serialized = JSON.stringify(presets);
    if (new TextEncoder().encode(serialized).length > MAX_BYTES)
      throw new Error('These settings are too large to save. Shorten the source ranges.');
    this.storage.setItem(STORAGE_KEY, serialized);
  }

  remove(recipeId: string): void {
    this.storage.setItem(
      STORAGE_KEY,
      JSON.stringify(this.list().filter((p) => p.recipeId !== recipeId)),
    );
  }

  clear(): void {
    this.storage.setItem(STORAGE_KEY, '[]');
  }
}

export function browserWorkflowPresets(): WorkflowPresetStore | undefined {
  try {
    return typeof window === 'undefined' ? undefined : new WorkflowPresetStore(window.localStorage);
  } catch {
    return undefined;
  }
}
