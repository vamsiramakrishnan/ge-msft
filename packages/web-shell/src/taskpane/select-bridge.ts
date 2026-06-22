import type { Surface } from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import { WordBridge } from '@ge/bridge-word';
import { ExcelBridge } from '@ge/bridge-excel';
import { OutlookBridge } from '@ge/bridge-outlook';
import { PowerPointBridge } from '@ge/bridge-powerpoint';
import { OneNoteBridge } from '@ge/bridge-onenote';
import { TeamsBridge, type TeamsBridgeOptions } from '@ge/teams';

/**
 * Per-surface construction options the bootstrap may pass through. Today only Teams needs one
 * (its transcript snapshot / TeamsJS handle); the Office surfaces take no constructor argument.
 */
export interface BridgeFactoryOptions {
  teams?: TeamsBridgeOptions;
}

/**
 * The ONE place the surface-agnostic shell makes a host-specific choice: pick the `DocBridge`
 * implementation for the detected surface. Everything downstream (`composeSession`,
 * `PanelController`, the React panel) talks only to the `DocBridge` interface, so this is the
 * entire surface seam. Pure + side-effect-free (bridge constructors do no host I/O) so it is
 * unit-testable without an Office host.
 *
 * Returns `undefined` for a surface we don't yet wire here — the caller renders an
 * unsupported-host message rather than crashing.
 */
export function selectBridge(
  surface: Surface | undefined,
  options: BridgeFactoryOptions = {},
): DocBridge | undefined {
  switch (surface) {
    case 'word':
      return new WordBridge();
    case 'excel':
      return new ExcelBridge();
    case 'outlook':
      return new OutlookBridge();
    case 'powerpoint':
      return new PowerPointBridge();
    case 'onenote':
      return new OneNoteBridge();
    case 'teams':
      return new TeamsBridge(options.teams);
    default:
      return undefined;
  }
}

/** The surfaces this shell can currently construct a bridge for. */
export const SUPPORTED_SURFACES: readonly Surface[] = [
  'word',
  'excel',
  'outlook',
  'powerpoint',
  'onenote',
  'teams',
];

/** True when `selectBridge` will return a bridge for this surface. */
export function isSupportedSurface(surface: Surface | undefined): surface is Surface {
  return surface !== undefined && SUPPORTED_SURFACES.includes(surface);
}
