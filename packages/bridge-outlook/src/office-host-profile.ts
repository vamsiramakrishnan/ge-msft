import type { HostCapabilityProfile, Surface } from '@ge/contracts';
import { hostCapabilityProfile } from '@ge/runtime';

// Native Office context stays in the bridge; runtime negotiation is pure.
export function officeHostProfile(surface: Surface): HostCapabilityProfile | undefined {
  if (typeof Office === 'undefined') return undefined;
  if (!Office.context?.requirements?.isSetSupported)
    return hostCapabilityProfile(surface, { isSetSupported: () => false });
  const context = Office.context;
  const item = context.mailbox?.item;
  const mode =
    surface === 'outlook'
      ? item && typeof item.subject !== 'string'
        ? 'compose'
        : 'read'
      : context.document?.mode === undefined
        ? 'unknown'
        : context.document.mode === Office.DocumentMode?.ReadOnly
          ? 'read'
          : context.document.mode === Office.DocumentMode?.ReadWrite
            ? 'write'
            : 'unknown';
  return hostCapabilityProfile(surface, {
    platform: String(context.platform ?? 'unknown'),
    build: context.diagnostics?.version ?? context.mailbox?.diagnostics?.hostVersion,
    mode,
    isSetSupported: (name, version) => context.requirements.isSetSupported(name, version),
  });
}
