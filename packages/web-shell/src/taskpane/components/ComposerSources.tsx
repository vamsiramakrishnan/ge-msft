import type { GroundSource } from '@ge/contracts';
import type { ComposerMention } from './Composer.js';

export type MentionOptions = Partial<Record<GroundSource, { id: string; label: string }[]>>;
export function mentionKey(mention: ComposerMention): string {
  return JSON.stringify([mention.kind, mention.ref ?? '']);
}
export function mentionLabel(mention: ComposerMention, options?: MentionOptions): string {
  if (mention.kind === 'this') return 'Current context';
  if (mention.kind === 'unit') return 'Attached sources';
  return (
    options?.[mention.kind]?.find((option) => option.id === mention.ref)?.label ??
    mention.ref ??
    mention.kind
  );
}
