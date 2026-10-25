import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DEFAULT_BRAND, brand } from './brand.js';

const defaultJson = JSON.parse(
  readFileSync(resolve(__dirname, '../../../brands/default/brand.json'), 'utf8'),
) as Record<string, string>;

describe('brand', () => {
  it('falls back to the upstream default when no build-time brand is set', () => {
    expect(brand).toEqual(DEFAULT_BRAND);
  });

  it('matches brands/default/brand.json', () => {
    expect(DEFAULT_BRAND).toEqual({
      name: defaultJson.name,
      assistantName: defaultJson.assistantName,
      fullName: defaultJson.fullName,
    });
  });
});
