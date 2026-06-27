import { describe, it, expect } from 'vitest';
import { commentAdded, deriveOrigin, documentChanged, selectionChanged } from './events.js';

describe('excel event mappers (pure)', () => {
  describe('deriveOrigin', () => {
    it('maps the Remote enum value to remote', () => {
      expect(deriveOrigin('Remote')).toBe('remote');
    });

    it('is case-insensitive for remote', () => {
      expect(deriveOrigin('remote')).toBe('remote');
      expect(deriveOrigin('REMOTE')).toBe('remote');
    });

    it('maps the Local enum value to local', () => {
      expect(deriveOrigin('Local')).toBe('local');
    });

    it('defaults to local for missing or unknown source', () => {
      expect(deriveOrigin(undefined)).toBe('local');
      expect(deriveOrigin('')).toBe('local');
      expect(deriveOrigin('something-else')).toBe('local');
    });
  });

  describe('selectionChanged', () => {
    it('is always local and carries the address as preview', () => {
      expect(selectionChanged('Sheet1!A1:B3')).toEqual({
        type: 'selection-changed',
        surface: 'excel',
        origin: 'local',
        preview: 'Sheet1!A1:B3',
      });
    });

    it('omits preview when address is absent or empty', () => {
      expect(selectionChanged()).toEqual({
        type: 'selection-changed',
        surface: 'excel',
        origin: 'local',
      });
      expect(selectionChanged('')).toEqual({
        type: 'selection-changed',
        surface: 'excel',
        origin: 'local',
      });
    });
  });

  describe('documentChanged', () => {
    it('carries the derived origin', () => {
      expect(documentChanged('local')).toEqual({
        type: 'document-changed',
        surface: 'excel',
        origin: 'local',
      });
      expect(documentChanged('remote')).toEqual({
        type: 'document-changed',
        surface: 'excel',
        origin: 'remote',
      });
    });
  });

  describe('commentAdded', () => {
    it('builds one event per comment id with origin', () => {
      expect(commentAdded('cmt-1', 'remote')).toEqual({
        type: 'comment-added',
        surface: 'excel',
        origin: 'remote',
        commentId: 'cmt-1',
      });
    });

    it('includes optional text when provided', () => {
      expect(commentAdded('cmt-2', 'local', 'looks good')).toEqual({
        type: 'comment-added',
        surface: 'excel',
        origin: 'local',
        commentId: 'cmt-2',
        text: 'looks good',
      });
    });
  });
});
