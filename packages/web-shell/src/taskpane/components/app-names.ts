/** Display names for the Office app a chat ran in, and what "this" means there. */
export const APP_NAME: Readonly<Record<string, string>> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  outlook: 'Outlook',
  onenote: 'OneNote',
  teams: 'Teams',
};

export const APP_ITEM: Readonly<Record<string, string>> = {
  word: 'document',
  excel: 'workbook',
  powerpoint: 'presentation',
  outlook: 'email',
  onenote: 'page',
  teams: 'meeting',
};

export const appName = (surface: string): string => APP_NAME[surface] ?? surface;
export const appItem = (surface: string): string => APP_ITEM[surface] ?? 'file';
