import type { TaskBlocker } from '@ge/contracts';
import { brand } from './brand.js';

/**
 * The pane follows the language the user started the conversation in (English or Bahasa
 * Indonesia). There is no language setting: the first request decides, and a new chat decides again.
 */
export type UiLanguage = 'en' | 'id';

/** Common Bahasa Indonesia words that rarely appear in English requests. */
const BAHASA =
  /\b(?:yang|dan|ini|itu|untuk|dengan|dari|tolong|buat(?:kan)?|tambah(?:kan)?|ubah|ganti|apa|bagaimana|berapa|mana|ringkas(?:kan)?|jelaskan|sebutkan|analisis|semua|setiap|baris|kolom|tabel|grafik|slide nya|surat|balas(?:kan)?|tidak|sudah|belum|saya|kami|kita|tahun|bulan|wilayah|cabang)\b/gi;

/** Bahasa when at least two Bahasa words appear (one can be a shared word such as "data"). */
export function detectLanguage(text: string): UiLanguage {
  return (text.match(BAHASA) ?? []).length >= 2 ? 'id' : 'en';
}

const STRINGS = {
  en: {
    doneOne: 'Done: 1 change',
    doneMany: (n: number) => `Done: ${n} changes`,
    showMe: 'Show me',
    undo: 'Undo',
    undoAll: 'Undo all',
    undone: 'Undone',
    undoInOffice: 'To undo, use Undo (Ctrl+Z) in Office.',
    undoRequest: 'Undo the last change',
    undoGone: 'Undo is no longer available for this change.',
    failedTitle: "That didn't finish",
    failedPartial: (made: number) =>
      made === 1
        ? '1 change was made before it stopped.'
        : `${made} changes were made before it stopped.`,
    tryAgain: 'Try again',
    blockedTitle: 'This request was blocked',
    blockedBody:
      "Your organisation's safety settings stopped this request. Try asking in a different way.",
    resumeTitle: 'Your last request was interrupted',
    resumeBody: (request: string, made: number) =>
      `"${request}": ${made === 0 ? 'nothing was changed yet' : made === 1 ? '1 change was made' : `${made} changes were made`} before the pane closed.`,
    resumeFinish: 'Finish it',
    resumeDismiss: 'Dismiss',
    planTitle: 'Here is what I plan to do',
    planGo: 'Go ahead',
    planChange: 'Change',
    planChangeAsk: 'What should change in this plan?',
    revealFailed: 'Could not show that place in the document. Select it in Office directly.',
    attachFailed: 'Could not attach that item. Try again, or select it in Office first.',
    planChangeTitle: 'Change the plan',
    clarifyTitle: 'One question before I plan',
    clarifyAnswer: 'Your answer',
    clarifyContinue: 'Continue',
    cancel: 'Cancel',
    planIn: 'In',
    planSelection: 'Using your selection',
    planKeep: 'Leave unchanged:',
    noChangeMade: 'No changes were made.',
    planUsing: 'Using',
    planAnEmail: 'an email',
    planAnItem: 'an item',
    reviewOne: 'Review this change',
    reviewMany: (n: number) => `Review ${n} changes`,
    apply: 'Apply',
    applyNote: 'Nothing changes until you press Apply.',
    applyOrder: 'Changes are applied in order. If one fails, the earlier ones stay.',
    check: 'Check',
    where: 'Where',
    becomes: 'Becomes',
    beforeAfter: 'Before → after',
    noPreview: 'No preview for this change. It runs as shown above.',
  },
  id: {
    doneOne: 'Selesai: 1 perubahan',
    doneMany: (n: number) => `Selesai: ${n} perubahan`,
    showMe: 'Tunjukkan',
    undo: 'Urungkan',
    undoAll: 'Urungkan semua',
    undone: 'Diurungkan',
    undoInOffice: 'Untuk mengurungkan, gunakan Undo (Ctrl+Z) di Office.',
    undoRequest: 'Urungkan perubahan terakhir',
    undoGone: 'Urungkan tidak tersedia lagi untuk perubahan ini.',
    failedTitle: 'Maaf, belum berhasil',
    failedPartial: (made: number) => `${made} perubahan sudah dibuat sebelum berhenti.`,
    tryAgain: 'Coba lagi',
    blockedTitle: 'Permintaan ini diblokir',
    blockedBody:
      'Pengaturan keamanan organisasi Anda menghentikan permintaan ini. Coba tanyakan dengan cara lain.',
    resumeTitle: 'Permintaan terakhir Anda terputus',
    resumeBody: (request: string, made: number) =>
      `"${request}": ${made === 0 ? 'belum ada yang diubah' : `${made} perubahan sudah dibuat`} sebelum panel ditutup.`,
    resumeFinish: 'Selesaikan',
    resumeDismiss: 'Tutup',
    planTitle: 'Ini rencana saya',
    planGo: 'Lanjutkan',
    planChange: 'Ubah',
    planChangeAsk: 'Apa yang perlu diubah dari rencana ini?',
    revealFailed: 'Tidak dapat menampilkan bagian itu di dokumen. Pilih langsung di Office.',
    attachFailed: 'Tidak dapat melampirkan item itu. Coba lagi, atau pilih dulu di Office.',
    planChangeTitle: 'Ubah rencana',
    clarifyTitle: 'Satu pertanyaan sebelum saya merencanakan',
    clarifyAnswer: 'Jawaban Anda',
    clarifyContinue: 'Lanjutkan',
    cancel: 'Batal',
    planIn: 'Di',
    planSelection: 'Menggunakan pilihan Anda',
    planKeep: 'Tidak diubah:',
    noChangeMade: 'Tidak ada perubahan yang dibuat.',
    planUsing: 'Menggunakan',
    planAnEmail: 'sebuah email',
    planAnItem: 'sebuah item',
    reviewOne: 'Periksa perubahan ini',
    reviewMany: (n: number) => `Periksa ${n} perubahan`,
    apply: 'Terapkan',
    applyNote: 'Tidak ada yang berubah sebelum Anda menekan Terapkan.',
    applyOrder: 'Perubahan diterapkan berurutan. Jika satu gagal, yang sebelumnya tetap.',
    check: 'Periksa',
    where: 'Lokasi',
    becomes: 'Menjadi',
    beforeAfter: 'Sebelum → sesudah',
    noPreview: 'Tidak ada pratinjau untuk perubahan ini. Akan dijalankan seperti di atas.',
  },
} as const;

export type UiStrings = (typeof STRINGS)['en'];

export function uiStrings(lang: UiLanguage): UiStrings {
  return STRINGS[lang] as UiStrings;
}

const KIND_LABEL: Record<UiLanguage, Record<string, string>> = {
  en: {
    'write-cells': 'Cells',
    'format-cells': 'Format',
    'create-table': 'Table',
    'insert-chart': 'Chart',
    'format-conditional': 'Conditional format',
    'add-comment': 'Comment',
    'comment-reply': 'Reply to comment',
    'tracked-change': 'Suggested edit',
    'replace-selection': 'Selection text',
    'find-replace': 'Find and replace',
    'insert-text': 'Text',
    'insert-ooxml': 'Content',
    'insert-table': 'Table',
    'apply-style': 'Style',
    'insert-slide': 'New slide',
    'add-table-slide': 'Table slide',
    'set-shape-text': 'Slide text',
    'add-shape': 'Shape',
    'format-shape': 'Shape format',
    'reply-mail': 'Reply draft',
    'create-mail': 'New email',
    'set-recipients': 'Recipients',
    'insert-hyperlink': 'Link',
    'add-attachment': 'Attachment',
    'delete-slide': 'Slide removal',
    'manage-worksheet': 'Sheet',
    'resolve-revisions': 'Tracked changes',
    'insert-pivot': 'Pivot table',
    'sort-range': 'Sort',
    'filter-range': 'Filter',
    'format-chart': 'Chart format',
    'insert-image': 'Image',
    'apply-slide-layout': 'Slide layout',
  },
  id: {
    'write-cells': 'Sel',
    'format-cells': 'Format',
    'create-table': 'Tabel',
    'insert-chart': 'Grafik',
    'format-conditional': 'Format bersyarat',
    'add-comment': 'Komentar',
    'comment-reply': 'Balasan komentar',
    'tracked-change': 'Usulan perubahan',
    'replace-selection': 'Teks pilihan',
    'find-replace': 'Cari dan ganti',
    'insert-text': 'Teks',
    'insert-ooxml': 'Konten',
    'insert-table': 'Tabel',
    'apply-style': 'Gaya',
    'insert-slide': 'Slide baru',
    'add-table-slide': 'Slide tabel',
    'set-shape-text': 'Teks slide',
    'add-shape': 'Bentuk',
    'format-shape': 'Format bentuk',
    'reply-mail': 'Draf balasan',
    'create-mail': 'Email baru',
    'set-recipients': 'Penerima',
    'insert-hyperlink': 'Tautan',
    'add-attachment': 'Lampiran',
    'delete-slide': 'Penghapusan slide',
    'manage-worksheet': 'Sheet',
    'resolve-revisions': 'Perubahan terlacak',
    'insert-pivot': 'Pivot table',
    'sort-range': 'Urutkan',
    'filter-range': 'Filter',
    'format-chart': 'Format grafik',
    'insert-image': 'Gambar',
    'apply-slide-layout': 'Tata letak slide',
  },
};

/**
 * The one sentence a task that stopped for the user ends with: what blocked it and what to do.
 * Built from closed codes and the target address only, never from host or model text.
 */
export function blockerText(blocker: TaskBlocker, lang: UiLanguage): string {
  const id = lang === 'id';
  if (blocker.cause === 'not_supported') {
    const can = blocker.available.map((kind) => kindLabel(kind, lang).toLowerCase()).join(', ');
    if (blocker.kind === undefined)
      return id
        ? `${brand.name} belum bisa melakukan ini dengan perubahan yang tersedia di sini. Yang bisa saya bantu: ${can}.`
        : `${brand.name} can't do this yet with the changes available here. I can help with: ${can}.`;
    const what = kindLabel(blocker.kind, lang);
    return id
      ? `${brand.name} belum bisa melakukan ini: ${what}. Yang bisa saya bantu: ${can}.`
      : `${brand.name} can't do this yet: ${what}. I can help with: ${can}.`;
  }
  if (blocker.cause === 'no_progress')
    return id
      ? 'Saya tidak bisa menyelesaikan ini: saya sudah memeriksa dokumen, tetapi tidak menemukan cara melakukannya dengan perubahan yang bisa saya buat di sini. Coba jelaskan dengan cara lain, atau lakukan langkah ini langsung di Office.'
      : "I couldn't finish this: I checked the document but found no way to do it with the changes I can make here. Try asking in a different way, or do this step in Office directly.";
  const where = blocker.location ?? (id ? 'Bagian ini' : 'This part of the document');
  switch (blocker.reason) {
    case 'protected':
      return id
        ? `${where} dilindungi atau hanya-baca, jadi tidak bisa saya ubah. Matikan proteksi (di Excel: Review → Unprotect Sheet), lalu minta lagi.`
        : `${where} is protected or read-only, so I couldn't change it. Turn off protection (in Excel: Review → Unprotect Sheet), then ask again.`;
    case 'edit_mode':
      return id
        ? 'Excel sedang mengedit sebuah sel. Tekan Enter atau Esc untuk selesai mengedit, lalu minta lagi.'
        : 'Excel is still editing a cell. Press Enter or Esc to finish editing, then ask again.';
    case 'conflict':
      return id
        ? `${where} bertabrakan dengan sel gabungan, tabel, filter, atau PivotTable, jadi tidak bisa saya ubah. Pilih tempat lain, lalu minta lagi.`
        : `${where} overlaps merged cells, a table, a filter or a PivotTable, so I couldn't change it. Choose another place, then ask again.`;
    case 'not_found':
      return id
        ? `${where} tidak ditemukan. Mungkin sudah diganti nama atau dihapus. Periksa, lalu minta lagi.`
        : `I couldn't find ${where}. It may have been renamed or deleted. Check it, then ask again.`;
    case 'unsupported_api':
      return id
        ? `Versi Office ini belum mendukung perubahan ini (${kindLabel(blocker.kind, lang)}). Coba di Office versi web atau versi terbaru.`
        : `This version of Office can't make this change (${kindLabel(blocker.kind, lang)}). Try Office on the web or a newer version.`;
  }
}

/** A plain noun for a change kind (read before and after it runs, so never past tense), falling back to the kind with dashes as spaces. */
export function kindLabel(kind: string, lang: UiLanguage): string {
  return KIND_LABEL[lang][kind] ?? kind.replace(/-/g, ' ');
}
