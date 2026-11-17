/**
 * Utility for formatting document filenames for computer & Google Drive storage.
 * Converts '/' to '.', keeps Vietnamese diacritics and natural casing in the document title,
 * and formats as "<Số ký hiệu> - <Tên văn bản>".
 *
 * Examples:
 * - 09-KH/ĐTN + Kế hoạch tổ chức hoạt động chào mừng ngày 20-11 (.docx)
 *   => 09-KH.ĐTN - Kế hoạch tổ chức hoạt động chào mừng ngày 20-11.docx
 * - 1234/SGDĐT-GDTrH + Hướng dẫn tổ chức hoạt động giáo dục (.pdf)
 *   => 1234.SGDĐT-GDTrH - Hướng dẫn tổ chức hoạt động giáo dục.pdf
 */

export function removeVietnameseDiacritics(str: string): string {
  if (!str) return '';
  return str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D');
}

/**
 * Cleans string for safe filesystem filename
 */
export function sanitizeFileName(name: string): string {
  if (!name) return '';
  return name
    .replace(/[\/]/g, '.')
    .replace(/[\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Generates suggested document filename without extension:
 * "<Số ký hiệu (đổi / thành .)> - <Tên văn bản (giữ dấu tiếng Việt)>"
 */
export function generateSuggestedFileName(fullDocNumber: string, docName: string): string {
  const safeDocNumber = (fullDocNumber || '')
    .replace(/[\x00-\x1f\x7f]/g, '')
    .replace(/\//g, '.')
    .replace(/[\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '')
    .trim();

  const safeDocName = (docName || '')
    .replace(/[\x00-\x1f\x7f]/g, '')
    .replace(/\//g, '.')
    .replace(/[\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '')
    .trim();

  if (safeDocNumber && safeDocName) {
    return `${safeDocNumber} - ${safeDocName}`;
  } else if (safeDocNumber) {
    return safeDocNumber;
  } else if (safeDocName) {
    return safeDocName;
  }
  return '';
}

export const DEFAULT_MAX_WORD_MB = 10;

export function getWordFileExtension(fileName: string): '.doc' | '.docx' | null {
  if (!fileName) return null;
  const lower = fileName.trim().toLowerCase();
  if (lower.endsWith('.docx')) return '.docx';
  if (lower.endsWith('.doc')) return '.doc';
  return null;
}

export function getSupportedAttachmentExtension(fileNameOrExt: string): '.doc' | '.docx' | '.pdf' | null {
  if (!fileNameOrExt) return null;
  const lower = fileNameOrExt.trim().toLowerCase();
  if (lower === '.pdf' || lower.endsWith('.pdf')) return '.pdf';
  if (lower === '.docx' || lower.endsWith('.docx')) return '.docx';
  if (lower === '.doc' || lower.endsWith('.doc')) return '.doc';
  return null;
}

/**
 * Unified filename builder used for ALL uploads (internal Word attachments, PDF from URL, PDF from file).
 * Reuses the exact "<Số ký hiệu> - <Tên văn bản>.<ext>" template, converts '/' -> '.',
 * preserves Vietnamese diacritics, and keeps the target extension (.doc, .docx, .pdf).
 */
export function buildStandardDriveFileName(
  fullDocNumber: string,
  docName: string,
  originalNameOrExt: string
): string {
  const ext = getSupportedAttachmentExtension(originalNameOrExt) || '.docx';
  let baseName = generateSuggestedFileName(fullDocNumber, docName)
    .replace(/[. ]+$/g, '')
    .trim();

  if (!baseName) {
    baseName = 'VAN_BAN';
  }
  if (baseName.length > 180) {
    baseName = baseName.slice(0, 180).replace(/[-. ]+$/g, '').trim();
  }
  return `${baseName}${ext}`;
}

/**
 * Normalizes any date or date-time string into date-only "dd/MM/yyyy" (no HH:mm or HH:mm:ss).
 * Returns '' if input is empty.
 */
export function normalizeDateOnlyDdMmYyyy(raw: string | undefined | null): string {
  const trimmed = (raw || '').trim();
  if (!trimmed) return '';

  const formatValid = (dNum: number, mNum: number, yNum: number): string => {
    if (
      !Number.isFinite(dNum) ||
      !Number.isFinite(mNum) ||
      !Number.isFinite(yNum) ||
      dNum < 1 ||
      dNum > 31 ||
      mNum < 1 ||
      mNum > 12 ||
      yNum < 1900 ||
      yNum > 2100
    ) {
      return '';
    }
    return `${String(dNum).padStart(2, '0')}/${String(mNum).padStart(2, '0')}/${yNum}`;
  };

  // 1. "ngày 14 tháng 11 năm 2026" (with optional trailing time)
  const vnText = trimmed.match(/ng[àa]y\s+(\d{1,2})\s+th[áa]ng\s+(\d{1,2})\s+n[ăa]m\s+(\d{4})/i);
  if (vnText) {
    const valid = formatValid(parseInt(vnText[1], 10), parseInt(vnText[2], 10), parseInt(vnText[3], 10));
    if (valid) return valid;
  }

  // 2. ISO yyyy-MM-dd or yyyy/MM/dd (with optional T... or HH:mm:ss)
  const isoMatch = trimmed.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s,]+.*)?$/);
  if (isoMatch) {
    const valid = formatValid(parseInt(isoMatch[3], 10), parseInt(isoMatch[2], 10), parseInt(isoMatch[1], 10));
    if (valid) return valid;
  }

  // 3. dd/MM/yyyy, dd-MM-yyyy, dd.MM.yyyy (with optional HH:mm or HH:mm:ss)
  const dmyMatch = trimmed.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})(?:[T\s,]+\d{1,2}:\d{2}(?::\d{2})?.*)?$/);
  if (dmyMatch) {
    const rawYear = dmyMatch[3];
    const yNum = rawYear.length === 2 ? parseInt(`20${rawYear}`, 10) : parseInt(rawYear, 10);
    const valid = formatValid(parseInt(dmyMatch[1], 10), parseInt(dmyMatch[2], 10), yNum);
    if (valid) return valid;
  }

  // Fallback: strip any trailing time component if present
  return trimmed.replace(/[\s,T]+\d{1,2}:\d{2}(?::\d{2})?(?:\s*(?:AM|PM|UTC|GMT|[+-]\d{2}:?\d{2}))?$/i, '').trim();
}

/**
 * Backward-compatible alias for Word attachment callers
 */
export function buildWordDriveFileName(
  fullDocNumber: string,
  docName: string,
  originalFileName: string
): string {
  return buildStandardDriveFileName(fullDocNumber, docName, originalFileName);
}

export function appendConflictSuffix(fileName: string, index = 2): string {
  const ext = getSupportedAttachmentExtension(fileName) || '';
  const base = ext ? fileName.slice(0, -ext.length) : fileName;
  return `${base} (${index})${ext}`;
}

export function sanitizeCustomDriveFileName(
  inputName: string,
  requiredExt: '.doc' | '.docx' | '.pdf'
): string {
  let cleaned = (inputName || '')
    .replace(/[\x00-\x1f\x7f]/g, '')
    .replace(/\//g, '.')
    .replace(/[\\:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  if (cleaned.toLowerCase().endsWith(requiredExt)) {
    cleaned = cleaned.slice(0, -requiredExt.length).trim();
  } else {
    const anyExt = getSupportedAttachmentExtension(cleaned);
    if (anyExt) {
      cleaned = cleaned.slice(0, -anyExt.length).trim();
    }
  }
  cleaned = cleaned.replace(/[. ]+$/g, '').trim();
  if (!cleaned) cleaned = 'VAN_BAN';
  if (cleaned.length > 180) cleaned = cleaned.slice(0, 180).trim();
  return `${cleaned}${requiredExt}`;
}

export function formatFileSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

/**
 * Extracts Google Drive Folder ID from a URL or raw ID string.
 * Supports:
 * - https://drive.google.com/drive/folders/FOLDER_ID
 * - https://drive.google.com/drive/u/0/folders/FOLDER_ID?usp=sharing
 * - https://drive.google.com/open?id=FOLDER_ID
 * - Raw folder ID string
 */
export function extractGoogleDriveFolderId(input: string): string {
  const raw = (input || '').trim();
  if (!raw) return '';

  const folderMatch = raw.match(/\/folders\/([a-zA-Z0-9_-]{10,})/);
  if (folderMatch) return folderMatch[1];

  const idParamMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (idParamMatch) return idParamMatch[1];

  if (/^[a-zA-Z0-9_-]{10,}$/.test(raw) && !raw.includes('http')) {
    return raw;
  }

  return '';
}

export function normalizeCompareText(input: string | undefined | null): string {
  if (!input) return '';
  return removeVietnameseDiacritics(input)
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function normalizeDocNumberForCompare(input: string | undefined | null): string {
  if (!input) return '';
  return removeVietnameseDiacritics(input)
    .toLowerCase()
    .replace(/\s*([\/\-])\s*/g, '$1')
    .replace(/\s+/g, '')
    .trim();
}

export function getDocumentFilterYear(doc: {
  docCategory?: 'outgoing' | 'incoming';
  issueDate?: string;
  year?: number;
}): number | 'unknown' {
  if (doc.docCategory === 'incoming') {
    const cleanIssue = normalizeDateOnlyDdMmYyyy(doc.issueDate || '');
    if (!cleanIssue) return 'unknown';
    const parts = cleanIssue.split('/');
    if (parts.length === 3 && /^\d{4}$/.test(parts[2])) {
      const y = parseInt(parts[2], 10);
      if (Number.isFinite(y) && y >= 1900 && y <= 2100) {
        return y;
      }
    }
    return 'unknown';
  }
  return typeof doc.year === 'number' && Number.isFinite(doc.year) ? doc.year : 'unknown';
}


