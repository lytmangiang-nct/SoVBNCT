import crypto from 'crypto';
import type { ArchiveDriveCandidate, GoogleDriveConfig } from '../src/types.js';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  extractGoogleDriveFolderId,
  getSupportedAttachmentExtension,
  getWordFileExtension,
  sanitizeCustomDriveFileName,
} from '../src/utils/fileNaming.js';

export function getMaxWordFileSizeMB(): number {
  const envVal = process.env.MAX_WORD_UPLOAD_MB ? parseFloat(process.env.MAX_WORD_UPLOAD_MB) : DEFAULT_MAX_WORD_MB;
  if (!Number.isFinite(envVal) || envVal < 1) return DEFAULT_MAX_WORD_MB;
  return Math.min(Math.floor(envVal), 50);
}

export function getMaxWordFileSizeBytes(): number {
  return getMaxWordFileSizeMB() * 1024 * 1024;
}

export interface WordFileValidationResult {
  valid: boolean;
  error?: string;
  extension?: '.doc' | '.docx';
  normalizedMimeType?: string;
}

export function validateWordFileBuffer(
  originalName: string,
  mimeType: string,
  buffer: Buffer
): WordFileValidationResult {
  const ext = getWordFileExtension(originalName);
  if (!ext) {
    return {
      valid: false,
      error: 'Định dạng tệp không hợp lệ. Hệ thống chỉ chấp nhận tệp Word (.doc hoặc .docx).',
    };
  }

  if (!buffer || buffer.length === 0) {
    return {
      valid: false,
      error: 'Tệp Word được chọn đang bị rỗng (0 bytes). Vui lòng kiểm tra lại.',
    };
  }

  const maxBytes = getMaxWordFileSizeBytes();
  const maxMB = getMaxWordFileSizeMB();
  if (buffer.length > maxBytes) {
    return {
      valid: false,
      error: `Dung lượng tệp vượt quá giới hạn tối đa ${maxMB} MB của hệ thống.`,
    };
  }

  // Reject disguised non-Word files by magic bytes
  const head4 = buffer.subarray(0, 4);
  const headStr = head4.toString('latin1');
  if (headStr === '%PDF') {
    return { valid: false, error: 'Tệp này là PDF, không phải tệp Word (.doc/.docx).' };
  }
  if (headStr.startsWith('MZ')) {
    return { valid: false, error: 'Không chấp nhận tệp thực thi.' };
  }
  if (head4[0] === 0x89 && headStr.slice(1, 4) === 'PNG') {
    return { valid: false, error: 'Tệp này là ảnh PNG, không phải tệp Word.' };
  }
  if (head4[0] === 0xff && head4[1] === 0xd8 && head4[2] === 0xff) {
    return { valid: false, error: 'Tệp này là ảnh JPEG, không phải tệp Word.' };
  }

  if (ext === '.docx') {
    // .docx is an OpenXML ZIP archive starting with PK (0x50 0x4B)
    if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
      return {
        valid: false,
        error: 'Nội dung tệp .docx không hợp lệ (không đúng cấu trúc chuẩn Microsoft Word OpenXML).',
      };
    }
    return {
      valid: true,
      extension: '.docx',
      normalizedMimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    };
  }

  // .doc: OLE2 (D0 CF 11 E0), RTF ({\rtf), or Word XML/HTML (<)
  const isOle2 =
    buffer.length >= 8 &&
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0;
  const trimmedHead = buffer.subarray(0, 64).toString('utf8').trimStart();
  const isRtfOrXml = trimmedHead.startsWith('{\\rtf') || trimmedHead.startsWith('<');
  if (!isOle2 && !isRtfOrXml) {
    return {
      valid: false,
      error: 'Nội dung tệp .doc không hợp lệ (không đúng định dạng Microsoft Word .doc).',
    };
  }

  return {
    valid: true,
    extension: '.doc',
    normalizedMimeType: 'application/msword',
  };
}

export interface PdfFileValidationResult {
  valid: boolean;
  error?: string;
}

export function validatePdfFileBuffer(
  originalName: string,
  mimeType: string,
  buffer: Buffer
): PdfFileValidationResult {
  if (!buffer || buffer.length === 0) {
    return { valid: false, error: 'Tệp PDF đang bị rỗng (0 bytes).' };
  }
  const maxBytes = getMaxWordFileSizeBytes();
  const maxMB = getMaxWordFileSizeMB();
  if (buffer.length > maxBytes) {
    return {
      valid: false,
      error: `Dung lượng tệp PDF vượt quá giới hạn tối đa ${maxMB} MB của hệ thống.`,
    };
  }

  // Kiểm tra chữ ký %PDF- trong 1024 byte đầu tiên theo chuẩn PDF
  const headText = buffer.subarray(0, Math.min(buffer.length, 1024)).toString('latin1');
  if (!headText.includes('%PDF-')) {
    const ext = (originalName || '').toLowerCase();
    if (headText.trimStart().startsWith('<!DOCTYPE') || headText.trimStart().startsWith('<html') || mimeType.includes('text/html')) {
      return {
        valid: false,
        error: 'Nội dung nhận được là trang web/HTML (có thể yêu cầu đăng nhập hoặc không phải liên kết tải trực tiếp), không phải tệp PDF.',
      };
    }
    return {
      valid: false,
      error: `Nội dung tệp${ext ? ` (${originalName})` : ''} không phải định dạng PDF hợp lệ (thiếu chữ ký %PDF).`,
    };
  }

  return { valid: true };
}

export interface ParsedServiceAccount {
  client_email: string;
  private_key: string;
  token_uri?: string;
}

export function parseServiceAccountJson(rawJson: string): {
  parsed?: ParsedServiceAccount;
  error?: string;
} {
  if (!rawJson || !rawJson.trim()) {
    return { error: 'Chưa cung cấp nội dung JSON Tài khoản dịch vụ (Service Account).' };
  }
  try {
    const obj = JSON.parse(rawJson.trim());
    if (!obj.client_email || !obj.private_key) {
      return {
        error: 'Tệp JSON Tài khoản dịch vụ thiếu trường bắt buộc (client_email hoặc private_key).',
      };
    }
    if (!String(obj.private_key).includes('BEGIN PRIVATE KEY')) {
      return {
        error: 'Khóa private_key trong JSON Tài khoản dịch vụ không đúng định dạng PEM.',
      };
    }
    return {
      parsed: {
        client_email: String(obj.client_email).trim(),
        private_key: String(obj.private_key).replace(/\\n/g, '\n'),
        token_uri: obj.token_uri ? String(obj.token_uri).trim() : 'https://oauth2.googleapis.com/token',
      },
    };
  } catch {
    return { error: 'Nội dung Tài khoản dịch vụ không phải JSON hợp lệ.' };
  }
}

export const UNKNOWN_DRIVE_YEAR_FOLDER = 'Chưa xác định năm';

export interface DriveFolderCacheEntry {
  bookFolderId?: string;
  yearFolderId?: string;
  categoryFolderId: string;
  folderPath: string;
  updatedAt: string;
}

export function normalizeDriveArchiveYearLabel(year?: number | string | null): string {
  if (typeof year === 'number' && Number.isFinite(year) && year >= 1900 && year <= 2100) {
    return String(Math.trunc(year));
  }
  if (typeof year === 'string') {
    const trimmed = year.trim();
    if (!trimmed || trimmed === 'unknown' || trimmed.toLowerCase() === 'chưa xác định năm') {
      return UNKNOWN_DRIVE_YEAR_FOLDER;
    }
    if (/^\d{4}$/.test(trimmed)) {
      const n = parseInt(trimmed, 10);
      if (n >= 1900 && n <= 2100) return String(n);
    }
  }
  return UNKNOWN_DRIVE_YEAR_FOLDER;
}

export function buildDriveFolderHierarchy(
  bookName: string,
  year: number | string | null | undefined,
  docCategory: 'outgoing' | 'incoming'
): { bookFolder: string; yearFolder: string; categoryFolder: string; displayPath: string } {
  const cleanBook = (bookName || 'Sổ văn bản').trim();
  const yearStr = normalizeDriveArchiveYearLabel(year);
  const cat = docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi';
  return {
    bookFolder: cleanBook,
    yearFolder: yearStr,
    categoryFolder: cat,
    displayPath: `${cleanBook}/${yearStr}/${cat}`,
  };
}

export interface ResolvedGoogleDriveConfig {
  enabled: boolean;
  authMode: 'apps_script' | 'service_account';
  folderUrl: string;
  folderId: string;
  folderName?: string;
  appsScriptUrl: string;
  serviceAccountJson: string;
  serviceAccountEmail: string;
  isConfigured: boolean;
  missingReason?: string;
}

export function resolveEffectiveDriveConfig(
  stored?: GoogleDriveConfig,
  unitCode?: string,
  bookDriveConfig?: Partial<GoogleDriveConfig>,
  isDefaultBook: boolean = true
): ResolvedGoogleDriveConfig {
  const envEnabled =
    isDefaultBook && process.env.GOOGLE_DRIVE_ENABLED !== undefined
      ? process.env.GOOGLE_DRIVE_ENABLED === 'true'
      : undefined;
  const envAuthMode = process.env.GOOGLE_DRIVE_AUTH_MODE as 'apps_script' | 'service_account' | undefined;
  const envFolderId = isDefaultBook
    ? extractGoogleDriveFolderId(process.env.GOOGLE_DRIVE_FOLDER_ID || process.env.GOOGLE_DRIVE_FOLDER_URL || '')
    : '';
  const envAppsScriptUrl = (process.env.GOOGLE_DRIVE_APPS_SCRIPT_URL || '').trim();
  const envSaJson = (process.env.GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON || '').trim();

  // Sổ riêng (không phải mặc định) KHÔNG kế thừa enabled hoặc folderId của Đoàn trường để đảm bảo cô lập tuyệt đối thư mục & quyền Drive theo sổ
  const enabled =
    envEnabled !== undefined
      ? envEnabled
      : !isDefaultBook
      ? Boolean(bookDriveConfig?.enabled)
      : bookDriveConfig?.enabled !== undefined
      ? Boolean(bookDriveConfig.enabled)
      : Boolean(stored?.enabled);
  const authMode: 'apps_script' | 'service_account' =
    envAuthMode === 'service_account' || envAuthMode === 'apps_script'
      ? envAuthMode
      : bookDriveConfig?.authMode || stored?.authMode || 'apps_script';

  const bookFolderUrl = (bookDriveConfig?.folderUrl || '').trim();
  const bookFolderId = extractGoogleDriveFolderId(bookDriveConfig?.folderId || bookFolderUrl);
  const unitScope = isDefaultBook && unitCode && stored?.folderByUnit ? stored.folderByUnit[unitCode] : undefined;
  const folderUrl = !isDefaultBook
    ? bookFolderUrl
    : (
        bookFolderUrl ||
        unitScope?.folderUrl ||
        stored?.folderUrl ||
        process.env.GOOGLE_DRIVE_FOLDER_URL ||
        ''
      ).trim();
  const folderId = !isDefaultBook
    ? bookFolderId
    : envFolderId ||
      bookFolderId ||
      extractGoogleDriveFolderId(unitScope?.folderId || unitScope?.folderUrl || '') ||
      extractGoogleDriveFolderId(stored?.folderId || stored?.folderUrl || '');
  const folderName = !isDefaultBook
    ? bookDriveConfig?.folderName
    : (bookFolderId ? bookDriveConfig?.folderName : undefined) || unitScope?.folderName || stored?.folderName;

  const appsScriptUrl =
    envAppsScriptUrl ||
    (bookDriveConfig?.appsScriptUrl || '').trim() ||
    (stored?.appsScriptUrl || '').trim();
  const serviceAccountJson =
    envSaJson ||
    (bookDriveConfig?.serviceAccountJson || '').trim() ||
    (stored?.serviceAccountJson || '').trim();
  const parsedSa = serviceAccountJson ? parseServiceAccountJson(serviceAccountJson) : {};
  const serviceAccountEmail =
    parsedSa.parsed?.client_email ||
    bookDriveConfig?.serviceAccountEmail ||
    stored?.serviceAccountEmail ||
    '';

  let isConfigured = false;
  let missingReason: string | undefined;

  if (!enabled) {
    missingReason = 'Chưa bật lưu trữ tệp văn bản (Word/PDF) lên Google Drive cho sổ này.';
  } else if (!folderId) {
    missingReason = 'Chưa cấu hình URL hoặc ID thư mục Google Drive riêng cho sổ này.';
  } else if (authMode === 'apps_script') {
    if (!appsScriptUrl) {
      missingReason = 'Chưa cấu hình URL Web App Apps Script kết nối Google Drive.';
    } else {
      isConfigured = true;
    }
  } else {
    if (!serviceAccountJson || parsedSa.error) {
      missingReason = parsedSa.error || 'Chưa cấu hình khóa JSON Tài khoản dịch vụ (Service Account).';
    } else {
      isConfigured = true;
    }
  }

  return {
    enabled,
    authMode,
    folderUrl,
    folderId,
    folderName,
    appsScriptUrl,
    serviceAccountJson,
    serviceAccountEmail,
    isConfigured,
    missingReason,
  };
}

function maskUrl(url: string): string {
  if (!url) return '';
  if (url.length <= 42) return 'Đã cấu hình URL bảo mật';
  return `${url.slice(0, 36)}••••${url.slice(-10)}`;
}

export function sanitizeDriveConfigForClient(
  stored: GoogleDriveConfig | undefined,
  isAdmin: boolean,
  unitCode?: string,
  bookDriveConfig?: Partial<GoogleDriveConfig>,
  isDefaultBook: boolean = true
): GoogleDriveConfig {
  const effective = resolveEffectiveDriveConfig(stored, unitCode, bookDriveConfig, isDefaultBook);
  const maxFileSizeMB = getMaxWordFileSizeMB();
  const maxFileSizeBytes = getMaxWordFileSizeBytes();

  if (!isAdmin) {
    return {
      enabled: effective.enabled,
      authMode: effective.authMode,
      folderUrl: '',
      folderId: '',
      folderName: effective.folderName,
      isConfigured: effective.isConfigured,
      maxFileSizeMB,
      maxFileSizeBytes,
      lastCheckMessage: effective.missingReason,
    };
  }

  const lastCheckedAt = isDefaultBook
    ? bookDriveConfig?.lastCheckedAt || stored?.lastCheckedAt
    : bookDriveConfig?.lastCheckedAt;
  const lastCheckStatus = isDefaultBook
    ? bookDriveConfig?.lastCheckStatus || stored?.lastCheckStatus
    : bookDriveConfig?.lastCheckStatus;
  const lastCheckMessage = isDefaultBook
    ? bookDriveConfig?.lastCheckMessage || stored?.lastCheckMessage || effective.missingReason
    : bookDriveConfig?.lastCheckMessage || effective.missingReason;

  return {
    enabled: effective.enabled,
    authMode: effective.authMode,
    folderUrl: effective.folderUrl,
    folderId: effective.folderId,
    folderName: effective.folderName,
    folderByUnit: isDefaultBook ? stored?.folderByUnit : undefined,
    hasAppsScriptUrl: Boolean(effective.appsScriptUrl),
    appsScriptUrlMasked: maskUrl(effective.appsScriptUrl),
    hasServiceAccount: Boolean(effective.serviceAccountJson),
    serviceAccountEmail: effective.serviceAccountEmail,
    isConfigured: effective.isConfigured,
    maxFileSizeMB,
    maxFileSizeBytes,
    lastCheckedAt,
    lastCheckStatus,
    lastCheckMessage,
  };
}

// --- Service Account OAuth2 JWT Token Minting ---
async function getServiceAccountAccessToken(sa: ParsedServiceAccount): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const claim = Buffer.from(
    JSON.stringify({
      iss: sa.client_email,
      scope: 'https://www.googleapis.com/auth/drive',
      aud: sa.token_uri || 'https://oauth2.googleapis.com/token',
      iat: now,
      exp: now + 3600,
    })
  ).toString('base64url');

  const unsigned = `${header}.${claim}`;
  const signer = crypto.createSign('RSA-SHA256');
  signer.update(unsigned);
  const signature = signer.sign(sa.private_key, 'base64url');
  const jwt = `${unsigned}.${signature}`;

  const params = new URLSearchParams();
  params.set('grant_type', 'urn:ietf:params:oauth:grant-type:jwt-bearer');
  params.set('assertion', jwt);

  const tokenRes = await fetch(sa.token_uri || 'https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  const data: any = await tokenRes.json().catch(() => ({}));
  if (!tokenRes.ok || !data.access_token) {
    throw new Error(
      data.error_description || data.error || 'Không thể xác thực Tài khoản dịch vụ Google (JWT).'
    );
  }
  return data.access_token;
}

async function parseAppsScriptJsonResponse(res: Response): Promise<any> {
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    if (text.includes('<html') || text.includes('<!DOCTYPE') || text.includes('accounts.google.com')) {
      throw new Error(
        'Apps Script trả về trang HTML thay vì JSON. Hãy kiểm tra khi Triển khai (Deploy) mục "Ai có quyền truy cập" đã chọn "Bất kỳ ai (Anyone)" và sử dụng đúng URL kết thúc bằng /exec.'
      );
    }
    throw new Error(`Phản hồi không hợp lệ từ Apps Script (HTTP ${res.status}).`);
  }
}

export async function testGoogleDriveConnection(params: {
  authMode: 'apps_script' | 'service_account';
  folderId: string;
  appsScriptUrl?: string;
  serviceAccountJson?: string;
}): Promise<{
  success: boolean;
  folderId: string;
  folderName?: string;
  message: string;
}> {
  const { authMode, folderId, appsScriptUrl, serviceAccountJson } = params;
  if (!folderId) {
    return {
      success: false,
      folderId: '',
      message: 'URL hoặc ID thư mục Google Drive không hợp lệ.',
    };
  }

  if (authMode === 'apps_script') {
    if (!appsScriptUrl || !appsScriptUrl.trim()) {
      return {
        success: false,
        folderId,
        message: 'Chưa nhập URL Ứng dụng web Google Apps Script.',
      };
    }

    const res = await fetch(appsScriptUrl.trim(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'test',
        folderId,
      }),
    });

    const data = await parseAppsScriptJsonResponse(res);
    if (data.result !== 'success') {
      return {
        success: false,
        folderId,
        message: data.error || 'Kiểm tra thư mục Google Drive qua Apps Script thất bại.',
      };
    }

    return {
      success: true,
      folderId,
      folderName: data.folderName || '',
      message: `Kết nối thành công tới thư mục "${data.folderName || folderId}" (tài khoản Google: ${
        data.accountEmail || 'đã xác thực'
      }). Có đầy đủ quyền ghi tệp.`,
    };
  }

  // Service Account mode via Google Drive API v3
  const { parsed, error } = parseServiceAccountJson(serviceAccountJson || '');
  if (!parsed || error) {
    return {
      success: false,
      folderId,
      message: error || 'Tài khoản dịch vụ không hợp lệ.',
    };
  }

  const accessToken = await getServiceAccountAccessToken(parsed);

  // 1. Check folder existence and capabilities
  const metaRes = await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(
      folderId
    )}?supportsAllDrives=true&fields=id,name,mimeType,driveId,trashed,capabilities(canAddChildren,canEdit)`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );
  const metaData: any = await metaRes.json().catch(() => ({}));
  if (!metaRes.ok) {
    const apiMsg = metaData?.error?.message || metaRes.statusText;
    return {
      success: false,
      folderId,
      message: `Không truy cập được thư mục (${metaRes.status}): ${apiMsg}. Hãy chắc chắn đã chia sẻ thư mục cho email ${parsed.client_email}.`,
    };
  }

  if (metaData.mimeType !== 'application/vnd.google-apps.folder') {
    return {
      success: false,
      folderId,
      message: 'ID đã nhập không phải là thư mục Google Drive.',
    };
  }
  if (metaData.trashed) {
    return {
      success: false,
      folderId,
      message: `Thư mục "${metaData.name}" đang nằm trong Thùng rác.`,
    };
  }
  if (!metaData.capabilities?.canAddChildren) {
    return {
      success: false,
      folderId,
      folderName: metaData.name,
      message: `Tài khoản dịch vụ (${parsed.client_email}) chưa có quyền tải tệp lên thư mục "${metaData.name}". Vui lòng cấp quyền Cộng tác viên / Người chỉnh sửa.`,
    };
  }

  // 2. Verify actual write/quota capability using a tiny probe file and delete it immediately
  const boundary = `nct_probe_${Date.now()}`;
  const probeMeta = JSON.stringify({
    name: `.nct_check_write_${Date.now()}.txt`,
    parents: [folderId],
  });
  const multipartBody =
    `--${boundary}\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${probeMeta}\r\n` +
    `--${boundary}\r\n` +
    `Content-Type: text/plain\r\n\r\n` +
    `NCT Drive write check\r\n` +
    `--${boundary}--`;

  const probeRes = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body: multipartBody,
    }
  );
  const probeData: any = await probeRes.json().catch(() => ({}));
  if (!probeRes.ok || !probeData.id) {
    const errReason = probeData?.error?.errors?.[0]?.reason || '';
    const errMsg = probeData?.error?.message || probeRes.statusText;
    if (errReason === 'storageQuotaExceeded' || errMsg.includes('storage quota')) {
      return {
        success: false,
        folderId,
        folderName: metaData.name,
        message:
          'Tài khoản dịch vụ (Service Account) không có dung lượng lưu trữ trên "Drive của tôi". Hãy dùng thư mục trong "Bộ nhớ dùng chung (Shared Drive)" của Google Workspace hoặc chuyển sang phương thức Apps Script.',
      };
    }
    return {
      success: false,
      folderId,
      folderName: metaData.name,
      message: `Không thể ghi thử vào thư mục "${metaData.name}": ${errMsg}`,
    };
  }

  // Clean up probe file
  await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(probeData.id)}?supportsAllDrives=true`,
    {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  ).catch(() => {});

  return {
    success: true,
    folderId,
    folderName: metaData.name,
    message: `Kết nối thành công tới thư mục "${metaData.name}"${
      metaData.driveId ? ' (Bộ nhớ dùng chung)' : ''
    }. Tài khoản dịch vụ có đầy đủ quyền tải tệp lên.`,
  };
}

export interface DriveUploadResult {
  fileId: string;
  fileName: string;
  webViewLink: string;
  folderId?: string;
  folderPath?: string;
  deduplicated?: boolean;
}

function getGoogleDriveApiBaseUrl(): string {
  const custom = (process.env.GOOGLE_DRIVE_API_BASE || '').trim().replace(/\/+$/, '');
  return custom || 'https://www.googleapis.com';
}

// Khóa đồng bộ hóa theo key trong cùng tiến trình để nhiều người tải lên đồng thời không tạo trùng thư mục
const inFlightFolderCreations = new Map<string, Promise<any>>();
const folderKeyQueues = new Map<string, Promise<void>>();

async function runSerializedByKey<T>(key: string, fn: () => Promise<T>): Promise<T> {
  const previous = folderKeyQueues.get(key) || Promise.resolve();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  folderKeyQueues.set(
    key,
    previous
      .catch(() => {})
      .then(() => gate)
  );
  await previous.catch(() => {});
  try {
    return await fn();
  } finally {
    release();
    if (folderKeyQueues.get(key) === gate) {
      folderKeyQueues.delete(key);
    }
  }
}

async function runCoalescedByKey<T>(key: string, fn: () => Promise<T>): Promise<T> {
  const existing = inFlightFolderCreations.get(key) as Promise<T> | undefined;
  if (existing) {
    return existing;
  }
  const promise = fn().finally(() => {
    if (inFlightFolderCreations.get(key) === promise) {
      inFlightFolderCreations.delete(key);
    }
  });
  inFlightFolderCreations.set(key, promise);
  return promise;
}

const FOLDER_VERIFY_MEM_TTL_MS = 5 * 60 * 1000;
const verifiedFolderMemCache = new Map<string, number>();
const folderNameMemCache = new Map<string, { name: string; expiresAt: number }>();

function clearFolderMemCacheForId(folderId?: string) {
  if (!folderId) return;
  folderNameMemCache.delete(folderId);
  for (const key of Array.from(verifiedFolderMemCache.keys())) {
    if (key.startsWith(`${folderId}::`)) {
      verifiedFolderMemCache.delete(key);
    }
  }
}

async function getServiceAccountFolderMeta(
  accessToken: string,
  folderId: string
): Promise<{ id: string; name: string; mimeType?: string; trashed?: boolean; parents?: string[] } | null> {
  const apiBase = getGoogleDriveApiBaseUrl();
  const res = await fetch(
    `${apiBase}/drive/v3/files/${encodeURIComponent(folderId)}?supportsAllDrives=true&fields=id,name,mimeType,trashed,parents`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );
  if (!res.ok) return null;
  const data: any = await res.json().catch(() => null);
  if (!data || !data.id) return null;
  return data;
}

async function getServiceAccountFolderName(accessToken: string, folderId: string): Promise<string> {
  const cached = folderNameMemCache.get(folderId);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.name;
  }
  const meta = await getServiceAccountFolderMeta(accessToken, folderId);
  const name = typeof meta?.name === 'string' ? meta.name : '';
  if (name) {
    folderNameMemCache.set(folderId, { name, expiresAt: Date.now() + FOLDER_VERIFY_MEM_TTL_MS });
  }
  return name;
}

async function verifyCachedServiceAccountFolder(
  accessToken: string,
  folderId: string | undefined,
  expectedName: string,
  expectedParentId?: string
): Promise<boolean> {
  const cleanId = (folderId || '').trim();
  if (!cleanId) return false;
  const memKey = `${cleanId}::${expectedName.trim().toLowerCase()}::${expectedParentId || ''}`;
  const cachedExp = verifiedFolderMemCache.get(memKey);
  if (cachedExp && cachedExp > Date.now()) {
    return true;
  }
  const meta = await getServiceAccountFolderMeta(accessToken, cleanId);
  if (!meta || meta.trashed) {
    clearFolderMemCacheForId(cleanId);
    return false;
  }
  if (meta.mimeType !== 'application/vnd.google-apps.folder') return false;
  if ((meta.name || '').trim().toLowerCase() !== expectedName.trim().toLowerCase()) return false;
  if (expectedParentId && Array.isArray(meta.parents) && meta.parents.length > 0) {
    if (!meta.parents.includes(expectedParentId)) return false;
  }
  verifiedFolderMemCache.set(memKey, Date.now() + FOLDER_VERIFY_MEM_TTL_MS);
  return true;
}

async function ensureServiceAccountSubfolder(
  accessToken: string,
  parentId: string,
  folderName: string
): Promise<string> {
  const cleanName = folderName.trim();
  const coalesceKey = `sa_sub::${parentId}::${cleanName.toLowerCase()}`;

  return runCoalescedByKey(coalesceKey, () =>
    runSerializedByKey(coalesceKey, async () => {
      const apiBase = getGoogleDriveApiBaseUrl();
      const safeName = cleanName.replace(/'/g, "\\'");
      const q = `'${parentId}' in parents and name = '${safeName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
      const listUrl = `${apiBase}/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name)&q=${encodeURIComponent(
        q
      )}`;

      const listRes = await fetch(listUrl, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      if (!listRes.ok) {
        const listErr: any = await listRes.json().catch(() => ({}));
        throw new Error(
          `Không thể kiểm tra thư mục "${cleanName}" trên Google Drive (${listRes.status}): ${
            listErr?.error?.message || listRes.statusText
          }`
        );
      }
      const listData: any = await listRes.json().catch(() => ({}));
      if (listData?.files?.[0]?.id) {
        const foundId = String(listData.files[0].id);
        verifiedFolderMemCache.set(
          `${foundId}::${cleanName.toLowerCase()}::${parentId}`,
          Date.now() + FOLDER_VERIFY_MEM_TTL_MS
        );
        return foundId;
      }

      const createRes = await fetch(
        `${apiBase}/drive/v3/files?supportsAllDrives=true&fields=id,name`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            name: cleanName,
            mimeType: 'application/vnd.google-apps.folder',
            parents: [parentId],
          }),
        }
      );
      const createData: any = await createRes.json().catch(() => ({}));
      if (!createRes.ok || !createData.id) {
        throw new Error(
          `Không thể tạo thư mục "${cleanName}" trên Google Drive: ${
            createData?.error?.message || createRes.statusText
          }`
        );
      }
      const createdId = String(createData.id);
      verifiedFolderMemCache.set(
        `${createdId}::${cleanName.toLowerCase()}::${parentId}`,
        Date.now() + FOLDER_VERIFY_MEM_TTL_MS
      );
      return createdId;
    })
  );
}

async function resolveServiceAccountTargetFolder(
  accessToken: string,
  rootFolderId: string,
  hierarchy: { bookFolder: string; yearFolder: string; categoryFolder: string; displayPath: string },
  cachedEntry?: DriveFolderCacheEntry
): Promise<{
  bookFolderId?: string;
  yearFolderId: string;
  categoryFolderId: string;
}> {
  if (cachedEntry?.categoryFolderId) {
    const catOk = await verifyCachedServiceAccountFolder(
      accessToken,
      cachedEntry.categoryFolderId,
      hierarchy.categoryFolder,
      cachedEntry.yearFolderId
    );
    if (catOk) {
      return {
        bookFolderId: cachedEntry.bookFolderId,
        yearFolderId: cachedEntry.yearFolderId || '',
        categoryFolderId: cachedEntry.categoryFolderId,
      };
    }
  }

  let bookParentId = rootFolderId;
  let bookFolderId: string | undefined;
  if (
    cachedEntry?.bookFolderId &&
    (await verifyCachedServiceAccountFolder(accessToken, cachedEntry.bookFolderId, hierarchy.bookFolder, rootFolderId))
  ) {
    bookParentId = cachedEntry.bookFolderId;
    bookFolderId = cachedEntry.bookFolderId;
  } else {
    const rootName = await getServiceAccountFolderName(accessToken, rootFolderId);
    if (rootName.trim().toLowerCase() !== hierarchy.bookFolder.trim().toLowerCase()) {
      bookParentId = await ensureServiceAccountSubfolder(accessToken, rootFolderId, hierarchy.bookFolder);
      bookFolderId = bookParentId;
    }
  }

  let yearFolderId = '';
  if (
    cachedEntry?.yearFolderId &&
    (await verifyCachedServiceAccountFolder(accessToken, cachedEntry.yearFolderId, hierarchy.yearFolder, bookParentId))
  ) {
    yearFolderId = cachedEntry.yearFolderId;
  } else {
    yearFolderId = await ensureServiceAccountSubfolder(accessToken, bookParentId, hierarchy.yearFolder);
  }

  const categoryFolderId = await ensureServiceAccountSubfolder(
    accessToken,
    yearFolderId,
    hierarchy.categoryFolder
  );

  return {
    bookFolderId,
    yearFolderId,
    categoryFolderId,
  };
}

async function findExistingFileServiceAccount(
  accessToken: string,
  folderId: string,
  docId: string,
  expectedFileName?: string
): Promise<{ id: string; name: string; webViewLink: string } | null> {
  const apiBase = getGoogleDriveApiBaseUrl();
  const safeDocId = docId.replace(/'/g, "\\'");
  const qByDoc = `'${folderId}' in parents and trashed = false and appProperties has { key='nctDocId' and value='${safeDocId}' }`;
  const urlByDoc = `${apiBase}/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,webViewLink)&q=${encodeURIComponent(
    qByDoc
  )}`;

  const res = await fetch(urlByDoc, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (res.ok) {
    const data: any = await res.json().catch(() => ({}));
    const first = data?.files?.[0];
    if (first) {
      return {
        id: first.id,
        name: first.name,
        webViewLink: first.webViewLink || `https://drive.google.com/file/d/${first.id}/view`,
      };
    }
  }

  if (expectedFileName && expectedFileName.trim()) {
    const safeName = expectedFileName.trim().replace(/'/g, "\\'");
    const qByName = `'${folderId}' in parents and trashed = false and name = '${safeName}'`;
    const urlByName = `${apiBase}/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,webViewLink)&q=${encodeURIComponent(
      qByName
    )}`;
    const nameRes = await fetch(urlByName, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (nameRes.ok) {
      const nameData: any = await nameRes.json().catch(() => ({}));
      const firstByName = nameData?.files?.[0];
      if (firstByName) {
        return {
          id: firstByName.id,
          name: firstByName.name,
          webViewLink: firstByName.webViewLink || `https://drive.google.com/file/d/${firstByName.id}/view`,
        };
      }
    }
  }

  return null;
}

export async function uploadWordToGoogleDrive(params: {
  config: ReturnType<typeof resolveEffectiveDriveConfig>;
  docId: string;
  bookName?: string;
  year?: number | string | null;
  docCategory?: 'outgoing' | 'incoming';
  fullDocNumber: string;
  docName: string;
  originalFileName: string;
  mimeType: string;
  buffer: Buffer;
  overrideFileName?: string;
  replaceMode?: boolean;
  isRetry?: boolean;
  previousFileId?: string;
  keepPreviousFile?: boolean;
  cacheKey?: string;
  folderCache?: Record<string, DriveFolderCacheEntry>;
  onFolderCacheUpdate?: (key: string, entry: DriveFolderCacheEntry) => void;
}): Promise<DriveUploadResult> {
  const {
    config,
    docId,
    bookName = 'Đoàn trường',
    year,
    docCategory = 'outgoing',
    fullDocNumber,
    docName,
    originalFileName,
    mimeType,
    buffer,
    overrideFileName,
    replaceMode,
    isRetry = false,
    previousFileId,
    keepPreviousFile,
    cacheKey,
    folderCache,
    onFolderCacheUpdate,
  } = params;

  if (!config.isConfigured || !config.folderId) {
    throw new Error(config.missingReason || 'Chưa cấu hình thư mục Google Drive hợp lệ.');
  }

  const hierarchy = buildDriveFolderHierarchy(bookName, year, docCategory);
  const effectiveCacheKey =
    cacheKey || `${config.folderId}::${hierarchy.bookFolder}::${hierarchy.yearFolder}::${hierarchy.categoryFolder}`;

  const requiredExt =
    getSupportedAttachmentExtension(originalFileName) ||
    (mimeType === 'application/pdf' ? '.pdf' : '.docx');

  const finalFileName =
    overrideFileName && overrideFileName.trim()
      ? sanitizeCustomDriveFileName(overrideFileName, requiredExt)
      : buildStandardDriveFileName(fullDocNumber, docName, originalFileName);

  const saveCacheEntry = (resolved: {
    bookFolderId?: string;
    yearFolderId?: string;
    categoryFolderId?: string;
    folderPath?: string;
  }) => {
    const cleanCatId = (resolved.categoryFolderId || '').trim();
    if (!cleanCatId) return;
    const entry: DriveFolderCacheEntry = {
      bookFolderId: resolved.bookFolderId || undefined,
      yearFolderId: resolved.yearFolderId || undefined,
      categoryFolderId: cleanCatId,
      folderPath: resolved.folderPath || hierarchy.displayPath,
      updatedAt: new Date().toISOString(),
    };
    if (folderCache) {
      folderCache[effectiveCacheKey] = entry;
    }
    if (onFolderCacheUpdate) {
      onFolderCacheUpdate(effectiveCacheKey, entry);
    }
  };

  if (config.authMode === 'apps_script') {
    const runAppsScriptFlow = async (): Promise<DriveUploadResult> => {
      const currentCached = folderCache ? folderCache[effectiveCacheKey] : undefined;

      // Chỉ gọi bước kiểm tra trùng trước khi tải lên trong kịch bản thử lại / khôi phục để giảm 1 HTTP roundtrip khi cấp mới
      if (isRetry && !replaceMode && !keepPreviousFile) {
        const checkRes = await fetch(config.appsScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'find',
            folderId: config.folderId,
            bookName: hierarchy.bookFolder,
            year: hierarchy.yearFolder,
            categoryFolder: hierarchy.categoryFolder,
            cachedBookFolderId: currentCached?.bookFolderId || '',
            cachedYearFolderId: currentCached?.yearFolderId || '',
            cachedCategoryFolderId: currentCached?.categoryFolderId || '',
            docId,
            expectedFileName: finalFileName,
            checkNameFallback: true,
          }),
        });
        const checkData = await parseAppsScriptJsonResponse(checkRes);
        if (checkData.result === 'success') {
          const checkFolderId = checkData.targetFolderId || checkData.folderId;
          if (checkFolderId && checkFolderId !== config.folderId) {
            saveCacheEntry({
              bookFolderId: checkData.bookFolderId,
              yearFolderId: checkData.yearFolderId,
              categoryFolderId: checkFolderId,
              folderPath: checkData.folderPath,
            });
          }
          if (checkData.found && checkData.fileId) {
            return {
              fileId: checkData.fileId,
              fileName: checkData.fileName || finalFileName,
              webViewLink: checkData.webViewLink || `https://drive.google.com/file/d/${checkData.fileId}/view`,
              folderId: checkFolderId || currentCached?.categoryFolderId || config.folderId,
              folderPath: checkData.folderPath || hierarchy.displayPath,
              deduplicated: true,
            };
          }
        }
      }

      const freshCached = folderCache ? folderCache[effectiveCacheKey] : currentCached;

      // Upload new file into Tên sổ / Năm / Văn bản đi | Văn bản đến
      const uploadRes = await fetch(config.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'upload',
          folderId: config.folderId,
          bookName: hierarchy.bookFolder,
          year: hierarchy.yearFolder,
          categoryFolder: hierarchy.categoryFolder,
          cachedBookFolderId: freshCached?.bookFolderId || '',
          cachedYearFolderId: freshCached?.yearFolderId || '',
          cachedCategoryFolderId: freshCached?.categoryFolderId || '',
          docId,
          fullDocNumber,
          docName,
          fileName: finalFileName,
          mimeType,
          replaceMode: Boolean(replaceMode || keepPreviousFile),
          keepPreviousFile: Boolean(keepPreviousFile),
          previousFileId: keepPreviousFile ? '' : previousFileId || '',
          fileBase64: buffer.toString('base64'),
        }),
      });

      const uploadData = await parseAppsScriptJsonResponse(uploadRes);
      if (uploadData.result !== 'success' || !uploadData.fileId) {
        throw new Error(uploadData.error || 'Lưu tệp lên Google Drive qua Apps Script thất bại.');
      }

      const resolvedFolderId = uploadData.targetFolderId || uploadData.folderId || freshCached?.categoryFolderId || config.folderId;
      const resolvedFolderPath = uploadData.folderPath || hierarchy.displayPath;
      if (resolvedFolderId && (uploadData.targetFolderId || uploadData.folderId)) {
        saveCacheEntry({
          bookFolderId: uploadData.bookFolderId,
          yearFolderId: uploadData.yearFolderId,
          categoryFolderId: resolvedFolderId,
          folderPath: resolvedFolderPath,
        });
      }

      return {
        fileId: uploadData.fileId,
        fileName: uploadData.fileName || finalFileName,
        webViewLink: uploadData.webViewLink || `https://drive.google.com/file/d/${uploadData.fileId}/view`,
        folderId: resolvedFolderId,
        folderPath: resolvedFolderPath,
        deduplicated: Boolean(uploadData.deduplicated),
      };
    };

    // Nếu thư mục đích chưa có trong cache, tuần tự hóa lần tạo/tìm thư mục đầu tiên để tránh tạo trùng thư mục khi nhiều người tải cùng lúc
    if (!folderCache?.[effectiveCacheKey]?.categoryFolderId) {
      return runSerializedByKey(`apps_script_init::${effectiveCacheKey}`, runAppsScriptFlow);
    }
    return runAppsScriptFlow();
  }

  // Service Account mode (Google Drive API v3)
  const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
  if (!parsed || error) {
    throw new Error(error || 'Cấu hình Tài khoản dịch vụ không hợp lệ.');
  }

  const accessToken = await getServiceAccountAccessToken(parsed);
  const apiBase = getGoogleDriveApiBaseUrl();

  // Resolve/create folder layout: Tên sổ / Năm / Văn bản đi | Văn bản đến (có cache ID + TTL 5 phút trong RAM + khóa đồng bộ)
  const cachedEntry = folderCache ? folderCache[effectiveCacheKey] : undefined;
  let resolvedFolders = await resolveServiceAccountTargetFolder(
    accessToken,
    config.folderId,
    hierarchy,
    cachedEntry
  );
  let targetFolderId = resolvedFolders.categoryFolderId;
  saveCacheEntry({
    bookFolderId: resolvedFolders.bookFolderId,
    yearFolderId: resolvedFolders.yearFolderId,
    categoryFolderId: targetFolderId,
    folderPath: hierarchy.displayPath,
  });

  // Step 1: Đối soát tệp trùng trên Drive chỉ chạy khi retry / khôi phục để lượt tải đầu tiên giữ tốc độ tối đa
  if (isRetry && !replaceMode && !keepPreviousFile) {
    const existing = await findExistingFileServiceAccount(accessToken, targetFolderId, docId, finalFileName);
    if (existing) {
      if (existing.name !== finalFileName) {
        await fetch(
          `${apiBase}/drive/v3/files/${encodeURIComponent(existing.id)}?supportsAllDrives=true`,
          {
            method: 'PATCH',
            headers: {
              Authorization: `Bearer ${accessToken}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ name: finalFileName }),
          }
        ).catch(() => {});
      }
      return {
        fileId: existing.id,
        fileName: finalFileName,
        webViewLink: existing.webViewLink,
        folderId: targetFolderId,
        folderPath: hierarchy.displayPath,
        deduplicated: true,
      };
    }
  } else if (replaceMode && previousFileId && !keepPreviousFile) {
    // Move previous file to trash when replacing via Service Account (best-effort)
    await fetch(
      `${apiBase}/drive/v3/files/${encodeURIComponent(previousFileId)}?supportsAllDrives=true`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ trashed: true }),
      }
    ).catch(() => {});
  }

  // Step 2: Multipart upload preserving exact binary Word/PDF content and folder-inherited permissions
  const performMultipartUpload = async (folderIdToUse: string) => {
    const boundary = `nct_upload_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
    const metadata = JSON.stringify({
      name: finalFileName,
      parents: [folderIdToUse],
      description: `Sổ: ${hierarchy.bookFolder} | Số văn bản: ${fullDocNumber} | Tên văn bản: ${docName} | Mã hệ thống: ${docId}`,
      appProperties: {
        nctDocId: docId,
        nctDocNumber: fullDocNumber,
        nctBookName: hierarchy.bookFolder,
      },
    });

    const prefix = Buffer.from(
      `--${boundary}\r\n` +
        `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
        `${metadata}\r\n` +
        `--${boundary}\r\n` +
        `Content-Type: ${mimeType}\r\n\r\n`,
      'utf-8'
    );
    const suffix = Buffer.from(`\r\n--${boundary}--`, 'utf-8');
    const multipartPayload = Buffer.concat([prefix, buffer, suffix]);

    const uploadRes = await fetch(
      `${apiBase}/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': `multipart/related; boundary=${boundary}`,
          'Content-Length': String(multipartPayload.length),
        },
        body: multipartPayload,
      }
    );
    const uploadData: any = await uploadRes.json().catch(() => ({}));
    return { uploadRes, uploadData };
  };

  let { uploadRes, uploadData } = await performMultipartUpload(targetFolderId);

  // Nếu thư mục cache vừa bị xóa/di chuyển (404), vô hiệu hóa cache và phân giải lại 1 lần
  if (uploadRes.status === 404) {
    clearFolderMemCacheForId(targetFolderId);
    if (resolvedFolders.yearFolderId) clearFolderMemCacheForId(resolvedFolders.yearFolderId);
    if (resolvedFolders.bookFolderId) clearFolderMemCacheForId(resolvedFolders.bookFolderId);
    if (folderCache) delete folderCache[effectiveCacheKey];

    resolvedFolders = await resolveServiceAccountTargetFolder(
      accessToken,
      config.folderId,
      hierarchy,
      undefined
    );
    targetFolderId = resolvedFolders.categoryFolderId;
    saveCacheEntry({
      bookFolderId: resolvedFolders.bookFolderId,
      yearFolderId: resolvedFolders.yearFolderId,
      categoryFolderId: targetFolderId,
      folderPath: hierarchy.displayPath,
    });
    const retryRes = await performMultipartUpload(targetFolderId);
    uploadRes = retryRes.uploadRes;
    uploadData = retryRes.uploadData;
  }

  if (!uploadRes.ok || !uploadData.id) {
    const errReason = uploadData?.error?.errors?.[0]?.reason || '';
    const errMsg = uploadData?.error?.message || uploadRes.statusText;
    if (errReason === 'storageQuotaExceeded' || errMsg.includes('storage quota')) {
      throw new Error(
        'Tài khoản dịch vụ không có dung lượng trên Drive cá nhân. Hãy dùng Bộ nhớ dùng chung (Shared Drive) hoặc chế độ Apps Script.'
      );
    }
    throw new Error(`Tải tệp lên Google Drive thất bại (${uploadRes.status}): ${errMsg}`);
  }

  return {
    fileId: uploadData.id,
    fileName: uploadData.name || finalFileName,
    webViewLink: uploadData.webViewLink || `https://drive.google.com/file/d/${uploadData.id}/view`,
    folderId: targetFolderId,
    folderPath: hierarchy.displayPath,
    deduplicated: false,
  };
}

export const uploadFileToGoogleDrive = uploadWordToGoogleDrive;

export interface DriveFileCheckResult {
  status: 'ok' | 'not_found' | 'access_denied' | 'unverified';
  message?: string;
}

/**
 * Kiểm tra trạng thái tồn tại và quyền truy cập của 1 tệp theo driveFileId trên Google Drive.
 */
export async function checkGoogleDriveFileStatus(
  config: ResolvedGoogleDriveConfig,
  fileId: string
): Promise<DriveFileCheckResult> {
  const cleanId = (fileId || '').trim();
  if (!cleanId) {
    return { status: 'not_found', message: 'Thiếu mã tệp Google Drive.' };
  }
  if (!config.isConfigured) {
    return { status: 'unverified', message: config.missingReason || 'Chưa cấu hình Google Drive.' };
  }

  try {
    if (config.authMode === 'service_account') {
      const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
      if (!parsed || error) {
        return { status: 'unverified', message: error || 'Khóa Service Account không hợp lệ.' };
      }
      const accessToken = await getServiceAccountAccessToken(parsed);
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(cleanId)}?supportsAllDrives=true&fields=id,name,trashed,capabilities(canDownload)`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );
      if (res.status === 404) {
        return { status: 'not_found', message: 'Tệp không còn tồn tại trên Google Drive (404).' };
      }
      if (res.status === 401 || res.status === 403) {
        return {
          status: 'access_denied',
          message: 'Tài khoản tích hợp thiếu quyền truy cập tệp này trên Google Drive.',
        };
      }
      if (!res.ok) {
        return { status: 'unverified', message: `Google Drive trả về mã ${res.status}.` };
      }
      const data: any = await res.json().catch(() => ({}));
      if (data?.trashed) {
        return { status: 'not_found', message: 'Tệp đã bị chuyển vào thùng rác trên Google Drive.' };
      }
      return { status: 'ok' };
    }

    if (!config.appsScriptUrl) {
      return { status: 'unverified', message: 'Thiếu URL Apps Script.' };
    }

    const res = await fetch(config.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'check_file',
        folderId: config.folderId,
        fileId: cleanId,
      }),
    });
    const data = await parseAppsScriptJsonResponse(res);
    if (data.result === 'success') {
      if (data.status === 'not_found' || data.exists === false || data.trashed === true) {
        return { status: 'not_found', message: data.message || 'Tệp không tìm thấy trên Google Drive.' };
      }
      if (data.status === 'access_denied') {
        return {
          status: 'access_denied',
          message: data.message || 'Tài khoản Apps Script không có quyền truy cập tệp này.',
        };
      }
      return { status: 'ok' };
    }
    const errText = String(data.error || '').toLowerCase();
    if (errText.includes('thao tác không hợp lệ') || errText.includes('check_file')) {
      return {
        status: 'unverified',
        message: 'Apps Script hiện tại chưa cập nhật hàm kiểm tra check_file.',
      };
    }
    if (errText.includes('permission') || errText.includes('access denied') || errText.includes('quyền')) {
      return { status: 'access_denied', message: data.error };
    }
    if (errText.includes('no item with the given id') || errText.includes('not found') || errText.includes('không tìm thấy')) {
      return { status: 'not_found', message: 'Tệp không tìm thấy trên Google Drive.' };
    }
    return { status: 'unverified', message: data.error || 'Chưa xác minh được liên kết Drive.' };
  } catch (err: any) {
    return {
      status: 'unverified',
      message: err?.message || 'Lỗi kết nối khi kiểm tra tệp trên Google Drive.',
    };
  }
}

export function extractGoogleDriveFileId(input: string): string {
  const raw = (input || '').trim();
  if (!raw) return '';
  if (/\/folders\//i.test(raw)) return '';

  const fileDMatch = raw.match(/\/(?:file|document|spreadsheets|presentation)\/d\/([a-zA-Z0-9_-]{12,})/);
  if (fileDMatch) return fileDMatch[1];

  const idParamMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]{12,})/);
  if (idParamMatch) return idParamMatch[1];

  // Raw Drive ID without path separators, spaces, or file extensions
  if (/^[a-zA-Z0-9_-]{18,}$/.test(raw) && !raw.includes('http') && !raw.includes('.')) {
    return raw;
  }

  return '';
}

export interface DriveFileMetadataLookupResult {
  status: 'ok' | 'not_found' | 'access_denied' | 'unverified';
  file?: ArchiveDriveCandidate;
  message?: string;
}

export async function getDriveFileMetadata(
  config: ResolvedGoogleDriveConfig,
  fileId: string
): Promise<DriveFileMetadataLookupResult> {
  const cleanId = (fileId || '').trim();
  if (!cleanId) {
    return { status: 'not_found', message: 'Thiếu mã tệp Google Drive.' };
  }
  if (!config.enabled || (config.authMode === 'service_account' && !config.serviceAccountJson) || (config.authMode === 'apps_script' && !config.appsScriptUrl)) {
    return { status: 'unverified', message: config.missingReason || 'Chưa cấu hình kết nối Google Drive.' };
  }

  try {
    if (config.authMode === 'service_account') {
      const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
      if (!parsed || error) {
        return { status: 'unverified', message: error || 'Khóa Service Account không hợp lệ.' };
      }
      const accessToken = await getServiceAccountAccessToken(parsed);
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(cleanId)}?supportsAllDrives=true&fields=id,name,mimeType,webViewLink,trashed`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );
      if (res.status === 404) {
        return { status: 'not_found', message: 'Không tìm thấy tệp trên Google Drive (404).' };
      }
      if (res.status === 401 || res.status === 403) {
        return { status: 'access_denied', message: 'Thiếu quyền truy cập tệp này trên Google Drive.' };
      }
      if (!res.ok) {
        return { status: 'unverified', message: `Google Drive trả về mã ${res.status}.` };
      }
      const data: any = await res.json().catch(() => ({}));
      if (data?.trashed) {
        return { status: 'not_found', message: 'Tệp đã bị chuyển vào thùng rác trên Google Drive.' };
      }
      return {
        status: 'ok',
        file: {
          fileId: String(data.id || cleanId),
          fileName: String(data.name || cleanId),
          webViewLink: String(data.webViewLink || `https://drive.google.com/file/d/${cleanId}/view`),
          mimeType: data.mimeType ? String(data.mimeType) : undefined,
        },
      };
    }

    // Apps Script mode: try get_file first, fallback to check_file
    const res = await fetch(config.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'get_file',
        folderId: config.folderId || cleanId,
        fileId: cleanId,
      }),
    });
    const data = await parseAppsScriptJsonResponse(res);
    if (data.result === 'success' && data.status === 'ok' && data.file) {
      return {
        status: 'ok',
        file: {
          fileId: String(data.file.fileId || cleanId),
          fileName: String(data.file.fileName || cleanId),
          webViewLink: String(data.file.webViewLink || `https://drive.google.com/file/d/${cleanId}/view`),
          mimeType: data.file.mimeType ? String(data.file.mimeType) : undefined,
        },
      };
    }
    if (data.result === 'success' && (data.status === 'access_denied' || data.status === 'not_found')) {
      return {
        status: data.status,
        message: data.error || data.message || (data.status === 'access_denied' ? 'Không có quyền truy cập tệp Drive.' : 'Không tìm thấy tệp trên Google Drive.'),
      };
    }

    // Backward compatible fallback to check_file on older Apps Script deployments
    const errText = String(data.error || '').toLowerCase();
    if (errText.includes('thao tác không hợp lệ') || errText.includes('get_file')) {
      const checkRes = await fetch(config.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'check_file',
          folderId: config.folderId || cleanId,
          fileId: cleanId,
        }),
      });
      const checkData = await parseAppsScriptJsonResponse(checkRes);
      if (checkData.result === 'success' && checkData.exists) {
        return {
          status: 'ok',
          file: {
            fileId: cleanId,
            fileName: String(checkData.fileName || cleanId),
            webViewLink: String(checkData.webViewLink || `https://drive.google.com/file/d/${cleanId}/view`),
          },
        };
      }
      if (checkData.status === 'access_denied') {
        return { status: 'access_denied', message: checkData.message || 'Thiếu quyền truy cập tệp Drive.' };
      }
      return { status: 'not_found', message: checkData.message || 'Không tìm thấy tệp trên Google Drive.' };
    }

    if (errText.includes('permission') || errText.includes('access') || errText.includes('quyền')) {
      return { status: 'access_denied', message: data.error };
    }
    if (errText.includes('not found') || errText.includes('no item') || errText.includes('không tìm thấy')) {
      return { status: 'not_found', message: 'Không tìm thấy tệp trên Google Drive.' };
    }
    return { status: 'unverified', message: data.error || 'Chưa xác minh được tệp trên Google Drive.' };
  } catch (err: any) {
    return {
      status: 'unverified',
      message: err?.message || 'Lỗi kết nối Google Drive khi lấy thông tin tệp.',
    };
  }
}

export async function listDriveFilesInFolder(
  config: ResolvedGoogleDriveConfig,
  folderIdOverride?: string,
  maxDepth: number = 4
): Promise<{ files: ArchiveDriveCandidate[]; error?: string }> {
  const rootFolderId = (folderIdOverride || config.folderId || '').trim();
  if (!rootFolderId) {
    return { files: [], error: 'Chưa chỉ định thư mục Google Drive để quét tệp.' };
  }
  if (!config.enabled) {
    return { files: [], error: config.missingReason || 'Chưa bật tích hợp Google Drive.' };
  }

  try {
    if (config.authMode === 'service_account') {
      const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
      if (!parsed || error) {
        return { files: [], error: error || 'Cấu hình Tài khoản dịch vụ không hợp lệ.' };
      }
      const accessToken = await getServiceAccountAccessToken(parsed);
      const collected: ArchiveDriveCandidate[] = [];
      const queue: Array<{ id: string; path: string; depth: number }> = [
        { id: rootFolderId, path: config.folderName || '', depth: 1 },
      ];
      const visitedFolders = new Set<string>();

      while (queue.length > 0 && collected.length < 3000) {
        const current = queue.shift()!;
        if (visitedFolders.has(current.id)) continue;
        visitedFolders.add(current.id);

        let pageToken = '';
        do {
          const q = `'${current.id.replace(/'/g, "\\'")}' in parents and trashed = false`;
          const url = new URL('https://www.googleapis.com/drive/v3/files');
          url.searchParams.set('q', q);
          url.searchParams.set('pageSize', '1000');
          url.searchParams.set('supportsAllDrives', 'true');
          url.searchParams.set('includeItemsFromAllDrives', 'true');
          url.searchParams.set('fields', 'nextPageToken,files(id,name,mimeType,webViewLink)');
          if (pageToken) url.searchParams.set('pageToken', pageToken);

          const res = await fetch(url.toString(), {
            headers: { Authorization: `Bearer ${accessToken}` },
          });
          if (!res.ok) {
            if (current.depth === 1) {
              const errJson: any = await res.json().catch(() => ({}));
              return {
                files: collected,
                error: `Không thể quét thư mục Drive (${res.status}): ${errJson?.error?.message || res.statusText}`,
              };
            }
            break;
          }
          const data: any = await res.json().catch(() => ({}));
          const items: any[] = Array.isArray(data.files) ? data.files : [];
          for (const item of items) {
            if (item.mimeType === 'application/vnd.google-apps.folder') {
              if (current.depth < maxDepth) {
                const subPath = current.path ? `${current.path}/${item.name}` : String(item.name || '');
                queue.push({ id: String(item.id), path: subPath, depth: current.depth + 1 });
              }
            } else {
              collected.push({
                fileId: String(item.id),
                fileName: String(item.name || ''),
                webViewLink: String(item.webViewLink || `https://drive.google.com/file/d/${item.id}/view`),
                mimeType: item.mimeType ? String(item.mimeType) : undefined,
                folderPath: current.path || undefined,
              });
            }
          }
          pageToken = data.nextPageToken ? String(data.nextPageToken) : '';
        } while (pageToken && collected.length < 3000);
      }

      return { files: collected };
    }

    if (!config.appsScriptUrl) {
      return { files: [], error: 'Chưa cấu hình URL Google Apps Script.' };
    }

    const res = await fetch(config.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'list_files',
        folderId: rootFolderId,
        maxDepth,
      }),
    });
    const data = await parseAppsScriptJsonResponse(res);
    if (data.result !== 'success' || !Array.isArray(data.files)) {
      return {
        files: [],
        error:
          data.error ||
          'Apps Script chưa hỗ trợ liệt kê tệp trong thư mục (cần cập nhật mẫu Apps Script mới hoặc dán trực tiếp liên kết Drive vào từng dòng).',
      };
    }
    return {
      files: data.files.map((f: any) => ({
        fileId: String(f.fileId || f.id || ''),
        fileName: String(f.fileName || f.name || ''),
        webViewLink: String(f.webViewLink || `https://drive.google.com/file/d/${f.fileId || f.id}/view`),
        mimeType: f.mimeType ? String(f.mimeType) : undefined,
        folderPath: f.folderPath ? String(f.folderPath) : undefined,
      })),
    };
  } catch (err: any) {
    return {
      files: [],
      error: err?.message || 'Lỗi khi quét danh sách tệp trong thư mục Google Drive.',
    };
  }
}

export async function downloadDriveFileBuffer(
  config: ResolvedGoogleDriveConfig,
  fileId: string
): Promise<{ ok: boolean; buffer?: Buffer; mimeType?: string; error?: string }> {
  const cleanId = (fileId || '').trim();
  if (!cleanId) return { ok: false, error: 'Thiếu mã tệp Google Drive.' };
  if (!config.enabled || config.authMode !== 'service_account' || !config.serviceAccountJson) {
    return { ok: false, error: 'Tải trực tiếp nội dung tệp chỉ hỗ trợ ở chế độ Service Account.' };
  }
  try {
    const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
    if (!parsed || error) return { ok: false, error: error || 'Cấu hình Service Account không hợp lệ.' };
    const accessToken = await getServiceAccountAccessToken(parsed);
    const res = await fetch(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(cleanId)}?alt=media&supportsAllDrives=true`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );
    if (!res.ok) {
      return { ok: false, error: `Không thể tải tệp từ Drive (${res.status}).` };
    }
    const arr = await res.arrayBuffer();
    return {
      ok: true,
      buffer: Buffer.from(arr),
      mimeType: res.headers.get('content-type') || 'application/pdf',
    };
  } catch (err: any) {
    return { ok: false, error: err?.message || 'Lỗi tải tệp từ Google Drive.' };
  }
}
