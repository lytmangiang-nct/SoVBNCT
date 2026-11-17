import crypto from 'crypto';
import type { GoogleDriveConfig } from '../src/types.js';
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

export function buildDriveFolderHierarchy(
  bookName: string,
  year: number,
  docCategory: 'outgoing' | 'incoming'
): { bookFolder: string; yearFolder: string; categoryFolder: string; displayPath: string } {
  const cleanBook = (bookName || 'Sổ văn bản').trim();
  const yearStr = String(year || new Date().getFullYear());
  const cat = docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi';
  return {
    bookFolder: cleanBook,
    yearFolder: yearStr,
    categoryFolder: cat,
    displayPath: `${cleanBook} → ${yearStr} → ${cat}`,
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

async function getServiceAccountFolderName(accessToken: string, folderId: string): Promise<string> {
  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folderId)}?supportsAllDrives=true&fields=id,name`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );
  if (!res.ok) return '';
  const data: any = await res.json().catch(() => ({}));
  return typeof data.name === 'string' ? data.name : '';
}

async function ensureServiceAccountSubfolder(
  accessToken: string,
  parentId: string,
  folderName: string
): Promise<string> {
  const safeName = folderName.replace(/'/g, "\\'");
  const q = `'${parentId}' in parents and name = '${safeName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
  const listUrl = `https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name)&q=${encodeURIComponent(
    q
  )}`;

  const listRes = await fetch(listUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (listRes.ok) {
    const listData: any = await listRes.json().catch(() => ({}));
    if (listData?.files?.[0]?.id) {
      return listData.files[0].id;
    }
  }

  const createRes = await fetch(
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id,name',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId],
      }),
    }
  );
  const createData: any = await createRes.json().catch(() => ({}));
  if (!createRes.ok || !createData.id) {
    throw new Error(`Không thể tạo thư mục "${folderName}" trên Google Drive: ${createData?.error?.message || createRes.statusText}`);
  }
  return createData.id;
}

async function resolveServiceAccountTargetFolder(
  accessToken: string,
  rootFolderId: string,
  hierarchy: { bookFolder: string; yearFolder: string; categoryFolder: string }
): Promise<string> {
  const rootName = await getServiceAccountFolderName(accessToken, rootFolderId);
  let currentParentId = rootFolderId;
  if (rootName.trim().toLowerCase() !== hierarchy.bookFolder.trim().toLowerCase()) {
    currentParentId = await ensureServiceAccountSubfolder(accessToken, currentParentId, hierarchy.bookFolder);
  }
  currentParentId = await ensureServiceAccountSubfolder(accessToken, currentParentId, hierarchy.yearFolder);
  currentParentId = await ensureServiceAccountSubfolder(accessToken, currentParentId, hierarchy.categoryFolder);
  return currentParentId;
}

async function findExistingFileServiceAccount(
  accessToken: string,
  folderId: string,
  docId: string
): Promise<{ id: string; name: string; webViewLink: string } | null> {
  const safeDocId = docId.replace(/'/g, "\\'");
  const q = `'${folderId}' in parents and trashed = false and appProperties has { key='nctDocId' and value='${safeDocId}' }`;
  const url = `https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&fields=files(id,name,webViewLink)&q=${encodeURIComponent(
    q
  )}`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) return null;
  const data: any = await res.json().catch(() => ({}));
  const first = data?.files?.[0];
  if (!first) return null;
  return {
    id: first.id,
    name: first.name,
    webViewLink: first.webViewLink || `https://drive.google.com/file/d/${first.id}/view`,
  };
}

export async function uploadWordToGoogleDrive(params: {
  config: ReturnType<typeof resolveEffectiveDriveConfig>;
  docId: string;
  bookName?: string;
  year?: number;
  docCategory?: 'outgoing' | 'incoming';
  fullDocNumber: string;
  docName: string;
  originalFileName: string;
  mimeType: string;
  buffer: Buffer;
  overrideFileName?: string;
  replaceMode?: boolean;
  previousFileId?: string;
  keepPreviousFile?: boolean;
}): Promise<DriveUploadResult> {
  const {
    config,
    docId,
    bookName = 'Đoàn trường',
    year = new Date().getFullYear(),
    docCategory = 'outgoing',
    fullDocNumber,
    docName,
    originalFileName,
    mimeType,
    buffer,
    overrideFileName,
    replaceMode,
    previousFileId,
    keepPreviousFile,
  } = params;

  if (!config.isConfigured || !config.folderId) {
    throw new Error(config.missingReason || 'Chưa cấu hình thư mục Google Drive hợp lệ.');
  }

  const hierarchy = buildDriveFolderHierarchy(bookName, year, docCategory);

  const requiredExt =
    getSupportedAttachmentExtension(originalFileName) ||
    (mimeType === 'application/pdf' ? '.pdf' : '.docx');

  const finalFileName =
    overrideFileName && overrideFileName.trim()
      ? sanitizeCustomDriveFileName(overrideFileName, requiredExt)
      : buildStandardDriveFileName(fullDocNumber, docName, originalFileName);

  if (config.authMode === 'apps_script') {
    // Step 1: Check existing file by docId first (skip when replacing or adding versioned attachments)
    if (!replaceMode && !keepPreviousFile) {
      const checkRes = await fetch(config.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'find',
          folderId: config.folderId,
          bookName: hierarchy.bookFolder,
          year: hierarchy.yearFolder,
          categoryFolder: hierarchy.categoryFolder,
          docId,
          expectedFileName: finalFileName,
        }),
      });
      const checkData = await parseAppsScriptJsonResponse(checkRes);
      if (checkData.result === 'success' && checkData.found && checkData.fileId) {
        return {
          fileId: checkData.fileId,
          fileName: checkData.fileName || finalFileName,
          webViewLink: checkData.webViewLink || `https://drive.google.com/file/d/${checkData.fileId}/view`,
          folderId: checkData.folderId || config.folderId,
          folderPath: checkData.folderPath || hierarchy.displayPath,
          deduplicated: true,
        };
      }
    }

    // Step 2: Upload new file into Tên sổ → Năm → Văn bản đi / Văn bản đến
    const uploadRes = await fetch(config.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'upload',
        folderId: config.folderId,
        bookName: hierarchy.bookFolder,
        year: hierarchy.yearFolder,
        categoryFolder: hierarchy.categoryFolder,
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

    return {
      fileId: uploadData.fileId,
      fileName: uploadData.fileName || finalFileName,
      webViewLink: uploadData.webViewLink || `https://drive.google.com/file/d/${uploadData.fileId}/view`,
      folderId: uploadData.folderId || config.folderId,
      folderPath: uploadData.folderPath || hierarchy.displayPath,
      deduplicated: Boolean(uploadData.deduplicated),
    };
  }

  // Service Account mode (Google Drive API v3)
  const { parsed, error } = parseServiceAccountJson(config.serviceAccountJson);
  if (!parsed || error) {
    throw new Error(error || 'Cấu hình Tài khoản dịch vụ không hợp lệ.');
  }

  const accessToken = await getServiceAccountAccessToken(parsed);

  // Resolve/create folder layout: Tên sổ → Năm → Văn bản đi / Văn bản đến
  const targetFolderId = await resolveServiceAccountTargetFolder(accessToken, config.folderId, hierarchy);

  // Step 1: Deduplication lookup by appProperties.nctDocId (skip when replacing or keeping separate versions)
  if (!replaceMode && !keepPreviousFile) {
    const existing = await findExistingFileServiceAccount(accessToken, targetFolderId, docId);
    if (existing) {
      if (existing.name !== finalFileName) {
        await fetch(
          `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(existing.id)}?supportsAllDrives=true`,
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
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(previousFileId)}?supportsAllDrives=true`,
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
  const boundary = `nct_upload_${Date.now()}_${crypto.randomBytes(4).toString('hex')}`;
  const metadata = JSON.stringify({
    name: finalFileName,
    parents: [targetFolderId],
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
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink',
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
