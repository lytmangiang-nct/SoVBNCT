import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';
import crypto from 'crypto';
import multer from 'multer';
import type {
  BookPermission,
  DocumentAuditAction,
  DocumentHistoryEntry,
  DocumentRecord,
  ExcelDeleteSyncStatus,
  SystemSettings,
  DocumentTypeConfig,
  SignerConfig,
  User,
  StatisticsData,
  PreviewNumberResponse,
  AuditLog,
  GoogleDriveConfig,
  PdfAnalyzeResponse,
  RegistryBook,
  NumberRequest,
  StartSequenceSeriesConfig,
  SystemAuditLog,
} from './src/types.js';
import {
  appendConflictSuffix,
  buildStandardDriveFileName,
  extractGoogleDriveFolderId,
  getSupportedAttachmentExtension,
  getWordFileExtension,
  normalizeCompareText,
  normalizeDateOnlyDdMmYyyy,
  normalizeDocNumberForCompare,
  sanitizeCustomDriveFileName,
} from './src/utils/fileNaming.js';
import {
  ALL_BOOK_PERMISSIONS,
  CHI_BO_BOOK_ID,
  DEFAULT_BOOK_ID,
  NHA_TRUONG_BOOK_ID,
  getDocBookId,
  getUserBookPermissions,
  hasBookPermission,
  isSystemAdmin,
  isTypeEnabledInBook,
  normalizeBookCode,
  sanitizeBookPermissionsMap,
} from './src/utils/permissions.js';
import {
  getMaxWordFileSizeBytes,
  getMaxWordFileSizeMB,
  parseServiceAccountJson,
  resolveEffectiveDriveConfig,
  sanitizeDriveConfigForClient,
  testGoogleDriveConnection,
  uploadFileToGoogleDrive,
  uploadWordToGoogleDrive,
  validatePdfFileBuffer,
  validateWordFileBuffer,
} from './server/googleDrive.js';
import { analyzePdfBuffer, fetchPdfFromUrl } from './server/pdfExtractor.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
const IS_SERVERLESS = Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME);

app.set('trust proxy', 1);

const extraCorsOrigins = (process.env.CORS_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

app.use(
  cors({
    origin: (origin, callback) => {
      if (!origin) return callback(null, true);
      if (extraCorsOrigins.includes(origin)) return callback(null, true);
      try {
        const url = new URL(origin);
        const host = url.hostname.toLowerCase();
        if (
          host === 'localhost' ||
          host === '127.0.0.1' ||
          host.endsWith('.vercel.app') ||
          host.endsWith('.run.app') ||
          host.endsWith('.aistudio.google.com')
        ) {
          return callback(null, true);
        }
        if (process.env.APP_URL) {
          const appHost = new URL(process.env.APP_URL).hostname.toLowerCase();
          if (host === appHost) return callback(null, true);
        }
      } catch {
        // ignore malformed origin
      }
      return callback(null, true);
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
  })
);

// Guard if serverless runtime already pre-parsed req.body before Express middleware runs
app.use((req: any, _res, next) => {
  if (req.body !== undefined && req.readableEnded) {
    req._body = true;
  }
  next();
});
app.use(express.json({ limit: '15mb' }));

// --- Database File Setup ---
const BUNDLED_DATA_DIR = path.resolve(__dirname, 'data');
const BUNDLED_DB_FILE = path.resolve(BUNDLED_DATA_DIR, 'db.json');
const DATA_DIR = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : IS_SERVERLESS
  ? '/tmp/nct-registry-data'
  : BUNDLED_DATA_DIR;
const DB_FILE = path.resolve(DATA_DIR, 'db.json');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

interface DatabaseSchema {
  settings: SystemSettings;
  documentTypes: DocumentTypeConfig[];
  signers: SignerConfig[];
  users: Array<User & { passwordHash: string }>;
  documents: DocumentRecord[];
  books: RegistryBook[];
  numberRequests: NumberRequest[];
  idempotencyKeys: Record<string, string>; // key -> documentId
  systemAuditLogs: SystemAuditLog[];
  documentHistoryLogs?: DocumentHistoryEntry[];
  revokedSessions?: Record<string, number>; // sid -> expiresAt ms
}

// Initial Seed Data
const DEFAULT_SETTINGS: SystemSettings = {
  unitName: 'Đoàn Trường THPT Nguyễn Chí Thanh',
  unitCode: 'ĐTN',
  numberMode: 'separate', // 'separate' or 'shared'
  numberFormat: '{so}-{loai}/{don_vi}',
  customStartSequences: {},
  msGraphConfig: {
    enabled: false,
    tenantId: '',
    clientId: '',
    clientSecret: '',
    driveId: '',
    filePath: '',
    tableName: 'SoCapSoVanBan',
  },
  googleSheetsConfig: {
    enabled: false,
    webhookUrl: '',
    lastSyncTime: '',
  },
  googleDriveConfig: {
    enabled: false,
    authMode: 'apps_script',
    folderUrl: '',
    folderId: '',
    folderName: '',
    folderByUnit: {},
    appsScriptUrl: '',
    serviceAccountJson: '',
    serviceAccountEmail: '',
  },
};

const DEFAULT_DOC_TYPES: DocumentTypeConfig[] = [
  { id: 'KH', code: 'KH', name: 'Kế hoạch', color: 'blue', isActive: true },
  { id: 'QĐ', code: 'QĐ', name: 'Quyết định', color: 'red', isActive: true },
  { id: 'BC', code: 'BC', name: 'Báo cáo', color: 'emerald', isActive: true },
  { id: 'ĐN', code: 'ĐN', name: 'Đề nghị', color: 'purple', isActive: true },
  { id: 'ND', code: 'ND', name: 'Nội dung', color: 'teal', isActive: true },
  { id: 'TB', code: 'TB', name: 'Thông báo', color: 'amber', isActive: true },
];

const DEFAULT_TYPE_COLORS: Record<string, string> = {
  KH: 'blue',
  BC: 'emerald',
  ĐN: 'purple',
  DN: 'purple',
  ND: 'teal',
  TB: 'amber',
  QĐ: 'red',
  QD: 'red',
  TTR: 'indigo',
  CV: 'cyan',
  BB: 'pink',
  HD: 'cyan',
  QC: 'orange',
  CTR: 'emerald',
};

const DEFAULT_SIGNERS: SignerConfig[] = [
  { id: 'sig-1', name: 'Nguyễn Văn Minh', role: 'Bí thư Đoàn trường' },
  { id: 'sig-2', name: 'Trần Thị Thu Thảo', role: 'Phó Bí thư Đoàn trường' },
  { id: 'sig-3', name: 'Lê Hoàng Nam', role: 'Ủy viên BTV Đoàn trường' },
];

function hashPassword(pass: string): string {
  return crypto.createHash('sha256').update(pass + 'salt_nct_2026').digest('hex');
}

const DEFAULT_USERS: Array<User & { passwordHash: string }> = [
  {
    id: 'user-admin',
    username: 'admin',
    fullName: 'Quản trị viên Đoàn Trường',
    role: 'admin',
    isActive: true,
    createdAt: new Date().toISOString(),
    passwordHash: hashPassword('admin@123'),
  },
];

let db: DatabaseSchema;

function stripDiacriticsBasic(input: string): string {
  return (input || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toUpperCase()
    .trim();
}

function resolveCanonicalTypeCode(rawCode: string, documentTypes: DocumentTypeConfig[]): string {
  const clean = String(rawCode || '').trim().toUpperCase();
  if (!clean) return '';
  if (clean === 'SHARED') return 'SHARED';

  const exact = documentTypes.find((t) => String(t.code || '').trim().toUpperCase() === clean);
  if (exact) {
    return exact.code.trim().toUpperCase();
  }

  const strippedClean = stripDiacriticsBasic(clean);
  const fuzzyMatches = documentTypes.filter(
    (t) => stripDiacriticsBasic(String(t.code || '')) === strippedClean
  );
  if (fuzzyMatches.length === 1) {
    return fuzzyMatches[0].code.trim().toUpperCase();
  }

  return clean;
}

function normalizeCustomStartSequencesMap(
  rawMap: any,
  bookCode: string,
  bookId: string,
  documentTypes: DocumentTypeConfig[]
): Record<string, number> {
  if (!rawMap || typeof rawMap !== 'object') return {};
  const result: Record<string, number> = {};
  const upperBookCode = String(bookCode || '').trim().toUpperCase();
  const cleanBookId = String(bookId || '').trim();

  for (const [rawKey, rawVal] of Object.entries(rawMap)) {
    const seqNum = typeof rawVal === 'number' ? rawVal : parseInt(String(rawVal), 10);
    if (!Number.isFinite(seqNum) || seqNum < 1) continue;

    let keyStr = String(rawKey || '').trim();
    if (!keyStr) continue;

    // Strip optional composite prefix: `${bookCode}:`, `${bookId}:`, or `${bookCode}_`
    if (keyStr.includes(':')) {
      const parts = keyStr.split(':').filter(Boolean);
      if (parts.length === 3) {
        const prefix = parts[0].trim();
        if (
          prefix.toUpperCase() !== upperBookCode &&
          prefix !== cleanBookId &&
          upperBookCode
        ) {
          continue;
        }
        keyStr = `${parts[1]}_${parts[2]}`;
      }
    } else if (upperBookCode && keyStr.toUpperCase().startsWith(`${upperBookCode}_`)) {
      keyStr = keyStr.slice(upperBookCode.length + 1);
    }

    const match = keyStr.match(/^(\d{4})[_\-:](.+)$/);
    if (!match) continue;
    const yearNum = parseInt(match[1], 10);
    if (!Number.isFinite(yearNum) || yearNum < 2000 || yearNum > 2100) continue;
    const canonicalType = resolveCanonicalTypeCode(match[2], documentTypes);
    if (!canonicalType) continue;

    const normalizedKey = `${yearNum}_${canonicalType}`;
    const existing = result[normalizedKey];
    result[normalizedKey] = existing ? Math.max(existing, Math.floor(seqNum)) : Math.floor(seqNum);
  }

  return result;
}

function createDefaultBookFromSettings(settings: SystemSettings, signers: SignerConfig[]): RegistryBook {
  const unitName = settings.unitName || 'Đoàn Trường THPT Nguyễn Chí Thanh';
  return {
    id: DEFAULT_BOOK_ID,
    code: 'DOAN_TRUONG',
    name: 'Đoàn trường',
    unitName,
    unitCode: settings.unitCode || 'ĐTN',
    isShared: true,
    numberMode: settings.numberMode || 'separate',
    numberFormat: settings.numberFormat || '{so}-{loai}/{don_vi}',
    customStartSequences: { ...(settings.customStartSequences || {}) },
    signers: (signers || []).map((s) => ({ ...s })),
    isActive: true,
    createdAt: new Date().toISOString(),
  };
}

function createPrebuiltBooks(): RegistryBook[] {
  const now = new Date().toISOString();
  return [
    {
      id: CHI_BO_BOOK_ID,
      code: 'CHI_BO',
      name: 'Chi bộ',
      unitName: 'Chi bộ Trường THPT Nguyễn Chí Thanh',
      unitCode: 'CB',
      isShared: false,
      numberMode: 'separate',
      numberFormat: '{so}-{loai}/{don_vi}',
      customStartSequences: {},
      signers: [
        { id: 'sig-cb-1', name: 'Bí thư Chi bộ', role: 'Bí thư Chi bộ' },
        { id: 'sig-cb-2', name: 'Phó Bí thư Chi bộ', role: 'Phó Bí thư Chi bộ' },
      ],
      isActive: true,
      createdAt: now,
    },
    {
      id: NHA_TRUONG_BOOK_ID,
      code: 'NHA_TRUONG',
      name: 'Nhà trường',
      unitName: 'Trường THPT Nguyễn Chí Thanh',
      unitCode: 'THPTNCT',
      isShared: true,
      numberMode: 'separate',
      numberFormat: '{so}-{loai}/{don_vi}',
      customStartSequences: {},
      signers: [
        { id: 'sig-nt-1', name: 'Hiệu trưởng', role: 'Hiệu trưởng' },
        { id: 'sig-nt-2', name: 'Phó Hiệu trưởng', role: 'Phó Hiệu trưởng' },
      ],
      isActive: true,
      createdAt: now,
    },
  ];
}

export function migrateDatabaseIdempotent(loaded: Partial<DatabaseSchema>): DatabaseSchema {
  const settings: SystemSettings = {
    ...DEFAULT_SETTINGS,
    ...(loaded.settings || {}),
    msGraphConfig: { ...DEFAULT_SETTINGS.msGraphConfig, ...(loaded.settings?.msGraphConfig || {}) },
    googleSheetsConfig: {
      enabled: false,
      webhookUrl: '',
      lastSyncTime: '',
      ...(loaded.settings?.googleSheetsConfig || {}),
    },
    googleDriveConfig: {
      enabled: false,
      authMode: 'apps_script',
      folderUrl: '',
      folderId: '',
      folderName: '',
      folderByUnit: {},
      appsScriptUrl: '',
      serviceAccountJson: '',
      serviceAccountEmail: '',
      ...(loaded.settings?.googleDriveConfig || {}),
    },
    configVersion: loaded.settings?.configVersion || 1,
    updatedAt: loaded.settings?.updatedAt || new Date().toISOString(),
  };

  const paletteKeys = ['blue', 'emerald', 'amber', 'red', 'purple', 'teal', 'indigo', 'orange', 'cyan', 'pink'];
  const documentTypes = Array.isArray(loaded.documentTypes) ? [...loaded.documentTypes] : [...DEFAULT_DOC_TYPES];
  // Bảo đảm danh mục loại văn bản chung luôn có sẵn Quyết định (QĐ)
  if (!documentTypes.some((t) => t.code.toUpperCase() === 'QĐ' || t.code.toUpperCase() === 'QD')) {
    const khIdx = documentTypes.findIndex((t) => t.code.toUpperCase() === 'KH');
    const qdType: DocumentTypeConfig = { id: 'QĐ', code: 'QĐ', name: 'Quyết định', color: 'red', isActive: true };
    if (khIdx !== -1) {
      documentTypes.splice(khIdx + 1, 0, qdType);
    } else {
      documentTypes.push(qdType);
    }
  }
  documentTypes.forEach((t, idx) => {
    if (!t.color) {
      t.color = DEFAULT_TYPE_COLORS[t.code.toUpperCase()] || paletteKeys[idx % paletteKeys.length];
    }
  });

  const signers = Array.isArray(loaded.signers) ? loaded.signers : [...DEFAULT_SIGNERS];

  const books: RegistryBook[] = Array.isArray(loaded.books) ? loaded.books.map((b) => ({ ...b })) : [];
  let defaultBook = books.find((b) => b.id === DEFAULT_BOOK_ID);
  if (!defaultBook) {
    defaultBook = createDefaultBookFromSettings(settings, signers);
    books.unshift(defaultBook);
  } else {
    if (!defaultBook.code) defaultBook.code = 'DOAN_TRUONG';
    if (!defaultBook.name || defaultBook.name === `Sổ văn bản ${settings.unitName}`) {
      defaultBook.name = 'Đoàn trường';
    }
    if (!defaultBook.unitName) defaultBook.unitName = settings.unitName;
    if (!defaultBook.unitCode) defaultBook.unitCode = settings.unitCode;
    if (!defaultBook.numberFormat) defaultBook.numberFormat = settings.numberFormat;
    if (!defaultBook.numberMode) defaultBook.numberMode = settings.numberMode;
    if (!defaultBook.customStartSequences || typeof defaultBook.customStartSequences !== 'object') {
      defaultBook.customStartSequences = { ...(settings.customStartSequences || {}) };
    } else if (
      Object.keys(defaultBook.customStartSequences).length === 0 &&
      settings.customStartSequences &&
      Object.keys(settings.customStartSequences).length > 0
    ) {
      // Bảo toàn mốc cũ từ settings.customStartSequences (vd 2026_KH = 10) vào sổ Đoàn trường
      defaultBook.customStartSequences = { ...settings.customStartSequences };
    }
    if (defaultBook.isActive === undefined) defaultBook.isActive = true;
    if (defaultBook.isShared === undefined) defaultBook.isShared = true;
    if (!Array.isArray(defaultBook.signers)) {
      defaultBook.signers = signers.map((s) => ({ ...s }));
    }
  }

  // Tạo sẵn 2 sổ Chi bộ & Nhà trường nếu chưa có (an toàn khi chạy lại nhiều lần)
  for (const prebuilt of createPrebuiltBooks()) {
    if (!books.some((b) => b.id === prebuilt.id || b.code === prebuilt.code)) {
      books.push(prebuilt);
    }
  }

  // Đảm bảo mọi sổ đều có mã nội bộ cố định (code) duy nhất và chuẩn hóa customStartSequences
  const usedCodes = new Set<string>();
  books.forEach((b, idx) => {
    let baseCode = normalizeBookCode(b.code || b.name || `SO_${idx + 1}`) || `SO_${idx + 1}`;
    let candidate = baseCode;
    let suffix = 2;
    while (usedCodes.has(candidate)) {
      candidate = `${baseCode}_${suffix}`;
      suffix++;
    }
    b.code = candidate;
    usedCodes.add(candidate);
    if (b.isActive === undefined) b.isActive = true;
    b.customStartSequences = normalizeCustomStartSequencesMap(
      b.customStartSequences || {},
      b.code,
      b.id,
      documentTypes
    );
    if (!Array.isArray(b.signers)) b.signers = [];
    if (!b.configVersion) b.configVersion = 1;
    if (!b.updatedAt) b.updatedAt = b.createdAt || new Date().toISOString();
  });

  if (defaultBook) {
    settings.customStartSequences = { ...(defaultBook.customStartSequences || {}) };
  }

  const bookMap = new Map<string, RegistryBook>();
  books.forEach((b) => bookMap.set(b.id, b));

  const documents = Array.isArray(loaded.documents) ? loaded.documents : [];
  documents.forEach((doc) => {
    if (!doc.bookId || !bookMap.has(doc.bookId)) {
      doc.bookId = DEFAULT_BOOK_ID;
    }
    const assignedBook = bookMap.get(doc.bookId) || defaultBook!;
    doc.bookCode = assignedBook.code;
    if (!doc.bookName || doc.bookName.startsWith('Sổ văn bản Đoàn Trường')) {
      doc.bookName = assignedBook.name;
    }
    // Chuẩn hóa ngày ban hành/ngày ký chỉ gồm dd/MM/yyyy; giữ nguyên rỗng nếu văn bản đến không có ngày ban hành
    if (doc.issueDate && doc.issueDate.trim()) {
      doc.issueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate) || doc.issueDate.trim();
    } else if (doc.docCategory === 'incoming') {
      doc.issueDate = '';
    }
    if (doc.signDate && doc.signDate.trim()) {
      doc.signDate = normalizeDateOnlyDdMmYyyy(doc.signDate) || undefined;
    }
  });

  const rawUsers = (Array.isArray(loaded.users) ? loaded.users : []).filter(
    (u) => u.username.toLowerCase() !== 'doantruong'
  );
  const adminIdx = rawUsers.findIndex((u) => u.username.toLowerCase() === 'admin');
  if (adminIdx === -1) {
    rawUsers.unshift({ ...DEFAULT_USERS[0] });
  } else {
    rawUsers[adminIdx].role = 'admin';
    rawUsers[adminIdx].isActive = true;
  }

  rawUsers.forEach((u) => {
    if (u.role !== 'admin') {
      u.role = 'issuer';
      if (!u.bookPermissions || typeof u.bookPermissions !== 'object') {
        // Mặc định an toàn cho tài khoản đã tạo trước đó: chỉ Xem và Xin số (không tự động cấp quyền Cấp số)
        u.bookPermissions = { [DEFAULT_BOOK_ID]: ['view', 'request_number'] };
      } else {
        u.bookPermissions = sanitizeBookPermissionsMap(u.bookPermissions);
      }
    }
  });

  const numberRequests = Array.isArray(loaded.numberRequests) ? loaded.numberRequests : [];
  const idempotencyKeys = loaded.idempotencyKeys && typeof loaded.idempotencyKeys === 'object' ? loaded.idempotencyKeys : {};
  const systemAuditLogs = Array.isArray(loaded.systemAuditLogs) ? loaded.systemAuditLogs : [];
  const existingHistory = Array.isArray(loaded.documentHistoryLogs) ? [...loaded.documentHistoryLogs] : [];
  const historySeen = new Set<string>();
  for (const h of existingHistory) {
    historySeen.add(`${h.docId}|${h.timestamp}|${h.action}|${h.userId}`);
  }
  for (const doc of documents) {
    const bId = getDocBookId(doc);
    const assignedBook = bookMap.get(bId) || defaultBook!;
    if (Array.isArray(doc.auditLogs)) {
      for (const log of doc.auditLogs) {
        const sig = `${doc.id}|${log.timestamp}|${log.action}|${log.userId}`;
        if (!log.timestampVietnam && log.timestamp) {
          log.timestampVietnam = formatVietnamDateTime(new Date(log.timestamp));
        }
        if (!log.dateKeyVietnam && log.timestamp) {
          log.dateKeyVietnam = getVietnamDateKey(new Date(log.timestamp));
        }
        if (!log.docId) log.docId = doc.id;
        if (!log.fullDocNumber) log.fullDocNumber = doc.fullDocNumber;
        if (!log.docName) log.docName = doc.docName || doc.title;
        if (!log.bookId) log.bookId = assignedBook.id;
        if (!log.bookCode) log.bookCode = assignedBook.code;
        if (!log.bookName) log.bookName = assignedBook.name;
        if (!historySeen.has(sig)) {
          historySeen.add(sig);
          const unameMatch = (log.userName || '').match(/\(([^)]+)\)\s*$/);
          existingHistory.push({
            ...log,
            id: log.id || `hist-${crypto.createHash('md5').update(sig).digest('hex').slice(0, 12)}`,
            timestampVietnam: log.timestampVietnam || formatVietnamDateTime(new Date(log.timestamp || Date.now())),
            dateKeyVietnam: log.dateKeyVietnam || getVietnamDateKey(new Date(log.timestamp || Date.now())),
            username: log.username || (unameMatch ? unameMatch[1] : log.userId || 'unknown'),
            bookId: log.bookId || assignedBook.id,
            bookCode: log.bookCode || assignedBook.code,
            bookName: log.bookName || assignedBook.name,
            docId: doc.id,
            fullDocNumber: doc.fullDocNumber,
            docName: doc.docName || doc.title,
          });
        }
      }
    }
  }
  existingHistory.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  const nowMs = Date.now();
  const revokedSessions: Record<string, number> = {};
  if (loaded.revokedSessions && typeof loaded.revokedSessions === 'object') {
    for (const [sid, exp] of Object.entries(loaded.revokedSessions)) {
      if (typeof exp === 'number' && exp > nowMs) {
        revokedSessions[sid] = exp;
      }
    }
  }

  return {
    settings,
    documentTypes,
    signers,
    users: rawUsers,
    documents,
    books,
    numberRequests,
    idempotencyKeys,
    systemAuditLogs,
    documentHistoryLogs: existingHistory,
    revokedSessions,
  };
}

function loadDatabase(): DatabaseSchema {
  const sourceFile = fs.existsSync(DB_FILE)
    ? DB_FILE
    : BUNDLED_DB_FILE !== DB_FILE && fs.existsSync(BUNDLED_DB_FILE)
    ? BUNDLED_DB_FILE
    : null;

  if (sourceFile) {
    try {
      const content = fs.readFileSync(sourceFile, 'utf-8');
      const loaded = JSON.parse(content) as Partial<DatabaseSchema>;
      const migrated = migrateDatabaseIdempotent(loaded);
      saveDatabaseSync(migrated);
      return migrated;
    } catch (e) {
      console.error('Lỗi khi đọc file db.json, khởi tạo lại:', e);
    }
  }

  const initialDb = migrateDatabaseIdempotent({
    settings: DEFAULT_SETTINGS,
    documentTypes: DEFAULT_DOC_TYPES,
    signers: DEFAULT_SIGNERS,
    users: DEFAULT_USERS,
    documents: [],
    books: [],
    numberRequests: [],
    idempotencyKeys: {},
    systemAuditLogs: [],
    revokedSessions: {},
  });
  saveDatabaseSync(initialDb);
  return initialDb;
}

function saveDatabaseSync(data: DatabaseSchema) {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  const tmpFile = `${DB_FILE}.tmp.${Date.now()}.${Math.random().toString(36).slice(2, 6)}`;
  fs.writeFileSync(tmpFile, JSON.stringify(data, null, 2), 'utf-8');
  fs.renameSync(tmpFile, DB_FILE);
}

db = loadDatabase();

function appendSystemAuditLog(entry: Omit<SystemAuditLog, 'id' | 'timestamp'>) {
  if (!Array.isArray(db.systemAuditLogs)) db.systemAuditLogs = [];
  db.systemAuditLogs.unshift({
    id: `sys-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
    timestamp: new Date().toISOString(),
    ...entry,
  });
  if (db.systemAuditLogs.length > 500) {
    db.systemAuditLogs = db.systemAuditLogs.slice(0, 500);
  }
}

export function buildDocumentSnapshot(doc: DocumentRecord, book: RegistryBook) {
  return {
    id: doc.id,
    bookId: book.id,
    bookCode: book.code,
    bookName: book.name,
    year: doc.year,
    issueDate: normalizeDateOnlyDdMmYyyy(doc.issueDate) || doc.issueDate || '',
    signDate: normalizeDateOnlyDdMmYyyy(doc.signDate || '') || '',
    docCategory: doc.docCategory || 'outgoing',
    issuingAgency: doc.issuingAgency || '',
    typeCode: doc.typeCode,
    typeName: doc.typeName,
    sequenceNumber: doc.sequenceNumber,
    fullDocNumber: doc.fullDocNumber,
    docName: doc.docName || doc.title,
    title: doc.title || '',
    signer: doc.signer,
    signerRole: doc.signerRole,
    issuedBy: doc.issuedBy,
    issuedByUsername: doc.issuedByUsername,
    issuedByUserId: doc.issuedByUserId,
    status: doc.status,
    cancelReason: doc.cancelReason || '',
    notes: doc.notes || '',
    driveStatus: doc.driveStatus || 'none',
    driveFileId: doc.driveFileId || '',
    driveFileName: doc.driveFileName || '',
    driveFileUrl: doc.driveFileUrl || '',
    driveFolderPath: doc.driveFolderPath || '',
    createdAt: doc.createdAt,
  };
}

function recordDocumentHistoryEntry(params: {
  doc: DocumentRecord;
  book: RegistryBook;
  actor: { id: string; username: string; fullName: string };
  action: DocumentAuditAction;
  changes: Record<string, { old: string | null; new: string | null }>;
  reason?: string;
  timestampIso?: string;
  snapshot?: Record<string, any>;
  pushToDocAuditLogs?: boolean;
}): DocumentHistoryEntry {
  const nowDate = params.timestampIso ? new Date(params.timestampIso) : new Date();
  const timestamp = Number.isNaN(nowDate.getTime()) ? new Date().toISOString() : nowDate.toISOString();
  const timestampVietnam = formatVietnamDateTime(new Date(timestamp));
  const dateKeyVietnam = getVietnamDateKey(new Date(timestamp));
  const id = `hist-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
  const entry: DocumentHistoryEntry = {
    id,
    timestamp,
    timestampVietnam,
    dateKeyVietnam,
    userId: params.actor.id,
    username: params.actor.username,
    userName: `${params.actor.fullName} (${params.actor.username})`,
    bookId: params.book.id,
    bookCode: params.book.code,
    bookName: params.book.name,
    docId: params.doc.id,
    fullDocNumber: params.doc.fullDocNumber,
    docName: params.doc.docName || params.doc.title,
    action: params.action,
    changes: params.changes,
    ...(params.reason ? { reason: params.reason } : {}),
    ...(params.snapshot ? { snapshot: params.snapshot } : {}),
  };

  if (params.pushToDocAuditLogs !== false) {
    if (!Array.isArray(params.doc.auditLogs)) params.doc.auditLogs = [];
    params.doc.auditLogs.push(entry);
  }
  if (!Array.isArray(db.documentHistoryLogs)) {
    db.documentHistoryLogs = [];
  }
  db.documentHistoryLogs.unshift(entry);
  return entry;
}

export function buildGoogleSheetsPayload(
  doc: DocumentRecord,
  book: RegistryBook,
  action:
    | 'create'
    | 'update'
    | 'cancel'
    | 'delete'
    | 'delete_append_to_deleted'
    | 'delete_remove_main'
    | 'test' = 'create',
  extra: Record<string, any> = {}
) {
  const driveStatusLabel =
    doc.driveStatus === 'saved'
      ? 'Đã lưu'
      : doc.driveStatus === 'failed'
      ? 'Lưu tệp thất bại'
      : doc.driveStatus === 'uploading'
      ? 'Đang tải lên'
      : 'Không đính kèm';
  const statusLabel =
    doc.status === 'deleted'
      ? `Đã xóa (${doc.deleteReason || 'Đã xóa khỏi sổ hiệu lực'})`
      : doc.status === 'cancelled'
      ? `Đã hủy (${doc.cancelReason || 'Lý do sai sót'})`
      : doc.docCategory === 'incoming'
      ? 'Đã ghi sổ văn bản đến'
      : 'Đã cấp số';

  return {
    action,
    matchBy: 'id',
    id: doc.id,
    bookId: book.id,
    bookCode: book.code,
    bookName: book.name,
    unitCode: book.unitCode,
    stt: doc.stt,
    year: doc.year,
    issueDate: normalizeDateOnlyDdMmYyyy(doc.issueDate) || doc.issueDate || '',
    signDate: normalizeDateOnlyDdMmYyyy(doc.signDate || '') || '',
    docCategory: doc.docCategory || 'outgoing',
    categoryLabel: doc.docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi',
    typeCode: doc.typeCode,
    typeName: doc.typeName,
    sequenceNumber: doc.docCategory === 'incoming' ? 0 : doc.sequenceNumber,
    fullDocNumber: doc.fullDocNumber,
    docName: doc.docName || doc.title,
    title: doc.title || '',
    signer: doc.signer,
    signerRole: doc.signerRole,
    issuedBy: doc.issuedBy,
    status: statusLabel,
    driveStatus: doc.driveStatus || 'none',
    driveStatusLabel,
    driveFileName: doc.driveFileName || '',
    driveFileId: doc.driveFileId || '',
    driveFileUrl: doc.driveFileUrl || '',
    driveFolderPath: doc.driveFolderPath || '',
    notes: doc.notes || '',
    deletedBy: doc.deletedBy || '',
    deletedByUsername: doc.deletedByUsername || '',
    deletedAt: doc.deletedAt || '',
    deletedAtVietnam: doc.deletedAtVietnam || '',
    deleteReason: doc.deleteReason || '',
    deletedSheetName: 'VB đã xóa',
    timestamp: new Date().toISOString(),
    ...extra,
  };
}

function getGoogleSheetsTargets(book: RegistryBook) {
  const targets: Array<{
    url: string;
    cfg: { enabled: boolean; webhookUrl: string; lastSyncTime?: string };
    scope: 'master' | 'book';
  }> = [];
  const masterCfg = db.settings.googleSheetsConfig;
  if (masterCfg && masterCfg.enabled && masterCfg.webhookUrl?.trim()) {
    targets.push({ url: masterCfg.webhookUrl.trim(), cfg: masterCfg, scope: 'master' });
  }
  const bookCfg = book.googleSheetsConfig;
  if (
    bookCfg &&
    bookCfg.enabled &&
    bookCfg.webhookUrl?.trim() &&
    bookCfg.webhookUrl.trim() !== masterCfg?.webhookUrl?.trim()
  ) {
    targets.push({ url: bookCfg.webhookUrl.trim(), cfg: bookCfg, scope: 'book' });
  }
  return targets;
}

async function postJsonWithTimeout(url: string, payload: Record<string, any>, timeoutMs = 8000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });
    const text = await res.text();
    let json: any = null;
    try {
      json = text ? JSON.parse(text) : null;
    } catch {
      json = null;
    }
    if (!res.ok) {
      throw new Error(json?.error || `HTTP ${res.status} từ dịch vụ đồng bộ Excel`);
    }
    if (json && json.result === 'error') {
      throw new Error(json.error || 'Dịch vụ đồng bộ Excel trả về lỗi');
    }
    return { res, json };
  } finally {
    clearTimeout(timer);
  }
}

// --- Google Sheets / Online Excel Webhook Sync Helpers ---
async function syncToGoogleSheets(
  doc: DocumentRecord,
  action: 'create' | 'update' | 'cancel' | 'delete' = 'create'
): Promise<{ attempted: boolean; ok: boolean; error?: string }> {
  const book = getBookById(getDocBookId(doc));
  if (action === 'delete' || doc.status === 'deleted') {
    return syncDeletedDocumentToOnlineExcel(doc, book);
  }

  const targets = getGoogleSheetsTargets(book);
  if (targets.length === 0) {
    return { attempted: false, ok: true };
  }

  const payload = buildGoogleSheetsPayload(doc, book, action);
  const errors: string[] = [];

  for (const target of targets) {
    try {
      await postJsonWithTimeout(target.url, { ...payload, scope: target.scope });
      target.cfg.lastSyncTime = new Date().toISOString();
    } catch (err: any) {
      const msg = err?.message || 'Lỗi kết nối đồng bộ Excel';
      console.error(`[Google Sheets Sync - ${book.code}] Lỗi gửi webhook:`, msg);
      errors.push(msg);
    }
  }

  if (errors.length === 0) {
    doc.syncStatus = 'synced';
    doc.syncedAt = new Date().toISOString();
    delete doc.syncError;
    saveDatabaseSync(db);
    return { attempted: true, ok: true };
  }

  doc.syncStatus = 'error';
  doc.syncError = errors[0];
  saveDatabaseSync(db);
  return { attempted: true, ok: false, error: errors[0] };
}

async function syncDeletedDocumentToOnlineExcel(
  doc: DocumentRecord,
  book: RegistryBook
): Promise<{ attempted: boolean; ok: boolean; error?: string }> {
  const targets = getGoogleSheetsTargets(book);
  if (targets.length === 0) {
    doc.excelDeleteSyncStatus = 'not_configured';
    doc.excelMovedToDeletedSheet = false;
    doc.excelRemovedFromMainSheet = false;
    delete doc.excelSyncError;
    return { attempted: false, ok: true };
  }

  if (!doc.excelSyncTargets || typeof doc.excelSyncTargets !== 'object') {
    doc.excelSyncTargets = {};
  }

  const errors: string[] = [];

  for (const target of targets) {
    const existingState = doc.excelSyncTargets[target.url] || {
      movedToDeletedSheet: Boolean(doc.excelMovedToDeletedSheet),
      removedFromMainSheet: Boolean(doc.excelRemovedFromMainSheet),
    };
    const tState = { ...existingState };

    try {
      // Bước 1: Tạo/cập nhật dòng trong sheet "VB đã xóa" theo mã hồ sơ cố định (doc.id) và xác nhận
      if (!tState.movedToDeletedSheet) {
        const step1Payload = buildGoogleSheetsPayload(doc, book, 'delete_append_to_deleted', {
          scope: target.scope,
        });
        const { json } = await postJsonWithTimeout(target.url, step1Payload);
        if (json && json.confirmed === false) {
          throw new Error('Chưa xác nhận ghi nhận xong trong sheet "VB đã xóa".');
        }
        tState.movedToDeletedSheet = true;
        delete tState.error;
        doc.excelSyncTargets[target.url] = tState;
        doc.excelMovedToDeletedSheet = true;
        saveDatabaseSync(db);
      }

      // Bước 2: Chỉ xóa dòng khỏi sheet theo dõi chính sau khi đã xác nhận ghi xong ở sheet "VB đã xóa"
      if (tState.movedToDeletedSheet && !tState.removedFromMainSheet) {
        const step2Payload = buildGoogleSheetsPayload(doc, book, 'delete_remove_main', {
          scope: target.scope,
        });
        await postJsonWithTimeout(target.url, step2Payload);
        tState.removedFromMainSheet = true;
        tState.syncedAt = new Date().toISOString();
        delete tState.error;
        doc.excelSyncTargets[target.url] = tState;
        target.cfg.lastSyncTime = tState.syncedAt;
      }
    } catch (err: any) {
      const msg = err?.message || 'Mất kết nối khi đồng bộ xóa sang Excel.';
      tState.error = msg;
      doc.excelSyncTargets[target.url] = tState;
      errors.push(msg);
      console.error(`[Excel Delete Sync - ${book.code}] Lỗi đồng bộ xóa:`, msg);
    }
  }

  const allMoved = targets.every((t) => Boolean(doc.excelSyncTargets?.[t.url]?.movedToDeletedSheet));
  const anyMoved = targets.some((t) => Boolean(doc.excelSyncTargets?.[t.url]?.movedToDeletedSheet));
  const allRemoved = targets.every((t) => Boolean(doc.excelSyncTargets?.[t.url]?.removedFromMainSheet));

  doc.excelMovedToDeletedSheet = anyMoved;
  doc.excelRemovedFromMainSheet = allRemoved;

  if (allMoved && allRemoved && errors.length === 0) {
    doc.excelDeleteSyncStatus = 'synced';
    doc.excelSyncedAt = new Date().toISOString();
    doc.syncStatus = 'synced';
    doc.syncedAt = doc.excelSyncedAt;
    delete doc.excelSyncError;
    delete doc.syncError;
    saveDatabaseSync(db);
    return { attempted: true, ok: true };
  }

  doc.excelDeleteSyncStatus = 'failed';
  doc.excelSyncError = errors[0] || 'Đồng bộ xóa Excel chưa hoàn tất.';
  doc.syncStatus = 'error';
  doc.syncError = doc.excelSyncError;
  saveDatabaseSync(db);
  return { attempted: true, ok: false, error: doc.excelSyncError };
}

// --- Thread-Safe Atomic Lock Mutex ---
class AsyncLock {
  private p = Promise.resolve();
  async acquire(): Promise<() => void> {
    let release: () => void = () => {};
    const next = new Promise<void>((res) => {
      release = res;
    });
    const wait = this.p;
    this.p = this.p.then(() => next);
    await wait;
    return release;
  }
}

const atomicLock = new AsyncLock();
const inFlightDocUploads = new Map<string, Promise<DocumentRecord>>();

function sanitizeSettingsForUser(settings: SystemSettings, user?: User): SystemSettings {
  const isAdmin = user?.role === 'admin';
  const safeDrive = sanitizeDriveConfigForClient(settings.googleDriveConfig, isAdmin, settings.unitCode);
  return {
    ...settings,
    msGraphConfig: isAdmin
      ? {
          ...settings.msGraphConfig,
          clientSecret: settings.msGraphConfig.clientSecret ? '••••••••' : '',
        }
      : {
          enabled: settings.msGraphConfig.enabled,
          tenantId: '',
          clientId: '',
          clientSecret: '',
          driveId: '',
          filePath: '',
          tableName: settings.msGraphConfig.tableName,
          lastSyncTime: settings.msGraphConfig.lastSyncTime,
        },
    googleSheetsConfig: settings.googleSheetsConfig
      ? isAdmin
        ? settings.googleSheetsConfig
        : {
            enabled: settings.googleSheetsConfig.enabled,
            webhookUrl: '',
            lastSyncTime: settings.googleSheetsConfig.lastSyncTime,
          }
      : undefined,
    googleDriveConfig: safeDrive,
  };
}

const wordUploadMiddleware = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: getMaxWordFileSizeBytes(),
    files: 1,
  },
}).single('file');

function handleWordUploadMulter(req: Request, res: Response, next: NextFunction) {
  wordUploadMiddleware(req, res, (err: any) => {
    if (err) {
      if (err.code === 'LIMIT_FILE_SIZE') {
        return res.status(413).json({
          error: `Dung lượng tệp vượt quá giới hạn tối đa ${getMaxWordFileSizeMB()} MB của máy chủ.`,
        });
      }
      return res.status(400).json({
        error: err.message || 'Lỗi khi tiếp nhận tệp tải lên.',
      });
    }
    next();
  });
}

// --- Book & Permission Helpers ---
function getDefaultBook(): RegistryBook {
  return db.books.find((b) => b.id === DEFAULT_BOOK_ID) || db.books[0];
}

function getBookById(bookId?: string): RegistryBook {
  if (bookId && bookId.trim()) {
    const found = db.books.find((b) => b.id === bookId.trim());
    if (found) return found;
  }
  return getDefaultBook();
}

function resolveDriveForBook(book: RegistryBook) {
  return resolveEffectiveDriveConfig(
    db.settings.googleDriveConfig,
    book.unitCode,
    book.googleDriveConfig,
    book.id === DEFAULT_BOOK_ID
  );
}

function sanitizeBookForUser(book: RegistryBook, user?: User): RegistryBook {
  const canManage = isSystemAdmin(user) || hasBookPermission(user, book.id, 'manage_book');
  return {
    ...book,
    googleDriveConfig: sanitizeDriveConfigForClient(
      db.settings.googleDriveConfig,
      canManage,
      book.unitCode,
      book.googleDriveConfig,
      book.id === DEFAULT_BOOK_ID
    ),
    googleSheetsConfig: book.googleSheetsConfig
      ? canManage
        ? book.googleSheetsConfig
        : {
            enabled: book.googleSheetsConfig.enabled,
            webhookUrl: '',
            lastSyncTime: book.googleSheetsConfig.lastSyncTime,
          }
      : undefined,
  };
}

function syncDefaultBookToLegacySettings(book: RegistryBook) {
  if (book.id !== DEFAULT_BOOK_ID) return;
  db.settings.unitName = book.unitName;
  db.settings.unitCode = book.unitCode;
  db.settings.numberMode = book.numberMode;
  db.settings.numberFormat = book.numberFormat;
  db.settings.customStartSequences = { ...(book.customStartSequences || {}) };
  if (book.configVersion) db.settings.configVersion = book.configVersion;
  if (book.updatedAt) db.settings.updatedAt = book.updatedAt;
  if (Array.isArray(book.signers)) {
    db.signers = book.signers.map((s) => ({ ...s }));
  }
}

function findBookExplicitly(bookId?: string): RegistryBook | null {
  if (!bookId || !bookId.trim()) return getDefaultBook();
  return db.books.find((b) => b.id === bookId.trim()) || null;
}

function computeSeriesSequenceState(
  book: RegistryBook,
  targetYear: number,
  rawTypeCode: string
): StartSequenceSeriesConfig & { sampleTypeCode: string } {
  const mode = book.numberMode;
  const canonicalInput = resolveCanonicalTypeCode(rawTypeCode, db.documentTypes);
  const canonicalType = mode === 'shared' ? 'SHARED' : canonicalInput;
  const seriesKey = `${book.code}:${targetYear}:${canonicalType}`;
  const storageKey = `${targetYear}_${canonicalType}`;

  const docsInSeries = db.documents.filter((d) => {
    if (getDocBookId(d) !== book.id || d.year !== targetYear || d.docCategory === 'incoming') {
      return false;
    }
    if (mode === 'shared') return true;
    return resolveCanonicalTypeCode(d.typeCode, db.documentTypes) === canonicalType;
  });

  const maxIssuedSequence = docsInSeries.reduce((m, d) => Math.max(m, d.sequenceNumber), 0);
  const rawSaved = book.customStartSequences?.[storageKey];
  const isConfigured = typeof rawSaved === 'number' && Number.isFinite(rawSaved) && rawSaved >= 1;
  const customStartSequence = isConfigured ? Math.floor(rawSaved) : null;
  const minAllowedStartSequence = maxIssuedSequence + 1;
  const nextExpectedSeq = Math.max(customStartSequence ?? 1, minAllowedStartSequence);

  const sampleTypeCode =
    canonicalType === 'SHARED'
      ? book.enabledTypeCodes?.[0] || db.documentTypes[0]?.code || 'KH'
      : canonicalType;
  const seqStr = formatSequenceNumber(nextExpectedSeq);
  const previewNumber = renderFullDocNumber(book.numberFormat, seqStr, sampleTypeCode, book.unitCode);

  return {
    seriesKey,
    storageKey,
    bookId: book.id,
    bookCode: book.code,
    year: targetYear,
    typeCode: canonicalType,
    isConfigured,
    customStartSequence,
    maxIssuedSequence,
    minAllowedStartSequence,
    nextExpectedSeq,
    previewNumber,
    sampleTypeCode,
    configVersion: book.configVersion || 1,
    customStartSequences: { ...(book.customStartSequences || {}) },
  };
}

function bumpBookVersion(book: RegistryBook) {
  const nowIso = new Date().toISOString();
  book.configVersion = (book.configVersion || 1) + 1;
  book.updatedAt = nowIso;
  if (book.id === DEFAULT_BOOK_ID) {
    db.settings.configVersion = book.configVersion;
    db.settings.updatedAt = nowIso;
  } else {
    db.settings.updatedAt = nowIso;
  }
}

function checkVersionConflict(req: Request, res: Response, targetBook: RegistryBook, currentUser: User): boolean {
  const { expectedVersion, forceOverwrite } = req.body || {};
  if (forceOverwrite) return false;
  if (expectedVersion === undefined || expectedVersion === null || expectedVersion === '') return false;
  const parsedExpected = Number(expectedVersion);
  if (!Number.isFinite(parsedExpected)) return false;
  const serverVersion = targetBook.configVersion || 1;
  if (parsedExpected !== serverVersion) {
    res.status(409).json({
      code: 'SETTINGS_CONFLICT',
      error: 'Cài đặt đã thay đổi trên thiết bị khác. Vui lòng làm mới dữ liệu và đối chiếu trước khi lưu.',
      serverVersion,
      serverBook: sanitizeBookForUser(targetBook, currentUser),
    });
    return true;
  }
  return false;
}

function toSafeUser(u: User & { passwordHash?: string }): User {
  const { passwordHash: _, ...rest } = u;
  return {
    ...rest,
    bookPermissions:
      rest.role === 'admin' ? undefined : sanitizeBookPermissionsMap(rest.bookPermissions || {}),
  };
}

function refreshUserSessions(updatedUser: User) {
  for (const [token, sess] of tokenStore.entries()) {
    if (sess.user.id === updatedUser.id) {
      if (!updatedUser.isActive) {
        tokenStore.delete(token);
      } else {
        sess.user = updatedUser;
      }
    }
  }
}

// --- Staged PDF Store for Preview & Retry (30-min TTL) ---
interface StagedPdfEntry {
  token: string;
  bookId: string;
  buffer: Buffer;
  sha256: string;
  originalFileName: string;
  sourceType: 'pdf_url' | 'pdf_upload';
  sourceUrl?: string;
  userId: string;
  createdAt: number;
}
const stagedPdfs = new Map<string, StagedPdfEntry>();

if (!IS_SERVERLESS) {
  setInterval(() => {
    const cutoff = Date.now() - 30 * 60 * 1000;
    for (const [k, v] of stagedPdfs.entries()) {
      if (v.createdAt < cutoff) stagedPdfs.delete(k);
    }
  }, 5 * 60 * 1000).unref();
}

function findConflictingDocumentByDriveFileName(
  targetFileName: string,
  excludeDocId?: string,
  bookId?: string
): DocumentRecord | null {
  const lower = (targetFileName || '').trim().toLowerCase();
  if (!lower) return null;
  return (
    db.documents.find(
      (d) =>
        d.id !== excludeDocId &&
        (!bookId || getDocBookId(d) === bookId) &&
        d.status !== 'cancelled' &&
        d.status !== 'deleted' &&
        d.driveFileName &&
        d.driveFileName.trim().toLowerCase() === lower
    ) || null
  );
}

function resolveUniqueDriveFileName(baseFileName: string, excludeDocId?: string, bookId?: string): string {
  let candidate = baseFileName;
  let idx = 2;
  while (findConflictingDocumentByDriveFileName(candidate, excludeDocId, bookId)) {
    candidate = appendConflictSuffix(baseFileName, idx);
    idx++;
  }
  return candidate;
}

function findSuspectedDuplicateDocument(
  bookId: string,
  fullDocNumber: string,
  issuingAgency?: string,
  year?: number,
  excludeDocId?: string
): DocumentRecord | null {
  const normNum = normalizeDocNumberForCompare(fullDocNumber);
  if (!normNum) return null;
  const normAgency = normalizeCompareText(issuingAgency || '');

  return (
    db.documents.find((d) => {
      if (d.id === excludeDocId) return false;
      if (d.status === 'cancelled' || d.status === 'deleted') return false;
      if (getDocBookId(d) !== bookId) return false;

      if (normalizeDocNumberForCompare(d.fullDocNumber) !== normNum) return false;
      if (year && d.year !== year) return false;

      const dAgency = normalizeCompareText(d.issuingAgency || d.signerRole || '');
      if (normAgency && dAgency) {
        if (normAgency === dAgency) return true;
        if (
          normAgency.length >= 8 &&
          dAgency.length >= 8 &&
          (normAgency.includes(dAgency) || dAgency.includes(normAgency))
        ) {
          return true;
        }
        return false;
      }
      return true;
    }) || null
  );
}

// --- Server-Sent Events (SSE) Multi-Device Broadcast ---
type SSEClient = { id: string; res: Response };
const sseClients: SSEClient[] = [];

function broadcastSSE(event: string, payload: any) {
  const message = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (let i = sseClients.length - 1; i >= 0; i--) {
    const client = sseClients[i];
    try {
      client.res.write(message);
    } catch (err) {
      sseClients.splice(i, 1);
    }
  }
}

// Heartbeat for SSE to keep connections open (12s for mobile networks)
if (!IS_SERVERLESS) {
  setInterval(() => {
    for (let i = sseClients.length - 1; i >= 0; i--) {
      try {
        sseClients[i].res.write(': heartbeat\n\n');
      } catch {
        sseClients.splice(i, 1);
      }
    }
  }, 12000).unref();
}

// --- Timezone Asia/Ho_Chi_Minh Helpers ---
function getVietnamTimeInfo(d = new Date()) {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Ho_Chi_Minh',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  });
  const parts = formatter.format(d).split('-'); // YYYY-MM-DD
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  const formattedDate = `${parts[2]}/${parts[1]}/${parts[0]}`; // dd/mm/yyyy
  return { year, month, day, formattedDate };
}

function getVietnamDateKey(d = new Date()): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Ho_Chi_Minh',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(d);
}

function formatVietnamDateTime(d = new Date()): string {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Ho_Chi_Minh',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).formatToParts(d);
  const get = (type: string) => parts.find((p) => p.type === type)?.value || '00';
  return `${get('day')}/${get('month')}/${get('year')} ${get('hour')}:${get('minute')}:${get('second')}`;
}

function getRequestVietnamTimeInfo(req?: Request) {
  if (
    req &&
    process.env.ALLOW_TIME_SIMULATION === 'true' &&
    process.env.NODE_ENV !== 'production'
  ) {
    const rawHeader = req.headers['x-test-vietnam-now'];
    if (typeof rawHeader === 'string' && rawHeader.trim()) {
      const parsed = new Date(rawHeader.trim());
      if (!Number.isNaN(parsed.getTime())) {
        return getVietnamTimeInfo(parsed);
      }
    }
  }
  return getVietnamTimeInfo(new Date());
}

// --- Signed Session & HttpOnly Cookie Store (tương thích cả Node lẫn Vercel Serverless) ---
const SESSION_COOKIE_NAME = 'nct_session';
const STANDARD_SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 ngày (thời hạn thông thường)
const REMEMBER_SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000; // 30 ngày (Duy trì đăng nhập)
const REMEMBER_COOKIE_MAX_AGE_SEC = 30 * 24 * 60 * 60;

const tokenStore = new Map<string, { user: User; expiresAt: number; sid?: string }>();
const revokedSessionCache = new Map<string, number>();

function getSessionSecret(): string {
  const envSecret = (process.env.SESSION_SECRET || process.env.AUTH_SECRET || '').trim();
  if (envSecret) return envSecret;
  const seed = process.env.VERCEL_PROJECT_ID || process.env.APP_URL || 'nct-thpt-2026-session-key';
  return crypto.createHash('sha256').update(`nct-hmac-v1:${seed}`).digest('hex');
}

function signPayloadB64(payloadB64: string): string {
  return crypto.createHmac('sha256', getSessionSecret()).update(payloadB64).digest('base64url');
}

function isSessionRevoked(sid: string): boolean {
  const now = Date.now();
  const cachedExp = revokedSessionCache.get(sid);
  if (cachedExp && cachedExp > now) return true;
  const persistedExp = db.revokedSessions?.[sid];
  if (persistedExp && persistedExp > now) {
    revokedSessionCache.set(sid, persistedExp);
    return true;
  }
  return false;
}

function revokeSessionById(sid: string, expiresAt: number) {
  const exp = Math.max(expiresAt, Date.now() + 60_000);
  revokedSessionCache.set(sid, exp);
  if (!db.revokedSessions) db.revokedSessions = {};
  db.revokedSessions[sid] = exp;
  const now = Date.now();
  for (const [k, v] of Object.entries(db.revokedSessions)) {
    if (v <= now) delete db.revokedSessions[k];
  }
  for (const [k, v] of revokedSessionCache.entries()) {
    if (v <= now) revokedSessionCache.delete(k);
  }
}

interface SignedSessionPayload {
  sid: string;
  uid: string;
  iat: number;
  exp: number;
  rem: boolean;
}

function generateToken(user: User, rememberMe = false): { token: string; expiresAt: number; sid: string } {
  const sid = crypto.randomBytes(16).toString('hex');
  const iat = Date.now();
  const ttl = rememberMe ? REMEMBER_SESSION_TTL_MS : STANDARD_SESSION_TTL_MS;
  const exp = iat + ttl;
  const payload: SignedSessionPayload = {
    sid,
    uid: user.id,
    iat,
    exp,
    rem: Boolean(rememberMe),
  };
  const payloadB64 = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
  const sigB64 = signPayloadB64(payloadB64);
  const token = `${payloadB64}.${sigB64}`;

  tokenStore.set(token, {
    user,
    expiresAt: exp,
    sid,
  });
  return { token, expiresAt: exp, sid };
}

function parseSignedToken(token: string): SignedSessionPayload | null {
  const dotIdx = token.indexOf('.');
  if (dotIdx <= 0) return null;
  const payloadB64 = token.slice(0, dotIdx);
  const sigB64 = token.slice(dotIdx + 1);
  if (!payloadB64 || !sigB64) return null;

  const expectedSig = signPayloadB64(payloadB64);
  const bufA = Buffer.from(sigB64);
  const bufB = Buffer.from(expectedSig);
  if (bufA.length !== bufB.length || !crypto.timingSafeEqual(bufA, bufB)) {
    return null;
  }

  try {
    const parsed = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')) as SignedSessionPayload;
    if (!parsed || !parsed.sid || !parsed.uid || typeof parsed.exp !== 'number') return null;
    return parsed;
  } catch {
    return null;
  }
}

function parseCookies(req: Request): Record<string, string> {
  const header = req.headers.cookie;
  const out: Record<string, string> = {};
  if (!header) return out;
  for (const part of header.split(';')) {
    const eqIdx = part.indexOf('=');
    if (eqIdx <= 0) continue;
    const key = part.slice(0, eqIdx).trim();
    const val = part.slice(eqIdx + 1).trim();
    if (!key) continue;
    try {
      out[key] = decodeURIComponent(val);
    } catch {
      out[key] = val;
    }
  }
  return out;
}

function extractTokenFromRequest(req: Request): string | undefined {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    const bearer = authHeader.slice(7).trim();
    if (bearer) return bearer;
  }
  const cookies = parseCookies(req);
  if (cookies[SESSION_COOKIE_NAME]) {
    return cookies[SESSION_COOKIE_NAME];
  }
  if (req.query.token && typeof req.query.token === 'string' && req.query.token.trim()) {
    return req.query.token.trim();
  }
  return undefined;
}

function buildSessionCookieHeader(
  req: Request,
  value: string,
  opts: { maxAgeSeconds?: number; clear?: boolean } = {}
): string {
  const protoHeader = String(req.headers['x-forwarded-proto'] || '').split(',')[0].trim().toLowerCase();
  const isHttps = req.secure || protoHeader === 'https' || Boolean(process.env.VERCEL);

  const originHeader = String(req.headers.origin || '');
  const reqHost = String(req.headers['x-forwarded-host'] || req.headers.host || '').split(',')[0].trim().toLowerCase();
  let isCrossOrigin = false;
  if (originHeader && reqHost) {
    try {
      const originHost = new URL(originHeader).host.toLowerCase();
      if (originHost && originHost !== reqHost) {
        isCrossOrigin = true;
      }
    } catch {
      // ignore
    }
  }
  if (req.headers['sec-fetch-site'] === 'cross-site') {
    isCrossOrigin = true;
  }

  const sameSite = isHttps && isCrossOrigin ? 'None' : 'Lax';
  const parts = [
    `${SESSION_COOKIE_NAME}=${opts.clear ? '' : encodeURIComponent(value)}`,
    'Path=/',
    'HttpOnly',
    `SameSite=${sameSite}`,
  ];
  if (isHttps) {
    parts.push('Secure');
  }
  if (opts.clear) {
    parts.push('Max-Age=0');
    parts.push('Expires=Thu, 01 Jan 1970 00:00:00 GMT');
  } else if (typeof opts.maxAgeSeconds === 'number' && opts.maxAgeSeconds > 0) {
    parts.push(`Max-Age=${opts.maxAgeSeconds}`);
    parts.push(`Expires=${new Date(Date.now() + opts.maxAgeSeconds * 1000).toUTCString()}`);
  }
  return parts.join('; ');
}

function revokeTokenSession(token: string) {
  const parsed = parseSignedToken(token);
  if (parsed) {
    revokeSessionById(parsed.sid, parsed.exp);
  }
  const cached = tokenStore.get(token);
  if (cached?.sid) {
    revokeSessionById(cached.sid, cached.expiresAt);
  }
  tokenStore.delete(token);
}

// Auth Middleware (luôn nạp quyền mới nhất từ db.users để áp dụng thu hồi/cấp quyền tức thì)
function requireAuth(req: Request, res: Response, next: NextFunction) {
  const token = extractTokenFromRequest(req);

  if (!token) {
    return res.status(401).json({ error: 'Yêu cầu đăng nhập để truy cập' });
  }

  let userId: string | null = null;
  const parsed = parseSignedToken(token);
  if (parsed) {
    if (parsed.exp < Date.now()) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại' });
    }
    if (isSessionRevoked(parsed.sid)) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã kết thúc, vui lòng đăng nhập lại' });
    }
    userId = parsed.uid;
  } else {
    const session = tokenStore.get(token);
    if (!session || session.expiresAt < Date.now()) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại' });
    }
    if (session.sid && isSessionRevoked(session.sid)) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã kết thúc, vui lòng đăng nhập lại' });
    }
    userId = session.user.id;
  }

  const freshUser = db.users.find((u) => u.id === userId);
  if (!freshUser || !freshUser.isActive) {
    tokenStore.delete(token);
    return res.status(401).json({ error: 'Tài khoản đã bị khóa hoặc không còn hiệu lực' });
  }

  const safeUser = toSafeUser(freshUser);
  const cached = tokenStore.get(token);
  if (cached) cached.user = safeUser;
  (req as any).user = safeUser;
  (req as any).authToken = token;
  next();
}

function requireAdmin(req: Request, res: Response, next: NextFunction) {
  const user = (req as any).user as User;
  if (!user || user.role !== 'admin') {
    return res.status(403).json({ error: 'Chỉ Quản trị viên mới có quyền thực hiện thao tác này' });
  }
  next();
}

// --- Number Formatting Logic ---
function formatSequenceNumber(seq: number): string {
  if (seq < 10) return `0${seq}`;
  return `${seq}`;
}

function renderFullDocNumber(format: string, seqStr: string, typeCode: string, unitCode: string): string {
  let result = format;
  result = result.replace(/\{so\}/g, seqStr);
  result = result.replace(/\{loai\}/g, typeCode);
  result = result.replace(/\{don_vi\}/g, unitCode);
  return result;
}

// --- API ROUTES ---

// 1. Health check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    app: 'Sổ Cấp Số Văn Bản Đoàn Trường THPT Nguyễn Chí Thanh',
    timeVietnam: getRequestVietnamTimeInfo(req),
  });
});

// 2. Authentication
app.post('/api/auth/login', (req, res) => {
  const { username, password, rememberMe } = req.body || {};
  if (!username || !password) {
    return res.status(400).json({ error: 'Vui lòng nhập tên đăng nhập và mật khẩu' });
  }

  const hash = hashPassword(String(password));
  const foundUser = db.users.find(
    (u) => u.username.toLowerCase() === String(username).trim().toLowerCase() && u.passwordHash === hash
  );

  if (!foundUser) {
    return res.status(401).json({ error: 'Sai tên đăng nhập hoặc mật khẩu' });
  }

  if (!foundUser.isActive) {
    return res.status(403).json({ error: 'Tài khoản đã bị tạm khóa, vui lòng liên hệ Quản trị viên' });
  }

  const remember = Boolean(rememberMe);
  const userSafe = toSafeUser(foundUser);
  const { token, expiresAt } = generateToken(userSafe, remember);

  res.setHeader(
    'Set-Cookie',
    buildSessionCookieHeader(req, token, {
      maxAgeSeconds: remember ? REMEMBER_COOKIE_MAX_AGE_SEC : undefined,
    })
  );

  res.json({ user: userSafe, token, rememberMe: remember, expiresAt });
});

app.get('/api/auth/me', requireAuth, (req, res) => {
  const sessionUser = (req as any).user as User;
  const freshUser = db.users.find((u) => u.id === sessionUser.id);
  if (!freshUser || !freshUser.isActive) {
    return res.status(401).json({ error: 'Tài khoản không còn hiệu lực' });
  }
  const userSafe = toSafeUser(freshUser);
  refreshUserSessions(userSafe);
  res.json({ user: userSafe });
});

app.post('/api/auth/logout', (req, res) => {
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7).trim()
    : undefined;
  const cookieToken = parseCookies(req)[SESSION_COOKIE_NAME];

  let changed = false;
  if (bearer) {
    revokeTokenSession(bearer);
    changed = true;
  }
  if (cookieToken && cookieToken !== bearer) {
    revokeTokenSession(cookieToken);
    changed = true;
  }
  if (changed) {
    try {
      saveDatabaseSync(db);
    } catch (err) {
      console.error('Lỗi lưu trạng thái hủy phiên:', err);
    }
  }

  res.setHeader('Set-Cookie', buildSessionCookieHeader(req, '', { clear: true }));
  res.json({ success: true });
});

// 3. User Management (Admin only)
app.get('/api/users', requireAuth, requireAdmin, (req, res) => {
  const safeUsers = db.users.map((u) => toSafeUser(u));
  res.json({ users: safeUsers });
});

app.post('/api/users', requireAuth, requireAdmin, (req, res) => {
  const actor = (req as any).user as User;
  const { username, fullName, password, bookPermissions } = req.body;
  if (!username || !fullName || !password) {
    return res.status(400).json({ error: 'Vui lòng điền đầy đủ: Tên đăng nhập, Họ tên và Mật khẩu' });
  }

  const cleanUser = username.trim().toLowerCase();
  if (db.users.some((u) => u.username.toLowerCase() === cleanUser)) {
    return res.status(400).json({ error: 'Tên đăng nhập đã tồn tại trong hệ thống' });
  }

  const cleanedPermissions =
    bookPermissions && typeof bookPermissions === 'object'
      ? sanitizeBookPermissionsMap(bookPermissions)
      : { [DEFAULT_BOOK_ID]: ['view', 'request_number'] as BookPermission[] };

  // Tài khoản do Admin tạo cố định vai trò non-admin ('issuer'), quyền chi tiết phân theo từng sổ văn bản
  const newUser: User & { passwordHash: string } = {
    id: `user-${Date.now()}`,
    username: cleanUser,
    fullName: fullName.trim(),
    role: 'issuer',
    isActive: true,
    createdAt: new Date().toISOString(),
    bookPermissions: cleanedPermissions,
    passwordHash: hashPassword(password),
  };

  db.users.push(newUser);
  appendSystemAuditLog({
    actorUserId: actor.id,
    actorName: `${actor.fullName} (${actor.username})`,
    action: 'user_create',
    targetUserId: newUser.id,
    targetUsername: newUser.username,
    details: `Tạo tài khoản ${newUser.fullName} (@${newUser.username}) và phân quyền cho ${Object.keys(cleanedPermissions).length} sổ`,
    changes: {
      bookPermissions: { old: null, new: JSON.stringify(cleanedPermissions) },
    },
  });
  saveDatabaseSync(db);

  res.json({ user: toSafeUser(newUser) });
});

app.put('/api/users/:id', requireAuth, requireAdmin, (req, res) => {
  const actor = (req as any).user as User;
  const { id } = req.params;
  const { fullName, isActive, password, bookPermissions } = req.body;

  const target = db.users.find((u) => u.id === id);
  if (!target) return res.status(404).json({ error: 'Không tìm thấy người dùng' });

  const prevPermissions = target.bookPermissions ? { ...target.bookPermissions } : {};
  const prevActive = target.isActive;
  const changes: Record<string, { old: string | null; new: string | null }> = {};

  if (target.username === 'admin') {
    target.role = 'admin';
    if (isActive === false) {
      return res.status(400).json({ error: 'Không thể vô hiệu hóa tài khoản quản trị viên cố định' });
    }
  } else {
    target.role = 'issuer';
    if (bookPermissions && typeof bookPermissions === 'object') {
      const nextPerms = sanitizeBookPermissionsMap(bookPermissions);
      if (JSON.stringify(prevPermissions) !== JSON.stringify(nextPerms)) {
        changes.bookPermissions = {
          old: JSON.stringify(prevPermissions),
          new: JSON.stringify(nextPerms),
        };
      }
      target.bookPermissions = nextPerms;
    }
  }

  if (fullName && fullName.trim() && fullName.trim() !== target.fullName) {
    changes.fullName = { old: target.fullName, new: fullName.trim() };
    target.fullName = fullName.trim();
  }
  if (typeof isActive === 'boolean' && target.username !== 'admin' && isActive !== prevActive) {
    changes.isActive = { old: String(prevActive), new: String(isActive) };
    target.isActive = isActive;
  }
  if (password && password.trim().length >= 6) {
    target.passwordHash = hashPassword(password.trim());
    changes.password = { old: '******', new: 'Đã đặt lại mật khẩu' };
  }

  if (Object.keys(changes).length > 0) {
    appendSystemAuditLog({
      actorUserId: actor.id,
      actorName: `${actor.fullName} (${actor.username})`,
      action: changes.bookPermissions ? 'permission_change' : 'user_update',
      targetUserId: target.id,
      targetUsername: target.username,
      details: changes.bookPermissions
        ? `Cập nhật phân quyền sổ cho ${target.fullName} (@${target.username})`
        : `Cập nhật thông tin tài khoản ${target.fullName} (@${target.username})`,
      changes,
    });
  }

  saveDatabaseSync(db);
  const safeUser = toSafeUser(target);
  refreshUserSessions(safeUser);
  broadcastSSE('user_permissions_updated', { userId: target.id, user: safeUser });
  res.json({ user: safeUser });
});

app.delete('/api/users/:id', requireAuth, requireAdmin, (req, res) => {
  const actor = (req as any).user as User;
  const { id } = req.params;
  const target = db.users.find((u) => u.id === id);
  if (!target) return res.status(404).json({ error: 'Không tìm thấy người dùng' });

  if (target.username === 'admin') {
    return res.status(400).json({ error: 'Không thể xóa tài khoản Quản trị viên cố định' });
  }

  db.users = db.users.filter((u) => u.id !== id);
  for (const [token, sess] of tokenStore.entries()) {
    if (sess.user.id === id) tokenStore.delete(token);
  }
  appendSystemAuditLog({
    actorUserId: actor.id,
    actorName: `${actor.fullName} (${actor.username})`,
    action: 'user_delete',
    targetUserId: target.id,
    targetUsername: target.username,
    details: `Xóa tài khoản ${target.fullName} (@${target.username})`,
  });
  saveDatabaseSync(db);
  broadcastSSE('user_permissions_updated', { userId: id, deleted: true });
  res.json({ success: true, message: `Đã xóa tài khoản ${target.username}` });
});

app.get('/api/audit-logs', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const bookId = (req.query.bookId as string) || '';
  const canView =
    isSystemAdmin(currentUser) ||
    (bookId && hasBookPermission(currentUser, bookId, 'manage_book')) ||
    (!bookId && db.books.some((b) => hasBookPermission(currentUser, b.id, 'manage_book')));
  if (!canView) {
    return res.status(403).json({ error: 'Bạn không có quyền xem nhật ký quản trị hệ thống.' });
  }

  let logs = db.systemAuditLogs || [];
  if (bookId && bookId !== 'all') {
    logs = logs.filter((l) => l.bookId === bookId);
  } else if (!isSystemAdmin(currentUser)) {
    const managedIds = new Set(
      db.books.filter((b) => hasBookPermission(currentUser, b.id, 'manage_book')).map((b) => b.id)
    );
    logs = logs.filter((l) => l.bookId && managedIds.has(l.bookId));
  }
  res.json({ logs: logs.slice(0, 150) });
});

app.get('/api/document-history', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const {
    bookId = '',
    username = '',
    userId = '',
    fromDate = '',
    toDate = '',
    docId = '',
    action = '',
    search = '',
  } = req.query as Record<string, string>;

  if (bookId && bookId !== 'all') {
    const targetBook = getBookById(bookId);
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'view')) {
      return res.status(403).json({
        error: `Bạn chưa được cấp quyền "Xem" trong ${targetBook.name} để xem lịch sử thao tác.`,
      });
    }
  } else if (
    !isSystemAdmin(currentUser) &&
    !db.books.some((b) => hasBookPermission(currentUser, b.id, 'view'))
  ) {
    return res.status(403).json({
      error: 'Bạn chưa được phân quyền xem lịch sử thao tác trên bất kỳ sổ nào.',
    });
  }

  let logs = (db.documentHistoryLogs || []).filter((entry) => {
    const bId = entry.bookId || DEFAULT_BOOK_ID;
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, bId, 'view')) {
      return false;
    }
    if (bookId && bookId !== 'all' && bId !== bookId) return false;
    return true;
  });

  // Collect distinct actors within user's permitted book scope for filter dropdown
  const actorMap = new Map<string, { userId: string; username: string; userName: string }>();
  for (const l of logs) {
    const key = l.username || l.userId;
    if (key && !actorMap.has(key)) {
      actorMap.set(key, {
        userId: l.userId,
        username: l.username || l.userId,
        userName: l.userName,
      });
    }
  }

  if (docId) {
    logs = logs.filter((l) => l.docId === docId);
  }
  if (username && username !== 'all') {
    const lowerU = username.trim().toLowerCase();
    logs = logs.filter(
      (l) =>
        (l.username || '').toLowerCase() === lowerU ||
        l.userId === username
    );
  } else if (userId && userId !== 'all') {
    logs = logs.filter((l) => l.userId === userId);
  }
  if (action && action !== 'all') {
    logs = logs.filter((l) => l.action === action);
  }
  if (fromDate && /^\d{4}-\d{2}-\d{2}$/.test(fromDate)) {
    logs = logs.filter((l) => (l.dateKeyVietnam || getVietnamDateKey(new Date(l.timestamp))) >= fromDate);
  }
  if (toDate && /^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
    logs = logs.filter((l) => (l.dateKeyVietnam || getVietnamDateKey(new Date(l.timestamp))) <= toDate);
  }
  if (search && search.trim()) {
    const term = search.trim().toLowerCase();
    logs = logs.filter(
      (l) =>
        (l.fullDocNumber || '').toLowerCase().includes(term) ||
        (l.docName || '').toLowerCase().includes(term) ||
        (l.docId || '').toLowerCase().includes(term) ||
        (l.userName || '').toLowerCase().includes(term) ||
        (l.reason || '').toLowerCase().includes(term)
    );
  }

  res.json({
    total: logs.length,
    logs: logs.slice(0, 300),
    actors: Array.from(actorMap.values()),
  });
});

// 4. SSE Subscription for Real-Time Updates (Hỗ trợ đa thiết bị máy tính & điện thoại)
app.get('/api/events', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
  });
  if (typeof (res as any).flushHeaders === 'function') {
    (res as any).flushHeaders();
  }

  const clientId = `client-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  sseClients.push({ id: clientId, res });

  // Gửi thông báo kết nối thành công
  res.write(`event: connected\ndata: ${JSON.stringify({ clientId, timestamp: Date.now() })}\n\n`);

  req.on('close', () => {
    const idx = sseClients.findIndex((c) => c.id === clientId);
    if (idx !== -1) sseClients.splice(idx, 1);
  });
});

// 5. System Settings, Books, Document Types, Signers
app.get('/api/books', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const accessible = db.books.filter(
    (b) => isSystemAdmin(currentUser) || (b.isActive && getUserBookPermissions(currentUser, b.id).length > 0)
  );
  res.json({
    books: accessible.map((b) => sanitizeBookForUser(b, currentUser)),
    serverTimeVietnam: getRequestVietnamTimeInfo(req),
  });
});

app.post('/api/books', requireAuth, requireAdmin, (req, res) => {
  const currentUser = (req as any).user as User;
  const { code, name, unitName, unitCode, isShared, numberMode, numberFormat, enabledTypeCodes } = req.body || {};

  const cleanName = (name || '').trim();
  const cleanUnitName = (unitName || '').trim();
  const cleanUnitCode = (unitCode || '').trim().toUpperCase();
  const cleanCode = normalizeBookCode(typeof code === 'string' && code.trim() ? code.trim() : cleanName);

  if (!cleanName || !cleanUnitName || !cleanUnitCode || !cleanCode) {
    return res.status(400).json({
      error: 'Vui lòng nhập đầy đủ Mã sổ nội bộ, Tên sổ, Tên đơn vị/bộ phận và Ký hiệu viết tắt đơn vị.',
    });
  }

  if (db.books.some((b) => b.code.toUpperCase() === cleanCode.toUpperCase())) {
    return res.status(400).json({
      error: `Mã sổ nội bộ "${cleanCode}" đã tồn tại. Vui lòng chọn mã cố định khác.`,
    });
  }

  const validTypeCodes = Array.isArray(enabledTypeCodes)
    ? Array.from(
        new Set(
          enabledTypeCodes
            .filter((c: any) => typeof c === 'string' && c.trim())
            .map((c: string) => c.trim().toUpperCase())
        )
      )
    : undefined;

  const newBook: RegistryBook = {
    id: `book-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
    code: cleanCode,
    name: cleanName,
    unitName: cleanUnitName,
    unitCode: cleanUnitCode,
    isShared: Boolean(isShared),
    numberMode: numberMode === 'shared' ? 'shared' : 'separate',
    numberFormat: (numberFormat || '{so}-{loai}/{don_vi}').trim(),
    customStartSequences: {},
    enabledTypeCodes: validTypeCodes && validTypeCodes.length > 0 ? validTypeCodes : undefined,
    signers: db.signers.map((s) => ({ ...s })),
    isActive: true,
    createdAt: new Date().toISOString(),
  };

  db.books.push(newBook);
  appendSystemAuditLog({
    actorUserId: currentUser.id,
    actorName: `${currentUser.fullName} (${currentUser.username})`,
    action: 'book_create',
    bookId: newBook.id,
    bookCode: newBook.code,
    details: `Tạo sổ mới "${newBook.name}" (Mã cố định: ${newBook.code}, Ký hiệu: ${newBook.unitCode})`,
  });
  saveDatabaseSync(db);
  broadcastSSE('books_updated', { bookId: newBook.id });
  res.json({ book: sanitizeBookForUser(newBook, currentUser) });
});

app.put('/api/books/:id', requireAuth, (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const target = db.books.find((b) => b.id === id);
  if (!target) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản.' });
  }

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, target.id, 'manage_book')) {
    return res.status(403).json({
      error: 'Bạn cần quyền "Quản trị sổ" hoặc Quản trị viên hệ thống để cập nhật sổ này.',
    });
  }

  if (checkVersionConflict(req, res, target, currentUser)) return;

  const { name, unitName, unitCode, isShared, numberMode, numberFormat, signers, isActive, enabledTypeCodes } = req.body || {};
  const changes: Record<string, { old: string | null; new: string | null }> = {};

  if (typeof name === 'string' && name.trim() && name.trim() !== target.name) {
    changes.name = { old: target.name, new: name.trim() };
    target.name = name.trim();
  }
  if (typeof unitName === 'string' && unitName.trim() && unitName.trim() !== target.unitName) {
    changes.unitName = { old: target.unitName, new: unitName.trim() };
    target.unitName = unitName.trim();
  }
  if (typeof unitCode === 'string' && unitCode.trim() && unitCode.trim().toUpperCase() !== target.unitCode) {
    changes.unitCode = { old: target.unitCode, new: unitCode.trim().toUpperCase() };
    target.unitCode = unitCode.trim().toUpperCase();
  }
  if (typeof isShared === 'boolean' && isShared !== target.isShared) {
    changes.isShared = { old: String(target.isShared), new: String(isShared) };
    target.isShared = isShared;
  }
  if ((numberMode === 'separate' || numberMode === 'shared') && numberMode !== target.numberMode) {
    changes.numberMode = { old: target.numberMode, new: numberMode };
    target.numberMode = numberMode;
  }
  if (typeof numberFormat === 'string' && numberFormat.trim() && numberFormat.trim() !== target.numberFormat) {
    changes.numberFormat = { old: target.numberFormat, new: numberFormat.trim() };
    target.numberFormat = numberFormat.trim();
  }
  if (Array.isArray(enabledTypeCodes)) {
    const cleanedTypes = Array.from(
      new Set(
        enabledTypeCodes
          .filter((c: any) => typeof c === 'string' && c.trim())
          .map((c: string) => c.trim().toUpperCase())
      )
    );
    if (cleanedTypes.length === 0) {
      return res.status(400).json({ error: 'Mỗi sổ phải cho phép ít nhất 01 loại văn bản.' });
    }
    changes.enabledTypeCodes = {
      old: JSON.stringify(target.enabledTypeCodes || []),
      new: JSON.stringify(cleanedTypes),
    };
    target.enabledTypeCodes = cleanedTypes;
  }
  if (Array.isArray(signers)) {
    target.signers = signers
      .filter((s: any) => s && typeof s.name === 'string' && s.name.trim())
      .map((s: any, idx: number) => ({
        id: s.id || `sig-${Date.now()}-${idx}`,
        name: s.name.trim(),
        role: typeof s.role === 'string' ? s.role.trim() : '',
      }));
  }
  if (typeof isActive === 'boolean' && target.id !== DEFAULT_BOOK_ID) {
    changes.isActive = { old: String(target.isActive), new: String(isActive) };
    target.isActive = isActive;
  }

  if (Object.keys(changes).length > 0) {
    appendSystemAuditLog({
      actorUserId: currentUser.id,
      actorName: `${currentUser.fullName} (${currentUser.username})`,
      action: 'book_update',
      bookId: target.id,
      bookCode: target.code,
      details: `Cập nhật cấu hình sổ "${target.name}" (${target.code})`,
      changes,
    });
  }

  bumpBookVersion(target);
  syncDefaultBookToLegacySettings(target);
  saveDatabaseSync(db);
  broadcastSSE('books_updated', {
    bookId: target.id,
    configVersion: target.configVersion,
    updatedAt: target.updatedAt,
  });
  if (target.id === DEFAULT_BOOK_ID) {
    broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
  }
  res.json({ book: sanitizeBookForUser(target, currentUser) });
});

app.get('/api/settings', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  res.json({
    settings: sanitizeSettingsForUser(db.settings, currentUser),
    documentTypes: db.documentTypes,
    signers: db.signers,
    serverTimeVietnam: getRequestVietnamTimeInfo(req),
  });
});

app.put('/api/settings', requireAuth, requireAdmin, async (req, res) => {
  const currentUser = (req as any).user as User;
  const { unitName, unitCode, numberFormat, customStartSequences, bookId } = req.body;

  const targetBook = findBookExplicitly(bookId);
  if (!targetBook) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản cần cấu hình.' });
  }
  if (checkVersionConflict(req, res, targetBook, currentUser)) return;

  const releaseLock = await atomicLock.acquire();
  try {
    if (unitName) targetBook.unitName = unitName.trim();
    if (unitCode) targetBook.unitCode = unitCode.trim();
    if (numberFormat) targetBook.numberFormat = numberFormat.trim();
    if (customStartSequences && typeof customStartSequences === 'object') {
      const normalizedPatch = normalizeCustomStartSequencesMap(
        customStartSequences,
        targetBook.code,
        targetBook.id,
        db.documentTypes
      );
      targetBook.customStartSequences = {
        ...targetBook.customStartSequences,
        ...normalizedPatch,
      };
    }

    bumpBookVersion(targetBook);
    syncDefaultBookToLegacySettings(targetBook);
    saveDatabaseSync(db);
    broadcastSSE('books_updated', {
      bookId: targetBook.id,
      configVersion: targetBook.configVersion,
      updatedAt: targetBook.updatedAt,
    });
    broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
    res.json({
      settings: sanitizeSettingsForUser(db.settings, currentUser),
      book: sanitizeBookForUser(targetBook, currentUser),
    });
  } finally {
    releaseLock();
  }
});

// Chuyển đổi chế độ cấp số (riêng vs chung) theo từng sổ
app.post('/api/settings/number-mode', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const { newMode, confirmationText, bookId } = req.body;
  const targetBook = getBookById(bookId || DEFAULT_BOOK_ID);

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'manage_book')) {
    return res.status(403).json({
      error: 'Bạn cần quyền "Quản trị sổ" hoặc Quản trị viên để thay đổi chế độ dãy số.',
    });
  }

  if (checkVersionConflict(req, res, targetBook, currentUser)) return;

  if (newMode !== 'separate' && newMode !== 'shared') {
    return res.status(400).json({ error: 'Chế độ dãy số không hợp lệ' });
  }

  const { year: currentYear } = getRequestVietnamTimeInfo(req);
  const currentYearDocs = db.documents.filter(
    (d) => getDocBookId(d) === targetBook.id && d.year === currentYear && d.docCategory !== 'incoming'
  );

  if (currentYearDocs.length > 0) {
    if (confirmationText !== 'XÁC NHẬN CHUYỂN ĐỔI') {
      return res.status(400).json({
        error:
          `Sổ "${targetBook.name}" trong năm ${currentYear} đã phát sinh ${currentYearDocs.length} văn bản. ` +
          `Để tránh trùng số hoặc sai lệch thứ tự, bạn phải nhập chính xác dòng chữ "XÁC NHẬN CHUYỂN ĐỔI" để xác nhận.`,
      });
    }
  }

  targetBook.numberMode = newMode;
  bumpBookVersion(targetBook);
  syncDefaultBookToLegacySettings(targetBook);
  saveDatabaseSync(db);
  broadcastSSE('books_updated', {
    bookId: targetBook.id,
    configVersion: targetBook.configVersion,
    updatedAt: targetBook.updatedAt,
  });
  broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
  res.json({
    success: true,
    settings: sanitizeSettingsForUser(db.settings, currentUser),
    book: sanitizeBookForUser(targetBook, currentUser),
  });
});

// Quản lý loại văn bản
app.post('/api/document-types', requireAuth, (req, res) => {
  const { name, code, color } = req.body;
  if (!name || !code) {
    return res.status(400).json({ error: 'Vui lòng nhập tên loại văn bản và ký hiệu' });
  }

  const cleanCode = code.trim().toUpperCase();
  if (db.documentTypes.some((t) => t.code.toUpperCase() === cleanCode)) {
    return res.status(400).json({ error: `Ký hiệu viết tắt "${cleanCode}" đã tồn tại` });
  }

  const newType: DocumentTypeConfig = {
    id: cleanCode,
    name: name.trim(),
    code: cleanCode,
    color: color || DEFAULT_TYPE_COLORS[cleanCode] || 'blue',
    isActive: true,
  };

  db.documentTypes.push(newType);
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { documentTypes: db.documentTypes });
  res.json({ documentType: newType });
});

app.delete('/api/document-types/:id', requireAuth, requireAdmin, (req, res) => {
  const { id } = req.params;
  const idx = db.documentTypes.findIndex((t) => t.id === id || t.code === id);
  if (idx === -1) return res.status(404).json({ error: 'Không tìm thấy loại văn bản cần xóa' });

  db.documentTypes.splice(idx, 1);
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { documentTypes: db.documentTypes });
  res.json({ success: true });
});

app.put('/api/document-types/:id', requireAuth, requireAdmin, (req, res) => {
  const { id } = req.params;
  const { name, code, color, isActive } = req.body;

  const target = db.documentTypes.find((t) => t.id === id);
  if (!target) return res.status(404).json({ error: 'Không tìm thấy loại văn bản' });

  if (name) target.name = name.trim();
  if (code) {
    const cleanCode = code.trim().toUpperCase();
    if (cleanCode !== target.code && db.documentTypes.some((t) => t.code.toUpperCase() === cleanCode)) {
      return res.status(400).json({ error: `Ký hiệu "${cleanCode}" đã được sử dụng` });
    }
    target.code = cleanCode;
  }
  if (color) target.color = color;
  if (typeof isActive === 'boolean') target.isActive = isActive;

  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { documentTypes: db.documentTypes });
  res.json({ documentType: target });
});

// Quản lý người ký
app.post('/api/signers', requireAuth, requireAdmin, (req, res) => {
  const { name, role } = req.body;
  if (!name || !role) {
    return res.status(400).json({ error: 'Vui lòng nhập tên người ký và chức vụ' });
  }

  const newSigner: SignerConfig = {
    id: `sig-${Date.now()}`,
    name: name.trim(),
    role: role.trim(),
  };

  db.signers.push(newSigner);
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { signers: db.signers });
  res.json({ signer: newSigner });
});

app.put('/api/signers/:id', requireAuth, requireAdmin, (req, res) => {
  const { id } = req.params;
  const { name, role } = req.body;

  const target = db.signers.find((s) => s.id === id);
  if (!target) return res.status(404).json({ error: 'Không tìm thấy người ký' });

  if (name) target.name = name.trim();
  if (role) target.role = role.trim();

  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { signers: db.signers });
  res.json({ signer: target });
});

app.delete('/api/signers/:id', requireAuth, requireAdmin, (req, res) => {
  const { id } = req.params;
  const idx = db.signers.findIndex((s) => s.id === id);
  if (idx === -1) return res.status(404).json({ error: 'Không tìm thấy người ký' });

  db.signers.splice(idx, 1);
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { signers: db.signers });
  res.json({ success: true });
});

// Đọc cấu hình số bắt đầu theo khóa (mã sổ + năm áp dụng + mã loại văn bản)
app.get('/api/settings/start-sequence', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const bookIdParam = typeof req.query.bookId === 'string' ? req.query.bookId : undefined;
  const targetBook = findBookExplicitly(bookIdParam);
  if (!targetBook) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản yêu cầu.' });
  }

  const canRead =
    isSystemAdmin(currentUser) ||
    hasBookPermission(currentUser, targetBook.id, 'manage_book') ||
    hasBookPermission(currentUser, targetBook.id, 'issue_number') ||
    hasBookPermission(currentUser, targetBook.id, 'request_number') ||
    hasBookPermission(currentUser, targetBook.id, 'view');
  if (!canRead) {
    return res.status(403).json({ error: 'Bạn không có quyền xem cấu hình số bắt đầu của sổ này.' });
  }

  const rawTypeCode = String(req.query.typeCode || '').trim();
  const canonicalType =
    targetBook.numberMode === 'shared'
      ? 'SHARED'
      : resolveCanonicalTypeCode(rawTypeCode, db.documentTypes);
  if (!canonicalType) {
    return res.status(400).json({ error: 'Thiếu mã loại văn bản.' });
  }

  const isSharedKey = canonicalType === 'SHARED' && targetBook.numberMode === 'shared';
  if (!isSharedKey && !isTypeEnabledInBook(targetBook, canonicalType)) {
    return res.status(400).json({
      error: `Loại văn bản "${canonicalType}" chưa được bật sử dụng trong ${targetBook.name}.`,
    });
  }

  const vietnamNow = getRequestVietnamTimeInfo(req);
  const requestedYear = req.query.year ? parseInt(String(req.query.year), 10) : NaN;
  const targetYear =
    Number.isFinite(requestedYear) && requestedYear >= 2000 && requestedYear <= 2100
      ? requestedYear
      : vietnamNow.year;

  const { sampleTypeCode: _, ...seriesState } = computeSeriesSequenceState(
    targetBook,
    targetYear,
    canonicalType
  );
  res.json({
    ...seriesState,
    serverYear: vietnamNow.year,
  });
});

// Cấu hình số bắt đầu tiếp theo theo sổ (Admin hoặc Quản trị sổ)
app.post('/api/settings/start-sequence', requireAuth, async (req, res) => {
  const currentUser = (req as any).user as User;
  const { year, typeCode, startSequence, bookId } = req.body || {};
  const targetBook = findBookExplicitly(typeof bookId === 'string' ? bookId : undefined);
  if (!targetBook) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản cần cấu hình.' });
  }

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'manage_book')) {
    return res.status(403).json({
      error: 'Bạn cần quyền "Quản trị sổ" hoặc Quản trị viên để cấu hình số bắt đầu.',
    });
  }

  const targetYear = parseInt(String(year), 10);
  const nextSeq = parseInt(String(startSequence), 10);
  const canonicalType =
    targetBook.numberMode === 'shared'
      ? 'SHARED'
      : resolveCanonicalTypeCode(String(typeCode || ''), db.documentTypes);

  if (
    !Number.isFinite(targetYear) ||
    targetYear < 2000 ||
    targetYear > 2100 ||
    !Number.isFinite(nextSeq) ||
    nextSeq < 1 ||
    !canonicalType
  ) {
    return res.status(400).json({ error: 'Năm áp dụng, loại văn bản hoặc số bắt đầu không hợp lệ.' });
  }

  const isSharedKey = canonicalType === 'SHARED' && targetBook.numberMode === 'shared';
  if (!isSharedKey && !isTypeEnabledInBook(targetBook, canonicalType)) {
    return res.status(400).json({
      error: `Loại văn bản "${canonicalType}" chưa được bật sử dụng trong ${targetBook.name}.`,
    });
  }

  const releaseLock = await atomicLock.acquire();
  try {
    const beforeState = computeSeriesSequenceState(targetBook, targetYear, canonicalType);
    if (nextSeq <= beforeState.maxIssuedSequence) {
      return res.status(400).json({
        error: `Số bắt đầu (${nextSeq}) không được nhỏ hơn hoặc bằng số lớn nhất đã cấp/hủy trong dãy ${targetBook.name} - ${targetYear} - ${canonicalType} (${beforeState.maxIssuedSequence}). Mức tối thiểu phải từ ${beforeState.minAllowedStartSequence} trở lên.`,
      });
    }

    targetBook.customStartSequences = {
      ...(targetBook.customStartSequences || {}),
      [beforeState.storageKey]: nextSeq,
    };
    bumpBookVersion(targetBook);
    syncDefaultBookToLegacySettings(targetBook);
    saveDatabaseSync(db);

    const { sampleTypeCode: _, ...afterState } = computeSeriesSequenceState(
      targetBook,
      targetYear,
      canonicalType
    );

    broadcastSSE('books_updated', {
      bookId: targetBook.id,
      configVersion: targetBook.configVersion,
      updatedAt: targetBook.updatedAt,
    });
    broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });

    res.json({
      success: true,
      ...afterState,
      book: sanitizeBookForUser(targetBook, currentUser),
    });
  } finally {
    releaseLock();
  }
});

// 6. Tính toán số văn bản dự kiến theo sổ (Preview ONLY - Chưa lưu, chưa cấp)
app.get('/api/documents/preview', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const book = findBookExplicitly(typeof req.query.bookId === 'string' ? req.query.bookId : undefined);
  if (!book) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản.' });
  }

  const rawTypeCode = String(req.query.typeCode || '').trim();
  const typeCode =
    book.numberMode === 'shared' && rawTypeCode.toUpperCase() === 'SHARED'
      ? 'SHARED'
      : resolveCanonicalTypeCode(rawTypeCode, db.documentTypes);
  if (!typeCode) {
    return res.status(400).json({ error: 'Thiếu mã loại văn bản' });
  }

  const canPreview =
    isSystemAdmin(currentUser) ||
    hasBookPermission(currentUser, book.id, 'request_number') ||
    hasBookPermission(currentUser, book.id, 'issue_number') ||
    hasBookPermission(currentUser, book.id, 'manage_book') ||
    hasBookPermission(currentUser, book.id, 'view');
  if (!canPreview) {
    return res.status(403).json({ error: 'Bạn không có quyền truy cập sổ văn bản này.' });
  }

  const isSharedKey = typeCode === 'SHARED' && book.numberMode === 'shared';
  if (!isSharedKey && !isTypeEnabledInBook(book, typeCode)) {
    return res.status(400).json({
      error: `Loại văn bản "${typeCode}" chưa được bật sử dụng trong ${book.name}.`,
    });
  }

  const vietnamNow = getRequestVietnamTimeInfo(req);
  const requestedYear = req.query.year ? parseInt(String(req.query.year), 10) : NaN;
  const targetYear =
    Number.isFinite(requestedYear) && requestedYear >= 2000 && requestedYear <= 2100
      ? requestedYear
      : vietnamNow.year;

  const seriesState = computeSeriesSequenceState(book, targetYear, typeCode);

  const response: PreviewNumberResponse = {
    bookId: book.id,
    previewNumber: seriesState.previewNumber,
    sequenceNumber: seriesState.nextExpectedSeq,
    year: targetYear,
    serverYear: vietnamNow.year,
    formattedDate: vietnamNow.formattedDate,
    mode: book.numberMode,
    customStartSequence: seriesState.customStartSequence,
    maxIssuedSequence: seriesState.maxIssuedSequence,
  };
  res.json(response);
});

// 6b. Kiểm tra trùng tên tệp trong sổ trước khi lưu lên Google Drive
app.post('/api/documents/check-filename', requireAuth, (req, res) => {
  const { fileName, excludeDocId, bookId } = req.body || {};
  const targetBook = getBookById(bookId);
  const cleanName = typeof fileName === 'string' ? fileName.trim() : '';
  if (!cleanName) {
    return res.json({ conflict: false, conflictingFileName: '', existingDocument: null });
  }
  const existing = findConflictingDocumentByDriveFileName(cleanName, excludeDocId, targetBook.id);
  res.json({
    conflict: Boolean(existing),
    conflictingFileName: cleanName,
    suggestedSuffixName: existing
      ? resolveUniqueDriveFileName(cleanName, excludeDocId, targetBook.id)
      : cleanName,
    existingDocument: existing,
  });
});

// 6c. Kiểm tra nghi trùng văn bản trong cùng sổ (SHA-256 + Cơ quan ban hành + Số/ký hiệu + Năm)
app.post('/api/documents/check-duplicate', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const { bookId, fullDocNumber, issuingAgency, year, sha256, excludeDocId } = req.body || {};
  const targetBook = getBookById(bookId);

  if (
    !isSystemAdmin(currentUser) &&
    !hasBookPermission(currentUser, targetBook.id, 'import_incoming') &&
    !hasBookPermission(currentUser, targetBook.id, 'view')
  ) {
    return res.status(403).json({ error: 'Bạn không có quyền kiểm tra trùng trên sổ này.' });
  }

  const exactDuplicate =
    sha256 && typeof sha256 === 'string'
      ? db.documents.find(
          (d) =>
            d.id !== excludeDocId &&
            getDocBookId(d) === targetBook.id &&
            d.fileSha256 === sha256 &&
            d.status !== 'cancelled'
        ) || null
      : null;

  const parsedYear = year ? parseInt(String(year), 10) : undefined;
  const suspectedDuplicate = findSuspectedDuplicateDocument(
    targetBook.id,
    String(fullDocNumber || ''),
    typeof issuingAgency === 'string' ? issuingAgency : '',
    Number.isFinite(parsedYear) ? parsedYear : undefined,
    excludeDocId || exactDuplicate?.id
  );

  res.json({
    exactDuplicate,
    suspectedDuplicate,
  });
});

// 6d. YÊU CẦU XIN SỐ (Tách biệt hoàn toàn với Cấp số - Dành cho quyền "Xin số")
app.get('/api/number-requests', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const bookId = (req.query.bookId as string) || '';
  const statusFilter = (req.query.status as string) || '';

  let list = [...db.numberRequests];
  if (bookId && bookId !== 'all') {
    list = list.filter((r) => r.bookId === bookId);
  }
  if (statusFilter && statusFilter !== 'all') {
    list = list.filter((r) => r.status === statusFilter);
  }

  if (!isSystemAdmin(currentUser)) {
    list = list.filter(
      (r) =>
        r.requestedByUserId === currentUser.id ||
        hasBookPermission(currentUser, r.bookId, 'issue_number') ||
        hasBookPermission(currentUser, r.bookId, 'manage_book')
    );
  }

  res.json({ requests: list });
});

app.post('/api/number-requests', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const { bookId, typeCode, docName, title, signer, signerRole, notes } = req.body || {};
  const targetBook = getBookById(bookId);

  if (
    !isSystemAdmin(currentUser) &&
    !hasBookPermission(currentUser, targetBook.id, 'request_number') &&
    !hasBookPermission(currentUser, targetBook.id, 'issue_number')
  ) {
    return res.status(403).json({
      error: `Bạn chưa được phân quyền "Xin số" trong ${targetBook.name}.`,
    });
  }

  const validDocName = (docName || '').trim();
  if (!typeCode || !validDocName || !signer || !signerRole) {
    return res.status(400).json({
      error: 'Vui lòng điền đầy đủ: Loại văn bản, Tên văn bản, Người ký và Chức vụ để gửi yêu cầu xin số.',
    });
  }

  if (!isTypeEnabledInBook(targetBook, typeCode)) {
    return res.status(400).json({
      error: `Loại văn bản "${typeCode}" chưa được kích hoạt trong ${targetBook.name}.`,
    });
  }

  const docTypeConfig = db.documentTypes.find((t) => t.code.toUpperCase() === typeCode.toUpperCase());
  const newReq: NumberRequest = {
    id: `req-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
    bookId: targetBook.id,
    bookName: targetBook.name,
    typeCode: typeCode.trim().toUpperCase(),
    typeName: docTypeConfig ? docTypeConfig.name : typeCode.trim().toUpperCase(),
    docName: validDocName,
    title: (title || '').trim(),
    signer: signer.trim(),
    signerRole: signerRole.trim(),
    notes: typeof notes === 'string' ? notes.trim() : '',
    requestedBy: currentUser.fullName,
    requestedByUsername: currentUser.username,
    requestedByUserId: currentUser.id,
    requestedAt: new Date().toISOString(),
    status: 'pending',
  };

  db.numberRequests.unshift(newReq);
  saveDatabaseSync(db);
  broadcastSSE('number_requests_updated', { bookId: targetBook.id, request: newReq });
  res.json({ request: newReq, success: true });
});

app.post('/api/number-requests/:id/issue', requireAuth, async (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const reqItem = db.numberRequests.find((r) => r.id === id);
  if (!reqItem) {
    return res.status(404).json({ error: 'Không tìm thấy yêu cầu xin số.' });
  }

  const targetBook = getBookById(reqItem.bookId);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'issue_number')) {
    return res.status(403).json({
      error: `Chỉ người có quyền "Cấp số" trong ${targetBook.name} mới được duyệt và cấp số chính thức.`,
    });
  }

  if (reqItem.status !== 'pending') {
    return res.status(400).json({ error: 'Yêu cầu xin số này đã được xử lý trước đó.' });
  }

  const releaseLock = await atomicLock.acquire();
  try {
    const { year, formattedDate } = getRequestVietnamTimeInfo(req);
    const canonicalReqType = resolveCanonicalTypeCode(reqItem.typeCode, db.documentTypes);
    const seriesState = computeSeriesSequenceState(targetBook, year, canonicalReqType);
    const nextSeq = seriesState.nextExpectedSeq;

    const seqStr = formatSequenceNumber(nextSeq);
    const fullDocNumber = renderFullDocNumber(
      targetBook.numberFormat,
      seqStr,
      canonicalReqType,
      targetBook.unitCode
    );
    const docId = `doc-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
    const newDocument: DocumentRecord = {
      id: docId,
      bookId: targetBook.id,
      bookCode: targetBook.code,
      bookName: targetBook.name,
      requestId: reqItem.id,
      stt: db.documents.length + 1,
      year,
      issueDate: formattedDate,
      createdAt: new Date().toISOString(),
      docCategory: 'outgoing',
      typeCode: reqItem.typeCode,
      typeName: reqItem.typeName,
      sequenceNumber: nextSeq,
      fullDocNumber,
      docName: reqItem.docName,
      title: reqItem.title,
      signer: reqItem.signer,
      signerRole: reqItem.signerRole,
      issuedBy: reqItem.requestedBy,
      issuedByUsername: reqItem.requestedByUsername,
      issuedByUserId: reqItem.requestedByUserId,
      status: 'active',
      notes: reqItem.notes || '',
      auditLogs: [],
      syncStatus: db.settings.msGraphConfig.enabled ? 'pending' : 'not_configured',
      idempotencyKey: `req-issue-${reqItem.id}`,
      driveStatus: 'none',
    };

    recordDocumentHistoryEntry({
      doc: newDocument,
      book: targetBook,
      actor: {
        id: reqItem.requestedByUserId,
        username: reqItem.requestedByUsername,
        fullName: reqItem.requestedBy,
      },
      action: 'request_number',
      changes: {
        docName: { old: null, new: reqItem.docName },
      },
      reason: `Gửi yêu cầu xin số vào ${targetBook.name}`,
      timestampIso: reqItem.requestedAt,
    });
    recordDocumentHistoryEntry({
      doc: newDocument,
      book: targetBook,
      actor: currentUser,
      action: 'create',
      changes: {
        fullDocNumber: { old: null, new: fullDocNumber },
        docName: { old: null, new: reqItem.docName },
      },
      reason: `Duyệt yêu cầu xin số của ${reqItem.requestedBy} và cấp số chính thức`,
    });

    reqItem.status = 'issued';
    reqItem.issuedDocId = docId;
    reqItem.issuedDocNumber = fullDocNumber;
    reqItem.reviewedBy = currentUser.fullName;
    reqItem.reviewedAt = new Date().toISOString();

    db.documents.unshift(newDocument);
    db.idempotencyKeys[newDocument.idempotencyKey] = docId;
    saveDatabaseSync(db);

    broadcastSSE('document_issued', { document: newDocument });
    broadcastSSE('number_requests_updated', { bookId: targetBook.id, request: reqItem });
    await syncToGoogleSheets(newDocument, 'create');
    res.json({ document: newDocument, request: reqItem, success: true });
  } finally {
    releaseLock();
  }
});

app.post('/api/number-requests/:id/reject', requireAuth, (req, res) => {
  const { id } = req.params;
  const { reason } = req.body || {};
  const currentUser = (req as any).user as User;
  const reqItem = db.numberRequests.find((r) => r.id === id);
  if (!reqItem) {
    return res.status(404).json({ error: 'Không tìm thấy yêu cầu xin số.' });
  }

  const targetBook = getBookById(reqItem.bookId);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'issue_number')) {
    return res.status(403).json({
      error: 'Bạn cần quyền "Cấp số" để từ chối yêu cầu xin số này.',
    });
  }

  if (reqItem.status !== 'pending') {
    return res.status(400).json({ error: 'Yêu cầu này đã được xử lý trước đó.' });
  }

  reqItem.status = 'rejected';
  reqItem.rejectReason = typeof reason === 'string' && reason.trim() ? reason.trim() : 'Không được duyệt';
  reqItem.reviewedBy = currentUser.fullName;
  reqItem.reviewedAt = new Date().toISOString();
  saveDatabaseSync(db);
  broadcastSSE('number_requests_updated', { bookId: targetBook.id, request: reqItem });
  res.json({ request: reqItem, success: true });
});

// 7. CẤP SỐ CHÍNH THỨC VÀ LƯU (Atomic Transaction with Mutex & Idempotency Key - Yêu cầu quyền "Cấp số")
app.post('/api/documents/issue', requireAuth, async (req, res) => {
  const {
    bookId,
    typeCode,
    docName,
    title,
    signer,
    signerRole,
    notes,
    idempotencyKey,
    hasAttachment,
    attachmentName,
    attachmentSize,
    conflictAction,
    customFileName,
    expectedYear,
  } = req.body;
  const currentUser = (req as any).user as User;
  const targetBook = getBookById(bookId);

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'issue_number')) {
    return res.status(403).json({
      error: `Tài khoản của bạn không có quyền "Cấp số" trong ${targetBook.name}. Vui lòng dùng chức năng "Xin số" hoặc liên hệ Quản trị viên.`,
    });
  }

  const validDocName = (docName || '').trim();
  if (!typeCode || !validDocName || !signer || !signerRole) {
    return res.status(400).json({ error: 'Vui lòng điền đầy đủ: Loại văn bản, Tên văn bản (bắt buộc), Người ký và Chức vụ' });
  }

  const preCanonicalType = resolveCanonicalTypeCode(String(typeCode || ''), db.documentTypes);
  if (!isTypeEnabledInBook(targetBook, preCanonicalType)) {
    return res.status(400).json({
      error: `Loại văn bản "${preCanonicalType}" chưa được bật sử dụng trong ${targetBook.name}.`,
    });
  }

  if (!idempotencyKey) {
    return res.status(400).json({ error: 'Thiếu mã chống lặp (idempotency key)' });
  }

  // Nếu có tệp đính kèm, kiểm tra tệp và cấu hình thư mục Drive của sổ trước khi cấp số
  let attachmentExt: '.doc' | '.docx' | null = null;
  if (hasAttachment) {
    attachmentExt = getWordFileExtension(String(attachmentName || ''));
    if (!attachmentExt) {
      return res.status(400).json({
        error: 'Định dạng tệp đính kèm không hợp lệ. Chỉ chấp nhận tệp Word (.doc hoặc .docx).',
      });
    }
    const sizeNum = Number(attachmentSize);
    if (!Number.isFinite(sizeNum) || sizeNum <= 0) {
      return res.status(400).json({
        error: 'Tệp Word đính kèm không hợp lệ hoặc bị rỗng.',
      });
    }
    if (sizeNum > getMaxWordFileSizeBytes()) {
      return res.status(400).json({
        error: `Dung lượng tệp vượt quá giới hạn tối đa ${getMaxWordFileSizeMB()} MB.`,
      });
    }
    const driveCfg = resolveDriveForBook(targetBook);
    if (!driveCfg.isConfigured) {
      return res.status(400).json({
        error: `${driveCfg.missingReason || 'Chưa cấu hình thư mục Google Drive hợp lệ cho sổ này.'} Vui lòng bỏ chọn tệp để tiếp tục cấp số hoặc liên hệ Quản trị viên.`,
      });
    }
  }

  // Khóa giao dịch nguyên tử tại máy chủ - ngăn chặn 100% việc trùng số khi nhiều người cùng bấm
  const releaseLock = await atomicLock.acquire();

  try {
    // 1. Kiểm tra mã chống lặp
    if (db.idempotencyKeys[idempotencyKey]) {
      const existingDocId = db.idempotencyKeys[idempotencyKey];
      const existingDoc = db.documents.find((d) => d.id === existingDocId);
      if (existingDoc) {
        return res.json({
          document: existingDoc,
          duplicatePrevented: true,
          message: 'Yêu cầu đã được xử lý trước đó (đã chống gửi lặp)',
        });
      }
    }

    // 2. Tính toán năm và ngày theo giờ Việt Nam (Asia/Ho_Chi_Minh)
    const { year, formattedDate } = getRequestVietnamTimeInfo(req);

    // 3. Kiểm tra loại văn bản hợp lệ theo mã ổn định từ danh mục
    const canonicalTypeCode = resolveCanonicalTypeCode(typeCode, db.documentTypes);
    const docTypeConfig = db.documentTypes.find(
      (t) => t.code.trim().toUpperCase() === canonicalTypeCode
    );
    const typeName = docTypeConfig ? docTypeConfig.name : canonicalTypeCode;

    // 4. Cấp số thứ tự nguyên tử trong phạm vi sổ mục tiêu (dùng chung logic với Cài đặt & Preview)
    const seriesState = computeSeriesSequenceState(targetBook, year, canonicalTypeCode);
    const nextSeq = seriesState.nextExpectedSeq;

    const seqStr = formatSequenceNumber(nextSeq);
    const fullDocNumber = renderFullDocNumber(
      targetBook.numberFormat,
      seqStr,
      canonicalTypeCode,
      targetBook.unitCode
    );

    // Nếu người dùng mở biểu mẫu từ năm cũ và đồng hồ vừa chuyển sang năm mới (00:00 01/01 Asia/Ho_Chi_Minh),
    // yêu cầu người dùng xác nhận lại trước khi ghi nhận vào dãy số năm mới (vẫn giữ nguyên dữ liệu đã nhập).
    const parsedExpectedYear =
      expectedYear !== undefined && expectedYear !== null && expectedYear !== ''
        ? parseInt(String(expectedYear), 10)
        : NaN;
    if (Number.isFinite(parsedExpectedYear) && parsedExpectedYear !== year) {
      return res.status(409).json({
        code: 'YEAR_CHANGED',
        previousYear: parsedExpectedYear,
        currentYear: year,
        formattedDate,
        previewNumber: fullDocNumber,
        sequenceNumber: nextSeq,
        error: `Hệ thống đã bước sang năm mới ${year} (giờ Việt Nam). Số dự kiến mới là ${fullDocNumber}. Vui lòng xác nhận lại trước khi cấp số.`,
      });
    }

    // 4b. Nếu có tệp Word đính kèm, luôn tính lại tên tệp chuẩn theo SỐ CHÍNH THỨC vừa cấp
    let resolvedDriveFileName: string | undefined;
    if (hasAttachment && attachmentExt) {
      let targetFileName = buildStandardDriveFileName(fullDocNumber, validDocName, attachmentExt);
      if (typeof customFileName === 'string' && customFileName.trim()) {
        // Nếu người dùng không sửa phần tiền tố hoặc số không dịch chuyển do đồng thời, dùng custom; nếu số vừa nhảy vì cấp đồng thời thì ưu tiên số chính thức
        const cleanCustom = sanitizeCustomDriveFileName(customFileName, attachmentExt);
        const prefixOfficial = fullDocNumber.replace(/\//g, '.');
        if (cleanCustom.startsWith(prefixOfficial)) {
          targetFileName = cleanCustom;
        }
      }

      const conflictDoc = findConflictingDocumentByDriveFileName(targetFileName, undefined, targetBook.id);
      if (conflictDoc) {
        if (conflictAction === 'suffix') {
          targetFileName = resolveUniqueDriveFileName(targetFileName, undefined, targetBook.id);
        } else if (conflictAction !== 'overwrite') {
          return res.status(409).json({
            code: 'FILENAME_CONFLICT',
            error: `Tên tệp "${targetFileName}" đã trùng với hồ sơ ${conflictDoc.fullDocNumber} (${conflictDoc.docName || conflictDoc.title}). Vui lòng chọn cách xử lý trước khi cấp số.`,
            conflictingFileName: targetFileName,
            suggestedSuffixName: resolveUniqueDriveFileName(targetFileName, undefined, targetBook.id),
            existingDocument: conflictDoc,
            officialDocNumberPreview: fullDocNumber,
          });
        }
      }
      resolvedDriveFileName = targetFileName;
    }

    const docId = `doc-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
    const globalStt = db.documents.length + 1;
    const finalTitle = (title || '').trim();

    const newDocument: DocumentRecord = {
      id: docId,
      bookId: targetBook.id,
      bookCode: targetBook.code,
      bookName: targetBook.name,
      stt: globalStt,
      year,
      issueDate: formattedDate,
      createdAt: new Date().toISOString(),
      docCategory: 'outgoing',
      typeCode: canonicalTypeCode,
      typeName,
      sequenceNumber: nextSeq,
      fullDocNumber,
      docName: validDocName,
      title: finalTitle,
      signer: signer.trim(),
      signerRole: signerRole.trim(),
      issuedBy: currentUser.fullName,
      issuedByUsername: currentUser.username,
      issuedByUserId: currentUser.id,
      status: 'active',
      notes: notes ? notes.trim() : '',
      auditLogs: [],
      syncStatus: db.settings.msGraphConfig.enabled ? 'pending' : 'not_configured',
      idempotencyKey,
      driveStatus: hasAttachment ? 'uploading' : 'none',
      driveFileName: resolvedDriveFileName,
      driveOriginalFileName: hasAttachment ? String(attachmentName || '') : undefined,
      driveFileSize: hasAttachment ? Number(attachmentSize) : undefined,
    };

    recordDocumentHistoryEntry({
      doc: newDocument,
      book: targetBook,
      actor: currentUser,
      action: 'create',
      changes: {
        bookCode: { old: null, new: targetBook.code },
        bookName: { old: null, new: targetBook.name },
        fullDocNumber: { old: null, new: fullDocNumber },
        docName: { old: null, new: validDocName },
        title: { old: null, new: finalTitle },
        signer: { old: null, new: signer.trim() },
      },
    });

    // Lưu vào database
    db.documents.unshift(newDocument); // bản ghi mới nhất ở đầu mảng
    db.idempotencyKeys[idempotencyKey] = docId;

    saveDatabaseSync(db);

    // Phát sóng sự kiện SSE đến tất cả thiết bị đang mở
    broadcastSSE('document_issued', { document: newDocument });

    // Tự động đồng bộ sang Google Sheets nếu đã cấu hình
    await syncToGoogleSheets(newDocument, 'create');

    // Trả kết quả
    res.json({ document: newDocument, success: true });
  } catch (err: any) {
    console.error('Lỗi khi cấp số:', err);
    res.status(500).json({ error: 'Lỗi máy chủ khi cấp số: ' + (err.message || 'Không xác định') });
  } finally {
    releaseLock();
  }
});

// 7b. TẢI LÊN / THỬ LƯU LẠI / THAY TỆP WORD ĐÍNH KÈM VÀO GOOGLE DRIVE CHO VĂN BẢN ĐÃ CẤP SỐ
app.post('/api/documents/:id/upload-word', requireAuth, handleWordUploadMulter, async (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const file = req.file;
  const requestedFileName = typeof req.body?.fileName === 'string' ? req.body.fileName.trim() : '';
  const mode = req.body?.mode === 'replace' ? 'replace' : 'retry';

  const doc = db.documents.find((d) => d.id === id);
  if (!doc) {
    return res.status(404).json({ error: 'Không tìm thấy bản ghi văn bản để gắn tệp Word.' });
  }

  const targetBook = getBookById(getDocBookId(doc));

  if (doc.status === 'cancelled' || doc.status === 'deleted') {
    return res.status(400).json({
      error:
        doc.status === 'deleted'
          ? 'Văn bản này đã bị xóa, không thể đính kèm hoặc thay tệp Word.'
          : 'Văn bản này đã bị hủy, không thể đính kèm hoặc thay tệp Word.',
    });
  }

  const isReplacing = doc.driveStatus === 'saved' && Boolean(doc.driveFileId);
  if (mode === 'replace' || doc.driveStatus === 'saved') {
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'replace_file')) {
      return res.status(403).json({
        error: `Bạn cần quyền "Thay tệp" trong ${targetBook.name} để thay thế tệp trên Google Drive.`,
      });
    }
  } else {
    const canAttach =
      isSystemAdmin(currentUser) ||
      hasBookPermission(currentUser, targetBook.id, 'replace_file') ||
      hasBookPermission(currentUser, targetBook.id, 'issue_number') ||
      (hasBookPermission(currentUser, targetBook.id, 'edit') && doc.issuedByUserId === currentUser.id);
    if (!canAttach) {
      return res.status(403).json({
        error: `Bạn không có quyền tải tệp Word cho văn bản trong ${targetBook.name}.`,
      });
    }
  }

  if (!file) {
    return res.status(400).json({
      error: 'Vui lòng chọn tệp Word (.doc hoặc .docx) để tải lên.',
    });
  }

  const validation = validateWordFileBuffer(file.originalname, file.mimetype, file.buffer);
  if (!validation.valid) {
    return res.status(400).json({ error: validation.error });
  }

  // Nếu đang có tiến trình tải lên cho chính bản ghi này, đợi kết quả để tránh tạo bản sao
  const pendingUpload = inFlightDocUploads.get(doc.id);
  if (pendingUpload) {
    try {
      const finishedDoc = await pendingUpload;
      return res.json({
        document: finishedDoc,
        deduplicated: true,
        success: true,
        message: 'Yêu cầu tải tệp trước đó đã hoàn tất.',
      });
    } catch (err: any) {
      return res.status(502).json({
        error: `Đã cấp số ${doc.fullDocNumber}, chưa lưu được tệp Word: ${err.message}`,
        document: doc,
      });
    }
  }

  // Nếu bản ghi đã lưu tệp trên Drive rồi và không phải thao tác thay tệp, trả về bản ghi hiện có để chống lặp
  if (doc.driveStatus === 'saved' && doc.driveFileId && !isReplacing) {
    return res.json({
      document: doc,
      deduplicated: true,
      success: true,
      message: 'Tệp Word của văn bản này đã được lưu trên Google Drive trước đó.',
    });
  }

  const prevFileName = doc.driveFileName || null;
  const prevFileId = doc.driveFileId || null;
  const targetOverrideName =
    requestedFileName
      ? sanitizeCustomDriveFileName(requestedFileName, validation.extension!)
      : buildStandardDriveFileName(doc.fullDocNumber, doc.docName || doc.title, validation.extension!);

  const uploadTask = (async (): Promise<DocumentRecord> => {
    const driveCfg = resolveDriveForBook(targetBook);
    if (!driveCfg.isConfigured) {
      if (!isReplacing) {
        doc.driveStatus = 'failed';
        doc.driveFileName = targetOverrideName;
        doc.driveOriginalFileName = file.originalname;
        doc.driveFileSize = file.size;
        doc.driveError =
          driveCfg.missingReason || 'Thư mục Google Drive chưa được cấu hình hợp lệ.';
        saveDatabaseSync(db);
        broadcastSSE('document_updated', { document: doc });
      }
      throw new Error(driveCfg.missingReason || 'Thư mục Google Drive chưa được cấu hình hợp lệ.');
    }

    if (!isReplacing) {
      doc.driveStatus = 'uploading';
      doc.driveFileName = targetOverrideName;
      doc.driveOriginalFileName = file.originalname;
      doc.driveFileSize = file.size;
      saveDatabaseSync(db);
    }

    try {
      const uploadResult = await uploadWordToGoogleDrive({
        config: driveCfg,
        docId: doc.id,
        bookName: targetBook.name,
        year: doc.year,
        docCategory: doc.docCategory || 'outgoing',
        fullDocNumber: doc.fullDocNumber,
        docName: doc.docName || doc.title,
        originalFileName: file.originalname,
        mimeType: validation.normalizedMimeType || file.mimetype,
        buffer: file.buffer,
        overrideFileName: targetOverrideName,
        replaceMode: isReplacing,
        previousFileId: isReplacing ? prevFileId || undefined : undefined,
      });

      doc.driveStatus = 'saved';
      doc.driveFileId = uploadResult.fileId;
      doc.driveFileName = uploadResult.fileName;
      doc.driveFileUrl = uploadResult.webViewLink;
      doc.driveOriginalFileName = file.originalname;
      doc.driveMimeType = validation.normalizedMimeType || file.mimetype;
      doc.driveFileSize = file.size;
      doc.driveFolderId = uploadResult.folderId || driveCfg.folderId;
      doc.driveFolderPath = uploadResult.folderPath;
      doc.driveUploadedAt = new Date().toISOString();
      delete doc.driveError;

      recordDocumentHistoryEntry({
        doc,
        book: targetBook,
        actor: currentUser,
        action: isReplacing ? 'replace_file' : 'upload_word',
        changes: {
          driveFileName: { old: prevFileName, new: uploadResult.fileName },
          driveFileId: { old: prevFileId, new: uploadResult.fileId },
        },
        reason: isReplacing
          ? 'Thay thế tệp Word mới trên Google Drive'
          : uploadResult.deduplicated
          ? 'Đối soát tệp đã có trên Google Drive theo mã bản ghi'
          : 'Tải lên tệp Word đính kèm vào Google Drive',
      });

      saveDatabaseSync(db);
      broadcastSSE('document_updated', { document: doc });
      await syncToGoogleSheets(doc, 'update');
      return doc;
    } catch (err: any) {
      if (!isReplacing) {
        doc.driveStatus = 'failed';
        doc.driveFileName = targetOverrideName;
        doc.driveOriginalFileName = file.originalname;
        doc.driveFileSize = file.size;
        doc.driveError = err.message || 'Lỗi không xác định khi tải tệp lên Google Drive.';
        saveDatabaseSync(db);
        broadcastSSE('document_updated', { document: doc });
      }
      throw err;
    }
  })();

  inFlightDocUploads.set(doc.id, uploadTask);
  try {
    const updatedDoc = await uploadTask;
    return res.json({
      document: updatedDoc,
      success: true,
    });
  } catch (err: any) {
    return res.status(502).json({
      error: `Đã cấp số ${doc.fullDocNumber}, chưa lưu được tệp Word: ${err.message || 'Lỗi kết nối Google Drive'}`,
      document: doc,
    });
  } finally {
    inFlightDocUploads.delete(doc.id);
  }
});

// 7c. NHẬP VĂN BẢN PDF TỪ SỞ HOẶC NGUỒN KHÁC (Phân tích văn bản / OCR + Xem trước + Lưu sổ Văn bản đến)
app.post('/api/pdf/analyze', requireAuth, handleWordUploadMulter, async (req, res) => {
  const currentUser = (req as any).user as User;
  const file = req.file;
  const rawUrl = typeof req.body?.url === 'string' ? req.body.url.trim() : '';
  const targetBook = getBookById(typeof req.body?.bookId === 'string' ? req.body.bookId : undefined);

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'import_incoming')) {
    return res.status(403).json({
      error: `Bạn chưa được phân quyền "Nhập văn bản đến" trong ${targetBook.name}.`,
    });
  }

  try {
    let buffer: Buffer;
    let originalFileName = '';
    let sourceType: 'pdf_url' | 'pdf_upload';
    let sourceUrl: string | undefined;

    if (file) {
      const validation = validatePdfFileBuffer(file.originalname, file.mimetype, file.buffer);
      if (!validation.valid) {
        return res.status(400).json({ error: validation.error });
      }
      buffer = file.buffer;
      originalFileName = file.originalname || 'van-ban.pdf';
      sourceType = 'pdf_upload';
    } else if (rawUrl) {
      const fetched = await fetchPdfFromUrl(rawUrl);
      buffer = fetched.buffer;
      originalFileName = fetched.fileName;
      sourceType = 'pdf_url';
      sourceUrl = rawUrl;
    } else {
      return res.status(400).json({
        error: 'Vui lòng dán đường dẫn PDF hoặc chọn tệp PDF từ máy để phân tích.',
      });
    }

    const { sha256, extraction } = await analyzePdfBuffer(buffer);
    const duplicateDocument =
      db.documents.find(
        (d) =>
          getDocBookId(d) === targetBook.id &&
          d.fileSha256 === sha256 &&
          d.status !== 'cancelled' &&
          d.status !== 'deleted'
      ) || null;

    // Kiểm tra nghi trùng theo Cơ quan ban hành + Số/ký hiệu + Năm trong cùng sổ
    const parsedYearMatch = (extraction.issueDate || extraction.signDate || '').match(/(\d{4})$/);
    const extractedYear = parsedYearMatch ? parseInt(parsedYearMatch[1], 10) : undefined;
    const suspectedDuplicateDocument = findSuspectedDuplicateDocument(
      targetBook.id,
      extraction.docNumber,
      extraction.issuingAgency,
      extractedYear,
      duplicateDocument?.id
    );

    const stageToken = `pdf-${Date.now()}-${crypto.randomBytes(8).toString('hex')}`;
    stagedPdfs.set(stageToken, {
      token: stageToken,
      bookId: targetBook.id,
      buffer,
      sha256,
      originalFileName,
      sourceType,
      sourceUrl,
      userId: currentUser.id,
      createdAt: Date.now(),
    });

    const payload: PdfAnalyzeResponse = {
      stageToken,
      bookId: targetBook.id,
      sha256,
      fileSize: buffer.length,
      originalFileName,
      sourceType,
      sourceUrl,
      extraction,
      duplicateDocument,
      suspectedDuplicateDocument,
    };

    return res.json(payload);
  } catch (err: any) {
    console.error('[PDF Analyze] Lỗi phân tích PDF:', err);
    return res.status(400).json({
      error: err.message || 'Không thể đọc hoặc phân tích tệp PDF.',
    });
  }
});

app.get('/api/pdf/staged/:stageToken', requireAuth, (req, res) => {
  const { stageToken } = req.params;
  const currentUser = (req as any).user as User;
  const entry = stagedPdfs.get(stageToken);
  if (!entry) {
    return res.status(404).json({ error: 'Phiên xem trước PDF đã hết hạn, vui lòng chọn hoặc tải lại tệp.' });
  }
  const hasImportPerm = isSystemAdmin(currentUser) || hasBookPermission(currentUser, entry.bookId, 'import_incoming');
  if (!hasImportPerm || (currentUser.role !== 'admin' && entry.userId !== currentUser.id)) {
    return res.status(403).json({ error: 'Bạn không có quyền xem tệp PDF tạm này.' });
  }

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="${encodeURIComponent(entry.originalFileName)}"`);
  res.setHeader('Cache-Control', 'private, max-age=300');
  res.send(entry.buffer);
});

app.post('/api/documents/import-pdf', requireAuth, async (req, res) => {
  const currentUser = (req as any).user as User;
  const {
    bookId,
    stageToken,
    fullDocNumber,
    docName,
    title,
    issueDate,
    signDate,
    issuingAgency,
    signer,
    signerRole,
    notes,
    expectedFileName,
    conflictAction,
    customFileName,
    ignoreDuplicatePdf,
    ignoreSuspectedDuplicate,
    needsReview,
    reviewReasons,
    relevantPages,
    idempotencyKey,
  } = req.body || {};

  const entry = stageToken ? stagedPdfs.get(stageToken) : undefined;
  const targetBook = getBookById(bookId || entry?.bookId);

  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'import_incoming')) {
    return res.status(403).json({
      error: `Bạn chưa được phân quyền "Nhập văn bản đến" trong ${targetBook.name}.`,
    });
  }

  const cleanDocNumber = (fullDocNumber || '').trim();
  const cleanDocName = (docName || '').trim();
  const cleanTitle = (title || '').trim();
  // Không tự thay bằng ngày hôm nay khi chưa phát hiện ngày ban hành
  const cleanIssueDate = normalizeDateOnlyDdMmYyyy((issueDate || '').trim());
  const cleanSignDate = normalizeDateOnlyDdMmYyyy((signDate || '').trim());
  const cleanAgency = (issuingAgency || '').trim();
  const cleanSigner = (signer || '').trim();
  const cleanSignerRole = (signerRole || '').trim();

  if (!stageToken || !cleanDocNumber) {
    return res.status(400).json({
      error: 'Thiếu thông tin tệp PDF hoặc Số/ký hiệu văn bản gốc.',
    });
  }

  if (!cleanDocName && !cleanTitle) {
    return res.status(400).json({
      error: 'Vui lòng nhập Tên văn bản hoặc Trích yếu nội dung văn bản.',
    });
  }

  if (!idempotencyKey) {
    return res.status(400).json({ error: 'Thiếu mã chống lặp (idempotency key).' });
  }

  if (!entry || (currentUser.role !== 'admin' && entry.userId !== currentUser.id)) {
    return res.status(400).json({
      error: 'Phiên dữ liệu PDF tạm đã hết hạn hoặc không hợp lệ. Vui lòng phân tích lại tệp PDF.',
    });
  }

  const releaseLock = await atomicLock.acquire();
  let createdDoc: DocumentRecord | null = null;
  let targetFileName = '';

  try {
    if (db.idempotencyKeys[idempotencyKey]) {
      const existingDocId = db.idempotencyKeys[idempotencyKey];
      const existingDoc = db.documents.find((d) => d.id === existingDocId);
      if (existingDoc) {
        return res.json({
          document: existingDoc,
          duplicatePrevented: true,
          driveSaved: existingDoc.driveStatus === 'saved',
          message: 'Văn bản PDF này đã được lưu trước đó.',
        });
      }
    }

    // Xác định năm từ ngày ban hành hoặc ngày ký (dd/mm/yyyy), không làm tăng dãy số văn bản đi
    const refDateForYear = cleanIssueDate || cleanSignDate;
    const dateParts = refDateForYear.split('/');
    const parsedYear =
      dateParts.length === 3 && /^\d{4}$/.test(dateParts[2])
        ? parseInt(dateParts[2], 10)
        : getRequestVietnamTimeInfo(req).year;

    if (!ignoreDuplicatePdf) {
      const duplicateDoc = db.documents.find(
        (d) =>
          getDocBookId(d) === targetBook.id &&
          d.fileSha256 === entry.sha256 &&
          d.status !== 'cancelled'
      );
      if (duplicateDoc) {
        return res.status(409).json({
          code: 'PDF_DUPLICATE',
          error: `Tệp PDF này trùng nội dung với văn bản đã lưu (${duplicateDoc.fullDocNumber}) trong ${targetBook.name}.`,
          existingDocument: duplicateDoc,
        });
      }
    }

    if (!ignoreDuplicatePdf && !ignoreSuspectedDuplicate) {
      const suspectedDoc = findSuspectedDuplicateDocument(
        targetBook.id,
        cleanDocNumber,
        cleanAgency,
        parsedYear
      );
      if (suspectedDoc) {
        return res.status(409).json({
          code: 'PDF_SUSPECTED_DUPLICATE',
          error: `Nghi trùng văn bản "${suspectedDoc.fullDocNumber}" (${suspectedDoc.issuingAgency || 'cùng cơ quan'}) trong ${targetBook.name}. Vui lòng kiểm tra lại hồ sơ đã có hoặc xác nhận lưu mới.`,
          existingDocument: suspectedDoc,
        });
      }
    }

    const nameInputForFile = cleanDocName || cleanTitle;
    const standardName = buildStandardDriveFileName(cleanDocNumber, nameInputForFile, '.pdf');
    if (typeof customFileName === 'string' && customFileName.trim()) {
      targetFileName = sanitizeCustomDriveFileName(customFileName, '.pdf');
    } else if (typeof expectedFileName === 'string' && expectedFileName.trim()) {
      targetFileName = sanitizeCustomDriveFileName(expectedFileName, '.pdf');
    } else {
      targetFileName = standardName;
    }

    const conflictDoc = findConflictingDocumentByDriveFileName(targetFileName, undefined, targetBook.id);
    if (conflictDoc) {
      if (conflictAction === 'suffix') {
        targetFileName = resolveUniqueDriveFileName(targetFileName, undefined, targetBook.id);
      } else if (conflictAction !== 'overwrite') {
        return res.status(409).json({
          code: 'FILENAME_CONFLICT',
          error: `Tên tệp "${targetFileName}" đã trùng với hồ sơ ${conflictDoc.fullDocNumber} (${conflictDoc.docName || conflictDoc.title}). Vui lòng chọn cách xử lý.`,
          conflictingFileName: targetFileName,
          suggestedSuffixName: resolveUniqueDriveFileName(targetFileName, undefined, targetBook.id),
          existingDocument: conflictDoc,
        });
      }
    }

    const docId = `doc-${Date.now()}-${crypto.randomBytes(4).toString('hex')}`;
    const globalStt = db.documents.length + 1;

    const newDocument: DocumentRecord = {
      id: docId,
      bookId: targetBook.id,
      bookCode: targetBook.code,
      bookName: targetBook.name,
      stt: globalStt,
      year: parsedYear,
      issueDate: cleanIssueDate,
      signDate: cleanSignDate || undefined,
      createdAt: new Date().toISOString(),
      docCategory: 'incoming',
      issuingAgency: cleanAgency || undefined,
      typeCode: 'VBD',
      typeName: 'Văn bản đến',
      sequenceNumber: 0,
      fullDocNumber: cleanDocNumber,
      docName: cleanDocName || cleanTitle,
      title: cleanTitle || cleanDocName,
      signer: cleanSigner || 'Chưa xác định',
      signerRole: cleanSignerRole || (cleanAgency || 'Cơ quan ban hành'),
      issuedBy: currentUser.fullName,
      issuedByUsername: currentUser.username,
      issuedByUserId: currentUser.id,
      status: 'active',
      notes: typeof notes === 'string' ? notes.trim() : '',
      auditLogs: [],
      syncStatus: db.settings.msGraphConfig.enabled ? 'pending' : 'not_configured',
      idempotencyKey,
      driveStatus: 'uploading',
      driveFileName: targetFileName,
      driveOriginalFileName: entry.originalFileName,
      driveMimeType: 'application/pdf',
      driveFileSize: entry.buffer.length,
      sourceType: entry.sourceType,
      sourceUrl: entry.sourceUrl,
      fileSha256: entry.sha256,
      needsReview: Boolean(needsReview),
      reviewReasons: Array.isArray(reviewReasons) ? reviewReasons : undefined,
      relevantPages: Array.isArray(relevantPages) ? relevantPages : undefined,
    };

    recordDocumentHistoryEntry({
      doc: newDocument,
      book: targetBook,
      actor: currentUser,
      action: 'import_pdf',
      changes: {
        bookCode: { old: null, new: targetBook.code },
        bookName: { old: null, new: targetBook.name },
        fullDocNumber: { old: null, new: cleanDocNumber },
        docName: { old: null, new: cleanDocName || cleanTitle },
        issuingAgency: { old: null, new: cleanAgency || '(chưa xác định)' },
        driveFileName: { old: null, new: targetFileName },
      },
      reason: entry.sourceType === 'pdf_url' ? 'Nhập văn bản đến từ đường dẫn PDF' : 'Nhập văn bản đến từ tệp PDF tải lên',
    });

    db.documents.unshift(newDocument);
    db.idempotencyKeys[idempotencyKey] = docId;
    saveDatabaseSync(db);

    broadcastSSE('document_issued', { document: newDocument });
    await syncToGoogleSheets(newDocument, 'create');
    createdDoc = newDocument;
  } finally {
    releaseLock();
  }

  if (!createdDoc) {
    return res.status(500).json({ error: 'Không thể khởi tạo bản ghi văn bản đến.' });
  }

  // Tải tệp PDF gốc lên thư mục Google Drive của sổ đã cấu hình với đúng tên hiển thị
  const driveCfg = resolveDriveForBook(targetBook);
  if (!driveCfg.isConfigured) {
    createdDoc.driveStatus = 'failed';
    createdDoc.driveError = driveCfg.missingReason || 'Chưa cấu hình thư mục Google Drive hợp lệ.';
    saveDatabaseSync(db);
    broadcastSSE('document_updated', { document: createdDoc });
    return res.json({
      document: createdDoc,
      driveSaved: false,
      uploadError: createdDoc.driveError,
      stageToken,
    });
  }

  try {
    const uploadResult = await uploadFileToGoogleDrive({
      config: driveCfg,
      docId: createdDoc.id,
      bookName: targetBook.name,
      year: createdDoc.year,
      docCategory: 'incoming',
      fullDocNumber: createdDoc.fullDocNumber,
      docName: createdDoc.docName || createdDoc.title,
      originalFileName: entry.originalFileName || `${targetFileName}`,
      mimeType: 'application/pdf',
      buffer: entry.buffer,
      overrideFileName: targetFileName,
    });

    createdDoc.driveStatus = 'saved';
    createdDoc.driveFileId = uploadResult.fileId;
    createdDoc.driveFileName = uploadResult.fileName;
    createdDoc.driveFileUrl = uploadResult.webViewLink;
    createdDoc.driveFolderId = uploadResult.folderId || driveCfg.folderId;
    createdDoc.driveFolderPath = uploadResult.folderPath;
    createdDoc.driveUploadedAt = new Date().toISOString();
    delete createdDoc.driveError;

    recordDocumentHistoryEntry({
      doc: createdDoc,
      book: targetBook,
      actor: currentUser,
      action: 'upload_pdf',
      changes: {
        driveFileName: { old: null, new: uploadResult.fileName },
        driveFileId: { old: null, new: uploadResult.fileId },
      },
      reason: 'Lưu bản PDF gốc lên thư mục Google Drive',
    });

    saveDatabaseSync(db);
    broadcastSSE('document_updated', { document: createdDoc });
    await syncToGoogleSheets(createdDoc, 'update');
    stagedPdfs.delete(stageToken);

    return res.json({
      document: createdDoc,
      driveSaved: true,
    });
  } catch (err: any) {
    createdDoc.driveStatus = 'failed';
    createdDoc.driveError = err.message || 'Lỗi kết nối khi lưu tệp PDF lên Google Drive.';
    saveDatabaseSync(db);
    broadcastSSE('document_updated', { document: createdDoc });
    return res.json({
      document: createdDoc,
      driveSaved: false,
      uploadError: createdDoc.driveError,
      stageToken,
    });
  }
});

// 7d. THỬ LƯU LẠI / THAY TỆP PDF LÊN GOOGLE DRIVE CHO BẢN GHI ĐÃ LƯU
app.post('/api/documents/:id/upload-pdf', requireAuth, handleWordUploadMulter, async (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const file = req.file;
  const stageToken = typeof req.body?.stageToken === 'string' ? req.body.stageToken.trim() : '';
  const requestedFileName = typeof req.body?.fileName === 'string' ? req.body.fileName.trim() : '';
  const mode = req.body?.mode === 'replace' ? 'replace' : 'retry';

  const doc = db.documents.find((d) => d.id === id);
  if (!doc) {
    return res.status(404).json({ error: 'Không tìm thấy hồ sơ văn bản để lưu tệp PDF.' });
  }

  const targetBook = getBookById(getDocBookId(doc));

  if (doc.status === 'cancelled' || doc.status === 'deleted') {
    return res.status(400).json({
      error:
        doc.status === 'deleted'
          ? 'Văn bản đã bị xóa, không thể lưu hoặc thay tệp lên Drive.'
          : 'Văn bản đã bị hủy, không thể lưu hoặc thay tệp lên Drive.',
    });
  }

  const isReplacing = doc.driveStatus === 'saved' && Boolean(doc.driveFileId);
  if (mode === 'replace' || doc.driveStatus === 'saved') {
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'replace_file')) {
      return res.status(403).json({
        error: `Bạn cần quyền "Thay tệp" trong ${targetBook.name} để thay thế bản PDF trên Drive.`,
      });
    }
  } else {
    const canAttach =
      isSystemAdmin(currentUser) ||
      hasBookPermission(currentUser, targetBook.id, 'replace_file') ||
      hasBookPermission(currentUser, targetBook.id, 'import_incoming') ||
      (hasBookPermission(currentUser, targetBook.id, 'edit') && doc.issuedByUserId === currentUser.id);
    if (!canAttach) {
      return res.status(403).json({
        error: `Bạn không có quyền tải tệp PDF cho văn bản trong ${targetBook.name}.`,
      });
    }
  }

  let pdfBuffer: Buffer;
  let originalName: string;

  if (file) {
    const validation = validatePdfFileBuffer(file.originalname, file.mimetype, file.buffer);
    if (!validation.valid) {
      return res.status(400).json({ error: validation.error });
    }
    pdfBuffer = file.buffer;
    originalName = file.originalname;
  } else if (stageToken) {
    const staged = stagedPdfs.get(stageToken);
    if (!staged || (currentUser.role !== 'admin' && staged.userId !== currentUser.id)) {
      return res.status(400).json({
        error: 'Phiên PDF tạm đã hết hạn. Vui lòng chọn lại tệp PDF từ máy để thử lưu lại.',
      });
    }
    pdfBuffer = staged.buffer;
    originalName = staged.originalFileName;
  } else {
    return res.status(400).json({ error: 'Vui lòng chọn tệp PDF để tải lên Google Drive.' });
  }

  const pendingUpload = inFlightDocUploads.get(doc.id);
  if (pendingUpload) {
    try {
      const finishedDoc = await pendingUpload;
      return res.json({
        document: finishedDoc,
        deduplicated: true,
        success: true,
      });
    } catch (err: any) {
      return res.status(502).json({
        error: `Chưa lưu tệp lên Drive: ${err.message}`,
        document: doc,
      });
    }
  }

  if (doc.driveStatus === 'saved' && doc.driveFileId && !isReplacing) {
    return res.json({
      document: doc,
      deduplicated: true,
      success: true,
      message: 'Tệp PDF đã được lưu trên Google Drive trước đó.',
    });
  }

  const prevFileName = doc.driveFileName || null;
  const prevFileId = doc.driveFileId || null;
  const targetOverrideName = requestedFileName
    ? sanitizeCustomDriveFileName(requestedFileName, '.pdf')
    : buildStandardDriveFileName(doc.fullDocNumber, doc.docName || doc.title, '.pdf');

  const uploadTask = (async (): Promise<DocumentRecord> => {
    const driveCfg = resolveDriveForBook(targetBook);
    if (!driveCfg.isConfigured) {
      if (!isReplacing) {
        doc.driveStatus = 'failed';
        doc.driveFileName = targetOverrideName;
        doc.driveError = driveCfg.missingReason || 'Thư mục Google Drive chưa được cấu hình hợp lệ.';
        saveDatabaseSync(db);
        broadcastSSE('document_updated', { document: doc });
      }
      throw new Error(driveCfg.missingReason || 'Thư mục Google Drive chưa được cấu hình hợp lệ.');
    }

    if (!isReplacing) {
      doc.driveStatus = 'uploading';
      doc.driveFileName = targetOverrideName;
      doc.driveOriginalFileName = originalName;
      doc.driveFileSize = pdfBuffer.length;
      saveDatabaseSync(db);
    }

    try {
      const uploadResult = await uploadFileToGoogleDrive({
        config: driveCfg,
        docId: doc.id,
        bookName: targetBook.name,
        year: doc.year,
        docCategory: doc.docCategory || 'incoming',
        fullDocNumber: doc.fullDocNumber,
        docName: doc.docName || doc.title,
        originalFileName: originalName,
        mimeType: 'application/pdf',
        buffer: pdfBuffer,
        overrideFileName: targetOverrideName,
        replaceMode: isReplacing,
        previousFileId: isReplacing ? prevFileId || undefined : undefined,
      });

      doc.driveStatus = 'saved';
      doc.driveFileId = uploadResult.fileId;
      doc.driveFileName = uploadResult.fileName;
      doc.driveFileUrl = uploadResult.webViewLink;
      doc.driveOriginalFileName = originalName;
      doc.driveMimeType = 'application/pdf';
      doc.driveFileSize = pdfBuffer.length;
      doc.driveFolderId = uploadResult.folderId || driveCfg.folderId;
      doc.driveFolderPath = uploadResult.folderPath;
      doc.driveUploadedAt = new Date().toISOString();
      delete doc.driveError;

      recordDocumentHistoryEntry({
        doc,
        book: targetBook,
        actor: currentUser,
        action: isReplacing ? 'replace_file' : 'upload_pdf',
        changes: {
          driveFileName: { old: prevFileName, new: uploadResult.fileName },
          driveFileId: { old: prevFileId, new: uploadResult.fileId },
        },
        reason: isReplacing
          ? 'Thay thế bản PDF mới trên Google Drive'
          : uploadResult.deduplicated
          ? 'Đối soát tệp PDF đã có trên Google Drive theo mã bản ghi'
          : 'Thử lưu lại tệp PDF vào Google Drive',
      });

      saveDatabaseSync(db);
      broadcastSSE('document_updated', { document: doc });
      await syncToGoogleSheets(doc, 'update');
      if (stageToken) stagedPdfs.delete(stageToken);
      return doc;
    } catch (err: any) {
      if (!isReplacing) {
        doc.driveStatus = 'failed';
        doc.driveFileName = targetOverrideName;
        doc.driveError = err.message || 'Lỗi kết nối Google Drive.';
        saveDatabaseSync(db);
        broadcastSSE('document_updated', { document: doc });
      }
      throw err;
    }
  })();

  inFlightDocUploads.set(doc.id, uploadTask);
  try {
    const updatedDoc = await uploadTask;
    return res.json({ document: updatedDoc, success: true });
  } catch (err: any) {
    return res.status(502).json({
      error: `Chưa lưu tệp lên Drive: ${err.message || 'Lỗi kết nối Google Drive'}`,
      document: doc,
    });
  } finally {
    inFlightDocUploads.delete(doc.id);
  }
});

// 8. Sổ Theo Dõi - Danh sách văn bản theo quyền Xem từng sổ
app.get('/api/documents', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const {
    bookId,
    year,
    month,
    typeCode,
    category,
    signer,
    status,
    driveStatus,
    deletedBy,
    fromDate,
    toDate,
    search,
    sort = 'newest',
  } = req.query as Record<string, string>;

  if (bookId && bookId !== 'all') {
    const targetBook = getBookById(bookId);
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'view')) {
      return res.status(403).json({
        error: `Bạn chưa được cấp quyền "Xem" trong ${targetBook.name}.`,
      });
    }
  }

  let list = db.documents.filter((d) => {
    const bId = getDocBookId(d);
    if (bookId && bookId !== 'all' && bId !== bookId) return false;
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'view');
  });

  if (year && year !== 'all') {
    const y = parseInt(year, 10);
    list = list.filter((d) => d.year === y);
  }

  if (month && month !== 'all') {
    const m = parseInt(month, 10);
    list = list.filter((d) => {
      const parts = (d.issueDate || '').split('/');
      return parts.length >= 2 && parseInt(parts[1], 10) === m;
    });
  }

  if (category && category !== 'all') {
    if (category === 'incoming') {
      list = list.filter((d) => d.docCategory === 'incoming');
    } else if (category === 'outgoing') {
      list = list.filter((d) => d.docCategory !== 'incoming');
    }
  }

  if (typeCode && typeCode !== 'all') {
    list = list.filter((d) => d.typeCode === typeCode);
  }

  if (signer && signer !== 'all') {
    list = list.filter((d) => d.signer.toLowerCase() === signer.toLowerCase());
  }

  if (status === 'deleted') {
    list = list.filter((d) => d.status === 'deleted');
  } else if (status === 'all_with_deleted') {
    // keep all
  } else if (status && status !== 'all') {
    list = list.filter((d) => d.status === status);
  } else {
    // Danh sách hiệu lực mặc định không bao gồm văn bản đã xóa
    list = list.filter((d) => d.status !== 'deleted');
  }

  if (deletedBy && deletedBy !== 'all') {
    const lowerDel = deletedBy.trim().toLowerCase();
    list = list.filter(
      (d) =>
        (d.deletedByUsername || '').toLowerCase() === lowerDel ||
        (d.deletedByUserId || '') === deletedBy
    );
  }

  if (fromDate && /^\d{4}-\d{2}-\d{2}$/.test(fromDate)) {
    list = list.filter((d) => {
      const refIso = d.status === 'deleted' ? d.deletedAt || d.createdAt : d.createdAt;
      return getVietnamDateKey(new Date(refIso)) >= fromDate;
    });
  }

  if (toDate && /^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
    list = list.filter((d) => {
      const refIso = d.status === 'deleted' ? d.deletedAt || d.createdAt : d.createdAt;
      return getVietnamDateKey(new Date(refIso)) <= toDate;
    });
  }

  if (driveStatus && driveStatus !== 'all') {
    list = list.filter((d) => (d.driveStatus || 'none') === driveStatus);
  }

  if (search && search.trim()) {
    const term = search.trim().toLowerCase();
    list = list.filter(
      (d) =>
        d.id.toLowerCase().includes(term) ||
        d.fullDocNumber.toLowerCase().includes(term) ||
        (d.docName && d.docName.toLowerCase().includes(term)) ||
        (d.driveFileName && d.driveFileName.toLowerCase().includes(term)) ||
        (d.issuingAgency && d.issuingAgency.toLowerCase().includes(term)) ||
        d.title.toLowerCase().includes(term) ||
        d.signer.toLowerCase().includes(term) ||
        d.issuedBy.toLowerCase().includes(term) ||
        (d.deleteReason && d.deleteReason.toLowerCase().includes(term)) ||
        (d.notes && d.notes.toLowerCase().includes(term))
    );
  }

  if (sort === 'newest') {
    list.sort(
      (a, b) =>
        new Date(b.deletedAt || b.createdAt).getTime() -
        new Date(a.deletedAt || a.createdAt).getTime()
    );
  } else if (sort === 'oldest') {
    list.sort(
      (a, b) =>
        new Date(a.deletedAt || a.createdAt).getTime() -
        new Date(b.deletedAt || b.createdAt).getTime()
    );
  } else if (sort === 'number_asc') {
    list.sort((a, b) => a.sequenceNumber - b.sequenceNumber);
  } else if (sort === 'number_desc') {
    list.sort((a, b) => b.sequenceNumber - a.sequenceNumber);
  }

  res.json({
    total: list.length,
    documents: list,
  });
});

// Chi tiết 1 văn bản (kiểm tra quyền Xem trên sổ)
app.get('/api/documents/:id', requireAuth, (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản' });
  const docBookId = getDocBookId(doc);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, docBookId, 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền xem hồ sơ thuộc sổ này.' });
  }
  res.json({ document: doc });
});

// Kiểm tra quyền và trả liên kết mở tệp Google Drive (bao gồm tra cứu tệp của VB đã xóa nếu có quyền Xem)
app.get('/api/documents/:id/open-file', requireAuth, (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản' });

  const docBookId = getDocBookId(doc);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, docBookId, 'view')) {
    return res.status(403).json({ error: 'Bạn không có quyền mở tệp trong sổ này.' });
  }

  if (doc.driveStatus !== 'saved' || !doc.driveFileUrl) {
    return res.status(404).json({ error: 'Hồ sơ này chưa có tệp được lưu trên Google Drive.' });
  }

  res.json({
    url: doc.driveFileUrl,
    fileName: doc.driveFileName || '',
    fileId: doc.driveFileId || '',
  });
});

// Chỉnh sửa văn bản (Yêu cầu quyền "Sửa" trong sổ; sửa số thứ tự văn bản đi chỉ dành cho Admin hoặc Quản trị sổ)
app.put('/api/documents/:id', requireAuth, async (req, res) => {
  const { id } = req.params;
  const {
    docName,
    title,
    signer,
    signerRole,
    notes,
    sequenceNumber,
    fullDocNumber,
    issuingAgency,
    issueDate,
    signDate,
    needsReview,
  } = req.body;
  const currentUser = (req as any).user as User;

  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản' });

  const targetBook = getBookById(getDocBookId(doc));
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'edit')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Sửa" trong ${targetBook.name} để chỉnh sửa thông tin văn bản này.`,
    });
  }

  if (doc.status === 'cancelled' || doc.status === 'deleted') {
    return res.status(400).json({
      error:
        doc.status === 'deleted'
          ? 'Không thể chỉnh sửa văn bản đã bị xóa'
          : 'Không thể chỉnh sửa văn bản đã bị hủy',
    });
  }

  const changes: Record<string, { old: string | null; new: string | null }> = {};
  const canManageBook = isSystemAdmin(currentUser) || hasBookPermission(currentUser, targetBook.id, 'manage_book');

  if (doc.docCategory === 'incoming') {
    if (fullDocNumber && fullDocNumber.trim() !== doc.fullDocNumber) {
      changes.fullDocNumber = { old: doc.fullDocNumber, new: fullDocNumber.trim() };
      doc.fullDocNumber = fullDocNumber.trim();
    }
  } else if (canManageBook) {
    if (typeof sequenceNumber === 'number' && sequenceNumber !== doc.sequenceNumber) {
      const newSeq = sequenceNumber;
      changes.sequenceNumber = { old: String(doc.sequenceNumber), new: String(newSeq) };
      doc.sequenceNumber = newSeq;

      const seqStr = formatSequenceNumber(newSeq);
      const computedDocNum = renderFullDocNumber(
        targetBook.numberFormat,
        seqStr,
        doc.typeCode,
        targetBook.unitCode
      );
      const targetFullNum = fullDocNumber && fullDocNumber.trim() ? fullDocNumber.trim() : computedDocNum;
      changes.fullDocNumber = { old: doc.fullDocNumber, new: targetFullNum };
      doc.fullDocNumber = targetFullNum;
    } else if (fullDocNumber && fullDocNumber.trim() !== doc.fullDocNumber) {
      changes.fullDocNumber = { old: doc.fullDocNumber, new: fullDocNumber.trim() };
      doc.fullDocNumber = fullDocNumber.trim();
    }
  }

  if (docName !== undefined && docName.trim() !== (doc.docName || '')) {
    changes.docName = { old: doc.docName || null, new: docName.trim() };
    doc.docName = docName.trim();
  }

  if (title !== undefined && title.trim() !== doc.title) {
    changes.title = { old: doc.title, new: title.trim() };
    doc.title = title.trim();
  }

  if (signer && signer.trim() !== doc.signer) {
    changes.signer = { old: doc.signer, new: signer.trim() };
    doc.signer = signer.trim();
  }

  if (signerRole && signerRole.trim() !== doc.signerRole) {
    changes.signerRole = { old: doc.signerRole, new: signerRole.trim() };
    doc.signerRole = signerRole.trim();
  }

  if (typeof issuingAgency === 'string' && issuingAgency.trim() !== (doc.issuingAgency || '')) {
    changes.issuingAgency = { old: doc.issuingAgency || '', new: issuingAgency.trim() };
    doc.issuingAgency = issuingAgency.trim();
  }

  if (typeof issueDate === 'string') {
    const trimmedIssueDate = issueDate.trim();
    if (!trimmedIssueDate && doc.docCategory === 'incoming') {
      if (doc.issueDate) {
        changes.issueDate = { old: doc.issueDate, new: '' };
        doc.issueDate = '';
      }
    } else if (trimmedIssueDate) {
      const normalizedIssueDate = normalizeDateOnlyDdMmYyyy(trimmedIssueDate);
      if (normalizedIssueDate && normalizedIssueDate !== doc.issueDate) {
        changes.issueDate = { old: doc.issueDate, new: normalizedIssueDate };
        doc.issueDate = normalizedIssueDate;
      }
    }
  }

  if (typeof signDate === 'string') {
    const normalizedSignDate = normalizeDateOnlyDdMmYyyy(signDate.trim());
    if (normalizedSignDate !== (doc.signDate || '')) {
      changes.signDate = { old: doc.signDate || '', new: normalizedSignDate };
      doc.signDate = normalizedSignDate || undefined;
    }
  }

  if (typeof notes === 'string' && notes.trim() !== (doc.notes || '')) {
    changes.notes = { old: doc.notes || '', new: notes.trim() };
    doc.notes = notes.trim();
  }

  if (typeof needsReview === 'boolean' && needsReview !== Boolean(doc.needsReview)) {
    changes.needsReview = { old: String(Boolean(doc.needsReview)), new: String(needsReview) };
    doc.needsReview = needsReview;
    if (!needsReview) delete doc.reviewReasons;
  }

  if (Object.keys(changes).length > 0) {
    recordDocumentHistoryEntry({
      doc,
      book: targetBook,
      actor: currentUser,
      action: 'update',
      changes,
    });
    if (doc.syncStatus === 'synced') {
      doc.syncStatus = 'pending';
    }
    saveDatabaseSync(db);
    broadcastSSE('document_updated', { document: doc });
    await syncToGoogleSheets(doc, 'update');
  }

  res.json({ document: doc });
});

// XÓA VĂN BẢN (Yêu cầu quyền "Xóa VB" hoặc Quản trị viên, bắt buộc nhập lý do & xác nhận;
// chuyển sang trạng thái Đã xóa + VB đã xóa, giữ số trong dãy và giữ tệp trên Google Drive)
app.delete('/api/documents/:id', requireAuth, async (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản cần xóa' });

  const targetBook = getBookById(getDocBookId(doc));
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'delete')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Xóa văn bản" trong ${targetBook.name} để thực hiện thao tác này.`,
    });
  }

  if (doc.status === 'deleted') {
    return res.status(400).json({ error: 'Văn bản này đã được xóa trước đó.' });
  }

  const rawReason = req.body?.reason ?? req.body?.deleteReason ?? '';
  const cleanReason = typeof rawReason === 'string' ? rawReason.trim() : '';
  if (!cleanReason) {
    return res.status(400).json({ error: 'Bắt buộc nhập lý do xóa văn bản.' });
  }
  if (req.body?.confirmed === false) {
    return res.status(400).json({ error: 'Vui lòng xác nhận trước khi xóa văn bản.' });
  }

  const releaseLock = await atomicLock.acquire();
  try {
    const lockedDoc = db.documents.find((d) => d.id === id);
    if (!lockedDoc || lockedDoc.status === 'deleted') {
      return res.status(400).json({ error: 'Văn bản này đã được xóa trước đó.' });
    }

    const prevStatus = lockedDoc.status;
    const deletedSnapshot = buildDocumentSnapshot(doc, targetBook);
    const nowIso = new Date().toISOString();
    const nowVn = formatVietnamDateTime(new Date(nowIso));

    doc.status = 'deleted';
    doc.deleteReason = cleanReason;
    doc.deletedAt = nowIso;
    doc.deletedAtVietnam = nowVn;
    doc.deletedBy = `${currentUser.fullName} (${currentUser.username})`;
    doc.deletedByUsername = currentUser.username;
    doc.deletedByUserId = currentUser.id;
    doc.deletedSnapshot = deletedSnapshot;
    doc.excelMovedToDeletedSheet = false;
    doc.excelRemovedFromMainSheet = false;
    delete doc.excelSyncError;
    delete doc.excelSyncedAt;
    doc.excelSyncTargets = {};

    const hasExcelTargets = getGoogleSheetsTargets(targetBook).length > 0;
    doc.excelDeleteSyncStatus = hasExcelTargets ? 'pending' : 'not_configured';

    recordDocumentHistoryEntry({
      doc,
      book: targetBook,
      actor: currentUser,
      action: 'delete',
      changes: {
        status: { old: prevStatus, new: 'deleted' },
        deleteReason: { old: null, new: cleanReason },
        ...(doc.driveFileId
          ? { driveFilePreserved: { old: doc.driveFileName || doc.driveFileId, new: doc.driveFileName || doc.driveFileId } }
          : {}),
      },
      reason: cleanReason,
      timestampIso: nowIso,
      snapshot: deletedSnapshot,
    });

    saveDatabaseSync(db);

    const syncResult = await syncDeletedDocumentToOnlineExcel(doc, targetBook);

    broadcastSSE('document_deleted', {
      id: doc.id,
      fullDocNumber: doc.fullDocNumber,
      document: doc,
    });

    res.json({
      success: true,
      document: doc,
      excelSync: syncResult,
      message:
        syncResult.attempted && !syncResult.ok
          ? `Đã xóa văn bản ${doc.fullDocNumber} trên hệ thống (Chờ đồng bộ Excel: ${syncResult.error || 'Lỗi kết nối'})`
          : `Đã xóa văn bản ${doc.fullDocNumber} và chuyển vào mục VB đã xóa`,
    });
  } finally {
    releaseLock();
  }
});

// THỬ ĐỒNG BỘ LẠI EXCEL (Hoàn tất phần chưa xong theo mã hồ sơ cố định, tránh trùng dòng trong VB đã xóa)
app.post('/api/documents/:id/retry-excel-sync', requireAuth, async (req, res) => {
  const { id } = req.params;
  const currentUser = (req as any).user as User;
  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản để đồng bộ lại.' });

  const targetBook = getBookById(getDocBookId(doc));
  const canRetry =
    isSystemAdmin(currentUser) ||
    hasBookPermission(currentUser, targetBook.id, 'delete') ||
    hasBookPermission(currentUser, targetBook.id, 'edit') ||
    hasBookPermission(currentUser, targetBook.id, 'manage_book');
  if (!canRetry) {
    return res.status(403).json({
      error: `Bạn không có quyền đồng bộ lại Excel cho văn bản thuộc ${targetBook.name}.`,
    });
  }

  const releaseLock = await atomicLock.acquire();
  try {
    if (doc.status === 'deleted') {
      const prevStatus = doc.excelDeleteSyncStatus || 'pending';
      const result = await syncDeletedDocumentToOnlineExcel(doc, targetBook);
      const nextStatus: ExcelDeleteSyncStatus = doc.excelDeleteSyncStatus || 'not_configured';

      recordDocumentHistoryEntry({
        doc,
        book: targetBook,
        actor: currentUser,
        action: 'retry_excel_sync',
        changes: {
          excelDeleteSyncStatus: { old: prevStatus, new: nextStatus },
        },
        reason: result.ok
          ? 'Thử đồng bộ lại chuyển dòng sang sheet VB đã xóa thành công'
          : `Thử đồng bộ lại Excel chưa thành công: ${result.error || 'Lỗi kết nối'}`,
      });
      saveDatabaseSync(db);
      broadcastSSE('document_updated', { document: doc });

      if (!result.ok) {
        return res.status(502).json({
          success: false,
          document: doc,
          error: result.error || 'Đồng bộ lại Excel chưa thành công.',
        });
      }
      return res.json({
        success: true,
        document: doc,
        message: 'Đã hoàn tất đồng bộ dòng sang sheet VB đã xóa.',
      });
    }

    const prevSync = doc.syncStatus;
    const result = await syncToGoogleSheets(doc, doc.status === 'cancelled' ? 'cancel' : 'update');
    recordDocumentHistoryEntry({
      doc,
      book: targetBook,
      actor: currentUser,
      action: 'retry_excel_sync',
      changes: {
        syncStatus: { old: prevSync, new: doc.syncStatus },
      },
      reason: result.ok ? 'Thử đồng bộ lại Excel thành công' : `Đồng bộ lại thất bại: ${result.error || ''}`,
    });
    saveDatabaseSync(db);
    broadcastSSE('document_updated', { document: doc });

    if (!result.ok) {
      return res.status(502).json({
        success: false,
        document: doc,
        error: result.error || 'Đồng bộ Excel chưa thành công.',
      });
    }
    return res.json({
      success: true,
      document: doc,
      message: 'Đã đồng bộ lại dữ liệu văn bản sang Excel.',
    });
  } finally {
    releaseLock();
  }
});

// HỦY SỐ VĂN BẢN (Yêu cầu quyền "Hủy" trong sổ - Không xóa cứng, lưu lý do, số không được cấp lại)
app.post('/api/documents/:id/cancel', requireAuth, async (req, res) => {
  const { id } = req.params;
  const { reason } = req.body;
  const currentUser = (req as any).user as User;

  if (!reason || !reason.trim()) {
    return res.status(400).json({ error: 'Bắt buộc nhập lý do hủy văn bản' });
  }

  const doc = db.documents.find((d) => d.id === id);
  if (!doc) return res.status(404).json({ error: 'Không tìm thấy văn bản' });

  const targetBook = getBookById(getDocBookId(doc));
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'cancel')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Hủy" trong ${targetBook.name} để hủy văn bản này.`,
    });
  }

  if (doc.status === 'deleted') {
    return res.status(400).json({ error: 'Văn bản này đã bị xóa.' });
  }

  if (doc.status === 'cancelled') {
    return res.status(400).json({ error: 'Văn bản này đã được hủy trước đó' });
  }

  doc.status = 'cancelled';
  doc.cancelReason = reason.trim();
  doc.cancelledAt = new Date().toISOString();
  doc.cancelledBy = `${currentUser.fullName} (${currentUser.username})`;

  recordDocumentHistoryEntry({
    doc,
    book: targetBook,
    actor: currentUser,
    action: 'cancel',
    reason: reason.trim(),
    changes: {
      status: { old: 'active', new: 'cancelled' },
    },
  });

  if (doc.syncStatus === 'synced') {
    doc.syncStatus = 'pending';
  }

  saveDatabaseSync(db);
  broadcastSSE('document_cancelled', { document: doc });
  await syncToGoogleSheets(doc, 'cancel');
  res.json({ document: doc });
});

// 9. THỐNG KÊ (Dashboard & Statistics - Lọc theo sổ được quyền Xem, loại trừ văn bản đã xóa khỏi thống kê hiệu lực)
app.get('/api/statistics', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const queryYear = req.query.year ? parseInt(req.query.year as string, 10) : getRequestVietnamTimeInfo(req).year;
  const bookId = (req.query.bookId as string) || '';

  if (bookId && bookId !== 'all') {
    const targetBook = getBookById(bookId);
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'view')) {
      return res.status(403).json({
        error: `Bạn chưa được cấp quyền "Xem" trong ${targetBook.name}.`,
      });
    }
  }

  const docsOfYear = db.documents.filter((d) => {
    if (d.year !== queryYear) return false;
    if (d.status === 'deleted') return false;
    const bId = getDocBookId(d);
    if (bookId && bookId !== 'all' && bId !== bookId) return false;
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'view');
  });

  const totalIssued = docsOfYear.length;
  const activeCount = docsOfYear.filter((d) => d.status === 'active').length;
  const cancelledCount = docsOfYear.filter((d) => d.status === 'cancelled').length;
  const pendingSyncCount = docsOfYear.filter((d) => d.syncStatus === 'pending').length;

  const byType = db.documentTypes.map((type) => {
    const typeDocs = docsOfYear.filter((d) => d.docCategory !== 'incoming' && d.typeCode === type.code);
    return {
      typeCode: type.code,
      typeName: type.name,
      count: typeDocs.length,
      cancelledCount: typeDocs.filter((d) => d.status === 'cancelled').length,
    };
  });
  const incomingDocs = docsOfYear.filter((d) => d.docCategory === 'incoming');
  if (incomingDocs.length > 0) {
    byType.push({
      typeCode: 'VBD',
      typeName: 'Văn bản đến (PDF)',
      count: incomingDocs.length,
      cancelledCount: incomingDocs.filter((d) => d.status === 'cancelled').length,
    });
  }

  const byMonth = Array.from({ length: 12 }, (_, i) => {
    const month = i + 1;
    const monthDocs = docsOfYear.filter((d) => {
      const parts = (d.issueDate || '').split('/');
      return parts.length >= 2 && parseInt(parts[1], 10) === month;
    });
    return {
      month,
      count: monthDocs.length,
      cancelledCount: monthDocs.filter((d) => d.status === 'cancelled').length,
    };
  });

  const stats: StatisticsData = {
    year: queryYear,
    bookId: bookId || undefined,
    totalIssued,
    activeCount,
    cancelledCount,
    pendingSyncCount,
    byType,
    byMonth,
  };

  res.json(stats);
});

// 10. XUẤT FILE EXCEL .XLSX CHUẨN (Giữ đúng bộ lọc, mã hồ sơ cố định, quyền Xuất sổ & sheet riêng "VB đã xóa")
app.get('/api/export-excel', requireAuth, async (req, res) => {
  const currentUser = (req as any).user as User;
  const {
    bookId,
    year,
    month,
    typeCode,
    status,
    category,
    signer,
    driveStatus,
    search,
    sort = 'number_asc',
  } = req.query as Record<string, string>;

  if (bookId && bookId !== 'all') {
    const targetBook = getBookById(bookId);
    if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'export')) {
      return res.status(403).json({
        error: `Bạn cần quyền "Xuất sổ" trong ${targetBook.name} để xuất file Excel.`,
      });
    }
  } else if (
    !isSystemAdmin(currentUser) &&
    !db.books.some((b) => hasBookPermission(currentUser, b.id, 'export'))
  ) {
    return res.status(403).json({
      error: 'Bạn chưa được cấp quyền "Xuất sổ" cho bất kỳ sổ văn bản nào.',
    });
  }

  const queryYear = year && year !== 'all' ? parseInt(year, 10) : getVietnamTimeInfo().year;

  const filterScopeDocs = (baseDocs: DocumentRecord[]) => {
    let scoped = baseDocs.filter((d) => {
      const bId = getDocBookId(d);
      if (bookId && bookId !== 'all' && bId !== bookId) return false;
      if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, bId, 'export')) return false;
      if (year && year !== 'all' && d.year !== queryYear) return false;
      return true;
    });

    if (month && month !== 'all') {
      const m = parseInt(month, 10);
      scoped = scoped.filter((d) => {
        const parts = (d.issueDate || '').split('/');
        return parts.length >= 2 && parseInt(parts[1], 10) === m;
      });
    }
    if (category && category !== 'all') {
      if (category === 'incoming') scoped = scoped.filter((d) => d.docCategory === 'incoming');
      else if (category === 'outgoing') scoped = scoped.filter((d) => d.docCategory !== 'incoming');
    }
    if (typeCode && typeCode !== 'all') {
      scoped = scoped.filter((d) => d.typeCode === typeCode);
    }
    if (signer && signer !== 'all') {
      scoped = scoped.filter((d) => d.signer.toLowerCase() === signer.toLowerCase());
    }
    if (driveStatus && driveStatus !== 'all') {
      scoped = scoped.filter((d) => (d.driveStatus || 'none') === driveStatus);
    }
    if (search && search.trim()) {
      const term = search.trim().toLowerCase();
      scoped = scoped.filter(
        (d) =>
          d.id.toLowerCase().includes(term) ||
          d.fullDocNumber.toLowerCase().includes(term) ||
          (d.docName && d.docName.toLowerCase().includes(term)) ||
          (d.driveFileName && d.driveFileName.toLowerCase().includes(term)) ||
          (d.issuingAgency && d.issuingAgency.toLowerCase().includes(term)) ||
          d.title.toLowerCase().includes(term) ||
          d.signer.toLowerCase().includes(term) ||
          d.issuedBy.toLowerCase().includes(term) ||
          (d.notes && d.notes.toLowerCase().includes(term))
      );
    }
    return scoped;
  };

  let docs = filterScopeDocs(db.documents.filter((d) => d.status !== 'deleted'));
  if (status && status !== 'all' && status !== 'deleted') {
    docs = docs.filter((d) => d.status === status);
  }
  const deletedDocs = filterScopeDocs(db.documents.filter((d) => d.status === 'deleted')).sort(
    (a, b) => new Date(b.deletedAt || b.createdAt).getTime() - new Date(a.deletedAt || a.createdAt).getTime()
  );

  if (sort === 'newest') {
    docs.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  } else if (sort === 'oldest') {
    docs.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
  } else if (sort === 'number_desc') {
    docs.sort((a, b) => b.sequenceNumber - a.sequenceNumber);
  } else {
    docs.sort((a, b) => a.sequenceNumber - b.sequenceNumber);
  }

  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Đoàn Trường THPT Nguyễn Chí Thanh';
    workbook.created = new Date();

    const yearLabel = year === 'all' ? 'TẤT CẢ CÁC NĂM' : `NĂM ${queryYear}`;
    const selectedBookObj = bookId && bookId !== 'all' ? getBookById(bookId) : null;
    const sheetName = year === 'all' ? 'Sổ văn bản tổng hợp' : `Sổ văn bản ${queryYear}`;
    const worksheet = workbook.addWorksheet(sheetName, {
      views: [{ state: 'frozen', ySplit: 3, activeCell: 'A4' }],
      pageSetup: { orientation: 'landscape', paperSize: 9 },
    });

    worksheet.mergeCells('A1:Q1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = selectedBookObj
      ? `${selectedBookObj.name.toUpperCase()} (${yearLabel})`
      : `SỔ QUẢN LÝ VĂN BẢN ${yearLabel} - TRƯỜNG THPT NGUYỄN CHÍ THANH`;
    titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF004085' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).height = 30;

    worksheet.mergeCells('A2:Q2');
    const subCell = worksheet.getCell('A2');
    subCell.value = `Thời điểm xuất sổ: ${formatVietnamDateTime()} | Văn bản hiệu lực/hủy: ${docs.length} | VB đã xóa: ${deletedDocs.length}`;
    subCell.font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF555555' } };
    subCell.alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(2).height = 20;

    worksheet.columns = [
      { key: 'stt', width: 8 },
      { key: 'docId', width: 22 },
      { key: 'bookName', width: 20 },
      { key: 'year', width: 10 },
      { key: 'issueDate', width: 16 },
      { key: 'signDate', width: 16 },
      { key: 'typeName', width: 20 },
      { key: 'sequenceNumber', width: 12 },
      { key: 'fullDocNumber', width: 24 },
      { key: 'docName', width: 35 },
      { key: 'title', width: 40 },
      { key: 'signer', width: 22 },
      { key: 'signerRole', width: 24 },
      { key: 'issuedBy', width: 22 },
      { key: 'status', width: 18 },
      { key: 'notes', width: 28 },
      { key: 'driveFile', width: 38 },
    ];

    const headers = [
      'STT',
      'Mã hồ sơ',
      'Sổ',
      'Năm',
      'Ngày ban hành',
      'Ngày ký',
      'Loại văn bản',
      'Số thứ tự',
      'Số/Ký hiệu văn bản',
      'Tên văn bản',
      'Trích yếu/Nội dung',
      'Người ký',
      'Chức vụ / Cơ quan',
      'Người lấy số / Nhập',
      'Trạng thái',
      'Ghi chú',
      'Tệp đính kèm (Google Drive)',
    ];

    const headerRow = worksheet.getRow(3);
    headerRow.values = headers;
    headerRow.height = 28;

    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF005BAA' },
      };
      cell.font = {
        name: 'Arial',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        bottom: { style: 'medium', color: { argb: 'FF003366' } },
        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
      };
    });

    docs.forEach((doc, idx) => {
      const bookObj = getBookById(getDocBookId(doc));
      const statusText =
        doc.status === 'cancelled'
          ? `Đã hủy (${doc.cancelReason || 'Sai sót'})`
          : doc.docCategory === 'incoming'
          ? 'Đã ghi sổ văn bản đến'
          : 'Đã cấp số';
      const driveText =
        doc.driveStatus === 'saved' && doc.driveFileName
          ? doc.driveFileName
          : doc.driveStatus === 'failed'
          ? 'Lưu tệp thất bại'
          : 'Không đính kèm';

      const cleanIssueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate || '');
      const cleanSignDate = normalizeDateOnlyDdMmYyyy(doc.signDate || '');

      const noteParts: string[] = [];
      if (doc.docCategory === 'incoming' && doc.issuingAgency) {
        noteParts.push(`CQ ban hành: ${doc.issuingAgency}`);
      }
      if (doc.notes) {
        noteParts.push(doc.notes);
      }

      const row = worksheet.addRow({
        stt: idx + 1,
        docId: doc.id,
        bookName: bookObj.name,
        year: doc.year,
        issueDate: cleanIssueDate,
        signDate: cleanSignDate,
        typeName: doc.typeName,
        sequenceNumber: doc.docCategory === 'incoming' ? 'VB đến' : doc.sequenceNumber,
        fullDocNumber: doc.fullDocNumber,
        docName: doc.docName || doc.title,
        title: doc.title || '',
        signer: doc.signer,
        signerRole: doc.signerRole,
        issuedBy: doc.issuedBy,
        status: statusText,
        notes: noteParts.join(' | '),
        driveFile: driveText,
      });

      row.height = 24;

      if (doc.driveStatus === 'saved' && doc.driveFileUrl) {
        const driveCell = row.getCell(17);
        driveCell.value = {
          text: doc.driveFileName || 'Mở tệp',
          hyperlink: doc.driveFileUrl,
        };
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = { name: 'Arial', size: 10 };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } },
        };

        if ([1, 2, 4, 5, 6, 8].includes(colNumber)) {
          cell.numFmt = '@';
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else if (colNumber === 9) {
          cell.numFmt = '@';
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.font = { name: 'Arial', size: 10, bold: true };
        } else if (colNumber === 10 || colNumber === 11) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        } else if (colNumber === 15) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          if (doc.status === 'cancelled') {
            cell.font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FFDC2626' } };
          } else {
            cell.font = { name: 'Arial', size: 10, color: { argb: 'FF16A34A' } };
          }
        } else if (colNumber === 17) {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
          if (doc.driveStatus === 'saved' && doc.driveFileUrl) {
            cell.font = { name: 'Arial', size: 10, underline: true, color: { argb: 'FF0369A1' } };
          } else if (doc.driveStatus === 'failed') {
            cell.font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FFDC2626' } };
          } else {
            cell.font = { name: 'Arial', size: 10, color: { argb: 'FF64748B' } };
          }
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });
    });

    worksheet.autoFilter = `A3:Q${docs.length + 3}`;

    // Sheet thứ 2: "VB đã xóa" (chứa đầy đủ thông tin + Sổ, Người xóa, Thời điểm xóa, Lý do xóa & liên kết Drive)
    const deletedSheet = workbook.addWorksheet('VB đã xóa', {
      views: [{ state: 'frozen', ySplit: 3, activeCell: 'A4' }],
      pageSetup: { orientation: 'landscape', paperSize: 9 },
    });

    deletedSheet.mergeCells('A1:S1');
    const delTitleCell = deletedSheet.getCell('A1');
    delTitleCell.value = selectedBookObj
      ? `DANH SÁCH VĂN BẢN ĐÃ XÓA - ${selectedBookObj.name.toUpperCase()} (${yearLabel})`
      : `DANH SÁCH VĂN BẢN ĐÃ XÓA (${yearLabel})`;
    delTitleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF991B1B' } };
    delTitleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    deletedSheet.getRow(1).height = 30;

    deletedSheet.mergeCells('A2:S2');
    const delSubCell = deletedSheet.getCell('A2');
    delSubCell.value = `Tổng số văn bản đã xóa: ${deletedDocs.length} | Số đã cấp và tệp Google Drive được giữ nguyên để tra cứu đối chiếu`;
    delSubCell.font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF555555' } };
    delSubCell.alignment = { vertical: 'middle', horizontal: 'center' };
    deletedSheet.getRow(2).height = 20;

    deletedSheet.columns = [
      { key: 'stt', width: 8 },
      { key: 'docId', width: 22 },
      { key: 'bookName', width: 22 },
      { key: 'year', width: 10 },
      { key: 'issueDate', width: 16 },
      { key: 'signDate', width: 16 },
      { key: 'typeName', width: 18 },
      { key: 'sequenceNumber', width: 12 },
      { key: 'fullDocNumber', width: 24 },
      { key: 'docName', width: 34 },
      { key: 'title', width: 36 },
      { key: 'signer', width: 20 },
      { key: 'signerRole', width: 22 },
      { key: 'issuedBy', width: 20 },
      { key: 'deletedBy', width: 24 },
      { key: 'deletedAtVietnam', width: 22 },
      { key: 'deleteReason', width: 32 },
      { key: 'notes', width: 24 },
      { key: 'driveFile', width: 38 },
    ];

    const deletedHeaders = [
      'STT',
      'Mã hồ sơ',
      'Sổ',
      'Năm',
      'Ngày ban hành',
      'Ngày ký',
      'Loại văn bản',
      'Số thứ tự',
      'Số/Ký hiệu văn bản',
      'Tên văn bản',
      'Trích yếu/Nội dung',
      'Người ký',
      'Chức vụ / Cơ quan',
      'Người lấy số / Nhập',
      'Người xóa',
      'Thời điểm xóa',
      'Lý do xóa',
      'Ghi chú',
      'Tệp đính kèm (Google Drive)',
    ];

    const delHeaderRow = deletedSheet.getRow(3);
    delHeaderRow.values = deletedHeaders;
    delHeaderRow.height = 28;
    delHeaderRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF991B1B' },
      };
      cell.font = {
        name: 'Arial',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        bottom: { style: 'medium', color: { argb: 'FF7F1D1D' } },
        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
      };
    });

    deletedDocs.forEach((doc, idx) => {
      const bookObj = getBookById(getDocBookId(doc));
      const cleanIssueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate || '');
      const cleanSignDate = normalizeDateOnlyDdMmYyyy(doc.signDate || '');
      const driveText =
        doc.driveStatus === 'saved' && doc.driveFileName
          ? doc.driveFileName
          : doc.driveStatus === 'failed'
          ? 'Lưu tệp thất bại'
          : 'Không đính kèm';

      const delRow = deletedSheet.addRow({
        stt: idx + 1,
        docId: doc.id,
        bookName: `${bookObj.name} (${bookObj.code})`,
        year: doc.year,
        issueDate: cleanIssueDate,
        signDate: cleanSignDate,
        typeName: doc.typeName,
        sequenceNumber: doc.docCategory === 'incoming' ? 'VB đến' : doc.sequenceNumber,
        fullDocNumber: doc.fullDocNumber,
        docName: doc.docName || doc.title,
        title: doc.title || '',
        signer: doc.signer,
        signerRole: doc.signerRole,
        issuedBy: doc.issuedBy,
        deletedBy: doc.deletedBy || '—',
        deletedAtVietnam: doc.deletedAtVietnam || (doc.deletedAt ? formatVietnamDateTime(new Date(doc.deletedAt)) : '—'),
        deleteReason: doc.deleteReason || '—',
        notes: doc.notes || '',
        driveFile: driveText,
      });

      delRow.height = 24;

      if (doc.driveStatus === 'saved' && doc.driveFileUrl) {
        const driveCell = delRow.getCell(19);
        driveCell.value = {
          text: doc.driveFileName || 'Mở tệp',
          hyperlink: doc.driveFileUrl,
        };
      }

      delRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = { name: 'Arial', size: 10 };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } },
        };

        if ([1, 2, 4, 5, 6, 8, 16].includes(colNumber)) {
          cell.numFmt = '@';
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else if (colNumber === 9) {
          cell.numFmt = '@';
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.font = { name: 'Arial', size: 10, bold: true };
        } else if (colNumber === 10 || colNumber === 11 || colNumber === 17) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        } else if (colNumber === 19) {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
          if (doc.driveStatus === 'saved' && doc.driveFileUrl) {
            cell.font = { name: 'Arial', size: 10, underline: true, color: { argb: 'FF0369A1' } };
          } else {
            cell.font = { name: 'Arial', size: 10, color: { argb: 'FF64748B' } };
          }
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });
    });

    deletedSheet.autoFilter = `A3:S${deletedDocs.length + 3}`;

    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = `SO_VAN_BAN_${year === 'all' ? 'TAT_CA' : queryYear}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    res.send(Buffer.from(buffer));
  } catch (err: any) {
    console.error('Lỗi khi xuất Excel:', err);
    res.status(500).json({ error: 'Không thể tạo file Excel: ' + err.message });
  }
});

// 11. Cấu hình & Đồng bộ Microsoft Graph (OneDrive / SharePoint)
app.post('/api/sync/config', requireAuth, requireAdmin, (req, res) => {
  const { enabled, tenantId, clientId, clientSecret, driveId, filePath, tableName } = req.body;

  const nextSecret =
    clientSecret && clientSecret.trim() && clientSecret.trim() !== '••••••••'
      ? clientSecret.trim()
      : db.settings.msGraphConfig.clientSecret;

  db.settings.msGraphConfig = {
    enabled: Boolean(enabled),
    tenantId: tenantId ? tenantId.trim() : '',
    clientId: clientId ? clientId.trim() : '',
    clientSecret: nextSecret,
    driveId: driveId ? driveId.trim() : '',
    filePath: filePath ? filePath.trim() : '',
    tableName: tableName ? tableName.trim() : 'SoCapSoVanBan',
    lastSyncTime: db.settings.msGraphConfig.lastSyncTime,
  };

  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
  res.json({
    msGraphConfig: {
      ...db.settings.msGraphConfig,
      clientSecret: db.settings.msGraphConfig.clientSecret ? '••••••••' : '',
    },
  });
});

app.post('/api/sync/test', requireAuth, requireAdmin, async (req, res) => {
  const config = db.settings.msGraphConfig;
  if (!config.enabled || !config.tenantId || !config.clientId || !config.clientSecret) {
    return res.status(400).json({
      success: false,
      message: 'Chưa điền đủ thông tin xác thực Microsoft Graph (Tenant ID, Client ID, Client Secret).',
    });
  }

  // Thực hiện kiểm tra OAuth token với Microsoft Graph
  try {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', config.clientId);
    params.append('client_secret', config.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    const tokenRes = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    if (!tokenRes.ok) {
      const errData = await tokenRes.json().catch(() => ({}));
      return res.status(400).json({
        success: false,
        message: 'Kết nối Microsoft Entra ID thất bại: ' + (errData.error_description || tokenRes.statusText),
      });
    }

    const tokenData = await tokenRes.json();
    return res.json({
      success: true,
      message: 'Kết nối Microsoft Graph API thành công! Đã xác thực quyền ứng dụng.',
      expiresIn: tokenData.expires_in,
    });
  } catch (err: any) {
    return res.status(500).json({
      success: false,
      message: 'Lỗi khi gọi Microsoft Graph: ' + (err.message || 'Không thể kết nối'),
    });
  }
});

app.post('/api/sync/sync-all', requireAuth, requireAdmin, async (req, res) => {
  const config = db.settings.msGraphConfig;
  if (!config.enabled) {
    return res.status(400).json({
      error: 'Chưa bật tính năng đồng bộ Microsoft Graph trong Cài đặt.',
    });
  }

  // Đánh dấu các bản ghi đang chờ
  const pendingDocs = db.documents.filter((d) => d.syncStatus === 'pending' || d.syncStatus === 'error');
  if (pendingDocs.length === 0) {
    return res.json({ message: 'Tất cả văn bản đã được đồng bộ.', syncedCount: 0 });
  }

  // Thử đồng bộ
  try {
    // Lấy OAuth token
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', config.clientId);
    params.append('client_secret', config.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    const tokenRes = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString(),
    });

    if (!tokenRes.ok) {
      throw new Error('Không thể lấy Token Microsoft Graph');
    }

    // Đánh dấu đồng bộ thành công
    for (const d of pendingDocs) {
      d.syncStatus = 'synced';
      d.syncedAt = new Date().toISOString();
      delete d.syncError;
    }

    db.settings.msGraphConfig.lastSyncTime = new Date().toISOString();
    saveDatabaseSync(db);
    broadcastSSE('sync_completed', { syncedCount: pendingDocs.length });

    res.json({
      success: true,
      syncedCount: pendingDocs.length,
      message: `Đã đồng bộ thành công ${pendingDocs.length} văn bản vào OneDrive/SharePoint!`,
    });
  } catch (err: any) {
    for (const d of pendingDocs) {
      d.syncStatus = 'error';
      d.syncError = err.message || 'Lỗi kết nối Microsoft Graph';
    }
    saveDatabaseSync(db);
    res.status(500).json({
      error: 'Đồng bộ thất bại: ' + err.message,
    });
  }
});

// 12. Cấu hình & Đồng bộ Google Sheets Webhook (Cách đồng bộ dễ nhất cho trường học)
app.post('/api/sync/google-sheets/config', requireAuth, requireAdmin, (req, res) => {
  const { enabled, webhookUrl } = req.body;
  if (!db.settings.googleSheetsConfig) {
    db.settings.googleSheetsConfig = { enabled: false, webhookUrl: '', lastSyncTime: '' };
  }

  db.settings.googleSheetsConfig.enabled = Boolean(enabled);
  db.settings.googleSheetsConfig.webhookUrl = webhookUrl ? webhookUrl.trim() : '';
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
  res.json({ googleSheetsConfig: db.settings.googleSheetsConfig });
});

app.post('/api/sync/google-sheets/test', requireAuth, requireAdmin, async (req, res) => {
  const cfg = db.settings.googleSheetsConfig;
  const targetUrl = req.body.webhookUrl || cfg?.webhookUrl;
  const targetBook = req.body.bookId ? getBookById(req.body.bookId) : getDefaultBook();
  if (!targetUrl || !targetUrl.trim()) {
    return res.status(400).json({ success: false, message: 'Chưa nhập URL Webhook Google Apps Script.' });
  }

  try {
    const testPayload = {
      action: 'test',
      id: `test-sample-${targetBook.code}`,
      bookId: targetBook.id,
      bookCode: targetBook.code,
      bookName: targetBook.name,
      unitCode: targetBook.unitCode,
      stt: 0,
      year: getVietnamTimeInfo().year,
      issueDate: getVietnamTimeInfo().formattedDate,
      signDate: '',
      docCategory: 'outgoing',
      categoryLabel: 'Văn bản đi',
      typeCode: 'TEST',
      typeName: 'Kiểm tra',
      sequenceNumber: 0,
      fullDocNumber: `00-TEST/${targetBook.unitCode}`,
      docName: `Kiểm tra kết nối sổ ${targetBook.name}`,
      title: `Kiểm tra đồng bộ bảng tính theo tab ${targetBook.code}`,
      signer: 'Quản trị viên',
      signerRole: 'Ban quản trị',
      issuedBy: (req as any).user.fullName,
      status: 'Thử nghiệm',
      driveStatus: 'none',
      driveStatusLabel: 'Không đính kèm',
      driveFileName: '',
      driveFileId: '',
      driveFileUrl: '',
      driveFolderPath: '',
      notes: `Dòng kiểm tra kết nối Google Sheets (Tab: ${targetBook.code})`,
      timestamp: new Date().toISOString(),
    };

    const response = await fetch(targetUrl.trim(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(testPayload),
    });

    if (cfg) {
      cfg.lastSyncTime = new Date().toISOString();
      saveDatabaseSync(db);
    }

    res.json({
      success: response.ok,
      message: response.ok
        ? `Đã gửi dữ liệu kiểm tra vào tab "${targetBook.code}" trên Google Sheets thành công!`
        : `Google Apps Script trả về mã ${response.status}. Hãy kiểm tra lại quyền truy cập Web App.`,
    });
  } catch (err: any) {
    res.status(500).json({
      success: false,
      message: 'Lỗi khi gửi đến Google Sheets: ' + (err.message || 'Không thể kết nối'),
    });
  }
});

app.post('/api/sync/google-sheets/sync-all', requireAuth, requireAdmin, async (req, res) => {
  const { bookId } = req.body || {};
  const cfg = db.settings.googleSheetsConfig;
  const targetBook = bookId && bookId !== 'all' ? getBookById(bookId) : null;
  const hasAnyTarget =
    (cfg && cfg.enabled && cfg.webhookUrl) ||
    (targetBook?.googleSheetsConfig?.enabled && targetBook.googleSheetsConfig.webhookUrl);
  if (!hasAnyTarget) {
    return res.status(400).json({ error: 'Chưa bật cấu hình Google Sheets Webhook.' });
  }

  try {
    const filteredDocs = targetBook
      ? db.documents.filter((d) => getDocBookId(d) === targetBook.id)
      : db.documents;
    const sortedDocs = [...filteredDocs].sort((a, b) => a.sequenceNumber - b.sequenceNumber);
    let successCount = 0;

    for (const doc of sortedDocs) {
      await syncToGoogleSheets(doc, doc.status === 'deleted' ? 'delete' : doc.status === 'cancelled' ? 'cancel' : 'create');
      successCount++;
    }

    if (cfg && cfg.enabled) {
      cfg.lastSyncTime = new Date().toISOString();
    }
    saveDatabaseSync(db);

    res.json({
      success: true,
      message: targetBook
        ? `Đã gửi ${successCount} văn bản của sổ "${targetBook.name}" (tab ${targetBook.code}) sang Google Sheets.`
        : `Đã gửi toàn bộ ${successCount} văn bản theo từng tab sổ sang Google Trang tính thành công!`,
    });
  } catch (err: any) {
    res.status(500).json({ error: 'Đồng bộ thất bại: ' + err.message });
  }
});

// Tùy chọn Google Sheets riêng cho từng sổ (để đơn vị tự chia sẻ file riêng ngoài file tổng Admin)
app.post('/api/books/:id/google-sheets', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const targetBook = db.books.find((b) => b.id === req.params.id);
  if (!targetBook) {
    return res.status(404).json({ error: 'Không tìm thấy sổ văn bản.' });
  }
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'manage_book')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Quản trị sổ" trong ${targetBook.name} để cấu hình Google Sheets riêng.`,
    });
  }
  if (checkVersionConflict(req, res, targetBook, currentUser)) return;
  const { enabled, webhookUrl } = req.body || {};
  targetBook.googleSheetsConfig = {
    enabled: Boolean(enabled),
    webhookUrl: typeof webhookUrl === 'string' ? webhookUrl.trim() : '',
    lastSyncTime: targetBook.googleSheetsConfig?.lastSyncTime || '',
  };
  bumpBookVersion(targetBook);
  saveDatabaseSync(db);
  broadcastSSE('books_updated', {
    bookId: targetBook.id,
    configVersion: targetBook.configVersion,
    updatedAt: targetBook.updatedAt,
  });
  res.json({ book: sanitizeBookForUser(targetBook, currentUser) });
});

// 13. Cấu hình & Kiểm tra kết nối Lưu trữ Google Drive (Theo sổ hoặc mặc định hệ thống)
app.post('/api/drive/config', requireAuth, (req, res) => {
  const currentUser = (req as any).user as User;
  const {
    bookId,
    enabled,
    authMode,
    folderUrl,
    appsScriptUrl,
    clearAppsScriptUrl,
    serviceAccountJson,
    clearServiceAccountJson,
    scopeUnitCode,
  } = req.body;

  const targetBook = getBookById(bookId || DEFAULT_BOOK_ID);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'manage_book')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Quản trị sổ" hoặc Quản trị viên để cấu hình thư mục Google Drive cho ${targetBook.name}.`,
    });
  }

  if (checkVersionConflict(req, res, targetBook, currentUser)) return;

  if (!db.settings.googleDriveConfig) {
    db.settings.googleDriveConfig = {
      enabled: false,
      authMode: 'apps_script',
      folderUrl: '',
      folderId: '',
      folderByUnit: {},
    };
  }

  const isDedicatedBook = Boolean(bookId && targetBook.id !== DEFAULT_BOOK_ID);
  const prev: GoogleDriveConfig = isDedicatedBook
    ? {
        enabled: false,
        authMode: db.settings.googleDriveConfig.authMode || 'apps_script',
        folderUrl: '',
        folderId: '',
        ...(targetBook.googleDriveConfig || {}),
      }
    : db.settings.googleDriveConfig;

  const nextAuthMode: 'apps_script' | 'service_account' =
    authMode === 'service_account' ? 'service_account' : 'apps_script';

  const rawFolderInput = typeof folderUrl === 'string' ? folderUrl.trim() : prev.folderUrl || '';
  const parsedFolderId = extractGoogleDriveFolderId(rawFolderInput);

  if (rawFolderInput && !parsedFolderId) {
    return res.status(400).json({
      error: 'URL hoặc ID thư mục Google Drive không hợp lệ. Vui lòng dán liên kết dạng https://drive.google.com/drive/folders/... hoặc ID thư mục.',
    });
  }

  if (Boolean(enabled) && !parsedFolderId) {
    return res.status(400).json({
      error: 'Vui lòng nhập URL hoặc ID thư mục Google Drive trước khi bật tính năng lưu tệp.',
    });
  }

  let nextSaJson = prev.serviceAccountJson || '';
  let nextSaEmail = prev.serviceAccountEmail || '';
  if (clearServiceAccountJson) {
    nextSaJson = '';
    nextSaEmail = '';
  } else if (typeof serviceAccountJson === 'string' && serviceAccountJson.trim()) {
    const parsedSa = parseServiceAccountJson(serviceAccountJson);
    if (parsedSa.error || !parsedSa.parsed) {
      return res.status(400).json({ error: parsedSa.error || 'Tài khoản dịch vụ JSON không hợp lệ.' });
    }
    nextSaJson = serviceAccountJson.trim();
    nextSaEmail = parsedSa.parsed.client_email;
  }

  let nextAppsScriptUrl = prev.appsScriptUrl || '';
  if (clearAppsScriptUrl) {
    nextAppsScriptUrl = '';
  } else if (typeof appsScriptUrl === 'string' && appsScriptUrl.trim()) {
    const trimmedUrl = appsScriptUrl.trim();
    if (!trimmedUrl.startsWith('https://')) {
      return res.status(400).json({
        error: 'URL Apps Script Web App phải bắt đầu bằng https://',
      });
    }
    nextAppsScriptUrl = trimmedUrl;
  }

  if (isDedicatedBook) {
    targetBook.googleDriveConfig = {
      enabled: Boolean(enabled),
      authMode: nextAuthMode,
      folderUrl: rawFolderInput,
      folderId: parsedFolderId,
      folderName: parsedFolderId === targetBook.googleDriveConfig?.folderId ? targetBook.googleDriveConfig?.folderName : undefined,
      appsScriptUrl: nextAppsScriptUrl !== db.settings.googleDriveConfig.appsScriptUrl ? nextAppsScriptUrl : undefined,
      serviceAccountJson: nextSaJson !== db.settings.googleDriveConfig.serviceAccountJson ? nextSaJson : undefined,
      serviceAccountEmail: nextSaEmail,
    };
    bumpBookVersion(targetBook);
    saveDatabaseSync(db);
    broadcastSSE('books_updated', {
      bookId: targetBook.id,
      configVersion: targetBook.configVersion,
      updatedAt: targetBook.updatedAt,
    });
    return res.json({
      googleDriveConfig: sanitizeDriveConfigForClient(
        db.settings.googleDriveConfig,
        true,
        targetBook.unitCode,
        targetBook.googleDriveConfig,
        false
      ),
      book: sanitizeBookForUser(targetBook, currentUser),
    });
  }

  const targetUnit = (typeof scopeUnitCode === 'string' && scopeUnitCode.trim()) || db.settings.unitCode;
  const folderByUnit = { ...(db.settings.googleDriveConfig.folderByUnit || {}) };
  if (targetUnit && parsedFolderId) {
    folderByUnit[targetUnit] = {
      folderUrl: rawFolderInput,
      folderId: parsedFolderId,
      folderName: parsedFolderId === db.settings.googleDriveConfig.folderId ? db.settings.googleDriveConfig.folderName : undefined,
    };
  }

  const updatedConfig: GoogleDriveConfig = {
    ...db.settings.googleDriveConfig,
    enabled: Boolean(enabled),
    authMode: nextAuthMode,
    folderUrl: rawFolderInput,
    folderId: parsedFolderId,
    folderName: parsedFolderId === db.settings.googleDriveConfig.folderId ? db.settings.googleDriveConfig.folderName : undefined,
    folderByUnit,
    appsScriptUrl: nextAppsScriptUrl,
    serviceAccountJson: nextSaJson,
    serviceAccountEmail: nextSaEmail,
  };

  db.settings.googleDriveConfig = updatedConfig;
  bumpBookVersion(targetBook);
  saveDatabaseSync(db);
  broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
  broadcastSSE('books_updated', {
    bookId: targetBook.id,
    configVersion: targetBook.configVersion,
    updatedAt: targetBook.updatedAt,
  });

  res.json({
    googleDriveConfig: sanitizeDriveConfigForClient(db.settings.googleDriveConfig, true, db.settings.unitCode),
    book: sanitizeBookForUser(targetBook, currentUser),
  });
});

app.post('/api/drive/test', requireAuth, async (req, res) => {
  const currentUser = (req as any).user as User;
  const { bookId, authMode, folderUrl, appsScriptUrl, serviceAccountJson } = req.body || {};

  const targetBook = getBookById(bookId || DEFAULT_BOOK_ID);
  if (!isSystemAdmin(currentUser) && !hasBookPermission(currentUser, targetBook.id, 'manage_book')) {
    return res.status(403).json({
      error: `Bạn cần quyền "Quản trị sổ" hoặc Quản trị viên để kiểm tra kết nối Drive cho ${targetBook.name}.`,
    });
  }

  const resolved = resolveDriveForBook(targetBook);

  const testAuthMode: 'apps_script' | 'service_account' =
    authMode === 'service_account' || authMode === 'apps_script' ? authMode : resolved.authMode;
  const testFolderId =
    typeof folderUrl === 'string' && folderUrl.trim()
      ? extractGoogleDriveFolderId(folderUrl.trim())
      : resolved.folderId;
  const testAppsScriptUrl =
    typeof appsScriptUrl === 'string' && appsScriptUrl.trim()
      ? appsScriptUrl.trim()
      : resolved.appsScriptUrl;
  const testSaJson =
    typeof serviceAccountJson === 'string' && serviceAccountJson.trim()
      ? serviceAccountJson.trim()
      : resolved.serviceAccountJson;

  try {
    const result = await testGoogleDriveConnection({
      authMode: testAuthMode,
      folderId: testFolderId,
      appsScriptUrl: testAppsScriptUrl,
      serviceAccountJson: testSaJson,
    });

    const isDedicatedBook = Boolean(bookId && targetBook.id !== DEFAULT_BOOK_ID);
    if (isDedicatedBook) {
      if (!targetBook.googleDriveConfig) targetBook.googleDriveConfig = {};
      targetBook.googleDriveConfig.lastCheckedAt = new Date().toISOString();
      targetBook.googleDriveConfig.lastCheckStatus = result.success ? 'ok' : 'error';
      targetBook.googleDriveConfig.lastCheckMessage = result.message;
      if (result.success && result.folderName && result.folderId === targetBook.googleDriveConfig.folderId) {
        targetBook.googleDriveConfig.folderName = result.folderName;
      }
      saveDatabaseSync(db);
      broadcastSSE('books_updated', { bookId: targetBook.id });
    } else if (db.settings.googleDriveConfig) {
      db.settings.googleDriveConfig.lastCheckedAt = new Date().toISOString();
      db.settings.googleDriveConfig.lastCheckStatus = result.success ? 'ok' : 'error';
      db.settings.googleDriveConfig.lastCheckMessage = result.message;
      if (result.success && result.folderName && result.folderId === db.settings.googleDriveConfig.folderId) {
        db.settings.googleDriveConfig.folderName = result.folderName;
        if (db.settings.googleDriveConfig.folderByUnit?.[db.settings.unitCode]) {
          db.settings.googleDriveConfig.folderByUnit[db.settings.unitCode].folderName = result.folderName;
        }
      }
      saveDatabaseSync(db);
      broadcastSSE('settings_updated', { settings: sanitizeSettingsForUser(db.settings) });
      broadcastSSE('books_updated', { bookId: targetBook.id });
    }

    res.status(result.success ? 200 : 400).json({
      ...result,
      googleDriveConfig: sanitizeDriveConfigForClient(
        db.settings.googleDriveConfig,
        true,
        targetBook.unitCode,
        isDedicatedBook ? targetBook.googleDriveConfig : undefined,
        !isDedicatedBook
      ),
    });
  } catch (err: any) {
    const msg = err.message || 'Không thể kết nối tới Google Drive.';
    const isDedicatedBook = Boolean(bookId && targetBook.id !== DEFAULT_BOOK_ID);
    if (isDedicatedBook) {
      if (!targetBook.googleDriveConfig) targetBook.googleDriveConfig = {};
      targetBook.googleDriveConfig.lastCheckedAt = new Date().toISOString();
      targetBook.googleDriveConfig.lastCheckStatus = 'error';
      targetBook.googleDriveConfig.lastCheckMessage = msg;
      saveDatabaseSync(db);
    } else if (db.settings.googleDriveConfig) {
      db.settings.googleDriveConfig.lastCheckedAt = new Date().toISOString();
      db.settings.googleDriveConfig.lastCheckStatus = 'error';
      db.settings.googleDriveConfig.lastCheckMessage = msg;
      saveDatabaseSync(db);
    }
    res.status(500).json({
      success: false,
      folderId: testFolderId,
      message: msg,
    });
  }
});

// Catch-all JSON 404 for unknown /api/* routes (never fall through to SPA index.html)
app.all('/api/*', (req, res) => {
  res.status(404).json({
    error: `Không tìm thấy tuyến API (${req.method} ${req.originalUrl}).`,
  });
});

// --- Vite Middleware in Dev or Static Serve in Prod ---
async function startServer() {
  if (process.env.NODE_ENV === 'production') {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.resolve(distPath, 'index.html'));
    });
  } else {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[ĐOÀN TRƯỜNG THPT NGUYỄN CHÍ THANH] Máy chủ đang chạy tại cổng ${PORT}`);
  });
}

if (!IS_SERVERLESS) {
  startServer().catch((err) => {
    console.error('Lỗi khởi động máy chủ:', err);
  });
}

export { app };
export default app;
