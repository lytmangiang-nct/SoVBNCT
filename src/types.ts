export type NumberMode = 'separate' | 'shared';

export type DocumentStatus = 'active' | 'cancelled' | 'deleted';

export type SyncStatus = 'not_configured' | 'pending' | 'synced' | 'error';

export type ExcelDeleteSyncStatus = 'not_configured' | 'pending' | 'synced' | 'failed';

export type BookPermission =
  | 'view'
  | 'import_incoming'
  | 'request_number'
  | 'issue_number'
  | 'edit'
  | 'replace_file'
  | 'cancel'
  | 'delete'
  | 'export'
  | 'manage_book';

export interface DocumentTypeConfig {
  id: string;
  name: string;
  code: string; // e.g. 'KH', 'BC', 'ĐN', 'ND', 'TB'
  color?: string; // e.g. 'blue', 'emerald', 'amber', 'red', 'purple', 'teal'
  isActive: boolean;
}

export interface SignerConfig {
  id: string;
  name: string;
  role: string;
}

export interface MsGraphConfig {
  enabled: boolean;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  driveId: string;
  filePath: string;
  tableName: string;
  lastSyncTime?: string;
}

export interface GoogleSheetsConfig {
  enabled: boolean;
  webhookUrl: string; // URL Google Apps Script Web App (https://script.google.com/macros/s/.../exec)
  lastSyncTime?: string;
}

export type DriveUploadStatus = 'none' | 'uploading' | 'saved' | 'failed';

export type GoogleDriveAuthMode = 'apps_script' | 'service_account';

export interface GoogleDriveFolderScope {
  folderUrl: string;
  folderId: string;
  folderName?: string;
}

export interface GoogleDriveConfig {
  enabled: boolean;
  authMode: GoogleDriveAuthMode;
  folderUrl: string;
  folderId: string;
  folderName?: string;
  folderByUnit?: Record<string, GoogleDriveFolderScope>;
  // Server-side only secrets (omitted/masked in API responses)
  appsScriptUrl?: string;
  serviceAccountJson?: string;
  // Safe status fields returned to client
  hasAppsScriptUrl?: boolean;
  appsScriptUrlMasked?: string;
  hasServiceAccount?: boolean;
  serviceAccountEmail?: string;
  isConfigured?: boolean;
  maxFileSizeMB?: number;
  maxFileSizeBytes?: number;
  lastCheckedAt?: string;
  lastCheckStatus?: 'ok' | 'error';
  lastCheckMessage?: string;
}

export interface RegistryBook {
  id: string;
  code: string; // Mã sổ nội bộ cố định (dùng làm tên tab Google Sheets, vd: DOAN_TRUONG, CHI_BO, NHA_TRUONG)
  name: string;
  unitName: string;
  unitCode: string;
  isShared: boolean; // true = Sổ dùng chung toàn trường, false = Sổ riêng theo bộ phận/đơn vị
  numberMode: NumberMode;
  numberFormat: string;
  customStartSequences: Record<string, number>;
  enabledTypeCodes?: string[]; // Danh sách mã loại văn bản được phép dùng trong sổ này
  signers?: SignerConfig[];
  googleDriveConfig?: Partial<GoogleDriveConfig>;
  googleSheetsConfig?: GoogleSheetsConfig; // Tùy chọn tệp Google Sheets riêng cho bộ phận nếu cần chia sẻ trực tiếp
  isActive: boolean;
  createdAt: string;
  configVersion?: number;
  updatedAt?: string;
}

export interface SystemSettings {
  unitName: string; // "Đoàn Trường THPT Nguyễn Chí Thanh"
  unitCode: string; // "ĐTN"
  numberMode: NumberMode; // "separate" (mỗi loại dãy riêng) hoặc "shared" (dãy chung)
  numberFormat: string; // Mẫu số: default "{so}-{loai}/{don_vi}"
  customStartSequences: Record<string, number>; // key: `${year}_${typeCode}` hoặc `${year}_SHARED`
  msGraphConfig: MsGraphConfig;
  googleSheetsConfig?: GoogleSheetsConfig;
  googleDriveConfig?: GoogleDriveConfig;
  configVersion?: number;
  updatedAt?: string;
}

export type DocumentCategory = 'outgoing' | 'incoming';

export type DocumentAuditAction =
  | 'create'
  | 'update'
  | 'cancel'
  | 'delete'
  | 'retry_excel_sync'
  | 'upload_word'
  | 'import_pdf'
  | 'upload_pdf'
  | 'request_number'
  | 'replace_file';

export interface AuditLog {
  id?: string;
  timestamp: string;
  timestampVietnam?: string;
  dateKeyVietnam?: string;
  userId: string;
  username?: string;
  userName: string;
  bookId?: string;
  bookCode?: string;
  bookName?: string;
  docId?: string;
  fullDocNumber?: string;
  docName?: string;
  action: DocumentAuditAction;
  changes: Record<string, { old: string | null; new: string | null }>;
  reason?: string;
  snapshot?: Record<string, any>;
}

export interface DocumentHistoryEntry extends AuditLog {
  id: string;
  timestampVietnam: string;
  dateKeyVietnam: string;
  username: string;
  bookId: string;
  bookCode: string;
  bookName: string;
  docId: string;
  fullDocNumber: string;
  docName: string;
}

export interface SystemAuditLog {
  id: string;
  timestamp: string;
  actorUserId: string;
  actorName: string;
  action:
    | 'permission_change'
    | 'user_create'
    | 'user_update'
    | 'user_delete'
    | 'book_create'
    | 'book_update';
  bookId?: string;
  bookCode?: string;
  targetUserId?: string;
  targetUsername?: string;
  details: string;
  changes?: Record<string, { old: string | null; new: string | null }>;
}

export interface DocumentRecord {
  id: string;
  bookId?: string;
  bookCode?: string;
  bookName?: string;
  requestId?: string;
  stt: number;
  year: number;
  issueDate: string; // dd/mm/yyyy hoặc '' nếu văn bản đến không xác định được ngày ban hành
  signDate?: string; // Ngày ký riêng (nếu xác định được tách biệt với ngày ban hành)
  createdAt: string; // ISO
  docCategory?: DocumentCategory; // 'outgoing' (Đoàn trường ban hành) | 'incoming' (Văn bản đến PDF)
  issuingAgency?: string; // Cơ quan ban hành gốc (đối với văn bản đến)
  typeCode: string;
  typeName: string;
  sequenceNumber: number;
  fullDocNumber: string; // e.g. "01-KH/ĐTN" hoặc "1234/SGDĐT-GDTrH"
  docName?: string; // Tên văn bản
  title: string; // Trích yếu / Nội dung văn bản
  signer: string;
  signerRole: string;
  issuedBy: string; // Tên hiển thị người lấy số / người nhập
  issuedByUsername: string;
  issuedByUserId: string;
  status: DocumentStatus;
  cancelReason?: string;
  cancelledAt?: string;
  cancelledBy?: string;
  deleteReason?: string;
  deletedAt?: string;
  deletedAtVietnam?: string;
  deletedBy?: string;
  deletedByUsername?: string;
  deletedByUserId?: string;
  deletedSnapshot?: Record<string, any>;
  excelDeleteSyncStatus?: ExcelDeleteSyncStatus;
  excelMovedToDeletedSheet?: boolean;
  excelRemovedFromMainSheet?: boolean;
  excelSyncError?: string;
  excelSyncedAt?: string;
  excelSyncTargets?: Record<
    string,
    {
      movedToDeletedSheet: boolean;
      removedFromMainSheet: boolean;
      syncedAt?: string;
      error?: string;
    }
  >;
  notes?: string;
  auditLogs: AuditLog[];
  syncStatus: SyncStatus;
  syncError?: string;
  syncedAt?: string;
  idempotencyKey: string;
  // Attachment / Google Drive fields (Word or PDF)
  driveStatus?: DriveUploadStatus;
  driveFileId?: string;
  driveFileName?: string;
  driveFileUrl?: string;
  driveOriginalFileName?: string;
  driveMimeType?: string;
  driveFileSize?: number;
  driveFolderId?: string;
  driveFolderPath?: string;
  driveUploadedAt?: string;
  driveError?: string;
  // Incoming PDF provenance & review flags
  sourceType?: 'pdf_url' | 'pdf_upload';
  sourceUrl?: string;
  fileSha256?: string;
  needsReview?: boolean;
  reviewReasons?: string[];
  relevantPages?: number[];
}

export type NumberRequestStatus = 'pending' | 'issued' | 'rejected';

export interface NumberRequest {
  id: string;
  bookId: string;
  bookName?: string;
  typeCode: string;
  typeName: string;
  docName: string;
  title: string;
  signer: string;
  signerRole: string;
  notes?: string;
  requestedBy: string;
  requestedByUsername: string;
  requestedByUserId: string;
  requestedAt: string;
  status: NumberRequestStatus;
  issuedDocId?: string;
  issuedDocNumber?: string;
  reviewedBy?: string;
  reviewedAt?: string;
  rejectReason?: string;
}

export interface NormalizedBBox {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export type ExtractableFieldKey =
  | 'docNumber'
  | 'issuingAgency'
  | 'docName'
  | 'title'
  | 'issueDate'
  | 'signDate'
  | 'signer'
  | 'signerRole';

export interface FieldLocation {
  page: number;
  bbox?: NormalizedBBox;
  estimated?: boolean;
}

export interface PdfExtractionResult {
  docNumber: string;
  docName: string;
  title: string;
  issueDate: string;
  signDate: string;
  signer: string;
  signerRole: string;
  issuingAgency: string;
  needsReview: boolean;
  reviewReasons: string[];
  relevantPages: number[];
  usedOcr: boolean;
  ocrConfidence?: number;
  pageCount: number;
  signerCandidates?: string[];
  fieldLocations?: Partial<Record<ExtractableFieldKey, FieldLocation>>;
}

export interface PdfAnalyzeResponse {
  stageToken: string;
  bookId?: string;
  sha256: string;
  fileSize: number;
  originalFileName: string;
  sourceType: 'pdf_url' | 'pdf_upload';
  sourceUrl?: string;
  extraction: PdfExtractionResult;
  duplicateDocument: DocumentRecord | null;
  suspectedDuplicateDocument?: DocumentRecord | null;
  filenameConflictDocument?: DocumentRecord | null;
}

export interface User {
  id: string;
  username: string;
  fullName: string;
  role: 'admin' | 'issuer';
  isActive: boolean;
  createdAt: string;
  bookPermissions?: Record<string, BookPermission[]>;
}

export interface AuthResponse {
  user: User;
  token: string;
  rememberMe?: boolean;
  expiresAt?: number;
}

export interface StatisticsData {
  year: number;
  bookId?: string;
  totalIssued: number;
  activeCount: number;
  cancelledCount: number;
  pendingSyncCount: number;
  byType: Array<{
    typeCode: string;
    typeName: string;
    count: number;
    cancelledCount: number;
  }>;
  byMonth: Array<{
    month: number;
    count: number;
    cancelledCount: number;
  }>;
}

export type StorageMode =
  | 'upstash_kv'
  | 'vercel_blob'
  | 'remote_webhook'
  | 'local_disk'
  | 'ephemeral_tmp';

export interface PersistenceStatus {
  durable: boolean;
  storageMode: StorageMode;
  storageLabel: string;
  requiresRemoteEnvOnVercel?: boolean;
  lastSavedAt?: string;
  lastVerifiedAt?: string;
  warning?: string;
}

export interface PreviewNumberResponse {
  bookId?: string;
  bookCode?: string;
  typeCode?: string;
  seriesKey?: string;
  storageKey?: string;
  previewNumber: string;
  sequenceNumber: number;
  year: number;
  serverYear?: number;
  formattedDate?: string;
  mode: NumberMode;
  customStartSequence?: number | null;
  maxIssuedSequence?: number;
  configVersion?: number;
  serverTimeMs?: number;
  persistenceStatus?: PersistenceStatus;
}

export interface StartSequenceSeriesConfig {
  seriesKey: string;
  storageKey: string;
  bookId: string;
  bookCode: string;
  year: number;
  typeCode: string;
  isConfigured: boolean;
  customStartSequence: number | null;
  maxIssuedSequence: number;
  minAllowedStartSequence: number;
  nextExpectedSeq: number;
  previewNumber: string;
  configVersion: number;
  serverYear?: number;
  customStartSequences?: Record<string, number>;
  persistenceStatus?: PersistenceStatus;
  legacyReconciliationWarnings?: string[];
  verifiedFromPrimary?: boolean;
}

