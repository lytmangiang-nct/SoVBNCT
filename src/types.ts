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
  | 'import_archive'
  | 'close_book'
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

/** @deprecated Giữ kiểu tùy chọn để đọc dữ liệu sao lưu cũ có trường msGraphConfig */
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
  msGraphConfig?: MsGraphConfig;
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
  | 'replace_file'
  | 'template_manage'
  | 'word_generate'
  | 'attachment_add'
  | 'attachment_replace'
  | 'attachment_set_official'
  | 'task_create'
  | 'task_update'
  | 'reconcile_sync'
  | 'confirm_date'
  | 'archive_import'
  | 'book_close'
  | 'book_amendment';

export type AttachmentGroup = 'word_draft' | 'signed_pdf' | 'appendix';

export interface DocumentAttachmentVersion {
  version: number;
  fileName: string;
  originalFileName?: string;
  mimeType?: string;
  fileSize?: number;
  sha256?: string;
  driveFileId?: string;
  driveFileUrl?: string;
  driveFolderId?: string;
  driveFolderPath?: string;
  driveStatus: DriveUploadStatus;
  driveError?: string;
  uploadedBy: string;
  uploadedByUsername?: string;
  uploadedAt: string;
}

export interface DocumentAttachment {
  id: string;
  group: AttachmentGroup;
  label?: string;
  isOfficial: boolean;
  currentVersion: number;
  fileName: string;
  originalFileName?: string;
  mimeType?: string;
  fileSize?: number;
  sha256?: string;
  driveFileId?: string;
  driveFileUrl?: string;
  driveFolderId?: string;
  driveFolderPath?: string;
  driveStatus: DriveUploadStatus;
  driveError?: string;
  uploadedBy?: string;
  uploadedByUsername?: string;
  uploadedAt?: string;
  versions: DocumentAttachmentVersion[];
}

export interface WordTemplate {
  id: string;
  bookId: string;
  bookCode?: string;
  typeCode: string;
  name: string;
  originalFileName: string;
  fileSize: number;
  base64Docx: string;
  placeholders: string[];
  createdBy: string;
  createdByUserId: string;
  createdAt: string;
  updatedAt: string;
  updatedBy?: string;
}

export interface PdfPageContent {
  page: number;
  text: string;
}

export interface PdfTaskSuggestion {
  id: string;
  taskTitle: string;
  dueDate: string; // dd/MM/yyyy hoặc '' nếu là hạn tương đối cần xác nhận
  rawDeadlineText?: string;
  needsConfirmation: boolean;
  recipient?: string;
  sourceExcerpt: string;
  page: number;
}

export type DocumentTaskStatus = 'pending' | 'in_progress' | 'done';

export interface DocumentTask {
  id: string;
  bookId: string;
  bookCode?: string;
  bookName?: string;
  docId: string;
  fullDocNumber: string;
  docName: string;
  suggestionId?: string;
  title: string;
  dueDate: string;
  rawDeadlineText?: string;
  needsConfirmation: boolean;
  recipient?: string;
  sourceExcerpt?: string;
  sourcePage?: number;
  assigneeUserId: string;
  assigneeName: string;
  assigneeUsername: string;
  status: DocumentTaskStatus;
  createdByUserId: string;
  createdByName: string;
  createdAt: string;
  updatedAt: string;
}

export interface PdfContentSearchMatch {
  docId: string;
  bookId: string;
  bookCode?: string;
  bookName?: string;
  fullDocNumber: string;
  docName: string;
  docCategory?: DocumentCategory;
  issueDate?: string;
  year: number;
  typeCode: string;
  status: DocumentStatus;
  page: number;
  snippet: string;
  attachmentId?: string;
  pdfTextStatus: 'indexed' | 'no_text';
}

export type DriveVerificationState =
  | 'no_attachment'
  | 'unsaved'
  | 'not_found'
  | 'access_denied'
  | 'unverified'
  | 'ok';

export type ReconciliationIssueType =
  | 'drive_failed'
  | 'drive_link_issue'
  | 'sheets_unsynced'
  | 'duplicate_number';

export interface ReconciliationItem {
  id: string;
  issueType: ReconciliationIssueType;
  severity: 'error' | 'warning';
  bookId: string;
  bookCode: string;
  bookName: string;
  year: number;
  docId: string;
  fullDocNumber: string;
  docName: string;
  status: DocumentStatus;
  driveState: DriveVerificationState;
  driveStateLabel: string;
  driveFileId?: string;
  driveFileUrl?: string;
  driveError?: string;
  sheetsSyncState: string;
  sheetsSyncError?: string;
  duplicateDocIds?: string[];
  detailMessage: string;
  canRetry: boolean;
  retryKind?: 'sheets' | 'word_drive';
}

export interface ReconciliationReport {
  checkedAt: string;
  bookId: string;
  year: string;
  driveChecked: boolean;
  summary: {
    totalDocumentsScanned: number;
    totalIssues: number;
    driveFailedCount: number;
    driveLinkIssueCount: number;
    sheetsUnsyncedCount: number;
    duplicateNumberCount: number;
  };
  items: ReconciliationItem[];
}

export interface PreIssueCheckResult {
  bookId: string;
  bookName: string;
  year: number;
  formattedDate: string;
  previewNumber: string;
  sequenceNumber: number;
  nearDuplicateDocs: Array<{
    id: string;
    fullDocNumber: string;
    docName: string;
    issueDate: string;
    signer: string;
    status: DocumentStatus;
    canView: boolean;
  }>;
  sameFileDocs: Array<{
    id: string;
    fullDocNumber: string;
    docName: string;
    issueDate: string;
    driveFileName?: string;
    status: DocumentStatus;
    canView: boolean;
  }>;
}

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
    | 'book_update'
    | 'history_cleanup'
    | 'deleted_list_archive'
    | 'archive_import'
    | 'book_close'
    | 'book_amendment';
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
  receivedAt?: string; // ISO thời điểm tiếp nhận vào hệ thống
  receivedDate?: string; // dd/mm/yyyy ngày tiếp nhận theo giờ Việt Nam
  createdAt: string; // ISO
  docCategory?: DocumentCategory; // 'outgoing' (Đoàn trường ban hành) | 'incoming' (Văn bản đến PDF)
  parentAgency?: string; // Cơ quan chủ quản cấp trên (nếu có)
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
  deletedListArchived?: boolean;
  deletedArchivedAt?: string;
  deletedArchivedBy?: string;
  deletedArchivedByUserId?: string;
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
  // Incoming PDF provenance, date confirmation & review flags
  sourceType?: 'pdf_url' | 'pdf_upload';
  sourceUrl?: string;
  fileSha256?: string;
  needsReview?: boolean;
  reviewReasons?: string[];
  relevantPages?: number[];
  pendingDateConfirmation?: boolean;
  dateComparisonStatus?: DateComparisonStatus;
  dateMismatchSources?: {
    topRight?: ExtractedDateSource;
    signature?: ExtractedDateSource;
  };
  dateConfirmedBy?: string;
  dateConfirmedByUserId?: string;
  dateConfirmedAt?: string;
  // Multi-attachment & versioning
  attachments?: DocumentAttachment[];
  // Auto Word generation from template
  templateId?: string;
  wordGenerationStatus?: 'none' | 'generated' | 'failed';
  wordGenerationError?: string;
  // Indexed PDF page contents (and optional legacy task suggestions)
  pdfPageContents?: PdfPageContent[];
  pdfContentVersion?: number;
  pdfTextStatus?: 'indexed' | 'no_text';
  pdfIndexedPagesCount?: number;
  pdfTaskSuggestions?: PdfTaskSuggestion[];
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

export type DateComparisonStatus = 'single' | 'matched' | 'mismatch' | 'missing';

export interface ExtractedDateSource {
  date: string; // dd/MM/yyyy
  positionLabel: string;
  page: number;
  bbox?: NormalizedBBox;
  estimated?: boolean;
  rawText?: string;
}

export type ExtractableFieldKey =
  | 'docNumber'
  | 'parentAgency'
  | 'issuingAgency'
  | 'docName'
  | 'title'
  | 'documentDate'
  | 'issueDate'
  | 'signDate'
  | 'topRightDate'
  | 'signatureDate'
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
  documentDate?: string;
  issueDate: string;
  signDate: string;
  dateComparisonStatus?: DateComparisonStatus;
  topRightDate?: ExtractedDateSource | null;
  signatureDate?: ExtractedDateSource | null;
  signer: string;
  signerRole: string;
  parentAgency?: string;
  issuingAgency: string;
  needsReview: boolean;
  reviewReasons: string[];
  relevantPages: number[];
  usedOcr: boolean;
  aiUsed?: boolean;
  ocrConfidence?: number;
  pageCount: number;
  signerCandidates?: string[];
  fieldLocations?: Partial<Record<ExtractableFieldKey, FieldLocation>>;
  taskSuggestions?: PdfTaskSuggestion[];
  indexedPageCount?: number;
  docNumberNeedsConfirmation?: boolean;
  docNumberReviewReason?: string;
  filenameDocNumberHint?: string;
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

export interface IssuedUploadTask {
  docId: string;
  bookId: string;
  bookName: string;
  typeCode: string;
  fullDocNumber: string;
  docName: string;
  driveFileName?: string;
  driveOriginalFileName?: string;
  hasFileOrTemplate: boolean;
  templateId?: string;
  driveStatus: DriveUploadStatus;
  driveFileId?: string;
  driveFileUrl?: string;
  driveError?: string;
  issuedAtMs: number;
  issueMs?: number;
  driveCompletedAtMs?: number;
}

export interface User {
  id: string;
  username: string;
  fullName: string;
  displayName?: string;
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

export type StorageConnectionState =
  | 'connected'
  | 'missing_config'
  | 'conflict'
  | 'connection_error'
  | 'awaiting_migration'
  | 'corrupt_data';

export interface LegacySnapshotSummary {
  usersCount: number;
  booksCount: number;
  documentsCount: number;
  deletedDocumentsCount: number;
  startMilestonesCount: number;
  driveLinkedCount: number;
  bookNames: string[];
}

export interface PersistenceStatus {
  durable: boolean;
  storageMode: StorageMode;
  storageLabel: string;
  connectionState: StorageConnectionState;
  connectionLabel: string;
  envPairUsed?:
    | 'UPSTASH_REDIS_REST'
    | 'KV_REST_API'
    | 'VERCEL_BLOB'
    | 'REMOTE_WEBHOOK'
    | 'LOCAL_DISK'
    | 'NONE';
  dbKey?: string;
  maskedHost?: string;
  schemaVersion?: number;
  hasLegacyBundledSnapshot?: boolean;
  legacySnapshotSummary?: LegacySnapshotSummary;
  sessionSecretConfigured?: boolean;
  requiresRemoteEnvOnVercel?: boolean;
  lastSavedAt?: string;
  lastVerifiedAt?: string;
  warning?: string;
  errorDetail?: string;
}

export interface StorageTestResult {
  success: boolean;
  readOk: boolean;
  writeOk: boolean;
  deleteOk: boolean;
  keyExists: boolean;
  schemaValid?: boolean;
  latencyMs: number;
  probeKey?: string;
  message: string;
  persistenceStatus: PersistenceStatus;
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

export type HistoryCleanupScope = 'today' | 'last_10_days' | 'before_date' | 'all';

export interface HistoryCleanupPreviewResponse {
  bookId: string;
  bookCode: string;
  bookName: string;
  scope: HistoryCleanupScope;
  beforeDate?: string;
  rangeLabel: string;
  fromDateKey?: string | null;
  toDateKey?: string | null;
  matchCount: number;
  requiresTypedConfirmation: boolean;
  confirmationPhrase: string;
}

export interface HistoryCleanupResultResponse extends HistoryCleanupPreviewResponse {
  success: boolean;
  deletedCount: number;
  remainingCount: number;
  message: string;
}

export interface BulkSyncDeletedResult {
  success: boolean;
  bookId: string;
  bookName: string;
  totalTargeted: number;
  succeededCount: number;
  failedCount: number;
  succeededIds: string[];
  failedItems: Array<{
    id: string;
    fullDocNumber: string;
    docName: string;
    error: string;
  }>;
  documents: DocumentRecord[];
  message: string;
}

export interface ArchiveDeletedResult {
  success: boolean;
  bookId: string;
  bookName: string;
  archivedCount: number;
  skippedUnsyncedCount: number;
  archivedIds: string[];
  documents: DocumentRecord[];
  message: string;
}

// ==========================================
// 1. CHAT TRA CỨU CÓ DẪN NGUỒN (GROUNDED Q&A)
// ==========================================
export type DocumentQaCitationSourceType = 'pdf_page' | 'record_metadata';

export interface DocumentQaCitation {
  docId: string;
  bookId: string;
  bookCode?: string;
  bookName: string;
  docCategory?: DocumentCategory;
  typeCode: string;
  typeName: string;
  fullDocNumber: string;
  docName: string;
  issueDate?: string;
  issuingAgency?: string;
  sourceType: DocumentQaCitationSourceType;
  page?: number;
  excerpt: string;
  attachmentId?: string;
  hasPdfFile?: boolean;
  hasDriveFile?: boolean;
}

export interface DocumentQaResponse {
  question: string;
  answer: string;
  insufficientEvidence: boolean;
  mismatchNote?: string;
  citations: DocumentQaCitation[];
  searchedBooks: Array<{ id: string; code: string; name: string }>;
  usedAi: boolean;
  indexedPdfCount: number;
  totalSearchedDocs: number;
}

// ==========================================
// 2. NHẬP KHO VĂN BẢN CŨ TỪ EXCEL & DRIVE
// ==========================================
export type ArchiveImportRowStatus =
  | 'ready'
  | 'possible_duplicate'
  | 'missing_info'
  | 'unmatched_file'
  | 'ambiguous_file';

export interface ArchiveDriveCandidate {
  fileId: string;
  fileName: string;
  webViewLink: string;
  mimeType?: string;
  folderPath?: string;
}

export interface ArchiveImportColumnMapping {
  docCategory?: string;
  fullDocNumber?: string;
  docName?: string;
  issueDate?: string;
  year?: string;
  issuingAgency?: string;
  signer?: string;
  signerRole?: string;
  typeCode?: string;
  parentAgency?: string;
  title?: string;
  notes?: string;
  driveLinkOrFileName?: string;
}

export interface ArchiveImportEditableRow {
  rowNumber: number;
  docCategory: 'outgoing' | 'incoming';
  fullDocNumber: string;
  docName: string;
  issueDate: string;
  year: string;
  issuingAgency: string;
  signer: string;
  signerRole?: string;
  typeCode?: string;
  parentAgency?: string;
  title?: string;
  notes?: string;
  driveLinkOrFileName: string;
  driveHyperlink?: string;
  selectedDriveFileId?: string;
}

export interface ArchiveImportPreviewRow extends ArchiveImportEditableRow {
  rowId: string;
  rowFingerprint: string;
  status: ArchiveImportRowStatus;
  errors: string[];
  warnings: string[];
  yearMismatchWarning?: string;
  alreadyCommittedDocId?: string;
  duplicateMatches: Array<{
    docId: string;
    fullDocNumber: string;
    docName: string;
    issueDate: string;
    issuingAgency?: string;
    status: DocumentStatus;
  }>;
  batchDuplicateRows: number[];
  driveMatchStatus:
    | 'matched_by_id'
    | 'matched_exact_name'
    | 'matched_suggested_name'
    | 'ambiguous'
    | 'not_found'
    | 'access_denied'
    | 'none';
  matchedDriveFile?: ArchiveDriveCandidate;
  driveCandidates: ArchiveDriveCandidate[];
  resolvedTypeCode?: string;
  parsedSequenceNumber?: number;
  seriesUpdateInfo?: string;
}

export interface ArchiveImportPreviewResponse {
  bookId: string;
  bookName: string;
  bookCode: string;
  folderId: string;
  folderSource: 'custom' | 'book' | 'none';
  driveScanError?: string;
  driveFileCount: number;
  summary: {
    total: number;
    ready: number;
    possibleDuplicate: number;
    missingInfo: number;
    unmatchedFile: number;
    ambiguousFile: number;
  };
  rows: ArchiveImportPreviewRow[];
}

export interface ArchiveParseExcelResponse {
  fileName: string;
  sheets: Array<{
    name: string;
    rowCount: number;
    headers: string[];
    suggestedMapping: ArchiveImportColumnMapping;
    rows: Array<Record<string, { value: string; hyperlink?: string }>>;
  }>;
}

export interface ArchiveImportCommitRowResult {
  rowId: string;
  rowNumber: number;
  fullDocNumber: string;
  docName: string;
  status: 'imported' | 'skipped_duplicate' | 'error';
  docId?: string;
  message: string;
}

export interface ArchiveImportCommitResponse {
  batchId: string;
  totalRequested: number;
  importedCount: number;
  skippedDuplicateCount: number;
  errorCount: number;
  results: ArchiveImportCommitRowResult[];
  createdDocuments: DocumentRecord[];
  bookUpdated?: RegistryBook;
}

// ==========================================
// 3. CHỐT SỔ & BÀN GIAO CUỐI NĂM
// ==========================================
export type ClosingMetricKey =
  | 'outgoing'
  | 'incoming'
  | 'cancelled'
  | 'missingFile'
  | 'unsyncedSheets';

export interface ClosingSummaryCounts {
  totalActiveAndCancelled: number;
  outgoing: number;
  incoming: number;
  cancelled: number;
  missingFile: number;
  unsyncedSheets: number;
  pendingDateConfirmation: number;
}

export interface ClosingSnapshotDocument {
  docId: string;
  bookId: string;
  bookCode: string;
  bookName: string;
  docCategory: DocumentCategory;
  typeCode: string;
  typeName: string;
  sequenceNumber: number;
  fullDocNumber: string;
  docName: string;
  title: string;
  issueDate: string;
  signDate?: string;
  year: number;
  issuingAgency?: string;
  parentAgency?: string;
  signer: string;
  signerRole: string;
  issuedBy: string;
  status: DocumentStatus;
  cancelReason?: string;
  driveStatus: DriveUploadStatus;
  driveFileId?: string;
  driveFileName?: string;
  driveFileUrl?: string;
  syncStatus: SyncStatus;
  pendingDateConfirmation?: boolean;
  notes?: string;
  createdAt: string;
}

export interface ClosingAmendmentChange {
  docId?: string;
  fullDocNumber?: string;
  fieldLabel: string;
  beforeValue: string;
  afterValue: string;
}

export interface ClosingAmendmentRecord {
  id: string;
  reason: string;
  amendedBy: string;
  amendedByUserId: string;
  amendedByUsername: string;
  amendedAt: string;
  amendedAtVietnam: string;
  changes: ClosingAmendmentChange[];
}

export interface YearEndClosingSnapshot {
  id: string;
  closingCode: string;
  bookId: string;
  bookCode: string;
  bookName: string;
  year: number;
  closedBy: string;
  closedByUserId: string;
  closedByUsername: string;
  closedAt: string;
  closedAtVietnam: string;
  note?: string;
  summary: ClosingSummaryCounts;
  documents: ClosingSnapshotDocument[];
  amendments: ClosingAmendmentRecord[];
}

export interface ClosingPreviewResponse {
  bookId: string;
  bookCode: string;
  bookName: string;
  year: number;
  summary: ClosingSummaryCounts;
  documents: ClosingSnapshotDocument[];
  existingSnapshotsCount: number;
}



