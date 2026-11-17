import type {
  AuthResponse,
  BookPermission,
  DocumentHistoryEntry,
  DocumentRecord,
  DocumentTypeConfig,
  GoogleDriveConfig,
  LegacySnapshotSummary,
  NumberRequest,
  PdfAnalyzeResponse,
  PersistenceStatus,
  PreviewNumberResponse,
  RegistryBook,
  SignerConfig,
  StartSequenceSeriesConfig,
  StatisticsData,
  StorageTestResult,
  SystemAuditLog,
  SystemSettings,
  User,
} from './types';

const TOKEN_KEY = 'nct_doc_token';
const USER_KEY = 'nct_doc_user';
const RAW_API_BASE = (import.meta.env.VITE_API_BASE_URL || '').trim().replace(/\/+$/, '');

export function buildApiUrl(endpoint: string): string {
  if (!RAW_API_BASE) return endpoint;
  if (/^https?:\/\//i.test(endpoint)) return endpoint;
  return `${RAW_API_BASE}${endpoint.startsWith('/') ? endpoint : `/${endpoint}`}`;
}

export function getStoredToken(): string | null {
  try {
    return sessionStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

export function getStoredUser(): User | null {
  try {
    const data = sessionStorage.getItem(USER_KEY);
    if (!data) return null;
    return JSON.parse(data);
  } catch {
    return null;
  }
}

export function saveAuth(response: AuthResponse) {
  try {
    // Không lưu thông tin xác thực dài hạn vào localStorage; phiên 30 ngày dùng cookie HttpOnly phía máy chủ
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(USER_KEY);
    sessionStorage.setItem(TOKEN_KEY, response.token);
    sessionStorage.setItem(USER_KEY, JSON.stringify(response.user));
  } catch {
    // ignore storage errors
  }
}

export function clearAuth() {
  try {
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(USER_KEY);
    sessionStorage.removeItem(TOKEN_KEY);
    sessionStorage.removeItem(USER_KEY);
  } catch {
    // ignore storage errors
  }
}

type ApiRequestOptions = RequestInit & { timeoutMs?: number };

async function request<T>(endpoint: string, options: ApiRequestOptions = {}): Promise<T> {
  const token = getStoredToken();
  const headers = new Headers(options.headers || {});

  if (!headers.has('Content-Type') && !(options.body instanceof FormData)) {
    headers.set('Content-Type', 'application/json');
  }

  if (token) {
    headers.set('Authorization', `Bearer ${token}`);
  }

  const { timeoutMs: customTimeoutMs, signal: callerSignal, ...fetchInit } = options;
  const timeoutMs =
    typeof customTimeoutMs === 'number' && customTimeoutMs > 0
      ? customTimeoutMs
      : fetchInit.body instanceof FormData
      ? 60000
      : 15000;

  const controller = new AbortController();
  let timedOut = false;

  const onCallerAbort = () => {
    controller.abort(callerSignal?.reason);
  };
  if (callerSignal) {
    if (callerSignal.aborted) {
      controller.abort(callerSignal.reason);
    } else {
      callerSignal.addEventListener('abort', onCallerAbort, { once: true });
    }
  }

  const timeoutHandle = window.setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, timeoutMs);

  const targetUrl = buildApiUrl(endpoint);
  let res: Response;
  try {
    res = await fetch(targetUrl, {
      ...fetchInit,
      cache: fetchInit.cache || 'no-store',
      headers,
      credentials: 'include',
      signal: controller.signal,
    });
  } catch (fetchErr: any) {
    window.clearTimeout(timeoutHandle);
    if (callerSignal) callerSignal.removeEventListener('abort', onCallerAbort);

    if (timedOut) {
      const timeoutErr: any = new Error(
        'Quá thời gian chờ phản hồi từ máy chủ. Vui lòng thử lại.'
      );
      timeoutErr.status = 408;
      timeoutErr.code = 'API_TIMEOUT';
      throw timeoutErr;
    }

    if (callerSignal?.aborted || fetchErr?.name === 'AbortError') {
      throw fetchErr;
    }

    const networkErr: any = new Error(
      'Không kết nối được máy chủ. Vui lòng kiểm tra kết nối mạng hoặc trạng thái dịch vụ backend.'
    );
    networkErr.status = 0;
    networkErr.code = 'NETWORK_ERROR';
    throw networkErr;
  }

  window.clearTimeout(timeoutHandle);
  if (callerSignal) callerSignal.removeEventListener('abort', onCallerAbort);

  const isLoginEndpoint = endpoint.startsWith('/api/auth/login');

  if (res.status === 401) {
    if (isLoginEndpoint) {
      const errorData = await res.json().catch(() => ({}));
      const authErr: any = new Error(
        errorData.error || errorData.message || 'Sai tên đăng nhập hoặc mật khẩu'
      );
      authErr.status = 401;
      throw authErr;
    }
    clearAuth();
    window.dispatchEvent(new CustomEvent('auth_expired'));
    const expiredErr: any = new Error('Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.');
    expiredErr.status = 401;
    throw expiredErr;
  }

  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    let fallbackMsg = `Yêu cầu thất bại (Mã lỗi: ${res.status})`;
    if (res.status === 404) {
      fallbackMsg = isLoginEndpoint
        ? 'Không tìm thấy dịch vụ đăng nhập (404). Vui lòng kiểm tra cấu hình backend hoặc đường dẫn API.'
        : 'Không tìm thấy dịch vụ API yêu cầu (404).';
    } else if (res.status === 502 || res.status === 503 || res.status === 504) {
      fallbackMsg = 'Không kết nối được máy chủ. Dịch vụ backend hiện tạm thời không phản hồi.';
    }

    const message =
      res.status === 404 && isLoginEndpoint && !errorData.error
        ? fallbackMsg
        : errorData.error || errorData.message || fallbackMsg;

    const err: any = new Error(message);
    err.status = res.status;
    if (errorData.code) err.code = errorData.code;
    if (errorData.persistenceStatus) err.persistenceStatus = errorData.persistenceStatus;
    if (errorData.document) err.document = errorData.document;
    if (errorData.existingDocument) err.existingDocument = errorData.existingDocument;
    if (errorData.conflictingFileName) err.conflictingFileName = errorData.conflictingFileName;
    if (errorData.suggestedSuffixName) err.suggestedSuffixName = errorData.suggestedSuffixName;
    if (errorData.officialDocNumberPreview) err.officialDocNumberPreview = errorData.officialDocNumberPreview;
    if (errorData.serverVersion !== undefined) err.serverVersion = errorData.serverVersion;
    if (errorData.serverBook) err.serverBook = errorData.serverBook;
    if (errorData.previousYear !== undefined) err.previousYear = errorData.previousYear;
    if (errorData.currentYear !== undefined) err.currentYear = errorData.currentYear;
    if (errorData.formattedDate) err.formattedDate = errorData.formattedDate;
    if (errorData.previewNumber) err.previewNumber = errorData.previewNumber;
    if (errorData.sequenceNumber !== undefined) err.sequenceNumber = errorData.sequenceNumber;
    if (errorData.expectedNextSeq !== undefined) err.expectedNextSeq = errorData.expectedNextSeq;
    if (errorData.currentNextSeq !== undefined) err.currentNextSeq = errorData.currentNextSeq;
    if (errorData.documentType) err.documentType = errorData.documentType;
    if (errorData.bookId) err.bookId = errorData.bookId;
    if (errorData.bookName) err.bookName = errorData.bookName;
    throw err;
  }

  return res.json();
}

export interface ExportFilterParams {
  bookId?: string;
  year?: number | string;
  month?: number | string;
  typeCode?: string;
  status?: string;
  category?: string;
  signer?: string;
  driveStatus?: string;
  search?: string;
  sort?: string;
}

function buildExportQuery(params: ExportFilterParams, token: string | null): string {
  const q = new URLSearchParams();
  if (params.bookId) q.append('bookId', params.bookId);
  if (params.year) q.append('year', String(params.year));
  if (params.month) q.append('month', String(params.month));
  if (params.typeCode) q.append('typeCode', params.typeCode);
  if (params.status) q.append('status', params.status);
  if (params.category) q.append('category', params.category);
  if (params.signer) q.append('signer', params.signer);
  if (params.driveStatus) q.append('driveStatus', params.driveStatus);
  if (params.search) q.append('search', params.search);
  if (params.sort) q.append('sort', params.sort);
  if (token) q.append('token', token);
  return q.toString();
}

export const api = {
  getHealth: async (): Promise<{
    status: string;
    app: string;
    timeVietnam: { year: number; month: number; day: number; formattedDate: string };
    durable?: boolean;
    storageMode?: string;
    connectionState?: string;
    redisReachable?: boolean;
    persistenceStatus?: PersistenceStatus;
  }> => {
    return request('/api/health');
  },

  // Storage & Upstash Redis Management (Admin)
  getStorageStatus: async (): Promise<{ persistenceStatus: PersistenceStatus }> => {
    return request<{ persistenceStatus: PersistenceStatus }>('/api/storage/status');
  },

  testStorageConnection: async (): Promise<StorageTestResult> => {
    return request<StorageTestResult>('/api/storage/test-connection', {
      method: 'POST',
    });
  },

  exportLegacySnapshot: async (): Promise<{
    summary: LegacySnapshotSummary;
    snapshot: Record<string, any>;
  }> => {
    return request<{ summary: LegacySnapshotSummary; snapshot: Record<string, any> }>(
      '/api/storage/export-legacy-snapshot'
    );
  },

  importBundledLegacyToStorage: async (forceOverwrite = false): Promise<{
    success: boolean;
    summary: LegacySnapshotSummary;
    persistenceStatus: PersistenceStatus;
    message: string;
  }> => {
    return request('/api/storage/import-bundled-legacy', {
      method: 'POST',
      body: JSON.stringify({ forceOverwrite }),
    });
  },

  importSnapshotToStorage: async (
    snapshot: Record<string, any>,
    forceOverwrite = false
  ): Promise<{
    success: boolean;
    summary: LegacySnapshotSummary;
    persistenceStatus: PersistenceStatus;
    message: string;
  }> => {
    return request('/api/storage/import', {
      method: 'POST',
      body: JSON.stringify({ snapshot, forceOverwrite }),
    });
  },

  initializeNewStorage: async (confirmationText: string): Promise<{
    success: boolean;
    summary: LegacySnapshotSummary;
    persistenceStatus: PersistenceStatus;
    message: string;
  }> => {
    return request('/api/storage/initialize-new', {
      method: 'POST',
      body: JSON.stringify({ confirmationText }),
    });
  },

  // Auth
  login: async (
    username: string,
    password: string,
    rememberMe = false
  ): Promise<AuthResponse & { persistenceStatus?: PersistenceStatus }> => {
    return request<AuthResponse & { persistenceStatus?: PersistenceStatus }>('/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ username, password, rememberMe }),
    });
  },

  getCurrentUser: async (): Promise<{ user: User; persistenceStatus?: PersistenceStatus }> => {
    return request<{ user: User; persistenceStatus?: PersistenceStatus }>('/api/auth/me');
  },

  logout: async (): Promise<void> => {
    const token = getStoredToken();
    const headers = new Headers();
    if (token) {
      headers.set('Authorization', `Bearer ${token}`);
    }
    try {
      await fetch(buildApiUrl('/api/auth/logout'), {
        method: 'POST',
        headers,
        credentials: 'include',
      });
    } catch {
      // ignore network failure during logout
    } finally {
      clearAuth();
    }
  },

  // Users
  getUsers: async (): Promise<{ users: User[] }> => {
    return request<{ users: User[] }>('/api/users');
  },

  createUser: async (payload: {
    username: string;
    fullName: string;
    password: string;
    role?: 'admin' | 'issuer';
    bookPermissions?: Record<string, BookPermission[]>;
  }): Promise<{ user: User }> => {
    return request<{ user: User }>('/api/users', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  updateUser: async (
    id: string,
    payload: Partial<{
      fullName: string;
      role: string;
      isActive: boolean;
      password?: string;
      bookPermissions?: Record<string, BookPermission[]>;
    }>
  ): Promise<{ user: User }> => {
    return request<{ user: User }>(`/api/users/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  deleteUser: async (id: string): Promise<{ success: boolean; message: string }> => {
    return request<{ success: boolean; message: string }>(`/api/users/${id}`, {
      method: 'DELETE',
    });
  },

  // Registry Books
  getBooks: async (): Promise<{
    books: RegistryBook[];
    serverTimeVietnam?: { year: number; month: number; day: number; formattedDate: string };
  }> => {
    return request<{
      books: RegistryBook[];
      serverTimeVietnam?: { year: number; month: number; day: number; formattedDate: string };
    }>('/api/books');
  },

  createBook: async (payload: {
    code?: string;
    name: string;
    unitName: string;
    unitCode: string;
    isShared?: boolean;
    numberMode?: 'separate' | 'shared';
    numberFormat?: string;
    enabledTypeCodes?: string[];
  }): Promise<{ book: RegistryBook }> => {
    return request<{ book: RegistryBook }>('/api/books', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  updateBook: async (
    id: string,
    payload: Partial<{
      name: string;
      unitName: string;
      unitCode: string;
      isShared: boolean;
      numberMode: 'separate' | 'shared';
      numberFormat: string;
      enabledTypeCodes: string[];
      signers: SignerConfig[];
      isActive: boolean;
      expectedVersion: number;
      forceOverwrite: boolean;
    }>
  ): Promise<{ book: RegistryBook }> => {
    return request<{ book: RegistryBook }>(`/api/books/${encodeURIComponent(id)}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  getAuditLogs: async (bookId?: string): Promise<{ logs: SystemAuditLog[] }> => {
    const q = bookId ? `?bookId=${encodeURIComponent(bookId)}` : '';
    return request<{ logs: SystemAuditLog[] }>(`/api/audit-logs${q}`);
  },

  // Number Requests (Quyền "Xin số" tách biệt với "Cấp số")
  getNumberRequests: async (params: { bookId?: string; status?: string } = {}): Promise<{ requests: NumberRequest[] }> => {
    const q = new URLSearchParams();
    if (params.bookId) q.append('bookId', params.bookId);
    if (params.status) q.append('status', params.status);
    const qs = q.toString();
    return request<{ requests: NumberRequest[] }>(`/api/number-requests${qs ? `?${qs}` : ''}`);
  },

  createNumberRequest: async (payload: {
    bookId: string;
    typeCode: string;
    docName: string;
    title?: string;
    signer: string;
    signerRole: string;
    notes?: string;
  }): Promise<{ request: NumberRequest; success: boolean }> => {
    return request<{ request: NumberRequest; success: boolean }>('/api/number-requests', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  approveNumberRequest: async (id: string): Promise<{ document: DocumentRecord; request: NumberRequest; success: boolean }> => {
    return request<{ document: DocumentRecord; request: NumberRequest; success: boolean }>(
      `/api/number-requests/${encodeURIComponent(id)}/issue`,
      { method: 'POST' }
    );
  },

  rejectNumberRequest: async (id: string, reason?: string): Promise<{ request: NumberRequest; success: boolean }> => {
    return request<{ request: NumberRequest; success: boolean }>(
      `/api/number-requests/${encodeURIComponent(id)}/reject`,
      {
        method: 'POST',
        body: JSON.stringify({ reason }),
      }
    );
  },

  // Settings
  getSettings: async (): Promise<{
    settings: SystemSettings;
    documentTypes: DocumentTypeConfig[];
    signers: SignerConfig[];
    serverTimeVietnam?: { year: number; month: number; day: number; formattedDate: string };
    persistenceStatus?: PersistenceStatus;
    legacyReconciliationWarnings?: string[];
  }> => {
    return request<{
      settings: SystemSettings;
      documentTypes: DocumentTypeConfig[];
      signers: SignerConfig[];
      serverTimeVietnam?: { year: number; month: number; day: number; formattedDate: string };
      persistenceStatus?: PersistenceStatus;
      legacyReconciliationWarnings?: string[];
    }>('/api/settings');
  },

  updateSettings: async (
    settings: Partial<SystemSettings> & { bookId?: string; expectedVersion?: number; forceOverwrite?: boolean }
  ): Promise<{ settings: SystemSettings; book?: RegistryBook }> => {
    return request<{ settings: SystemSettings; book?: RegistryBook }>('/api/settings', {
      method: 'PUT',
      body: JSON.stringify(settings),
    });
  },

  updateNumberMode: async (
    newMode: 'separate' | 'shared',
    confirmationText?: string,
    bookId?: string,
    expectedVersion?: number,
    forceOverwrite?: boolean
  ): Promise<{ success: boolean; settings: SystemSettings; book?: RegistryBook }> => {
    return request<{ success: boolean; settings: SystemSettings; book?: RegistryBook }>('/api/settings/number-mode', {
      method: 'POST',
      body: JSON.stringify({ newMode, confirmationText, bookId, expectedVersion, forceOverwrite }),
    });
  },

  getStartSequenceConfig: async (
    bookId: string,
    year: number,
    typeCode: string,
    signal?: AbortSignal
  ): Promise<StartSequenceSeriesConfig> => {
    const q = new URLSearchParams({
      bookId,
      year: String(year),
      typeCode,
    });
    return request<StartSequenceSeriesConfig>(`/api/settings/start-sequence?${q.toString()}`, {
      signal,
      cache: 'no-store',
    });
  },

  setStartSequence: async (
    year: number,
    typeCode: string,
    startSequence: number,
    bookId?: string,
    expectedVersion?: number,
    forceOverwrite?: boolean
  ): Promise<
    StartSequenceSeriesConfig & {
      success: boolean;
      book?: RegistryBook;
    }
  > => {
    return request<
      StartSequenceSeriesConfig & {
        success: boolean;
        book?: RegistryBook;
      }
    >('/api/settings/start-sequence', {
      method: 'POST',
      body: JSON.stringify({ year, typeCode, startSequence, bookId, expectedVersion, forceOverwrite }),
    });
  },

  // Document Types
  createDocumentType: async (payload: {
    name: string;
    code: string;
    color?: string;
    bookId?: string;
    enableExistingForBook?: boolean;
  }): Promise<{ documentType: DocumentTypeConfig; book?: RegistryBook; enabledExisting?: boolean }> => {
    return request<{ documentType: DocumentTypeConfig; book?: RegistryBook; enabledExisting?: boolean }>(
      '/api/document-types',
      {
        method: 'POST',
        body: JSON.stringify(payload),
      }
    );
  },

  updateDocumentType: async (id: string, payload: Partial<DocumentTypeConfig>): Promise<{ documentType: DocumentTypeConfig }> => {
    return request<{ documentType: DocumentTypeConfig }>(`/api/document-types/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  deleteDocumentType: async (id: string): Promise<{ success: boolean }> => {
    return request<{ success: boolean }>(`/api/document-types/${id}`, {
      method: 'DELETE',
    });
  },

  // Signers
  createSigner: async (payload: { name: string; role: string }): Promise<{ signer: SignerConfig }> => {
    return request<{ signer: SignerConfig }>('/api/signers', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  updateSigner: async (id: string, payload: { name: string; role: string }): Promise<{ signer: SignerConfig }> => {
    return request<{ signer: SignerConfig }>(`/api/signers/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  deleteSigner: async (id: string): Promise<{ success: boolean }> => {
    return request<{ success: boolean }>(`/api/signers/${id}`, {
      method: 'DELETE',
    });
  },

  // Number Preview
  previewNumber: async (
    typeCode: string,
    bookId?: string,
    year?: number,
    signal?: AbortSignal
  ): Promise<PreviewNumberResponse> => {
    const q = new URLSearchParams({ typeCode });
    if (bookId) q.append('bookId', bookId);
    if (year) q.append('year', String(year));
    return request<PreviewNumberResponse>(`/api/documents/preview?${q.toString()}`, {
      signal,
      cache: 'no-store',
    });
  },

  // Check Drive filename conflict
  checkFilenameConflict: async (
    fileName: string,
    excludeDocId?: string,
    bookId?: string
  ): Promise<{
    conflict: boolean;
    conflictingFileName: string;
    suggestedSuffixName: string;
    existingDocument: DocumentRecord | null;
  }> => {
    return request('/api/documents/check-filename', {
      method: 'POST',
      body: JSON.stringify({ fileName, excludeDocId, bookId }),
    });
  },

  // Check suspected or exact duplicate within target book
  checkDuplicate: async (payload: {
    bookId?: string;
    fullDocNumber?: string;
    issuingAgency?: string;
    year?: number;
    sha256?: string;
    excludeDocId?: string;
  }): Promise<{
    exactDuplicate: DocumentRecord | null;
    suspectedDuplicate: DocumentRecord | null;
  }> => {
    return request('/api/documents/check-duplicate', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  // Issue Number (Quyền "Cấp số")
  issueDocument: async (payload: {
    bookId?: string;
    typeCode: string;
    docName: string;
    title?: string;
    signer: string;
    signerRole: string;
    notes?: string;
    idempotencyKey: string;
    hasAttachment?: boolean;
    attachmentName?: string;
    attachmentSize?: number;
    conflictAction?: 'suffix' | 'overwrite';
    customFileName?: string;
    expectedYear?: number;
    expectedNextSeq?: number;
    confirmUpdatedSequence?: boolean;
  }): Promise<{ document: DocumentRecord; nextPreview?: PreviewNumberResponse; duplicatePrevented?: boolean }> => {
    return request<{ document: DocumentRecord; nextPreview?: PreviewNumberResponse; duplicatePrevented?: boolean }>(
      '/api/documents/issue',
      {
        method: 'POST',
        body: JSON.stringify(payload),
      }
    );
  },

  // Upload, Retry or Replace Word Attachment on Google Drive
  uploadWordToDocument: async (
    docId: string,
    file: File,
    fileName?: string,
    mode?: 'retry' | 'replace'
  ): Promise<{ document: DocumentRecord; deduplicated?: boolean; success?: boolean; message?: string }> => {
    const formData = new FormData();
    formData.append('file', file);
    if (fileName) {
      formData.append('fileName', fileName);
    }
    if (mode) {
      formData.append('mode', mode);
    }
    return request<{ document: DocumentRecord; deduplicated?: boolean; success?: boolean; message?: string }>(
      `/api/documents/${encodeURIComponent(docId)}/upload-word`,
      {
        method: 'POST',
        body: formData,
      }
    );
  },

  // Incoming PDF Analysis & Import
  analyzePdfFile: async (file: File, bookId?: string): Promise<PdfAnalyzeResponse> => {
    const formData = new FormData();
    formData.append('file', file);
    if (bookId) {
      formData.append('bookId', bookId);
    }
    return request<PdfAnalyzeResponse>('/api/pdf/analyze', {
      method: 'POST',
      body: formData,
    });
  },

  analyzePdfUrl: async (url: string, bookId?: string): Promise<PdfAnalyzeResponse> => {
    return request<PdfAnalyzeResponse>('/api/pdf/analyze', {
      method: 'POST',
      body: JSON.stringify({ url, bookId }),
      timeoutMs: 45000,
    });
  },

  getStagedPdfUrl: (stageToken: string): string => {
    const token = getStoredToken();
    const q = token ? `?token=${encodeURIComponent(token)}` : '';
    return buildApiUrl(`/api/pdf/staged/${encodeURIComponent(stageToken)}${q}`);
  },

  importPdfDocument: async (payload: {
    bookId?: string;
    stageToken: string;
    fullDocNumber: string;
    docName: string;
    title?: string;
    issueDate: string;
    signDate?: string;
    issuingAgency?: string;
    signer?: string;
    signerRole?: string;
    notes?: string;
    expectedFileName: string;
    conflictAction?: 'suffix' | 'overwrite';
    customFileName?: string;
    ignoreDuplicatePdf?: boolean;
    ignoreSuspectedDuplicate?: boolean;
    needsReview?: boolean;
    reviewReasons?: string[];
    relevantPages?: number[];
    idempotencyKey: string;
  }): Promise<{
    document: DocumentRecord;
    driveSaved: boolean;
    uploadError?: string;
    stageToken?: string;
    duplicatePrevented?: boolean;
  }> => {
    return request('/api/documents/import-pdf', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  uploadPdfToDocument: async (
    docId: string,
    options: { file?: File; stageToken?: string; fileName?: string; mode?: 'retry' | 'replace' }
  ): Promise<{ document: DocumentRecord; deduplicated?: boolean; success?: boolean; message?: string }> => {
    const formData = new FormData();
    if (options.file) {
      formData.append('file', options.file);
    }
    if (options.stageToken) {
      formData.append('stageToken', options.stageToken);
    }
    if (options.fileName) {
      formData.append('fileName', options.fileName);
    }
    if (options.mode) {
      formData.append('mode', options.mode);
    }
    return request<{ document: DocumentRecord; deduplicated?: boolean; success?: boolean; message?: string }>(
      `/api/documents/${encodeURIComponent(docId)}/upload-pdf`,
      {
        method: 'POST',
        body: formData,
      }
    );
  },

  // Documents
  getDocuments: async (params: {
    bookId?: string;
    year?: string | number;
    month?: string | number;
    typeCode?: string;
    category?: string;
    signer?: string;
    status?: string;
    driveStatus?: string;
    deletedBy?: string;
    fromDate?: string;
    toDate?: string;
    search?: string;
    sort?: string;
  } = {}): Promise<{ total: number; documents: DocumentRecord[] }> => {
    const searchParams = new URLSearchParams();
    Object.entries(params).forEach(([k, v]) => {
      if (v !== undefined && v !== null && v !== '') {
        searchParams.append(k, String(v));
      }
    });
    return request<{ total: number; documents: DocumentRecord[] }>(`/api/documents?${searchParams.toString()}`);
  },

  getDocumentHistory: async (params: {
    bookId?: string;
    username?: string;
    userId?: string;
    fromDate?: string;
    toDate?: string;
    docId?: string;
    action?: string;
    search?: string;
  } = {}): Promise<{
    total: number;
    logs: DocumentHistoryEntry[];
    actors: Array<{ userId: string; username: string; userName: string }>;
  }> => {
    const searchParams = new URLSearchParams();
    Object.entries(params).forEach(([k, v]) => {
      if (v !== undefined && v !== null && v !== '' && v !== 'all') {
        searchParams.append(k, String(v));
      }
    });
    const qs = searchParams.toString();
    return request<{
      total: number;
      logs: DocumentHistoryEntry[];
      actors: Array<{ userId: string; username: string; userName: string }>;
    }>(`/api/document-history${qs ? `?${qs}` : ''}`);
  },

  getDocument: async (id: string): Promise<{ document: DocumentRecord }> => {
    return request<{ document: DocumentRecord }>(`/api/documents/${id}`);
  },

  openDocumentFile: async (id: string): Promise<{ url: string; fileName: string; fileId: string }> => {
    return request<{ url: string; fileName: string; fileId: string }>(
      `/api/documents/${encodeURIComponent(id)}/open-file`
    );
  },

  updateDocument: async (
    id: string,
    payload: {
      docName?: string;
      title?: string;
      signer?: string;
      signerRole?: string;
      notes?: string;
      sequenceNumber?: number;
      fullDocNumber?: string;
      issuingAgency?: string;
      issueDate?: string;
      signDate?: string;
      needsReview?: boolean;
    }
  ): Promise<{ document: DocumentRecord }> => {
    return request<{ document: DocumentRecord }>(`/api/documents/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  deleteDocument: async (
    id: string,
    reason: string,
    confirmed = true
  ): Promise<{ success: boolean; document: DocumentRecord; message: string }> => {
    return request<{ success: boolean; document: DocumentRecord; message: string }>(`/api/documents/${encodeURIComponent(id)}`, {
      method: 'DELETE',
      body: JSON.stringify({ reason, confirmed }),
    });
  },

  retryDocumentExcelSync: async (
    id: string
  ): Promise<{ success: boolean; document: DocumentRecord; message?: string }> => {
    return request<{ success: boolean; document: DocumentRecord; message?: string }>(
      `/api/documents/${encodeURIComponent(id)}/retry-excel-sync`,
      {
        method: 'POST',
      }
    );
  },

  cancelDocument: async (id: string, reason: string): Promise<{ document: DocumentRecord }> => {
    return request<{ document: DocumentRecord }>(`/api/documents/${id}/cancel`, {
      method: 'POST',
      body: JSON.stringify({ reason }),
    });
  },

  // Statistics
  getStatistics: async (year?: number, bookId?: string): Promise<StatisticsData> => {
    const q = new URLSearchParams();
    if (year) q.append('year', String(year));
    if (bookId && bookId !== 'all') q.append('bookId', bookId);
    const qs = q.toString();
    return request<StatisticsData>(`/api/statistics${qs ? `?${qs}` : ''}`);
  },

  // Excel Export URL generator
  getExportUrl: (params: ExportFilterParams = {}): string => {
    const token = getStoredToken();
    return buildApiUrl(`/api/export-excel?${buildExportQuery(params, token)}`);
  },

  // Reliable Blob-based Excel Download
  downloadExcel: async (params: ExportFilterParams = {}): Promise<void> => {
    const token = getStoredToken();
    const endpoint = buildApiUrl(`/api/export-excel?${buildExportQuery(params, token)}`);
    const headers = new Headers();
    if (token) {
      headers.set('Authorization', `Bearer ${token}`);
    }

    const res = await fetch(endpoint, { headers, credentials: 'include' });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Tải file Excel thất bại (Mã phản hồi: ${res.status})`);
    }

    const blob = await res.blob();
    const yearStr = params.year && params.year !== 'all' ? String(params.year) : new Date().getFullYear().toString();
    const fileName = `SO_CAP_SO_VAN_BAN_DOAN_TRUONG_${yearStr}.xlsx`;

    const downloadUrl = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = downloadUrl;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 5000);
  },

  // Google Sheets Webhook sync
  saveGoogleSheetsConfig: async (config: { enabled: boolean; webhookUrl: string }): Promise<any> => {
    return request('/api/sync/google-sheets/config', {
      method: 'POST',
      body: JSON.stringify(config),
    });
  },

  testGoogleSheetsConnection: async (
    webhookUrl?: string,
    bookId?: string
  ): Promise<{ success: boolean; message: string }> => {
    return request<{ success: boolean; message: string }>('/api/sync/google-sheets/test', {
      method: 'POST',
      body: JSON.stringify({ webhookUrl, bookId }),
    });
  },

  syncAllGoogleSheets: async (bookId?: string): Promise<{ success: boolean; message: string }> => {
    return request<{ success: boolean; message: string }>('/api/sync/google-sheets/sync-all', {
      method: 'POST',
      body: JSON.stringify(bookId ? { bookId } : {}),
    });
  },

  saveBookGoogleSheetsConfig: async (
    bookId: string,
    config: { enabled: boolean; webhookUrl: string; expectedVersion?: number; forceOverwrite?: boolean }
  ): Promise<{ book: RegistryBook }> => {
    return request<{ book: RegistryBook }>(`/api/books/${encodeURIComponent(bookId)}/google-sheets`, {
      method: 'POST',
      body: JSON.stringify(config),
    });
  },

  // Google Drive Word/PDF Attachment Config & Test (hỗ trợ từng sổ hoặc mặc định)
  saveGoogleDriveConfig: async (payload: {
    bookId?: string;
    enabled: boolean;
    authMode: 'apps_script' | 'service_account';
    folderUrl: string;
    appsScriptUrl?: string;
    clearAppsScriptUrl?: boolean;
    serviceAccountJson?: string;
    clearServiceAccountJson?: boolean;
    scopeUnitCode?: string;
    expectedVersion?: number;
    forceOverwrite?: boolean;
  }): Promise<{ googleDriveConfig: GoogleDriveConfig; book?: RegistryBook }> => {
    return request<{ googleDriveConfig: GoogleDriveConfig; book?: RegistryBook }>('/api/drive/config', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  testGoogleDriveConnection: async (payload?: {
    bookId?: string;
    authMode?: 'apps_script' | 'service_account';
    folderUrl?: string;
    appsScriptUrl?: string;
    serviceAccountJson?: string;
  }): Promise<{
    success: boolean;
    folderId: string;
    folderName?: string;
    message: string;
    googleDriveConfig?: GoogleDriveConfig;
  }> => {
    return request<{
      success: boolean;
      folderId: string;
      folderName?: string;
      message: string;
      googleDriveConfig?: GoogleDriveConfig;
    }>('/api/drive/test', {
      method: 'POST',
      body: JSON.stringify(payload || {}),
    });
  },
};
