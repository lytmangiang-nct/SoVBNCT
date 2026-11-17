import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  Search,
  Copy,
  Check,
  Eye,
  Edit2,
  Trash2,
  X,
  History,
  AlertTriangle,
  FileSpreadsheet,
  RefreshCw,
  ExternalLink,
  Ban,
  Loader2,
  FileText,
  LayoutList,
  Table,
  FileUp,
  FolderCheck,
  Upload,
  BookOpen,
  RotateCw,
  Clock,
  Archive,
  ChevronLeft,
  ChevronRight,
  Eraser,
  CheckCircle2,
  MoreHorizontal,
  SlidersHorizontal,
  Plus,
  Paperclip,
  ListChecks,
  CalendarClock,
} from 'lucide-react';
import { api } from '../api';
import type {
  ArchiveDeletedResult,
  AttachmentGroup,
  BulkSyncDeletedResult,
  DocumentAuditAction,
  DocumentHistoryEntry,
  DocumentRecord,
  DocumentTask,
  DocumentTaskStatus,
  DocumentTypeConfig,
  HistoryCleanupPreviewResponse,
  HistoryCleanupScope,
  PdfContentSearchMatch,
  PdfTaskSuggestion,
  RegistryBook,
  SystemSettings,
  User,
} from '../types';
import { ATTACHMENT_GROUP_LABELS, getEffectiveAttachments } from '../utils/attachments';
import { getTypeColorStyle, SEMANTIC_STATUS_STYLES } from '../utils/documentColors';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  formatFileSize,
  generateSuggestedFileName,
  getDocumentFilterYear,
  getWordFileExtension,
  normalizeDateOnlyDdMmYyyy,
} from '../utils/fileNaming';
import { getDocBookId, hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface DocumentRegistryViewProps {
  documents: DocumentRecord[];
  deletedDocuments?: DocumentRecord[];
  documentTypes: DocumentTypeConfig[];
  books: RegistryBook[];
  activeBookId: string;
  onSelectBookId: (bookId: string) => void;
  currentUser: User;
  settings?: SystemSettings;
  serverYear?: number;
  section?: 'active' | 'deleted' | 'history';
  onSectionChange?: (section: 'active' | 'deleted' | 'history') => void;
  onRefresh: () => void;
  onDocumentUpdated: (doc: DocumentRecord) => void;
  initialOpenDocId?: string | null;
  onClearInitialOpenDocId?: () => void;
  presetCategory?: 'all' | 'outgoing' | 'incoming' | null;
  onClearPresetCategory?: () => void;
}

function getRecordStatusMeta(doc: DocumentRecord): { label: string; className: string } {
  if (doc.status === 'deleted') {
    return {
      label: 'Đã xóa',
      className: SEMANTIC_STATUS_STYLES.danger.badge,
    };
  }
  if (doc.status === 'cancelled') {
    return {
      label: 'Đã hủy',
      className: SEMANTIC_STATUS_STYLES.danger.badge,
    };
  }
  if (doc.docCategory === 'incoming') {
    return {
      label: 'Văn bản đến',
      className: SEMANTIC_STATUS_STYLES.incoming.badge,
    };
  }
  return {
    label: 'Đã cấp số',
    className: SEMANTIC_STATUS_STYLES.outgoing.badge,
  };
}

function getDriveStatusLabel(doc: DocumentRecord): string {
  switch (doc.driveStatus) {
    case 'saved':
      return 'Đã lưu lên Drive';
    case 'failed':
      return 'Lưu tệp thất bại';
    case 'uploading':
      return 'Đang lưu';
    default:
      return 'Không đính kèm';
  }
}

function formatVietnamTimestampDate(isoOrDate?: string): string {
  if (!isoOrDate) return '—';
  const normalized = normalizeDateOnlyDdMmYyyy(isoOrDate);
  if (normalized) return normalized;
  try {
    const dt = new Date(isoOrDate);
    if (Number.isNaN(dt.getTime())) return '—';
    return new Intl.DateTimeFormat('vi-VN', {
      timeZone: 'Asia/Ho_Chi_Minh',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    }).format(dt);
  } catch {
    return '—';
  }
}

function getExcelDeleteSyncMeta(doc: DocumentRecord): {
  label: string;
  className: string;
  needsRetry: boolean;
} {
  if (doc.deletedListArchived) {
    return {
      label: 'Đã lưu trữ · Đã đồng bộ Google Trang tính',
      className: 'text-slate-700 bg-slate-100 border-slate-300',
      needsRetry: false,
    };
  }
  if (doc.excelDeleteSyncStatus === 'synced') {
    return {
      label: 'Đã chuyển sang sheet VB đã xóa',
      className: 'text-emerald-800 bg-emerald-50 border-emerald-200',
      needsRetry: false,
    };
  }
  if (doc.excelDeleteSyncStatus === 'failed' || doc.excelDeleteSyncStatus === 'pending') {
    return {
      label: 'Đã xóa trên hệ thống · Chờ đồng bộ Google Trang tính',
      className: 'text-amber-900 bg-amber-50 border-amber-300',
      needsRetry: true,
    };
  }
  return {
    label: 'Đã chuyển VB đã xóa (Chưa bật Google Trang tính)',
    className: 'text-slate-600 bg-slate-100 border-slate-200',
    needsRetry: false,
  };
}

function getClientVietnamTime() {
  try {
    const iso = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Ho_Chi_Minh',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    }).format(new Date());
    return {
      year: parseInt(iso.slice(0, 4), 10),
      dateKey: iso,
    };
  } catch {
    const now = new Date();
    return {
      year: now.getFullYear(),
      dateKey: now.toISOString().slice(0, 10),
    };
  }
}

interface PaginationControlsProps {
  totalItems: number;
  page: number;
  pageSize: 10 | 20 | 50;
  onPageChange: (nextPage: number) => void;
  onPageSizeChange: (nextSize: 10 | 20 | 50) => void;
}

const PaginationControls: React.FC<PaginationControlsProps> = ({
  totalItems,
  page,
  pageSize,
  onPageChange,
  onPageSizeChange,
}) => {
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const startIdx = totalItems === 0 ? 0 : (safePage - 1) * pageSize + 1;
  const endIdx = totalItems === 0 ? 0 : Math.min(safePage * pageSize, totalItems);
  const summaryText = totalItems === 0 ? '0 / 0 kết quả' : `${startIdx}–${endIdx} / ${totalItems} kết quả`;

  return (
    <div className="bg-white border border-slate-200 rounded-xl px-3.5 py-2.5 shadow-2xs flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-xs text-slate-600">
      <div className="flex items-center justify-between sm:justify-start gap-3 flex-wrap">
        <span className="font-semibold text-slate-800 tabular-nums">{summaryText}</span>
        <label className="inline-flex items-center gap-1.5 text-slate-500">
          <span>Hiển thị:</span>
          <select
            aria-label="Số dòng mỗi trang"
            value={pageSize}
            onChange={(e) => onPageSizeChange(Number(e.target.value) as 10 | 20 | 50)}
            className="min-h-[34px] h-8 px-2 text-xs font-semibold text-slate-800 bg-slate-50 border border-slate-200 rounded-md focus:outline-hidden focus:ring-2 focus:ring-sky-600"
          >
            <option value={10}>10 / trang</option>
            <option value={20}>20 / trang</option>
            <option value={50}>50 / trang</option>
          </select>
        </label>
      </div>

      <div className="flex items-center justify-end gap-2">
        <button
          type="button"
          onClick={() => onPageChange(Math.max(1, safePage - 1))}
          disabled={safePage <= 1 || totalItems === 0}
          className="inline-flex items-center gap-1 min-h-[34px] px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-semibold disabled:opacity-45 disabled:pointer-events-none transition-colors"
        >
          <ChevronLeft className="w-3.5 h-3.5" />
          <span>Trước</span>
        </button>

        <span className="font-mono text-xs font-semibold text-slate-700 px-2 tabular-nums">
          Trang {safePage} / {totalPages}
        </span>

        <button
          type="button"
          onClick={() => onPageChange(Math.min(totalPages, safePage + 1))}
          disabled={safePage >= totalPages || totalItems === 0}
          className="inline-flex items-center gap-1 min-h-[34px] px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-semibold disabled:opacity-45 disabled:pointer-events-none transition-colors"
        >
          <span>Sau</span>
          <ChevronRight className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
};

function getAuditActionMeta(action: DocumentAuditAction): { label: string; className: string } {
  switch (action) {
    case 'create':
      return { label: 'Cấp số', className: 'bg-emerald-50 text-emerald-800 border-emerald-200' };
    case 'import_pdf':
      return { label: 'Nhập PDF', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'update':
      return { label: 'Sửa văn bản', className: 'bg-amber-50 text-amber-800 border-amber-200' };
    case 'upload_word':
      return { label: 'Đính kèm Word', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'upload_pdf':
      return { label: 'Lưu PDF Drive', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'replace_file':
      return { label: 'Thay tệp Drive', className: 'bg-indigo-50 text-indigo-800 border-indigo-200' };
    case 'cancel':
      return { label: 'Hủy số', className: 'bg-rose-50 text-rose-800 border-rose-200' };
    case 'delete':
      return { label: 'Xóa văn bản', className: 'bg-rose-100 text-rose-900 border-rose-300' };
    case 'retry_excel_sync':
      return { label: 'Đồng bộ lại Google Trang tính', className: 'bg-amber-50 text-amber-900 border-amber-200' };
    case 'request_number':
      return { label: 'Xin số', className: 'bg-slate-100 text-slate-700 border-slate-200' };
    case 'template_manage':
      return { label: 'Quản lý mẫu Word', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'word_generate':
      return { label: 'Tạo Word từ mẫu', className: 'bg-emerald-50 text-emerald-800 border-emerald-200' };
    case 'attachment_add':
      return { label: 'Thêm tệp đính kèm', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'attachment_replace':
      return { label: 'Cập nhật phiên bản tệp', className: 'bg-indigo-50 text-indigo-800 border-indigo-200' };
    case 'attachment_set_official':
      return { label: 'Đặt bản chính thức', className: 'bg-emerald-50 text-emerald-800 border-emerald-200' };
    case 'task_create':
      return { label: 'Giao việc từ văn bản', className: 'bg-sky-50 text-sky-800 border-sky-200' };
    case 'task_update':
      return { label: 'Cập nhật tiến độ việc', className: 'bg-amber-50 text-amber-800 border-amber-200' };
    case 'reconcile_sync':
      return { label: 'Đối soát đồng bộ', className: 'bg-amber-50 text-amber-800 border-amber-200' };
    default:
      return { label: action, className: 'bg-slate-100 text-slate-700 border-slate-200' };
  }
}

const AUDIT_FIELD_LABELS: Record<string, string> = {
  docName: 'Tên văn bản',
  title: 'Trích yếu / Nội dung',
  signer: 'Người ký',
  signerRole: 'Chức vụ',
  notes: 'Ghi chú',
  sequenceNumber: 'Số thứ tự',
  fullDocNumber: 'Số / Ký hiệu',
  issuingAgency: 'Cơ quan ban hành',
  issueDate: 'Ngày ban hành',
  signDate: 'Ngày ký',
  driveFileName: 'Tên tệp Drive',
  driveFileId: 'ID tệp Drive',
  driveFileUrl: 'Liên kết Drive',
  driveFilePreserved: 'Tệp Drive được giữ nguyên',
  status: 'Trạng thái',
  deleteReason: 'Lý do xóa',
  excelDeleteSyncStatus: 'Trạng thái đồng bộ VB đã xóa',
  syncStatus: 'Trạng thái đồng bộ Google Trang tính',
  needsReview: 'Cờ cần kiểm tra',
  attachmentFile: 'Tệp đính kèm',
  officialAttachment: 'Bản chính thức',
  taskTitle: 'Nội dung việc',
  assignee: 'Người phụ trách',
  dueDate: 'Hạn nộp',
  taskStatus: 'Trạng thái việc',
};

function formatAuditFieldLabel(field: string): string {
  return AUDIT_FIELD_LABELS[field] || field;
}

export const DocumentRegistryView: React.FC<DocumentRegistryViewProps> = ({
  documents,
  deletedDocuments = [],
  documentTypes,
  books,
  activeBookId,
  onSelectBookId: _onSelectBookId,
  currentUser,
  settings,
  serverYear,
  section,
  onSectionChange,
  onRefresh,
  onDocumentUpdated,
  initialOpenDocId,
  onClearInitialOpenDocId,
  presetCategory,
  onClearPresetCategory,
}) => {
  const [internalSection, setInternalSection] = useState<'active' | 'deleted' | 'history'>(section || 'active');
  const registrySection = section ?? internalSection;
  const setRegistrySection = (next: 'active' | 'deleted' | 'history') => {
    setInternalSection(next);
    onSectionChange?.(next);
  };

  useEffect(() => {
    if (section) {
      setInternalSection(section);
    }
  }, [section]);
  const initialVnTime = useMemo(() => getClientVietnamTime(), []);
  const currentVnYear = serverYear || initialVnTime.year;
  const activeYearManualRef = React.useRef(false);
  const deletedYearManualRef = React.useRef(false);

  // Search & Filter State (Active)
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'outgoing' | 'incoming'>('all');
  const [selectedYear, setSelectedYear] = useState<string>(() => String(serverYear || initialVnTime.year));
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [selectedType, setSelectedType] = useState<string>('all');
  const [selectedSigner, setSelectedSigner] = useState<string>('all');
  const [selectedStatus, setSelectedStatus] = useState<string>('all');
  const [selectedDriveStatus, setSelectedDriveStatus] = useState<string>('all');
  const [sortBy, setSortBy] = useState<'newest' | 'oldest' | 'number_asc' | 'number_desc'>('newest');
  const [activePage, setActivePage] = useState<number>(1);
  const [activePageSize, setActivePageSize] = useState<10 | 20 | 50>(20);
  const [showAdvancedFilters, setShowAdvancedFilters] = useState<boolean>(false);

  // PDF Content Search State
  const [pdfContentMode, setPdfContentMode] = useState<boolean>(false);
  const [pdfSearchLoading, setPdfSearchLoading] = useState<boolean>(false);
  const [pdfSearchError, setPdfSearchError] = useState<string | null>(null);
  const [pdfSearchResults, setPdfSearchResults] = useState<{
    total: number;
    matches: PdfContentSearchMatch[];
    indexedCount: number;
    unindexedCount: number;
  } | null>(null);

  useEffect(() => {
    if (serverYear && !activeYearManualRef.current) {
      setSelectedYear(String(serverYear));
    }
  }, [serverYear]);

  // Interactive States
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [copiedDetailFileName, setCopiedDetailFileName] = useState<boolean>(false);
  const [mobileViewMode, setMobileViewMode] = useState<'cards' | 'table'>('cards');
  const [detailModalDoc, setDetailModalDoc] = useState<DocumentRecord | null>(null);
  const [editModalDoc, setEditModalDoc] = useState<DocumentRecord | null>(null);
  const [cancelModalDoc, setCancelModalDoc] = useState<DocumentRecord | null>(null);
  const [deleteModalDoc, setDeleteModalDoc] = useState<DocumentRecord | null>(null);
  const [openingFileId, setOpeningFileId] = useState<string | null>(null);
  const [fileActionError, setFileActionError] = useState<string | null>(null);
  const [openMenu, setOpenMenu] = useState<{
    doc: DocumentRecord;
    top: number;
    right: number;
    openUp: boolean;
  } | null>(null);

  useEffect(() => {
    if (!openMenu) return;
    const handleClose = () => setOpenMenu(null);
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpenMenu(null);
    };
    window.addEventListener('click', handleClose);
    window.addEventListener('scroll', handleClose, true);
    window.addEventListener('resize', handleClose);
    window.addEventListener('keydown', handleKey);
    return () => {
      window.removeEventListener('click', handleClose);
      window.removeEventListener('scroll', handleClose, true);
      window.removeEventListener('resize', handleClose);
      window.removeEventListener('keydown', handleKey);
    };
  }, [openMenu]);

  const toggleActionMenu = (e: React.MouseEvent<HTMLButtonElement>, doc: DocumentRecord) => {
    e.stopPropagation();
    if (openMenu?.doc.id === doc.id) {
      setOpenMenu(null);
      return;
    }
    const rect = e.currentTarget.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const openUp = spaceBelow < 220 && rect.top > 220;
    setOpenMenu({
      doc,
      top: openUp ? rect.top - 6 : rect.bottom + 6,
      right: Math.max(12, window.innerWidth - rect.right),
      openUp,
    });
  };

  // Edit form state
  const [editDocName, setEditDocName] = useState('');
  const [editTitle, setEditTitle] = useState('');
  const [editSigner, setEditSigner] = useState('');
  const [editSignerRole, setEditSignerRole] = useState('');
  const [editNotes, setEditNotes] = useState('');
  const [editIssuingAgency, setEditIssuingAgency] = useState('');
  const [editIssueDate, setEditIssueDate] = useState('');
  const [editSignDate, setEditSignDate] = useState('');
  const [editSequenceNumber, setEditSequenceNumber] = useState<number>(1);
  const [editFullDocNumber, setEditFullDocNumber] = useState<string>('');
  const [editSubmitting, setEditSubmitting] = useState(false);
  const [editError, setEditError] = useState<string | null>(null);

  // Cancel form state
  const [cancelReason, setCancelReason] = useState('');
  const [cancelSubmitting, setCancelSubmitting] = useState(false);
  const [cancelError, setCancelError] = useState<string | null>(null);

  // Delete form state
  const [deleteReason, setDeleteReason] = useState('');
  const [deleteConfirmed, setDeleteConfirmed] = useState(false);
  const [deleteSubmitting, setDeleteSubmitting] = useState(false);
  const [deleteError, setDeleteError] = useState<string | null>(null);

  // Deleted documents tab filters, pagination, bulk sync & archive state
  const [deletedSearch, setDeletedSearch] = useState('');
  const [deletedByFilter, setDeletedByFilter] = useState('all');
  const [deletedDocYear, setDeletedDocYear] = useState<string>(() => String(serverYear || initialVnTime.year));
  const [deletedFromDate, setDeletedFromDate] = useState('');
  const [deletedToDate, setDeletedToDate] = useState('');
  const [deletedSyncFilter, setDeletedSyncFilter] = useState<'all' | 'synced' | 'pending_sync' | 'archived'>('all');
  const [deletedPage, setDeletedPage] = useState<number>(1);
  const [deletedPageSize, setDeletedPageSize] = useState<10 | 20 | 50>(20);
  const [selectedDeletedIds, setSelectedDeletedIds] = useState<string[]>([]);
  const [archivedDeletedDocs, setArchivedDeletedDocs] = useState<DocumentRecord[]>([]);
  const [loadingArchivedDeleted, setLoadingArchivedDeleted] = useState<boolean>(false);
  const [archivedDeletedError, setArchivedDeletedError] = useState<string | null>(null);
  const [retryingSyncDocId, setRetryingSyncDocId] = useState<string | null>(null);
  const [syncRetryNotice, setSyncRetryNotice] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const [bulkSyncConfirmMode, setBulkSyncConfirmMode] = useState<'selected' | 'pending' | null>(null);
  const [bulkSyncSubmitting, setBulkSyncSubmitting] = useState<boolean>(false);
  const [bulkSyncResult, setBulkSyncResult] = useState<BulkSyncDeletedResult | null>(null);
  const [archiveModalOpen, setArchiveModalOpen] = useState<boolean>(false);
  const [archiveMode, setArchiveMode] = useState<'selected' | 'all_synced'>('all_synced');
  const [archiveSubmitting, setArchiveSubmitting] = useState<boolean>(false);
  const [archiveNotice, setArchiveNotice] = useState<ArchiveDeletedResult | null>(null);
  const [archiveError, setArchiveError] = useState<string | null>(null);

  useEffect(() => {
    if (serverYear && !deletedYearManualRef.current) {
      setDeletedDocYear(String(serverYear));
    }
  }, [serverYear]);

  // History tab state
  const [historyLogs, setHistoryLogs] = useState<DocumentHistoryEntry[]>([]);
  const [historyActors, setHistoryActors] = useState<Array<{ userId: string; username: string; userName: string }>>([]);
  const [historyTotal, setHistoryTotal] = useState<number>(0);
  const [historyLoading, setHistoryLoading] = useState<boolean>(false);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [historyActorFilter, setHistoryActorFilter] = useState<string>('all');
  const [historyActionFilter, setHistoryActionFilter] = useState<string>('all');
  const [historyFromDate, setHistoryFromDate] = useState<string>('');
  const [historyToDate, setHistoryToDate] = useState<string>('');
  const [historySearch, setHistorySearch] = useState<string>('');
  const [historyPage, setHistoryPage] = useState<number>(1);
  const [historyPageSize, setHistoryPageSize] = useState<10 | 20 | 50>(20);

  // History Cleanup (Global Admin only)
  const isAdmin = isSystemAdmin(currentUser);
  const [cleanupModalOpen, setCleanupModalOpen] = useState<boolean>(false);
  const [cleanupBookId, setCleanupBookId] = useState<string>('');
  const [cleanupScope, setCleanupScope] = useState<HistoryCleanupScope>('today');
  const [cleanupBeforeDate, setCleanupBeforeDate] = useState<string>(() => initialVnTime.dateKey);
  const [cleanupConfirmText, setCleanupConfirmText] = useState<string>('');
  const [cleanupPreview, setCleanupPreview] = useState<HistoryCleanupPreviewResponse | null>(null);
  const [cleanupPreviewLoading, setCleanupPreviewLoading] = useState<boolean>(false);
  const [cleanupPreviewError, setCleanupPreviewError] = useState<string | null>(null);
  const [cleanupSubmitting, setCleanupSubmitting] = useState<boolean>(false);
  const [cleanupNotice, setCleanupNotice] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

  // Attachment Upload / Retry / Replace modal state
  const [uploadModalDoc, setUploadModalDoc] = useState<DocumentRecord | null>(null);
  const [uploadModalMode, setUploadModalMode] = useState<'retry' | 'replace'>('retry');
  const [uploadModalFile, setUploadModalFile] = useState<File | null>(null);
  const [uploadModalSubmitting, setUploadModalSubmitting] = useState(false);
  const [uploadModalError, setUploadModalError] = useState<string | null>(null);

  // Export Excel loading state
  const [isExporting, setIsExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

  // Detail Modal: Multi-attachment & Tasks state
  const [detailTasks, setDetailTasks] = useState<DocumentTask[]>([]);
  const [detailTasksLoading, setDetailTasksLoading] = useState<boolean>(false);
  const [detailAssignees, setDetailAssignees] = useState<
    Array<{ id: string; username: string; fullName: string; displayName: string; role: string }>
  >([]);
  const [confirmingSugId, setConfirmingSugId] = useState<string | null>(null);
  const [sugAssigneeMap, setSugAssigneeMap] = useState<Record<string, string>>({});
  const [detailTaskError, setDetailTaskError] = useState<string | null>(null);
  const [updatingDetailTaskId, setUpdatingDetailTaskId] = useState<string | null>(null);

  const [showAddAttachmentForm, setShowAddAttachmentForm] = useState<boolean>(false);
  const [attTargetId, setAttTargetId] = useState<string | undefined>(undefined);
  const [attGroup, setAttGroup] = useState<AttachmentGroup>('signed_pdf');
  const [attLabel, setAttLabel] = useState<string>('');
  const [attSetOfficial, setAttSetOfficial] = useState<boolean>(false);
  const [attFile, setAttFile] = useState<File | null>(null);
  const [attUploading, setAttUploading] = useState<boolean>(false);
  const [attError, setAttError] = useState<string | null>(null);
  const [settingOfficialId, setSettingOfficialId] = useState<string | null>(null);
  const [retryingWordTemplate, setRetryingWordTemplate] = useState<boolean>(false);

  const viewableBooks = useMemo(
    () => books.filter((b) => isSystemAdmin(currentUser) || hasBookPermission(currentUser, b.id, 'view')),
    [books, currentUser]
  );

  const bookMap = useMemo(() => {
    const m = new Map<string, RegistryBook>();
    books.forEach((b) => m.set(b.id, b));
    return m;
  }, [books]);

  const driveConfig = settings?.googleDriveConfig;
  const maxWordFileSizeMB = driveConfig?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;
  const maxWordFileSizeBytes = driveConfig?.maxFileSizeBytes || maxWordFileSizeMB * 1024 * 1024;

  // Apply preset category (e.g., "Tra cứu văn bản đến")
  useEffect(() => {
    if (!presetCategory) return;
    setSelectedCategory(presetCategory);
    setRegistrySection('active');
    onClearPresetCategory?.();
  }, [presetCategory, onClearPresetCategory]);

  // Auto-open specific document when navigated from "Xem hồ sơ vừa nhập" / "Mở hồ sơ đã có"
  useEffect(() => {
    if (!initialOpenDocId) return;
    const alignViewForTarget = (target: DocumentRecord) => {
      if (target.status === 'deleted') {
        setRegistrySection('deleted');
      } else {
        setRegistrySection('active');
        const targetYear = String(getDocumentFilterYear(target));
        activeYearManualRef.current = true;
        setSelectedYear(targetYear);
        if (target.docCategory === 'incoming' && selectedCategory === 'outgoing') {
          setSelectedCategory('incoming');
        } else if (target.docCategory !== 'incoming' && selectedCategory === 'incoming') {
          setSelectedCategory('outgoing');
        }
      }
      setDetailModalDoc(target);
    };

    const target =
      documents.find((d) => d.id === initialOpenDocId) ||
      deletedDocuments.find((d) => d.id === initialOpenDocId);
    if (target) {
      alignViewForTarget(target);
      onClearInitialOpenDocId?.();
      return;
    }
    api
      .getDocument(initialOpenDocId)
      .then((res) => {
        if (res.document) {
          alignViewForTarget(res.document);
        }
      })
      .catch(() => {})
      .finally(() => {
        onClearInitialOpenDocId?.();
      });
  }, [initialOpenDocId, documents, deletedDocuments, onClearInitialOpenDocId, selectedCategory]);

  const effectiveBookScope = useMemo(() => {
    if (viewableBooks.length === 1) return viewableBooks[0].id;
    return activeBookId || 'all';
  }, [viewableBooks, activeBookId]);

  const fetchDocumentHistory = useCallback(async () => {
    setHistoryLoading(true);
    setHistoryError(null);
    try {
      const res = await api.getDocumentHistory({
        bookId: effectiveBookScope !== 'all' ? effectiveBookScope : undefined,
        username: historyActorFilter !== 'all' ? historyActorFilter : undefined,
        action: historyActionFilter !== 'all' ? historyActionFilter : undefined,
        fromDate: historyFromDate || undefined,
        toDate: historyToDate || undefined,
        search: historySearch.trim() || undefined,
      });
      setHistoryLogs(res.logs);
      setHistoryActors(res.actors);
      setHistoryTotal(res.total);
    } catch (err: any) {
      setHistoryError(err.message || 'Không tải được lịch sử thao tác.');
    } finally {
      setHistoryLoading(false);
    }
  }, [effectiveBookScope, historyActorFilter, historyActionFilter, historyFromDate, historyToDate, historySearch]);

  useEffect(() => {
    if (registrySection === 'history') {
      fetchDocumentHistory();
    }
  }, [registrySection, fetchDocumentHistory, documents.length, deletedDocuments.length]);

  // Debounced PDF Content Search
  useEffect(() => {
    if (registrySection !== 'active' || !pdfContentMode) {
      setPdfSearchResults(null);
      setPdfSearchError(null);
      return;
    }
    const q = searchTerm.trim();
    if (!q) {
      setPdfSearchResults(null);
      setPdfSearchError(null);
      setPdfSearchLoading(false);
      return;
    }
    let cancelled = false;
    setPdfSearchLoading(true);
    setPdfSearchError(null);
    const timer = setTimeout(() => {
      api
        .searchPdfContent({
          q,
          bookId: effectiveBookScope !== 'all' ? effectiveBookScope : undefined,
          year: selectedYear,
        })
        .then((res) => {
          if (!cancelled) setPdfSearchResults(res);
        })
        .catch((err: any) => {
          if (!cancelled) setPdfSearchError(err?.message || 'Không thể tìm trong nội dung PDF.');
        })
        .finally(() => {
          if (!cancelled) setPdfSearchLoading(false);
        });
    }, 300);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [registrySection, pdfContentMode, searchTerm, effectiveBookScope, selectedYear]);

  // Load Tasks & Assignees when Detail Modal opens
  useEffect(() => {
    if (!detailModalDoc?.id) {
      setDetailTasks([]);
      setDetailAssignees([]);
      setShowAddAttachmentForm(false);
      setAttTargetId(undefined);
      setAttFile(null);
      setAttError(null);
      setDetailTaskError(null);
      return;
    }
    let cancelled = false;
    setDetailTasksLoading(true);
    setShowAddAttachmentForm(false);
    setAttTargetId(undefined);
    setAttFile(null);
    setAttError(null);
    setDetailTaskError(null);

    Promise.all([
      api.getTasks({ docId: detailModalDoc.id }).catch(() => ({ total: 0, tasks: [] as DocumentTask[] })),
      api
        .getTaskAssignees(getDocBookId(detailModalDoc))
        .catch(() => ({ assignees: [] as Array<{ id: string; username: string; fullName: string; displayName: string; role: string }> })),
    ])
      .then(([taskRes, assRes]) => {
        if (cancelled) return;
        setDetailTasks(taskRes.tasks || []);
        setDetailAssignees(assRes.assignees || []);
      })
      .finally(() => {
        if (!cancelled) setDetailTasksLoading(false);
      });

    return () => {
      cancelled = true;
    };
  }, [detailModalDoc?.id]);

  const canRetryOrAttachDoc = (doc: DocumentRecord) => {
    if (doc.status === 'cancelled' || doc.status === 'deleted') return false;
    const bId = getDocBookId(doc);
    return (
      isSystemAdmin(currentUser) ||
      hasBookPermission(currentUser, bId, 'replace_file') ||
      hasBookPermission(currentUser, bId, 'edit') ||
      doc.issuedByUserId === currentUser.id
    );
  };

  const canReplaceDocFile = (doc: DocumentRecord) => {
    if (doc.status === 'cancelled' || doc.status === 'deleted') return false;
    const bId = getDocBookId(doc);
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'replace_file');
  };

  const canEditDoc = (doc: DocumentRecord) => {
    if (doc.status === 'cancelled' || doc.status === 'deleted') return false;
    const bId = getDocBookId(doc);
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'edit');
  };

  const canCancelDoc = (doc: DocumentRecord) => {
    if (doc.status === 'cancelled' || doc.status === 'deleted') return false;
    const bId = getDocBookId(doc);
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'cancel');
  };

  const canDeleteDoc = (doc: DocumentRecord) => {
    if (doc.status === 'deleted') return false;
    const bId = getDocBookId(doc);
    return isSystemAdmin(currentUser) || hasBookPermission(currentUser, bId, 'delete');
  };

  const canRetryExcelSync = (doc: DocumentRecord) => {
    const bId = getDocBookId(doc);
    return (
      isSystemAdmin(currentUser) ||
      hasBookPermission(currentUser, bId, 'delete') ||
      hasBookPermission(currentUser, bId, 'edit') ||
      hasBookPermission(currentUser, bId, 'manage_book')
    );
  };

  const canExportScope = useMemo(() => {
    if (isSystemAdmin(currentUser)) return true;
    if (activeBookId && activeBookId !== 'all') {
      return hasBookPermission(currentUser, activeBookId, 'export');
    }
    return books.some((b) => hasBookPermission(currentUser, b.id, 'export'));
  }, [currentUser, activeBookId, books]);

  const handleOpenDriveFile = async (doc: DocumentRecord, attachmentId?: string) => {
    setFileActionError(null);
    setOpeningFileId(doc.id);
    try {
      const res = await api.openDocumentFile(doc.id, attachmentId);
      if (res.url) {
        window.open(res.url, '_blank', 'noopener,noreferrer');
      }
    } catch (err: any) {
      setFileActionError(err.message || 'Không thể mở liên kết tệp Google Drive.');
    } finally {
      setOpeningFileId(null);
    }
  };

  const handleOpenDetailByDocId = async (docId: string) => {
    const localDoc =
      documents.find((d) => d.id === docId) || deletedDocuments.find((d) => d.id === docId);
    if (localDoc) {
      setDetailModalDoc(localDoc);
      return;
    }
    try {
      const res = await api.getDocument(docId);
      if (res.document) setDetailModalDoc(res.document);
    } catch (err: any) {
      setFileActionError(err?.message || 'Không thể mở chi tiết hồ sơ.');
    }
  };

  const handleUploadMultiAttachment = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!detailModalDoc || attUploading) return;
    if (!attFile) {
      setAttError('Vui lòng chọn tệp để tải lên.');
      return;
    }
    if (attFile.size <= 0) {
      setAttError('Tệp đang bị rỗng (0 bytes).');
      return;
    }
    if (attFile.size > maxWordFileSizeBytes) {
      setAttError(`Dung lượng tệp vượt quá giới hạn tối đa ${maxWordFileSizeMB} MB.`);
      return;
    }
    setAttUploading(true);
    setAttError(null);
    try {
      const res = await api.uploadDocumentAttachment(detailModalDoc.id, {
        file: attFile,
        group: attGroup,
        attachmentId: attTargetId,
        label: attLabel.trim() || undefined,
        setAsOfficial: attSetOfficial,
      });
      setDetailModalDoc(res.document);
      onDocumentUpdated(res.document);
      setShowAddAttachmentForm(false);
      setAttTargetId(undefined);
      setAttLabel('');
      setAttFile(null);
      setAttSetOfficial(false);
    } catch (err: any) {
      setAttError(err?.message || 'Không thể tải tệp lên Google Drive.');
    } finally {
      setAttUploading(false);
    }
  };

  const handleSetOfficialAttachment = async (doc: DocumentRecord, attachmentId: string) => {
    if (settingOfficialId) return;
    setSettingOfficialId(attachmentId);
    setAttError(null);
    try {
      const res = await api.setOfficialAttachment(doc.id, attachmentId);
      setDetailModalDoc(res.document);
      onDocumentUpdated(res.document);
    } catch (err: any) {
      setAttError(err?.message || 'Không thể đặt làm bản chính thức.');
    } finally {
      setSettingOfficialId(null);
    }
  };

  const handleRetryWordTemplateInDetail = async (doc: DocumentRecord) => {
    if (retryingWordTemplate) return;
    setRetryingWordTemplate(true);
    setAttError(null);
    try {
      const res = await api.retryWordDrive(doc.id, doc.templateId);
      setDetailModalDoc(res.document);
      onDocumentUpdated(res.document);
    } catch (err: any) {
      if (err.document) {
        setDetailModalDoc(err.document);
        onDocumentUpdated(err.document);
      }
      setAttError(err?.message || 'Chưa tạo & lưu lại được tệp Word từ mẫu.');
    } finally {
      setRetryingWordTemplate(false);
    }
  };

  const handleConfirmSuggestionInDetail = async (doc: DocumentRecord, sug: PdfTaskSuggestion) => {
    if (confirmingSugId) return;
    setConfirmingSugId(sug.id);
    setDetailTaskError(null);
    try {
      const assigneeId = sugAssigneeMap[sug.id] || currentUser.id;
      const res = await api.createTask({
        docId: doc.id,
        suggestionId: sug.id,
        title: sug.taskTitle,
        dueDate: sug.dueDate,
        rawDeadlineText: sug.rawDeadlineText,
        needsConfirmation: sug.needsConfirmation,
        recipient: sug.recipient,
        sourceExcerpt: sug.sourceExcerpt,
        sourcePage: sug.page,
        assigneeUserId: assigneeId,
      });
      setDetailTasks((prev) => {
        const exists = prev.some((t) => t.id === res.task.id);
        return exists ? prev.map((t) => (t.id === res.task.id ? res.task : t)) : [res.task, ...prev];
      });
    } catch (err: any) {
      setDetailTaskError(err?.message || 'Không thể xác nhận tạo việc.');
    } finally {
      setConfirmingSugId(null);
    }
  };

  const handleUpdateDetailTaskStatus = async (task: DocumentTask, status: DocumentTaskStatus) => {
    if (updatingDetailTaskId || task.status === status) return;
    setUpdatingDetailTaskId(task.id);
    setDetailTaskError(null);
    try {
      const res = await api.updateTask(task.id, { status });
      setDetailTasks((prev) => prev.map((t) => (t.id === task.id ? res.task : t)));
    } catch (err: any) {
      setDetailTaskError(err?.message || 'Không thể cập nhật trạng thái việc.');
    } finally {
      setUpdatingDetailTaskId(null);
    }
  };

  const validateAttachmentFile = (file: File, isIncomingPdf: boolean): string | null => {
    if (isIncomingPdf) {
      if (!file.name.toLowerCase().endsWith('.pdf')) {
        return 'Định dạng tệp không hợp lệ. Vui lòng chọn tệp PDF (.pdf).';
      }
    } else {
      const ext = getWordFileExtension(file.name);
      if (!ext) {
        return 'Định dạng tệp không hợp lệ. Vui lòng chỉ chọn tệp Word có đuôi .doc hoặc .docx.';
      }
    }
    if (file.size <= 0) {
      return 'Tệp đang bị rỗng (0 bytes). Vui lòng chọn tệp khác.';
    }
    if (file.size > maxWordFileSizeBytes) {
      return `Dung lượng tệp (${formatFileSize(file.size)}) vượt quá giới hạn tối đa ${maxWordFileSizeMB} MB.`;
    }
    return null;
  };

  const openUploadModal = (doc: DocumentRecord, mode: 'retry' | 'replace' = 'retry') => {
    setUploadModalDoc(doc);
    setUploadModalMode(mode);
    setUploadModalFile(null);
    setUploadModalError(null);
  };

  const handleUploadAttachmentSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!uploadModalDoc || uploadModalSubmitting) return;

    const isIncoming = uploadModalDoc.docCategory === 'incoming';
    if (!uploadModalFile) {
      setUploadModalError(
        isIncoming
          ? 'Vui lòng chọn tệp PDF (.pdf) để tải lên.'
          : 'Vui lòng chọn tệp Word (.doc hoặc .docx) để tải lên.'
      );
      return;
    }

    const validationErr = validateAttachmentFile(uploadModalFile, isIncoming);
    if (validationErr) {
      setUploadModalError(validationErr);
      return;
    }

    setUploadModalSubmitting(true);
    setUploadModalError(null);
    try {
      const targetName =
        uploadModalMode === 'replace'
          ? buildStandardDriveFileName(
              uploadModalDoc.fullDocNumber,
              uploadModalDoc.docName || uploadModalDoc.title,
              uploadModalFile.name
            )
          : uploadModalDoc.driveFileName;

      const res = isIncoming
        ? await api.uploadPdfToDocument(uploadModalDoc.id, {
            file: uploadModalFile,
            fileName: targetName,
            mode: uploadModalMode,
          })
        : await api.uploadWordToDocument(
            uploadModalDoc.id,
            uploadModalFile,
            targetName,
            uploadModalMode
          );
      onDocumentUpdated(res.document);
      if (detailModalDoc && detailModalDoc.id === res.document.id) {
        setDetailModalDoc(res.document);
      }
      setUploadModalDoc(null);
      setUploadModalFile(null);
    } catch (err: any) {
      const failedDoc: DocumentRecord = err.document || {
        ...uploadModalDoc,
        driveStatus: uploadModalMode === 'replace' ? uploadModalDoc.driveStatus : 'failed',
        driveError: err.message || 'Lỗi tải tệp lên Google Drive',
      };
      onDocumentUpdated(failedDoc);
      setUploadModalDoc(failedDoc);
      if (detailModalDoc && detailModalDoc.id === failedDoc.id) {
        setDetailModalDoc(failedDoc);
      }
      setUploadModalError(
        isIncoming
          ? `Lưu tệp thất bại: ${err.message || 'Lỗi kết nối Google Drive'}`
          : `Đã cấp số ${uploadModalDoc.fullDocNumber} · Chưa lưu tệp lên Drive: ${err.message || 'Lỗi kết nối Google Drive'}`
      );
    } finally {
      setUploadModalSubmitting(false);
    }
  };

  // Derive unique years and signers (using outgoing sequence year or incoming issueDate year)
  const availableYears = useMemo(() => {
    const yearSet = new Set<number>([currentVnYear]);
    documents.forEach((d) => {
      const fy = getDocumentFilterYear(d);
      if (typeof fy === 'number') yearSet.add(fy);
    });
    return Array.from(yearSet).sort((a, b) => b - a);
  }, [documents, currentVnYear]);

  const availableSigners = useMemo(() => {
    return Array.from(new Set(documents.map((d) => d.signer))).filter(Boolean);
  }, [documents]);

  // Client-side filtering & sorting (matches /api/export-excel 1-to-1)
  const filteredDocuments = useMemo(() => {
    let result = [...documents];

    const effectiveBookFilter =
      viewableBooks.length === 1
        ? viewableBooks[0].id
        : activeBookId;
    if (effectiveBookFilter && effectiveBookFilter !== 'all') {
      result = result.filter((d) => getDocBookId(d) === effectiveBookFilter);
    }

    if (selectedCategory === 'incoming') {
      result = result.filter((d) => d.docCategory === 'incoming');
    } else if (selectedCategory === 'outgoing') {
      result = result.filter((d) => d.docCategory !== 'incoming');
    }

    if (selectedYear === 'unknown') {
      result = result.filter((d) => getDocumentFilterYear(d) === 'unknown');
    } else if (selectedYear !== 'all') {
      const targetYr = parseInt(selectedYear, 10);
      if (Number.isFinite(targetYr)) {
        result = result.filter((d) => getDocumentFilterYear(d) === targetYr);
      }
    }

    if (selectedMonth !== 'all') {
      const m = parseInt(selectedMonth, 10);
      result = result.filter((d) => {
        const parts = (d.issueDate || '').split('/');
        return parts.length >= 2 && parseInt(parts[1], 10) === m;
      });
    }

    if (selectedType !== 'all') {
      result = result.filter((d) => d.typeCode === selectedType);
    }

    if (selectedSigner !== 'all') {
      result = result.filter((d) => d.signer.toLowerCase() === selectedSigner.toLowerCase());
    }

    if (selectedStatus !== 'all') {
      result = result.filter((d) => d.status === selectedStatus);
    }

    if (selectedDriveStatus !== 'all') {
      result = result.filter((d) => (d.driveStatus || 'none') === selectedDriveStatus);
    }

    if (searchTerm.trim()) {
      const term = searchTerm.trim().toLowerCase();
      result = result.filter(
        (d) =>
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

    if (sortBy === 'newest') {
      result.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
    } else if (sortBy === 'oldest') {
      result.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
    } else if (sortBy === 'number_asc') {
      result.sort((a, b) => a.sequenceNumber - b.sequenceNumber);
    } else if (sortBy === 'number_desc') {
      result.sort((a, b) => b.sequenceNumber - a.sequenceNumber);
    }

    return result;
  }, [
    documents,
    viewableBooks,
    activeBookId,
    selectedCategory,
    selectedYear,
    selectedMonth,
    selectedType,
    selectedSigner,
    selectedStatus,
    selectedDriveStatus,
    searchTerm,
    sortBy,
  ]);

  const bookYearTotalCount = useMemo(() => {
    const effectiveBookFilter =
      viewableBooks.length === 1
        ? viewableBooks[0].id
        : activeBookId;
    return documents.filter((d) => {
      if (effectiveBookFilter && effectiveBookFilter !== 'all' && getDocBookId(d) !== effectiveBookFilter) {
        return false;
      }
      if (selectedYear === 'unknown') {
        return getDocumentFilterYear(d) === 'unknown';
      }
      if (selectedYear !== 'all') {
        const targetYr = parseInt(selectedYear, 10);
        if (Number.isFinite(targetYr)) {
          return getDocumentFilterYear(d) === targetYr;
        }
      }
      return true;
    }).length;
  }, [documents, viewableBooks, activeBookId, selectedYear]);

  const activeExtraFilterCount = useMemo(() => {
    let count = 0;
    if (selectedMonth !== 'all') count += 1;
    if (selectedType !== 'all') count += 1;
    if (selectedSigner !== 'all') count += 1;
    if (selectedStatus !== 'all') count += 1;
    if (selectedDriveStatus !== 'all') count += 1;
    if (sortBy !== 'newest') count += 1;
    return count;
  }, [selectedMonth, selectedType, selectedSigner, selectedStatus, selectedDriveStatus, sortBy]);

  const activeTotalPages = Math.max(1, Math.ceil(filteredDocuments.length / activePageSize));
  const safeActivePage = Math.min(Math.max(1, activePage), activeTotalPages);
  const paginatedActiveDocuments = useMemo(() => {
    const start = (safeActivePage - 1) * activePageSize;
    return filteredDocuments.slice(start, start + activePageSize);
  }, [filteredDocuments, safeActivePage, activePageSize]);

  useEffect(() => {
    setActivePage(1);
  }, [
    effectiveBookScope,
    searchTerm,
    selectedCategory,
    selectedYear,
    selectedMonth,
    selectedType,
    selectedSigner,
    selectedStatus,
    selectedDriveStatus,
    sortBy,
    activePageSize,
  ]);

  useEffect(() => {
    if (activePage > activeTotalPages) {
      setActivePage(activeTotalPages);
    }
  }, [activePage, activeTotalPages]);

  const handleCopy = (id: string, text: string) => {
    navigator.clipboard?.writeText(text).catch(() => {});
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const openEditModal = (doc: DocumentRecord) => {
    setEditModalDoc(doc);
    setEditDocName(doc.docName || doc.title);
    setEditTitle(doc.title);
    setEditSigner(doc.signer);
    setEditSignerRole(doc.signerRole);
    setEditNotes(doc.notes || '');
    setEditIssuingAgency(doc.issuingAgency || '');
    setEditIssueDate(normalizeDateOnlyDdMmYyyy(doc.issueDate || ''));
    setEditSignDate(normalizeDateOnlyDdMmYyyy(doc.signDate || ''));
    setEditSequenceNumber(doc.sequenceNumber);
    setEditFullDocNumber(doc.fullDocNumber);
    setEditError(null);
  };

  const handleSaveEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editModalDoc) return;

    if (!editDocName.trim()) {
      setEditError('Vui lòng nhập tên văn bản (thông tin bắt buộc)');
      return;
    }

    const docBookId = getDocBookId(editModalDoc);
    const canManageBook = isSystemAdmin(currentUser) || hasBookPermission(currentUser, docBookId, 'manage_book');

    setEditSubmitting(true);
    setEditError(null);
    try {
      const payload: any = {
        docName: editDocName.trim(),
        title: editTitle.trim(),
        signer: editSigner.trim(),
        signerRole: editSignerRole.trim(),
        notes: editNotes.trim(),
      };

      if (editModalDoc.docCategory === 'incoming') {
        payload.issuingAgency = editIssuingAgency.trim();
        payload.issueDate = normalizeDateOnlyDdMmYyyy(editIssueDate.trim());
        payload.signDate = normalizeDateOnlyDdMmYyyy(editSignDate.trim());
        payload.needsReview = false;
      }

      if (canManageBook) {
        if (editModalDoc.docCategory !== 'incoming') {
          payload.sequenceNumber = editSequenceNumber;
        }
        payload.fullDocNumber = editFullDocNumber.trim();
      } else if (editModalDoc.docCategory === 'incoming') {
        payload.fullDocNumber = editFullDocNumber.trim();
      }

      const res = await api.updateDocument(editModalDoc.id, payload);
      onDocumentUpdated(res.document);
      if (detailModalDoc && detailModalDoc.id === res.document.id) {
        setDetailModalDoc(res.document);
      }
      setEditModalDoc(null);
      onRefresh();
    } catch (err: any) {
      setEditError(err.message || 'Lỗi khi cập nhật văn bản');
    } finally {
      setEditSubmitting(false);
    }
  };

  const openCancelModal = (doc: DocumentRecord) => {
    setCancelModalDoc(doc);
    setCancelReason('');
    setCancelError(null);
  };

  const handleSaveCancel = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!cancelModalDoc) return;
    if (!cancelReason.trim()) {
      setCancelError('Bắt buộc nhập lý do hủy văn bản');
      return;
    }

    setCancelSubmitting(true);
    setCancelError(null);
    try {
      const res = await api.cancelDocument(cancelModalDoc.id, cancelReason.trim());
      onDocumentUpdated(res.document);
      if (detailModalDoc && detailModalDoc.id === res.document.id) {
        setDetailModalDoc(res.document);
      }
      setCancelModalDoc(null);
    } catch (err: any) {
      setCancelError(err.message || 'Lỗi khi hủy số văn bản');
    } finally {
      setCancelSubmitting(false);
    }
  };

  const openDeleteModal = (doc: DocumentRecord) => {
    setDeleteModalDoc(doc);
    setDeleteReason('');
    setDeleteConfirmed(false);
    setDeleteError(null);
  };

  const handleConfirmDelete = async (e?: React.FormEvent) => {
    e?.preventDefault();
    if (!deleteModalDoc || deleteSubmitting) return;
    if (!deleteReason.trim()) {
      setDeleteError('Vui lòng nhập lý do xóa văn bản.');
      return;
    }
    if (!deleteConfirmed) {
      setDeleteError('Vui lòng đánh dấu xác nhận trước khi xóa văn bản.');
      return;
    }
    setDeleteSubmitting(true);
    setDeleteError(null);
    try {
      const res = await api.deleteDocument(deleteModalDoc.id, deleteReason.trim(), true);
      if (res.document) {
        onDocumentUpdated(res.document);
        if (detailModalDoc && detailModalDoc.id === res.document.id) {
          setDetailModalDoc(res.document);
        }
      }
      setDeleteModalDoc(null);
      setDeleteReason('');
      setDeleteConfirmed(false);
      onRefresh();
    } catch (err: any) {
      setDeleteError(err.message || 'Lỗi khi xóa văn bản');
    } finally {
      setDeleteSubmitting(false);
    }
  };

  const handleRetryExcelSync = async (doc: DocumentRecord) => {
    if (retryingSyncDocId) return;
    setRetryingSyncDocId(doc.id);
    setSyncRetryNotice(null);
    try {
      const res = await api.retryDocumentExcelSync(doc.id);
      if (res.document) {
        onDocumentUpdated(res.document);
        if (detailModalDoc && detailModalDoc.id === res.document.id) {
          setDetailModalDoc(res.document);
        }
      }
      setSyncRetryNotice({
        type: 'success',
        message: res.message || `Đã đồng bộ lại văn bản ${doc.fullDocNumber} sang sheet VB đã xóa trên Google Trang tính.`,
      });
    } catch (err: any) {
      if (err.document) {
        onDocumentUpdated(err.document);
        if (detailModalDoc && detailModalDoc.id === err.document.id) {
          setDetailModalDoc(err.document);
        }
      }
      setSyncRetryNotice({
        type: 'error',
        message: err.message || 'Đồng bộ lại Google Trang tính chưa thành công. Vui lòng thử lại sau.',
      });
    } finally {
      setRetryingSyncDocId(null);
    }
  };

  // Load archived deleted docs when Admin selects "Đã lưu trữ"
  const fetchArchivedDeletedDocs = useCallback(async () => {
    if (!isAdmin) return;
    setLoadingArchivedDeleted(true);
    setArchivedDeletedError(null);
    try {
      const res = await api.getDocuments({
        status: 'deleted',
        deletedArchive: 'archived',
      });
      setArchivedDeletedDocs(res.documents);
    } catch (err: any) {
      setArchivedDeletedError(err.message || 'Không tải được danh sách văn bản đã lưu trữ.');
    } finally {
      setLoadingArchivedDeleted(false);
    }
  }, [isAdmin]);

  useEffect(() => {
    if (registrySection === 'deleted' && deletedSyncFilter === 'archived' && isAdmin) {
      fetchArchivedDeletedDocs();
    }
  }, [registrySection, deletedSyncFilter, isAdmin, fetchArchivedDeletedDocs]);

  const unarchivedDeletedInBook = useMemo(() => {
    return deletedDocuments.filter((d) => {
      if (d.deletedListArchived) return false;
      if (effectiveBookScope && effectiveBookScope !== 'all') {
        return getDocBookId(d) === effectiveBookScope;
      }
      return true;
    });
  }, [deletedDocuments, effectiveBookScope]);

  const pendingExcelSyncCount = useMemo(
    () =>
      unarchivedDeletedInBook.filter(
        (d) => d.excelDeleteSyncStatus === 'failed' || d.excelDeleteSyncStatus === 'pending'
      ).length,
    [unarchivedDeletedInBook]
  );

  const syncedUnarchivedInBookCount = useMemo(
    () => unarchivedDeletedInBook.filter((d) => d.excelDeleteSyncStatus === 'synced').length,
    [unarchivedDeletedInBook]
  );

  const deletedSourceList = useMemo(() => {
    if (deletedSyncFilter === 'archived' && isAdmin) {
      return archivedDeletedDocs;
    }
    return deletedDocuments.filter((d) => !d.deletedListArchived);
  }, [deletedSyncFilter, isAdmin, archivedDeletedDocs, deletedDocuments]);

  const deletedAvailableYears = useMemo(() => {
    const yearSet = new Set<number>([currentVnYear]);
    deletedDocuments.forEach((d) => {
      const fy = getDocumentFilterYear(d);
      if (typeof fy === 'number') yearSet.add(fy);
    });
    archivedDeletedDocs.forEach((d) => {
      const fy = getDocumentFilterYear(d);
      if (typeof fy === 'number') yearSet.add(fy);
    });
    return Array.from(yearSet).sort((a, b) => b - a);
  }, [deletedDocuments, archivedDeletedDocs, currentVnYear]);

  const availableDeletedActors = useMemo(() => {
    const map = new Map<string, string>();
    deletedSourceList.forEach((d) => {
      const key = d.deletedByUsername || d.deletedByUserId || '';
      if (key && !map.has(key)) {
        map.set(key, d.deletedBy || key);
      }
    });
    return Array.from(map.entries()).map(([username, label]) => ({ username, label }));
  }, [deletedSourceList]);

  const filteredDeletedDocuments = useMemo(() => {
    let list = [...deletedSourceList];
    if (effectiveBookScope && effectiveBookScope !== 'all') {
      list = list.filter((d) => getDocBookId(d) === effectiveBookScope);
    }
    if (deletedDocYear === 'unknown') {
      list = list.filter((d) => getDocumentFilterYear(d) === 'unknown');
    } else if (deletedDocYear !== 'all') {
      const targetYr = parseInt(deletedDocYear, 10);
      if (Number.isFinite(targetYr)) {
        list = list.filter((d) => getDocumentFilterYear(d) === targetYr);
      }
    }
    if (deletedByFilter !== 'all') {
      list = list.filter(
        (d) => (d.deletedByUsername || d.deletedByUserId || '') === deletedByFilter
      );
    }
    if (deletedSyncFilter === 'synced') {
      list = list.filter((d) => d.excelDeleteSyncStatus === 'synced');
    } else if (deletedSyncFilter === 'pending_sync') {
      list = list.filter(
        (d) => d.excelDeleteSyncStatus === 'failed' || d.excelDeleteSyncStatus === 'pending'
      );
    }
    if (deletedFromDate) {
      list = list.filter((d) => {
        const key = d.deletedAt
          ? new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Ho_Chi_Minh' }).format(new Date(d.deletedAt))
          : '';
        return key >= deletedFromDate;
      });
    }
    if (deletedToDate) {
      list = list.filter((d) => {
        const key = d.deletedAt
          ? new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Ho_Chi_Minh' }).format(new Date(d.deletedAt))
          : '';
        return key <= deletedToDate;
      });
    }
    if (deletedSearch.trim()) {
      const term = deletedSearch.trim().toLowerCase();
      list = list.filter(
        (d) =>
          d.fullDocNumber.toLowerCase().includes(term) ||
          (d.docName || d.title || '').toLowerCase().includes(term) ||
          (d.deleteReason || '').toLowerCase().includes(term) ||
          (d.deletedBy || '').toLowerCase().includes(term) ||
          d.id.toLowerCase().includes(term)
      );
    }
    list.sort((a, b) => new Date(b.deletedAt || b.createdAt).getTime() - new Date(a.deletedAt || a.createdAt).getTime());
    return list;
  }, [
    deletedSourceList,
    effectiveBookScope,
    deletedDocYear,
    deletedByFilter,
    deletedSyncFilter,
    deletedFromDate,
    deletedToDate,
    deletedSearch,
  ]);

  const deletedTotalPages = Math.max(1, Math.ceil(filteredDeletedDocuments.length / deletedPageSize));
  const safeDeletedPage = Math.min(Math.max(1, deletedPage), deletedTotalPages);
  const paginatedDeletedDocuments = useMemo(() => {
    const start = (safeDeletedPage - 1) * deletedPageSize;
    return filteredDeletedDocuments.slice(start, start + deletedPageSize);
  }, [filteredDeletedDocuments, safeDeletedPage, deletedPageSize]);

  useEffect(() => {
    setDeletedPage(1);
  }, [
    effectiveBookScope,
    deletedSearch,
    deletedByFilter,
    deletedDocYear,
    deletedSyncFilter,
    deletedFromDate,
    deletedToDate,
    deletedPageSize,
  ]);

  useEffect(() => {
    if (deletedPage > deletedTotalPages) {
      setDeletedPage(deletedTotalPages);
    }
  }, [deletedPage, deletedTotalPages]);

  useEffect(() => {
    setSelectedDeletedIds([]);
  }, [effectiveBookScope, deletedSyncFilter]);

  const selectedDeletedDocs = useMemo(() => {
    const idSet = new Set(selectedDeletedIds);
    return unarchivedDeletedInBook.filter((d) => idSet.has(d.id));
  }, [selectedDeletedIds, unarchivedDeletedInBook]);

  const selectedSyncedCount = useMemo(
    () => selectedDeletedDocs.filter((d) => d.excelDeleteSyncStatus === 'synced').length,
    [selectedDeletedDocs]
  );

  const selectedUnsyncedCount = selectedDeletedDocs.length - selectedSyncedCount;

  const canBulkSyncScope = useMemo(() => {
    if (isAdmin) return true;
    if (effectiveBookScope && effectiveBookScope !== 'all') {
      return (
        hasBookPermission(currentUser, effectiveBookScope, 'delete') ||
        hasBookPermission(currentUser, effectiveBookScope, 'edit') ||
        hasBookPermission(currentUser, effectiveBookScope, 'manage_book')
      );
    }
    return books.some(
      (b) =>
        hasBookPermission(currentUser, b.id, 'delete') ||
        hasBookPermission(currentUser, b.id, 'edit') ||
        hasBookPermission(currentUser, b.id, 'manage_book')
    );
  }, [isAdmin, effectiveBookScope, currentUser, books]);

  const currentScopeBookName = useMemo(() => {
    if (effectiveBookScope && effectiveBookScope !== 'all') {
      return bookMap.get(effectiveBookScope)?.name || 'Sổ đang chọn';
    }
    return `Tất cả sổ có quyền (${viewableBooks.length})`;
  }, [effectiveBookScope, bookMap, viewableBooks.length]);

  const allOnPageSelected = useMemo(() => {
    if (paginatedDeletedDocuments.length === 0) return false;
    const idSet = new Set(selectedDeletedIds);
    return paginatedDeletedDocuments.every((d) => idSet.has(d.id));
  }, [paginatedDeletedDocuments, selectedDeletedIds]);

  const toggleSelectAllOnDeletedPage = () => {
    const pageIds = paginatedDeletedDocuments.map((d) => d.id);
    if (allOnPageSelected) {
      const pageSet = new Set(pageIds);
      setSelectedDeletedIds((prev) => prev.filter((id) => !pageSet.has(id)));
    } else {
      setSelectedDeletedIds((prev) => Array.from(new Set([...prev, ...pageIds])));
    }
  };

  const toggleSelectDeletedDoc = (docId: string) => {
    setSelectedDeletedIds((prev) =>
      prev.includes(docId) ? prev.filter((id) => id !== docId) : [...prev, docId]
    );
  };

  const handleConfirmBulkSync = async () => {
    if (!bulkSyncConfirmMode || bulkSyncSubmitting) return;
    setBulkSyncSubmitting(true);
    setBulkSyncResult(null);
    setSyncRetryNotice(null);
    try {
      const res = await api.bulkSyncDeletedDocuments({
        bookId: effectiveBookScope,
        mode: bulkSyncConfirmMode,
        docIds: bulkSyncConfirmMode === 'selected' ? selectedDeletedIds : undefined,
      });
      if (Array.isArray(res.documents)) {
        res.documents.forEach((updated) => {
          onDocumentUpdated(updated);
          if (detailModalDoc && detailModalDoc.id === updated.id) {
            setDetailModalDoc(updated);
          }
        });
      }
      setBulkSyncResult(res);
      setBulkSyncConfirmMode(null);
      if (res.succeededIds.length > 0) {
        const succeededSet = new Set(res.succeededIds);
        setSelectedDeletedIds((prev) => prev.filter((id) => !succeededSet.has(id)));
      }
      onRefresh();
    } catch (err: any) {
      setSyncRetryNotice({
        type: 'error',
        message: err.message || 'Đồng bộ hàng loạt Google Trang tính thất bại.',
      });
      setBulkSyncConfirmMode(null);
    } finally {
      setBulkSyncSubmitting(false);
    }
  };

  const handleConfirmArchiveDeleted = async () => {
    if (!isAdmin || archiveSubmitting) return;
    setArchiveSubmitting(true);
    setArchiveError(null);
    setArchiveNotice(null);
    try {
      const res = await api.archiveDeletedDocuments({
        bookId: effectiveBookScope,
        mode: archiveMode,
        docIds: archiveMode === 'selected' ? selectedDeletedIds : undefined,
      });
      setArchiveNotice(res);
      setArchiveModalOpen(false);
      if (res.archivedIds.length > 0) {
        const archivedSet = new Set(res.archivedIds);
        setSelectedDeletedIds((prev) => prev.filter((id) => !archivedSet.has(id)));
      }
      onRefresh();
      if (deletedSyncFilter === 'archived') {
        fetchArchivedDeletedDocs();
      }
    } catch (err: any) {
      setArchiveError(err.message || 'Không thể dọn danh sách VB đã xóa.');
    } finally {
      setArchiveSubmitting(false);
    }
  };

  // History pagination & Admin History Cleanup handlers
  const historyTotalPages = Math.max(1, Math.ceil(historyLogs.length / historyPageSize));
  const safeHistoryPage = Math.min(Math.max(1, historyPage), historyTotalPages);
  const paginatedHistoryLogs = useMemo(() => {
    const start = (safeHistoryPage - 1) * historyPageSize;
    return historyLogs.slice(start, start + historyPageSize);
  }, [historyLogs, safeHistoryPage, historyPageSize]);

  useEffect(() => {
    setHistoryPage(1);
  }, [
    effectiveBookScope,
    historySearch,
    historyActorFilter,
    historyActionFilter,
    historyFromDate,
    historyToDate,
    historyPageSize,
  ]);

  useEffect(() => {
    if (historyPage > historyTotalPages) {
      setHistoryPage(historyTotalPages);
    }
  }, [historyPage, historyTotalPages]);

  const openHistoryCleanupModal = () => {
    if (!isAdmin) return;
    const defaultBook =
      effectiveBookScope && effectiveBookScope !== 'all'
        ? effectiveBookScope
        : books[0]?.id || '';
    setCleanupBookId(defaultBook);
    setCleanupScope('today');
    setCleanupBeforeDate(initialVnTime.dateKey);
    setCleanupConfirmText('');
    setCleanupPreview(null);
    setCleanupPreviewError(null);
    setCleanupModalOpen(true);
  };

  useEffect(() => {
    if (!cleanupModalOpen || !isAdmin || !cleanupBookId) return;
    if (cleanupScope === 'before_date' && !cleanupBeforeDate) {
      setCleanupPreview(null);
      return;
    }
    let cancelled = false;
    setCleanupPreviewLoading(true);
    setCleanupPreviewError(null);
    api
      .previewHistoryCleanup({
        bookId: cleanupBookId,
        scope: cleanupScope,
        beforeDate: cleanupScope === 'before_date' ? cleanupBeforeDate : undefined,
      })
      .then((res) => {
        if (!cancelled) {
          setCleanupPreview(res);
        }
      })
      .catch((err: any) => {
        if (!cancelled) {
          setCleanupPreview(null);
          setCleanupPreviewError(err.message || 'Không thể xem trước số bản ghi lịch sử sẽ xóa.');
        }
      })
      .finally(() => {
        if (!cancelled) {
          setCleanupPreviewLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [cleanupModalOpen, isAdmin, cleanupBookId, cleanupScope, cleanupBeforeDate]);

  const handleConfirmHistoryCleanup = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin || cleanupSubmitting || !cleanupBookId) return;
    if (cleanupScope === 'all' && cleanupConfirmText.trim() !== 'XÓA LỊCH SỬ') {
      setCleanupPreviewError('Vui lòng nhập chính xác cụm từ xác nhận: XÓA LỊCH SỬ');
      return;
    }
    setCleanupSubmitting(true);
    setCleanupPreviewError(null);
    try {
      const res = await api.cleanupDocumentHistory({
        bookId: cleanupBookId,
        scope: cleanupScope,
        beforeDate: cleanupScope === 'before_date' ? cleanupBeforeDate : undefined,
        confirmationText: cleanupScope === 'all' ? cleanupConfirmText.trim() : undefined,
      });
      setCleanupModalOpen(false);
      setCleanupNotice({
        type: 'success',
        message: res.message,
      });
      await fetchDocumentHistory();
      onRefresh();
    } catch (err: any) {
      setCleanupPreviewError(err.message || 'Dọn lịch sử thao tác thất bại.');
    } finally {
      setCleanupSubmitting(false);
    }
  };

  // Export Excel preserving ALL active filters + sort order + permissions (ignores pagination)
  const handleExportExcel = async () => {
    setIsExporting(true);
    setExportError(null);
    try {
      await api.downloadExcel({
        bookId: activeBookId && activeBookId !== 'all' ? activeBookId : undefined,
        year: selectedYear,
        month: selectedMonth !== 'all' ? selectedMonth : undefined,
        category: selectedCategory !== 'all' ? selectedCategory : undefined,
        typeCode: selectedType !== 'all' ? selectedType : undefined,
        signer: selectedSigner !== 'all' ? selectedSigner : undefined,
        status: selectedStatus !== 'all' ? selectedStatus : undefined,
        driveStatus: selectedDriveStatus !== 'all' ? selectedDriveStatus : undefined,
        search: searchTerm.trim() || undefined,
        sort: sortBy,
      });
    } catch (err: any) {
      setExportError(err.message || 'Không thể tải file Excel. Vui lòng kiểm tra quyền hoặc kết nối.');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="max-w-7xl mx-auto py-5 sm:py-8 px-3 sm:px-6 lg:px-8 space-y-5 sm:space-y-6">
      {/* Header bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
            Sổ Theo Dõi Văn Bản
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
            Tra cứu nhanh văn bản đi và văn bản đến, theo dõi trạng thái tệp trên Google Drive và xuất báo cáo Excel chuẩn.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0 w-full sm:w-auto">
          <span className="inline-flex items-center gap-1.5 px-3 py-2 min-h-[42px] bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700">
            <BookOpen className="w-4 h-4 text-blue-700 shrink-0" />
            <span>{currentScopeBookName}</span>
          </span>

          {canExportScope && (
            <button
              type="button"
              onClick={handleExportExcel}
              disabled={isExporting}
              className="flex-1 sm:flex-initial justify-center flex items-center gap-2 min-h-[42px] px-4 py-2 text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 disabled:opacity-60 text-white rounded-lg shadow-xs transition-colors whitespace-nowrap"
              title="Xuất đúng danh sách đang lọc & sắp xếp hiện tại ra Excel .xlsx"
            >
              {isExporting ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin shrink-0" />
                  <span>Đang xuất Excel...</span>
                </>
              ) : (
                <>
                  <FileSpreadsheet className="w-4 h-4 shrink-0" />
                  <span>Tải Excel .xlsx</span>
                </>
              )}
            </button>
          )}
        </div>
      </div>

      {(exportError || fileActionError) && (
        <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
            <span className="break-words min-w-0">{exportError || fileActionError}</span>
          </div>
          <button
            type="button"
            onClick={() => {
              setExportError(null);
              setFileActionError(null);
            }}
            className="text-rose-600 hover:text-rose-800 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Registry Sub-navigation: Danh sách hiệu lực / VB đã xóa / Lịch sử thao tác */}
      <div className="flex items-center gap-2 overflow-x-auto pb-1">
        <button
          type="button"
          onClick={() => setRegistrySection('active')}
          className={`inline-flex items-center gap-2 px-3.5 py-2 min-h-[40px] rounded-lg text-xs font-bold border transition-colors shrink-0 ${
            registrySection === 'active'
              ? 'bg-sky-900 text-white border-sky-900 shadow-xs'
              : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
          }`}
        >
          <FileText className="w-3.5 h-3.5 shrink-0" />
          <span>Sổ văn bản hiệu lực</span>
          <span
            className={`font-mono text-[11px] px-1.5 py-0.5 rounded ${
              registrySection === 'active' ? 'bg-sky-800 text-sky-100' : 'bg-slate-100 text-slate-700'
            }`}
          >
            {documents.length}
          </span>
        </button>

        <button
          type="button"
          onClick={() => setRegistrySection('deleted')}
          className={`inline-flex items-center gap-2 px-3.5 py-2 min-h-[40px] rounded-lg text-xs font-bold border transition-colors shrink-0 ${
            registrySection === 'deleted'
              ? 'bg-rose-800 text-white border-rose-800 shadow-xs'
              : 'bg-white text-slate-700 border-slate-200 hover:bg-rose-50/50'
          }`}
        >
          <Trash2 className="w-3.5 h-3.5 shrink-0" />
          <span>VB đã xóa</span>
          <span
            className={`font-mono text-[11px] px-1.5 py-0.5 rounded ${
              registrySection === 'deleted' ? 'bg-rose-700 text-rose-100' : 'bg-slate-100 text-slate-700'
            }`}
          >
            {unarchivedDeletedInBook.length}
          </span>
          {pendingExcelSyncCount > 0 && (
            <span
              className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300"
              title="Có văn bản chờ đồng bộ sang sheet VB đã xóa trên Google Trang tính"
            >
              Chờ đồng bộ Google Trang tính: {pendingExcelSyncCount}
            </span>
          )}
        </button>

        <button
          type="button"
          onClick={() => setRegistrySection('history')}
          className={`inline-flex items-center gap-2 px-3.5 py-2 min-h-[40px] rounded-lg text-xs font-bold border transition-colors shrink-0 ${
            registrySection === 'history'
              ? 'bg-sky-900 text-white border-sky-900 shadow-xs'
              : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
          }`}
        >
          <History className="w-3.5 h-3.5 shrink-0" />
          <span>Lịch sử thao tác</span>
          {historyTotal > 0 && (
            <span
              className={`font-mono text-[11px] px-1.5 py-0.5 rounded ${
                registrySection === 'history' ? 'bg-sky-800 text-sky-100' : 'bg-slate-100 text-slate-700'
              }`}
            >
              {historyTotal}
            </span>
          )}
        </button>
      </div>

      {registrySection === 'active' && (
        <>
      {/* Compact Primary Filter Bar + Expandable Bộ lọc */}
      <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
        {/* Row 1: Tìm kiếm + Tìm trong PDF + Nhóm văn bản + Nút Bộ lọc */}
        <div className="flex flex-col lg:flex-row lg:items-center gap-2.5">
          <div className="relative flex-1 min-w-0">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <input
              type="text"
              aria-label="Tìm kiếm văn bản"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder={
                pdfContentMode
                  ? 'Tìm từ khóa trong nội dung bên trong tệp PDF...'
                  : 'Tìm số VB, tên VB, cơ quan, tên tệp...'
              }
              className="w-full min-h-[44px] h-11 sm:h-10 pl-9 pr-9 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
            />
            {searchTerm && (
              <button
                type="button"
                aria-label="Xóa từ khóa tìm kiếm"
                onClick={() => setSearchTerm('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-700"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          <button
            type="button"
            onClick={() => setPdfContentMode((prev) => !prev)}
            className={`inline-flex items-center justify-center gap-1.5 min-h-[44px] sm:min-h-[40px] px-3 py-2 rounded-lg border text-xs font-semibold transition-colors shrink-0 whitespace-nowrap ${
              pdfContentMode
                ? 'bg-purple-700 text-white border-purple-700 shadow-2xs'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-purple-50 hover:text-purple-900'
            }`}
          >
            <FileText className="w-3.5 h-3.5 shrink-0" />
            <span>Tìm trong nội dung PDF</span>
          </button>

          <div
            role="group"
            aria-label="Chọn nhóm văn bản"
            className="grid grid-cols-3 sm:inline-flex bg-slate-100 p-1 rounded-lg border border-slate-200 shrink-0"
          >
            <button
              type="button"
              onClick={() => setSelectedCategory('all')}
              className={`min-h-[38px] px-3 py-1.5 rounded-md text-xs font-semibold transition-colors whitespace-nowrap ${
                selectedCategory === 'all'
                  ? 'bg-white text-slate-900 shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Tất cả
            </button>
            <button
              type="button"
              onClick={() => setSelectedCategory('outgoing')}
              className={`min-h-[38px] px-3 py-1.5 rounded-md text-xs font-semibold transition-colors whitespace-nowrap inline-flex items-center justify-center gap-1.5 ${
                selectedCategory === 'outgoing'
                  ? 'bg-blue-700 text-white shadow-2xs'
                  : 'text-blue-800 hover:bg-blue-50'
              }`}
            >
              <span className={`w-1.5 h-1.5 rounded-full ${selectedCategory === 'outgoing' ? 'bg-white' : 'bg-blue-600'}`} />
              <span>Văn bản đi</span>
            </button>
            <button
              type="button"
              onClick={() => setSelectedCategory('incoming')}
              className={`min-h-[38px] px-3 py-1.5 rounded-md text-xs font-semibold transition-colors whitespace-nowrap inline-flex items-center justify-center gap-1.5 ${
                selectedCategory === 'incoming'
                  ? 'bg-purple-700 text-white shadow-2xs'
                  : 'text-purple-800 hover:bg-purple-50'
              }`}
            >
              <span className={`w-1.5 h-1.5 rounded-full ${selectedCategory === 'incoming' ? 'bg-white' : 'bg-purple-600'}`} />
              <span>Văn bản đến</span>
            </button>
          </div>

          <button
            type="button"
            onClick={() => setShowAdvancedFilters((prev) => !prev)}
            className={`inline-flex items-center justify-center gap-1.5 min-h-[44px] sm:min-h-[40px] px-3.5 py-2 rounded-lg border text-xs font-semibold transition-colors shrink-0 whitespace-nowrap ${
              showAdvancedFilters || activeExtraFilterCount > 0
                ? 'bg-blue-50 text-blue-900 border-blue-300'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
          >
            <SlidersHorizontal className="w-3.5 h-3.5 shrink-0" />
            <span>Bộ lọc</span>
            {activeExtraFilterCount > 0 && (
              <span className="font-mono text-[11px] font-bold px-1.5 py-0.5 rounded bg-blue-700 text-white">
                {activeExtraFilterCount}
              </span>
            )}
          </button>
        </div>

        {/* Row 2: Năm văn bản (Năm hiện tại + 2 năm trước, Năm cũ hơn, Tất cả các năm, Chưa xác định năm) */}
        {(() => {
          const quickYears = [currentVnYear, currentVnYear - 1, currentVnYear - 2];
          const olderYears = availableYears.filter((y) => !quickYears.includes(y));
          const isOlderYearSelected =
            selectedYear !== 'all' &&
            selectedYear !== 'unknown' &&
            !quickYears.includes(parseInt(selectedYear, 10));

          const selectYearVal = (val: string) => {
            activeYearManualRef.current = true;
            setSelectedYear(val);
          };

          return (
            <div className="flex flex-wrap items-center gap-1.5 pt-1">
              <span className="text-xs font-bold text-slate-700 mr-1 whitespace-nowrap">Năm văn bản:</span>
              {quickYears.map((yr) => {
                const active = selectedYear === String(yr);
                return (
                  <button
                    key={yr}
                    type="button"
                    onClick={() => selectYearVal(String(yr))}
                    className={`min-h-[34px] px-2.5 py-1 rounded-md border text-xs font-semibold transition-colors tabular-nums whitespace-nowrap ${
                      active
                        ? 'bg-blue-700 text-white border-blue-700 shadow-2xs'
                        : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                    }`}
                  >
                    {yr}
                  </button>
                );
              })}

              {(olderYears.length > 0 || isOlderYearSelected) && (
                <select
                  aria-label="Năm cũ hơn"
                  value={isOlderYearSelected ? selectedYear : ''}
                  onChange={(e) => {
                    if (e.target.value) selectYearVal(e.target.value);
                  }}
                  className={`min-h-[34px] px-2.5 py-1 rounded-md border text-xs font-semibold cursor-pointer ${
                    isOlderYearSelected
                      ? 'bg-blue-50 text-blue-900 border-blue-400'
                      : 'bg-slate-50 text-slate-700 border-slate-200'
                  }`}
                >
                  <option value="">Năm cũ hơn...</option>
                  {Array.from(
                    new Set([
                      ...olderYears,
                      ...(isOlderYearSelected ? [parseInt(selectedYear, 10)] : []),
                    ])
                  )
                    .filter((y) => !Number.isNaN(y))
                    .sort((a, b) => b - a)
                    .map((yr) => (
                      <option key={yr} value={String(yr)}>
                        Năm {yr}
                      </option>
                    ))}
                </select>
              )}

              <button
                type="button"
                onClick={() => selectYearVal('all')}
                className={`min-h-[34px] px-2.5 py-1 rounded-md border text-xs font-semibold transition-colors whitespace-nowrap ${
                  selectedYear === 'all'
                    ? 'bg-slate-800 text-white border-slate-800'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Tất cả các năm
              </button>

              <button
                type="button"
                onClick={() => selectYearVal('unknown')}
                className={`min-h-[34px] px-2.5 py-1 rounded-md border text-xs font-semibold transition-colors whitespace-nowrap ${
                  selectedYear === 'unknown'
                    ? 'bg-amber-600 text-white border-amber-600'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Chưa xác định năm
              </button>
            </div>
          );
        })()}

        {/* Collapsible Bộ lọc nâng cao (Tháng, Loại VB, Người ký, Trạng thái HS, Trạng thái tệp, Sắp xếp) */}
        {showAdvancedFilters && (
          <div className="pt-3 border-t border-slate-100 space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-2.5">
              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Tháng ban hành</label>
                <select
                  aria-label="Lọc theo tháng"
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
                >
                  <option value="all">Mọi tháng</option>
                  {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
                    <option key={m} value={m}>
                      Tháng {m}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Loại văn bản</label>
                <select
                  aria-label="Lọc theo loại văn bản"
                  value={selectedType}
                  onChange={(e) => setSelectedType(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
                >
                  <option value="all">Tất cả loại VB</option>
                  <option value="VBD">Văn bản đến (VBD)</option>
                  {documentTypes.map((t) => (
                    <option key={t.code} value={t.code}>
                      {t.name} ({t.code})
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Người ký</label>
                <select
                  aria-label="Lọc theo người ký"
                  value={selectedSigner}
                  onChange={(e) => setSelectedSigner(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
                >
                  <option value="all">Mọi người ký</option>
                  {availableSigners.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Trạng thái hồ sơ</label>
                <select
                  aria-label="Lọc trạng thái hồ sơ"
                  value={selectedStatus}
                  onChange={(e) => setSelectedStatus(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
                >
                  <option value="all">Tất cả hồ sơ</option>
                  <option value="active">Hiệu lực</option>
                  <option value="cancelled">Đã hủy</option>
                </select>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Trạng thái tệp</label>
                <select
                  aria-label="Lọc trạng thái tệp Drive"
                  value={selectedDriveStatus}
                  onChange={(e) => setSelectedDriveStatus(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600"
                >
                  <option value="all">Tất cả trạng thái tệp</option>
                  <option value="saved">Đã lưu lên Drive</option>
                  <option value="uploading">Đang lưu</option>
                  <option value="failed">Lưu tệp thất bại</option>
                  <option value="none">Không đính kèm</option>
                </select>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="text-[11px] font-bold text-slate-600 block">Sắp xếp</label>
                <select
                  aria-label="Sắp xếp danh sách"
                  value={sortBy}
                  onChange={(e: any) => setSortBy(e.target.value)}
                  className="w-full min-h-[40px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-blue-600 font-medium text-slate-700"
                >
                  <option value="newest">Mới nhất trước</option>
                  <option value="oldest">Cũ nhất trước</option>
                  <option value="number_asc">Số TT tăng dần</option>
                  <option value="number_desc">Số TT giảm dần</option>
                </select>
              </div>
            </div>

            {activeExtraFilterCount > 0 && (
              <div className="flex justify-end">
                <button
                  type="button"
                  onClick={() => {
                    setSelectedMonth('all');
                    setSelectedType('all');
                    setSelectedSigner('all');
                    setSelectedStatus('all');
                    setSelectedDriveStatus('all');
                    setSortBy('newest');
                  }}
                  className="text-xs font-semibold text-blue-700 hover:underline"
                >
                  Đặt lại bộ lọc phụ
                </button>
              </div>
            )}
          </div>
        )}

        {/* Sub-bar: Tổng của sổ theo năm + Số kết quả sau khi lọc + Mobile cards/table switch */}
        <div className="pt-2.5 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs text-slate-600">
          <div className="flex items-center gap-3 flex-wrap min-w-0">
            <span>
              Tổng sổ{' '}
              {selectedYear === 'all'
                ? 'tất cả các năm'
                : selectedYear === 'unknown'
                ? 'chưa xác định năm'
                : `năm ${selectedYear}`}
              : <strong className="text-slate-900 font-mono tabular-nums">{bookYearTotalCount}</strong> văn bản
            </span>
            <span className="text-slate-300">·</span>
            <span>
              Sau khi lọc: <strong className="text-blue-800 font-mono tabular-nums">{filteredDocuments.length}</strong> kết quả
            </span>
          </div>

          <div className="flex md:hidden items-center justify-end gap-2">
            <div className="flex items-center bg-slate-100 p-0.5 rounded-lg border border-slate-200">
              <button
                type="button"
                onClick={() => setMobileViewMode('cards')}
                className={`min-h-[34px] px-2.5 py-1 rounded text-[11px] font-semibold transition-colors flex items-center gap-1 ${
                  mobileViewMode === 'cards' ? 'bg-white text-blue-800 shadow-2xs' : 'text-slate-600'
                }`}
              >
                <LayoutList className="w-3.5 h-3.5" />
                <span>Thẻ</span>
              </button>
              <button
                type="button"
                onClick={() => setMobileViewMode('table')}
                className={`min-h-[34px] px-2.5 py-1 rounded text-[11px] font-semibold transition-colors flex items-center gap-1 ${
                  mobileViewMode === 'table' ? 'bg-white text-blue-800 shadow-2xs' : 'text-slate-600'
                }`}
              >
                <Table className="w-3.5 h-3.5" />
                <span>Bảng</span>
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* PDF Content Search Results Panel */}
      {pdfContentMode && (
        <div className="bg-white border border-purple-200 rounded-xl p-4 shadow-xs space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-purple-100 pb-2.5">
            <div className="flex items-center gap-2 text-xs font-bold text-purple-950">
              <FileText className="w-4 h-4 text-purple-700 shrink-0" />
              <span>Tra cứu nội dung bên trong PDF (theo từng trang)</span>
            </div>
            {pdfSearchResults && (
              <div className="flex flex-wrap items-center gap-2 text-[11px]">
                <span className="px-2 py-0.5 rounded bg-purple-50 text-purple-900 border border-purple-200 font-semibold tabular-nums">
                  {pdfSearchResults.total} đoạn khớp · {pdfSearchResults.indexedCount} PDF đã lập chỉ mục
                </span>
                {pdfSearchResults.unindexedCount > 0 && (
                  <span className="px-2 py-0.5 rounded bg-amber-50 text-amber-900 border border-amber-200 font-semibold">
                    {pdfSearchResults.unindexedCount} PDF cũ chưa có nội dung trích xuất
                  </span>
                )}
              </div>
            )}
          </div>

          {!searchTerm.trim() ? (
            <p className="text-xs text-slate-500">
              Nhập từ khóa vào ô tìm kiếm phía trên (hỗ trợ tiếng Việt có dấu và không dấu) để tìm chính xác trang và đoạn chứa nội dung trong các tệp PDF có quyền truy cập.
            </p>
          ) : pdfSearchLoading ? (
            <div className="flex items-center gap-2 text-xs text-slate-500 py-2">
              <Loader2 className="w-4 h-4 animate-spin text-purple-700" />
              <span>Đang tìm trong nội dung PDF...</span>
            </div>
          ) : pdfSearchError ? (
            <p className="text-xs text-rose-700 font-medium">{pdfSearchError}</p>
          ) : pdfSearchResults && pdfSearchResults.matches.length === 0 ? (
            <div className="text-xs text-slate-600 space-y-1 py-1">
              <p className="font-semibold">
                Không tìm thấy đoạn văn bản nào khớp từ khóa &ldquo;{searchTerm.trim()}&rdquo; trong các tệp PDF đã lập chỉ mục.
              </p>
              {pdfSearchResults.unindexedCount > 0 && (
                <p className="text-[11px] text-amber-800">
                  Lưu ý: Có {pdfSearchResults.unindexedCount} hồ sơ PDF cũ chưa có nội dung trích xuất — tải lại PDF nếu cần lập chỉ mục cho tệp đó.
                </p>
              )}
            </div>
          ) : pdfSearchResults ? (
            <div className="space-y-2.5">
              {pdfSearchResults.matches.map((m, idx) => (
                <div
                  key={`${m.docId}-${m.page}-${idx}`}
                  className="border border-slate-200 rounded-lg p-3 bg-slate-50/60 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                >
                  <div className="space-y-1 min-w-0 flex-1">
                    <div className="flex flex-wrap items-center gap-2 text-xs">
                      <span className="font-mono font-bold text-purple-900 bg-purple-50 border border-purple-200 px-2 py-0.5 rounded">
                        {m.fullDocNumber}
                      </span>
                      <span className="font-bold text-slate-900 truncate">{m.docName}</span>
                      <span className="text-[11px] font-bold px-2 py-0.5 rounded bg-sky-50 text-sky-900 border border-sky-200 tabular-nums">
                        Trang {m.page}
                      </span>
                      {m.bookName && (
                        <span className="text-[11px] text-slate-500">{m.bookName}</span>
                      )}
                    </div>
                    <p className="text-xs text-slate-700 bg-white border border-slate-200 rounded px-2.5 py-1.5 break-words">
                      {m.snippet}
                    </p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={() => handleOpenDetailByDocId(m.docId)}
                      className="inline-flex items-center gap-1 min-h-[36px] px-3 py-1.5 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 rounded-lg"
                    >
                      <Eye className="w-3.5 h-3.5 text-sky-700" />
                      <span>Xem chi tiết</span>
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        handleOpenDriveFile(
                          { id: m.docId } as DocumentRecord,
                          m.attachmentId
                        )
                      }
                      className="inline-flex items-center gap-1 min-h-[36px] px-3 py-1.5 text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      <span>Mở tệp</span>
                    </button>
                  </div>
                </div>
              ))}
            </div>
          ) : null}
        </div>
      )}

      {/* Mobile Default Card View (No horizontal scroll) */}
      {mobileViewMode === 'cards' && (
        <div className="md:hidden space-y-3">
          {paginatedActiveDocuments.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-8 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">Chưa có văn bản nào phù hợp</p>
              <span className="text-xs text-slate-400 mt-1 block">Hãy đổi tiêu chí tìm kiếm hoặc cấp số mới.</span>
            </div>
          ) : (
            paginatedActiveDocuments.map((doc) => {
              const isCancelled = doc.status === 'cancelled';
              const isIncoming = doc.docCategory === 'incoming';
              const categoryStyle = isIncoming ? SEMANTIC_STATUS_STYLES.incoming : SEMANTIC_STATUS_STYLES.outgoing;
              const recordStatus = getRecordStatusMeta(doc);
              const displayIssueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate) || '—';
              const bookObj = bookMap.get(getDocBookId(doc));
              const unitDisplay = isIncoming
                ? doc.issuingAgency || 'Cơ quan ngoài'
                : bookObj?.unitName || settings?.unitName || 'Đoàn trường';

              return (
                <div
                  key={doc.id}
                  className={`bg-white border rounded-xl p-3.5 sm:p-4 shadow-2xs space-y-3 min-w-0 ${
                    isCancelled ? 'border-rose-200 bg-rose-50/20' : 'border-slate-200'
                  }`}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex items-center gap-1.5 flex-wrap min-w-0">
                      <span className={`text-[11px] font-bold px-2 py-0.5 rounded border ${categoryStyle.badge}`}>
                        {isIncoming ? 'Văn bản đến' : `Văn bản đi · ${doc.typeName || doc.typeCode}`}
                      </span>
                      {doc.needsReview && (
                        <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${SEMANTIC_STATUS_STYLES.pending.badge}`}>
                          Cần kiểm tra
                        </span>
                      )}
                      {isCancelled && (
                        <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${recordStatus.className}`}>
                          {recordStatus.label}
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="flex items-center gap-2 min-w-0">
                    <span className={`font-mono font-bold text-sm break-all ${categoryStyle.text}`}>
                      {doc.fullDocNumber}
                    </span>
                  </div>

                  <div className="min-w-0">
                    <h3 className={`text-sm font-bold text-slate-900 leading-snug break-words ${isCancelled ? 'line-through text-slate-400' : ''}`}>
                      {doc.docName || doc.title}
                    </h3>
                    <p className="text-[11px] text-slate-600 mt-1 break-words">
                      <span className="text-slate-400">Cơ quan / Đơn vị:</span>{' '}
                      <strong className="text-slate-800">{unitDisplay}</strong> ·{' '}
                      <span className="text-slate-400">Ngày BH:</span>{' '}
                      <strong className="font-mono text-slate-800">{displayIssueDate}</strong>
                    </p>
                  </div>

                  {/* Drive File Status Badge */}
                  <div className="flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg border border-slate-200 bg-slate-50/70 text-[11px] flex-wrap">
                    <span
                      className={`font-semibold px-2 py-0.5 rounded border ${
                        doc.driveStatus === 'saved'
                          ? SEMANTIC_STATUS_STYLES.saved.badge
                          : doc.driveStatus === 'failed'
                          ? SEMANTIC_STATUS_STYLES.danger.badge
                          : doc.driveStatus === 'uploading'
                          ? SEMANTIC_STATUS_STYLES.pending.badge
                          : 'text-slate-600 bg-white border-slate-200'
                      }`}
                    >
                      {getDriveStatusLabel(doc)}
                    </span>
                    <span className="text-[11px] text-slate-500 truncate min-w-0">
                      {bookObj?.name || doc.bookName}
                    </span>
                  </div>

                  {/* Card Actions: Xem chi tiết + Mở tệp + ⋯ */}
                  <div className="pt-2 border-t border-slate-100 flex items-center justify-between gap-2 text-xs">
                    <div className="flex items-center gap-2 flex-wrap">
                      <button
                        type="button"
                        onClick={() => setDetailModalDoc(doc)}
                        className="inline-flex items-center gap-1.5 min-h-[40px] px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold transition-colors whitespace-nowrap"
                      >
                        <Eye className="w-3.5 h-3.5 text-blue-700" />
                        <span>Xem chi tiết</span>
                      </button>

                      {doc.driveStatus === 'saved' && doc.driveFileUrl && (
                        <button
                          type="button"
                          onClick={() => handleOpenDriveFile(doc)}
                          disabled={openingFileId === doc.id}
                          className="inline-flex items-center gap-1.5 min-h-[40px] px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 font-semibold transition-colors whitespace-nowrap"
                        >
                          <ExternalLink className="w-3.5 h-3.5" />
                          <span>Mở tệp</span>
                        </button>
                      )}
                    </div>

                    <button
                      type="button"
                      aria-label={`Thao tác khác cho ${doc.fullDocNumber}`}
                      onClick={(e) => toggleActionMenu(e, doc)}
                      className="inline-flex items-center justify-center min-w-[40px] min-h-[40px] rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 shrink-0"
                    >
                      <MoreHorizontal className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}

      {/* Streamlined Desktop Table */}
      <div
        className={`bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs ${
          mobileViewMode === 'cards' ? 'hidden md:block' : 'block'
        }`}
      >
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-sky-900 text-white text-[11px] font-bold uppercase tracking-wider divide-x divide-sky-800">
                <th className="py-3 px-3.5 w-48">Số / Ký hiệu</th>
                <th className="py-3 px-4 min-w-[240px]">Tên văn bản</th>
                <th className="py-3 px-3 w-32 text-center">Ngày ban hành</th>
                <th className="py-3 px-3 w-44">Cơ quan / Đơn vị</th>
                <th className="py-3 px-3 w-36 text-center">Trạng thái</th>
                <th className="py-3 px-3 w-44 text-center">Tệp</th>
                <th className="py-3 px-3 w-48 text-center">Thao tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
              {paginatedActiveDocuments.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    <p className="font-medium text-slate-500">Chưa có văn bản nào phù hợp với bộ lọc</p>
                    <span className="text-xs">Hãy thử đổi tiêu chí tìm kiếm hoặc tạo cấp số mới.</span>
                  </td>
                </tr>
              ) : (
                paginatedActiveDocuments.map((doc) => {
                  const isCancelled = doc.status === 'cancelled';
                  const isIncoming = doc.docCategory === 'incoming';
                  const categoryStyle = isIncoming ? SEMANTIC_STATUS_STYLES.incoming : SEMANTIC_STATUS_STYLES.outgoing;
                  const recordStatus = getRecordStatusMeta(doc);
                  const displayIssueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate) || '—';
                  const bookObj = bookMap.get(getDocBookId(doc));
                  const agencyOrUnit = isIncoming
                    ? doc.issuingAgency || 'Chưa xác định'
                    : bookObj?.unitName || settings?.unitName || 'Đoàn trường';

                  return (
                    <tr
                      key={doc.id}
                      className={`hover:bg-blue-50/30 transition-colors ${
                        isCancelled ? 'bg-rose-50/30' : ''
                      }`}
                    >
                      {/* 1. Số / Ký hiệu */}
                      <td className="py-3 px-3.5 align-top">
                        <div className="space-y-1">
                          <div className={`inline-flex items-center gap-1.5 font-mono font-bold px-2.5 py-0.5 rounded border shadow-2xs ${categoryStyle.badge}`}>
                            <span className={`w-1.5 h-1.5 rounded-full ${categoryStyle.dot}`} />
                            <span>{doc.fullDocNumber}</span>
                          </div>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className={`text-[11px] font-semibold ${categoryStyle.text}`}>
                              {isIncoming ? 'Văn bản đến' : `Văn bản đi · ${doc.typeName}`}
                            </span>
                            {doc.needsReview && (
                              <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${SEMANTIC_STATUS_STYLES.pending.badge}`}>
                                Cần kiểm tra
                              </span>
                            )}
                          </div>
                        </div>
                      </td>

                      {/* 2. Tên văn bản */}
                      <td className="py-3 px-4 align-top max-w-md">
                        <button
                          type="button"
                          onClick={() => setDetailModalDoc(doc)}
                          className={`text-left font-semibold text-slate-900 hover:text-blue-800 transition-colors leading-snug ${
                            isCancelled ? 'line-through text-slate-400' : ''
                          }`}
                        >
                          {doc.docName || doc.title}
                        </button>
                        {doc.docName && doc.title && doc.title !== doc.docName && (
                          <p className="text-[11px] text-slate-500 mt-0.5 line-clamp-1">
                            {doc.title}
                          </p>
                        )}
                        {isCancelled && doc.cancelReason && (
                          <p className="text-[11px] text-rose-600 italic mt-0.5">
                            Lý do hủy: {doc.cancelReason}
                          </p>
                        )}
                      </td>

                      {/* 3. Ngày ban hành */}
                      <td className="py-3 px-3 text-center font-mono whitespace-nowrap text-slate-800 align-top tabular-nums">
                        {displayIssueDate}
                      </td>

                      {/* 4. Cơ quan / Đơn vị */}
                      <td className="py-3 px-3 align-top">
                        <div className="font-medium text-slate-800 text-xs">{agencyOrUnit}</div>
                        {bookObj && (
                          <div className="text-[11px] text-slate-400 truncate max-w-[180px]">
                            {bookObj.name}
                          </div>
                        )}
                      </td>

                      {/* 5. Trạng thái (Hồ sơ) */}
                      <td className="py-3 px-3 text-center whitespace-nowrap align-top">
                        <span className={`inline-block text-[11px] font-semibold px-2.5 py-0.5 rounded border ${recordStatus.className}`}>
                          {recordStatus.label}
                        </span>
                      </td>

                      {/* 6. Tệp (Trạng thái Drive) */}
                      <td className="py-3 px-3 text-center align-top">
                        {doc.driveStatus === 'saved' ? (
                          <span className={`inline-block text-[11px] font-semibold px-2 py-0.5 rounded border ${SEMANTIC_STATUS_STYLES.saved.badge}`}>
                            Đã lưu lên Drive
                          </span>
                        ) : doc.driveStatus === 'failed' || doc.driveStatus === 'uploading' ? (
                          <span
                            className={`inline-block text-[11px] font-semibold px-2 py-0.5 rounded border ${
                              doc.driveStatus === 'uploading'
                                ? SEMANTIC_STATUS_STYLES.pending.badge
                                : SEMANTIC_STATUS_STYLES.danger.badge
                            }`}
                            title={doc.driveError || ''}
                          >
                            {getDriveStatusLabel(doc)}
                          </span>
                        ) : (
                          <span className="text-[11px] text-slate-400">Không đính kèm</span>
                        )}
                      </td>

                      {/* 7. Thao tác: Xem chi tiết + Mở tệp + ⋯ */}
                      <td className="py-3 px-3 text-center whitespace-nowrap align-top">
                        <div className="inline-flex items-center justify-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => setDetailModalDoc(doc)}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-semibold transition-colors whitespace-nowrap"
                          >
                            <Eye className="w-3.5 h-3.5 text-blue-700" />
                            <span>Xem chi tiết</span>
                          </button>

                          {doc.driveStatus === 'saved' && doc.driveFileUrl && (
                            <button
                              type="button"
                              onClick={() => handleOpenDriveFile(doc)}
                              disabled={openingFileId === doc.id}
                              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 text-[11px] font-semibold transition-colors whitespace-nowrap"
                              title={doc.driveFileName || 'Mở tệp'}
                            >
                              <ExternalLink className="w-3.5 h-3.5" />
                              <span>Mở tệp</span>
                            </button>
                          )}

                          <button
                            type="button"
                            aria-label={`Thao tác khác cho ${doc.fullDocNumber}`}
                            onClick={(e) => toggleActionMenu(e, doc)}
                            className="inline-flex items-center justify-center w-8 h-8 rounded-md border border-slate-200 bg-white hover:bg-slate-50 text-slate-600"
                            title="Thao tác khác"
                          >
                            <MoreHorizontal className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Floating Overflow Action Menu (⋯) */}
      {openMenu && (
        <div
          role="menu"
          onClick={(e) => e.stopPropagation()}
          style={{
            position: 'fixed',
            right: `${openMenu.right}px`,
            ...(openMenu.openUp
              ? { bottom: `${window.innerHeight - openMenu.top}px` }
              : { top: `${openMenu.top}px` }),
          }}
          className="z-50 w-52 bg-white border border-slate-200 rounded-xl shadow-lg py-1.5 text-xs text-slate-700 animate-scaleUp"
        >
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              handleCopy(openMenu.doc.id, openMenu.doc.fullDocNumber);
              setOpenMenu(null);
            }}
            className="w-full px-3.5 py-2 text-left hover:bg-slate-50 flex items-center gap-2 font-medium"
          >
            {copiedId === openMenu.doc.id ? (
              <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
            ) : (
              <Copy className="w-3.5 h-3.5 text-slate-500 shrink-0" />
            )}
            <span>{copiedId === openMenu.doc.id ? 'Đã sao chép số!' : 'Sao chép số văn bản'}</span>
          </button>

          {openMenu.doc.driveStatus === 'saved' && canReplaceDocFile(openMenu.doc) && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                const target = openMenu.doc;
                setOpenMenu(null);
                openUploadModal(target, 'replace');
              }}
              className="w-full px-3.5 py-2 text-left hover:bg-slate-50 flex items-center gap-2 font-medium"
            >
              <Upload className="w-3.5 h-3.5 text-blue-700 shrink-0" />
              <span>Thay tệp trên Drive</span>
            </button>
          )}

          {(openMenu.doc.driveStatus === 'failed' || openMenu.doc.driveStatus === 'uploading') &&
            canRetryOrAttachDoc(openMenu.doc) && (
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  const target = openMenu.doc;
                  setOpenMenu(null);
                  openUploadModal(target, 'retry');
                }}
                className="w-full px-3.5 py-2 text-left hover:bg-rose-50 text-rose-700 flex items-center gap-2 font-semibold"
              >
                <RefreshCw className="w-3.5 h-3.5 shrink-0" />
                <span>Thử lưu tệp lại</span>
              </button>
            )}

          {(!openMenu.doc.driveStatus || openMenu.doc.driveStatus === 'none') &&
            canRetryOrAttachDoc(openMenu.doc) && (
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  const target = openMenu.doc;
                  setOpenMenu(null);
                  openUploadModal(target, 'retry');
                }}
                className="w-full px-3.5 py-2 text-left hover:bg-slate-50 flex items-center gap-2 font-medium"
              >
                <FileUp className="w-3.5 h-3.5 text-blue-700 shrink-0" />
                <span>Đính kèm tệp lên Drive</span>
              </button>
            )}

          {canEditDoc(openMenu.doc) && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                const target = openMenu.doc;
                setOpenMenu(null);
                openEditModal(target);
              }}
              className="w-full px-3.5 py-2 text-left hover:bg-slate-50 flex items-center gap-2 font-medium"
            >
              <Edit2 className="w-3.5 h-3.5 text-amber-600 shrink-0" />
              <span>Sửa thông tin</span>
            </button>
          )}

          {canCancelDoc(openMenu.doc) && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                const target = openMenu.doc;
                setOpenMenu(null);
                openCancelModal(target);
              }}
              className="w-full px-3.5 py-2 text-left hover:bg-rose-50 text-rose-700 flex items-center gap-2 font-medium"
            >
              <Ban className="w-3.5 h-3.5 shrink-0" />
              <span>Hủy số văn bản</span>
            </button>
          )}

          {canDeleteDoc(openMenu.doc) && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                const target = openMenu.doc;
                setOpenMenu(null);
                openDeleteModal(target);
              }}
              className="w-full px-3.5 py-2 text-left hover:bg-rose-50 text-rose-700 flex items-center gap-2 font-semibold border-t border-slate-100 mt-1 pt-2"
            >
              <Trash2 className="w-3.5 h-3.5 shrink-0" />
              <span>Xóa văn bản</span>
            </button>
          )}
        </div>
      )}

      <PaginationControls
        totalItems={filteredDocuments.length}
        page={safeActivePage}
        pageSize={activePageSize}
        onPageChange={setActivePage}
        onPageSizeChange={setActivePageSize}
      />
        </>
      )}

      {/* SECTION 2: VB ĐÃ XÓA */}
      {registrySection === 'deleted' && (
        <div className="space-y-4">
          {syncRetryNotice && (
            <div
              className={`p-3 rounded-lg border text-xs flex items-center justify-between gap-2 ${
                syncRetryNotice.type === 'success'
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                  : 'bg-rose-50 border-rose-200 text-rose-900'
              }`}
            >
              <span className="break-words">{syncRetryNotice.message}</span>
              <button
                type="button"
                onClick={() => setSyncRetryNotice(null)}
                className="p-1 shrink-0"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          )}

          {bulkSyncResult && (
            <div
              className={`p-3.5 rounded-xl border text-xs space-y-2 ${
                bulkSyncResult.failedCount === 0
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-950'
                  : 'bg-amber-50 border-amber-300 text-amber-950'
              }`}
            >
              <div className="flex items-start justify-between gap-2">
                <div className="space-y-1">
                  <p className="font-bold">{bulkSyncResult.message}</p>
                  <p className="text-[11px]">
                    Phạm vi: <strong>{bulkSyncResult.bookName}</strong> · Thành công:{' '}
                    <strong className="font-mono text-emerald-800">{bulkSyncResult.succeededCount}</strong> · Thất bại:{' '}
                    <strong className="font-mono text-rose-700">{bulkSyncResult.failedCount}</strong>
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setBulkSyncResult(null)}
                  className="p-1 shrink-0"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              {bulkSyncResult.failedItems.length > 0 && (
                <ul className="space-y-1 pl-3 border-l-2 border-rose-300 text-[11px] text-rose-900">
                  {bulkSyncResult.failedItems.map((fi) => (
                    <li key={fi.id}>
                      <strong className="font-mono">{fi.fullDocNumber}</strong>: {fi.error}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          )}

          {archiveNotice && (
            <div className="p-3 rounded-lg border bg-emerald-50 border-emerald-200 text-emerald-950 text-xs flex items-center justify-between gap-2">
              <span className="break-words font-medium">{archiveNotice.message}</span>
              <button type="button" onClick={() => setArchiveNotice(null)} className="p-1 shrink-0">
                <X className="w-4 h-4" />
              </button>
            </div>
          )}

          {archivedDeletedError && (
            <div className="p-3 rounded-lg border bg-rose-50 border-rose-200 text-rose-900 text-xs">
              {archivedDeletedError}
            </div>
          )}

          <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-7 gap-2.5">
              <div className="lg:col-span-2 relative min-w-0">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                <input
                  type="text"
                  value={deletedSearch}
                  onChange={(e) => setDeletedSearch(e.target.value)}
                  placeholder="Tìm số VB, tên VB, lý do xóa, người xóa..."
                  className="w-full min-h-[42px] h-10 pl-9 pr-3 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-rose-600"
                />
              </div>

              <select
                aria-label="Lọc năm văn bản đã xóa"
                value={deletedDocYear}
                onChange={(e) => {
                  deletedYearManualRef.current = true;
                  setDeletedDocYear(e.target.value);
                }}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Tất cả các năm</option>
                {deletedAvailableYears.map((yr) => (
                  <option key={yr} value={String(yr)}>
                    Năm {yr}
                  </option>
                ))}
                <option value="unknown">Chưa xác định năm</option>
              </select>

              <select
                aria-label="Lọc theo người xóa"
                value={deletedByFilter}
                onChange={(e) => setDeletedByFilter(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Tất cả người xóa</option>
                {availableDeletedActors.map((a) => (
                  <option key={a.username} value={a.username}>
                    {a.label}
                  </option>
                ))}
              </select>

              <select
                aria-label="Lọc trạng thái đồng bộ Google Trang tính"
                value={deletedSyncFilter}
                onChange={(e) => setDeletedSyncFilter(e.target.value as any)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Mọi trạng thái Google Trang tính</option>
                <option value="pending_sync">Chờ đồng bộ Google Trang tính</option>
                <option value="synced">Đã đồng bộ VB đã xóa</option>
                {isAdmin && <option value="archived">Đã lưu trữ</option>}
              </select>

              <input
                type="date"
                aria-label="Từ ngày xóa"
                value={deletedFromDate}
                onChange={(e) => setDeletedFromDate(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              />

              <input
                type="date"
                aria-label="Đến ngày xóa"
                value={deletedToDate}
                onChange={(e) => setDeletedToDate(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              />
            </div>

            {/* Bulk Sync & Admin Archive Toolbar */}
            {deletedSyncFilter !== 'archived' && (
              <div className="pt-2.5 border-t border-slate-100 flex flex-col lg:flex-row lg:items-center justify-between gap-2.5">
                <div className="flex flex-wrap items-center gap-3 text-xs">
                  <label className="inline-flex items-center gap-2 cursor-pointer select-none font-semibold text-slate-700">
                    <input
                      type="checkbox"
                      checked={allOnPageSelected}
                      disabled={paginatedDeletedDocuments.length === 0}
                      onChange={toggleSelectAllOnDeletedPage}
                      className="w-4 h-4 rounded border-slate-300 text-rose-700 focus:ring-rose-600"
                    />
                    <span>Chọn tất cả trên trang này</span>
                  </label>
                  {selectedDeletedIds.length > 0 && (
                    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-900 font-semibold">
                      Đã chọn <strong className="font-mono">{selectedDeletedIds.length}</strong> mục
                      <button
                        type="button"
                        onClick={() => setSelectedDeletedIds([])}
                        className="text-sky-700 hover:underline ml-1"
                      >
                        Bỏ chọn
                      </button>
                    </span>
                  )}
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  {canBulkSyncScope && (
                    <>
                      <button
                        type="button"
                        disabled={selectedDeletedIds.length === 0 || bulkSyncSubmitting}
                        onClick={() => setBulkSyncConfirmMode('selected')}
                        className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold rounded-lg border border-sky-700 bg-sky-700 hover:bg-sky-800 text-white disabled:opacity-45 disabled:pointer-events-none transition-colors"
                      >
                        <RotateCw className="w-3.5 h-3.5" />
                        <span>Đồng bộ các mục đã chọn ({selectedDeletedIds.length})</span>
                      </button>

                      <button
                        type="button"
                        disabled={pendingExcelSyncCount === 0 || bulkSyncSubmitting}
                        onClick={() => setBulkSyncConfirmMode('pending')}
                        className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold rounded-lg border border-amber-600 bg-amber-600 hover:bg-amber-700 text-white disabled:opacity-45 disabled:pointer-events-none transition-colors"
                      >
                        <RotateCw className="w-3.5 h-3.5" />
                        <span>Đồng bộ tất cả mục đang chờ ({pendingExcelSyncCount})</span>
                      </button>
                    </>
                  )}

                  {isAdmin && (
                    <button
                      type="button"
                      disabled={
                        (selectedSyncedCount === 0 && syncedUnarchivedInBookCount === 0) ||
                        archiveSubmitting
                      }
                      onClick={() => {
                        setArchiveError(null);
                        setArchiveMode(selectedSyncedCount > 0 ? 'selected' : 'all_synced');
                        setArchiveModalOpen(true);
                      }}
                      className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold rounded-lg border border-slate-300 bg-white hover:bg-slate-50 text-slate-800 disabled:opacity-45 disabled:pointer-events-none transition-colors"
                      title="Ẩn các văn bản đã đồng bộ Google Trang tính khỏi danh sách xử lý mặc định (vẫn giữ hồ sơ, số đã cấp và link Drive)"
                    >
                      <Archive className="w-3.5 h-3.5 text-slate-700" />
                      <span>Dọn khỏi danh sách</span>
                    </button>
                  )}
                </div>
              </div>
            )}

            <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
              <span>
                Hiển thị <strong className="font-mono text-slate-800">{filteredDeletedDocuments.length}</strong>{' '}
                {deletedSyncFilter === 'archived' ? 'văn bản đã lưu trữ' : 'văn bản đã xóa'} (số văn bản trong dãy và tệp trên Google Drive vẫn được bảo toàn).
              </span>
              {(deletedSearch ||
                deletedByFilter !== 'all' ||
                deletedSyncFilter !== 'all' ||
                deletedDocYear !== String(currentVnYear) ||
                deletedFromDate ||
                deletedToDate) && (
                <button
                  type="button"
                  onClick={() => {
                    deletedYearManualRef.current = false;
                    setDeletedSearch('');
                    setDeletedByFilter('all');
                    setDeletedSyncFilter('all');
                    setDeletedDocYear(String(currentVnYear));
                    setDeletedFromDate('');
                    setDeletedToDate('');
                  }}
                  className="text-sky-700 hover:underline font-semibold"
                >
                  Đặt lại bộ lọc
                </button>
              )}
            </div>
          </div>

          {deletedSyncFilter === 'archived' && loadingArchivedDeleted ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-500 text-xs flex items-center justify-center gap-2">
              <Loader2 className="w-4 h-4 animate-spin text-sky-700" />
              <span>Đang tải danh sách văn bản đã lưu trữ...</span>
            </div>
          ) : paginatedDeletedDocuments.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">
                {deletedSyncFilter === 'archived'
                  ? 'Không có văn bản đã lưu trữ nào phù hợp'
                  : 'Không có văn bản đã xóa cần xử lý'}
              </p>
              <span className="text-xs text-slate-400 mt-1 block">
                {deletedSyncFilter === 'archived'
                  ? 'Các văn bản đã xóa sau khi đồng bộ Google Trang tính và được Admin dọn khỏi danh sách sẽ lưu tại đây.'
                  : 'Khi xóa văn bản khỏi danh sách hiệu lực, hồ sơ sẽ lưu tại đây và đồng bộ sang sheet "VB đã xóa" trên Google Trang tính.'}
              </span>
            </div>
          ) : (
            <div className="space-y-3">
              {paginatedDeletedDocuments.map((doc) => {
                const syncMeta = getExcelDeleteSyncMeta(doc);
                const bookObj = bookMap.get(getDocBookId(doc));
                const isSelected = selectedDeletedIds.includes(doc.id);
                return (
                  <div
                    key={doc.id}
                    className={`bg-white border rounded-xl p-4 shadow-2xs space-y-3 ${
                      doc.deletedListArchived
                        ? 'border-slate-300 bg-slate-50/40'
                        : isSelected
                        ? 'border-sky-500 ring-1 ring-sky-500/40'
                        : 'border-rose-200'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-2">
                      <div className="flex items-start gap-2.5 min-w-0">
                        {!doc.deletedListArchived && deletedSyncFilter !== 'archived' && (
                          <input
                            type="checkbox"
                            aria-label={`Chọn văn bản ${doc.fullDocNumber}`}
                            checked={isSelected}
                            onChange={() => toggleSelectDeletedDoc(doc.id)}
                            className="mt-1 w-4 h-4 rounded border-slate-300 text-rose-700 focus:ring-rose-600 shrink-0"
                          />
                        )}
                        <div className="space-y-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-mono font-bold text-rose-900 bg-rose-50 border border-rose-200 px-2.5 py-0.5 rounded text-xs">
                              {doc.fullDocNumber}
                            </span>
                            <span className="text-[11px] font-semibold px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
                              {bookObj?.name || doc.bookName || 'Đoàn trường'}
                            </span>
                            <span className="font-mono text-[10px] text-slate-400">
                              Mã hồ sơ: {doc.id}
                            </span>
                          </div>
                          <h3 className="text-sm font-bold text-slate-900 break-words">
                            {doc.docName || doc.title}
                          </h3>
                        </div>
                      </div>

                      <span className={`inline-flex items-center gap-1.5 text-[11px] font-semibold px-2.5 py-1 rounded border shrink-0 ${syncMeta.className}`}>
                        <span>{syncMeta.label}</span>
                      </span>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 p-3 bg-rose-50/50 border border-rose-100 rounded-lg text-xs">
                      <div>
                        <span className="text-slate-500 block text-[11px]">Người xóa:</span>
                        <span className="font-semibold text-slate-900">{doc.deletedBy || '—'}</span>
                      </div>
                      <div>
                        <span className="text-slate-500 block text-[11px]">Thời điểm xóa (giờ VN):</span>
                        <span className="font-mono font-semibold text-slate-900">
                          {doc.deletedAtVietnam || (doc.deletedAt ? new Date(doc.deletedAt).toLocaleString('vi-VN') : '—')}
                        </span>
                      </div>
                      <div>
                        <span className="text-slate-500 block text-[11px]">Lý do xóa:</span>
                        <span className="font-medium text-rose-900 break-words">{doc.deleteReason || '—'}</span>
                      </div>
                    </div>

                    {doc.deletedListArchived && (
                      <div className="text-[11px] text-slate-700 bg-slate-100 border border-slate-200 rounded-lg px-3 py-1.5">
                        Đã lưu trữ bởi <strong>{doc.deletedArchivedBy || 'Quản trị viên'}</strong>
                        {doc.deletedArchivedAt ? ` lúc ${new Date(doc.deletedArchivedAt).toLocaleString('vi-VN')}` : ''}. Số văn bản trong dãy, tệp Drive và dòng Google Trang tính vẫn được giữ nguyên.
                      </div>
                    )}

                    {doc.excelSyncError && syncMeta.needsRetry && (
                      <div className="text-[11px] text-amber-900 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                        <strong>Chi tiết lỗi đồng bộ Google Trang tính:</strong> {doc.excelSyncError}
                      </div>
                    )}

                    <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100 text-xs">
                      <div className="flex items-center gap-2 flex-wrap">
                        {doc.driveStatus === 'saved' && doc.driveFileUrl ? (
                          <button
                            type="button"
                            onClick={() => handleOpenDriveFile(doc)}
                            disabled={openingFileId === doc.id}
                            className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 font-semibold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                            <span>Mở tệp Drive ({doc.driveFileName || 'Tệp gốc'})</span>
                          </button>
                        ) : (
                          <span className="text-[11px] text-slate-400">Không có tệp Drive đính kèm</span>
                        )}
                      </div>

                      <div className="flex items-center gap-2 flex-wrap">
                        {syncMeta.needsRetry && canRetryExcelSync(doc) && (
                          <button
                            type="button"
                            onClick={() => handleRetryExcelSync(doc)}
                            disabled={retryingSyncDocId === doc.id}
                            className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold bg-amber-600 hover:bg-amber-700 disabled:opacity-60 text-white rounded-lg shadow-2xs"
                          >
                            <RotateCw className={`w-3.5 h-3.5 ${retryingSyncDocId === doc.id ? 'animate-spin' : ''}`} />
                            <span>{retryingSyncDocId === doc.id ? 'Đang đồng bộ...' : 'Thử đồng bộ lại'}</span>
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => setDetailModalDoc(doc)}
                          className="inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 rounded-lg"
                        >
                          <Eye className="w-3.5 h-3.5" />
                          <span>Xem chi tiết</span>
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          <PaginationControls
            totalItems={filteredDeletedDocuments.length}
            page={safeDeletedPage}
            pageSize={deletedPageSize}
            onPageChange={setDeletedPage}
            onPageSizeChange={setDeletedPageSize}
          />
        </div>
      )}

      {/* SECTION 3: LỊCH SỬ THAO TÁC (CHỈ ĐỌC, ADMIN HỆ THỐNG CÓ THỂ DỌN THEO THỜI GIAN) */}
      {registrySection === 'history' && (
        <div className="space-y-4">
          {cleanupNotice && (
            <div
              className={`p-3 rounded-lg border text-xs flex items-center justify-between gap-2 ${
                cleanupNotice.type === 'success'
                  ? 'bg-emerald-50 border-emerald-200 text-emerald-950'
                  : 'bg-rose-50 border-rose-200 text-rose-900'
              }`}
            >
              <span className="font-medium break-words">{cleanupNotice.message}</span>
              <button type="button" onClick={() => setCleanupNotice(null)} className="p-1 shrink-0">
                <X className="w-4 h-4" />
              </button>
            </div>
          )}

          <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2.5">
              <div className="lg:col-span-2 relative min-w-0">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                <input
                  type="text"
                  value={historySearch}
                  onChange={(e) => setHistorySearch(e.target.value)}
                  placeholder="Tìm số VB, mã hồ sơ, tên VB, tài khoản, lý do..."
                  className="w-full min-h-[42px] h-10 pl-9 pr-3 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              <select
                aria-label="Lọc theo tài khoản"
                value={historyActorFilter}
                onChange={(e) => setHistoryActorFilter(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Tất cả tài khoản</option>
                {historyActors.map((actor) => (
                  <option key={actor.username} value={actor.username}>
                    {actor.userName} ({actor.username})
                  </option>
                ))}
              </select>

              <select
                aria-label="Lọc theo loại thao tác"
                value={historyActionFilter}
                onChange={(e) => setHistoryActionFilter(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Tất cả thao tác</option>
                <option value="update">Sửa văn bản</option>
                <option value="delete">Xóa văn bản</option>
                <option value="replace_file">Thay tệp Drive</option>
                <option value="upload_word">Đính kèm Word</option>
                <option value="upload_pdf">Lưu PDF Drive</option>
                <option value="cancel">Hủy số</option>
                <option value="create">Cấp số mới</option>
                <option value="import_pdf">Nhập PDF</option>
                <option value="retry_excel_sync">Đồng bộ lại Google Trang tính</option>
              </select>

              <input
                type="date"
                aria-label="Từ ngày"
                value={historyFromDate}
                onChange={(e) => setHistoryFromDate(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              />

              <input
                type="date"
                aria-label="Đến ngày"
                value={historyToDate}
                onChange={(e) => setHistoryToDate(e.target.value)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              />
            </div>

            <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
              <span>
                Nhật ký bất biến trong phạm vi sổ được phân quyền · Tổng{' '}
                <strong className="font-mono text-slate-800">{historyTotal}</strong> bản ghi.
              </span>
              <div className="flex flex-wrap items-center gap-3">
                {(historySearch || historyActorFilter !== 'all' || historyActionFilter !== 'all' || historyFromDate || historyToDate) && (
                  <button
                    type="button"
                    onClick={() => {
                      setHistorySearch('');
                      setHistoryActorFilter('all');
                      setHistoryActionFilter('all');
                      setHistoryFromDate('');
                      setHistoryToDate('');
                    }}
                    className="text-sky-700 hover:underline font-semibold"
                  >
                    Đặt lại bộ lọc
                  </button>
                )}
                <button
                  type="button"
                  onClick={fetchDocumentHistory}
                  disabled={historyLoading}
                  className="inline-flex items-center gap-1 text-sky-700 hover:underline font-semibold"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${historyLoading ? 'animate-spin' : ''}`} />
                  <span>Tải lại nhật ký</span>
                </button>
                {isAdmin && (
                  <button
                    type="button"
                    onClick={openHistoryCleanupModal}
                    className="inline-flex items-center gap-1.5 min-h-[34px] px-3 py-1.5 text-xs font-semibold rounded-lg border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-800 transition-colors"
                  >
                    <Eraser className="w-3.5 h-3.5" />
                    <span>Dọn lịch sử</span>
                  </button>
                )}
              </div>
            </div>
          </div>

          {historyError && (
            <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800">
              {historyError}
            </div>
          )}

          {historyLoading && historyLogs.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-500 text-xs flex items-center justify-center gap-2">
              <Loader2 className="w-4 h-4 animate-spin text-sky-700" />
              <span>Đang tải lịch sử thao tác...</span>
            </div>
          ) : paginatedHistoryLogs.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">Chưa có bản ghi lịch sử nào phù hợp</p>
              <span className="text-xs text-slate-400 mt-1 block">
                Mọi thao tác cấp số, sửa thông tin, thay tệp, hủy số và xóa văn bản đều được lưu vết tự động theo tài khoản.
              </span>
            </div>
          ) : (
            <div className="space-y-2.5">
              {paginatedHistoryLogs.map((entry) => {
                const actionMeta = getAuditActionMeta(entry.action);
                const changeEntries = Object.entries(entry.changes || {});
                return (
                  <div
                    key={entry.id}
                    className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-2xs space-y-2 text-xs"
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div className="flex items-center gap-2 flex-wrap min-w-0">
                        <span className={`text-[11px] font-bold px-2 py-0.5 rounded border ${actionMeta.className}`}>
                          {actionMeta.label}
                        </span>
                        <span className="font-mono font-bold text-slate-900">
                          {entry.fullDocNumber || '—'}
                        </span>
                        <span className="text-[11px] font-semibold px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
                          {entry.bookName}
                        </span>
                        <span className="font-mono text-[10px] text-slate-400">
                          Mã hồ sơ: {entry.docId}
                        </span>
                      </div>

                      <div className="flex items-center gap-2 text-[11px] text-slate-500 shrink-0">
                        <Clock className="w-3.5 h-3.5 text-slate-400" />
                        <span className="font-mono font-semibold text-slate-700">
                          {entry.timestampVietnam || new Date(entry.timestamp).toLocaleString('vi-VN')}
                        </span>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-slate-700">
                      <span>
                        <strong>Tài khoản:</strong> {entry.userName}{' '}
                        <span className="font-mono text-slate-500">({entry.username})</span>
                      </span>
                      {entry.docName && (
                        <span className="truncate max-w-full">
                          <strong>Văn bản:</strong> {entry.docName}
                        </span>
                      )}
                    </div>

                    {entry.reason && (
                      <div className="p-2 rounded bg-slate-50 border border-slate-200 text-slate-800">
                        <strong>Lý do / Ghi nhận:</strong> {entry.reason}
                      </div>
                    )}

                    {changeEntries.length > 0 && (
                      <div className="p-2.5 rounded-lg bg-slate-50/80 border border-slate-200 space-y-1">
                        {changeEntries.map(([field, delta], fIdx) => (
                          <div key={fIdx} className="text-[11px] text-slate-700 break-words">
                            <span className="font-semibold text-slate-800">{formatAuditFieldLabel(field)}:</span>{' '}
                            {delta.old !== null && delta.old !== undefined ? (
                              <>
                                <del className="text-rose-600">{delta.old || '(trống)'}</del> →{' '}
                              </>
                            ) : null}
                            <ins className="text-emerald-700 no-underline font-semibold">
                              {delta.new || '(trống)'}
                            </ins>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          <PaginationControls
            totalItems={historyLogs.length}
            page={safeHistoryPage}
            pageSize={historyPageSize}
            onPageChange={setHistoryPage}
            onPageSizeChange={setHistoryPageSize}
          />
        </div>
      )}

      {/* MODAL 1: CHI TIẾT HỒ SƠ & LỊCH SỬ THAO TÁC */}
      {detailModalDoc && (() => {
        const recordStatus = getRecordStatusMeta(detailModalDoc);
        const isIncomingDetail = detailModalDoc.docCategory === 'incoming';
        const categoryStyleDetail = isIncomingDetail ? SEMANTIC_STATUS_STYLES.incoming : SEMANTIC_STATUS_STYLES.outgoing;
        const bookObj = bookMap.get(getDocBookId(detailModalDoc));
        const issueDateDisplay = normalizeDateOnlyDdMmYyyy(detailModalDoc.issueDate) || '—';
        const signDateDisplay = normalizeDateOnlyDdMmYyyy(detailModalDoc.signDate) || '—';
        const receivedDateDisplay =
          normalizeDateOnlyDdMmYyyy(detailModalDoc.receivedDate) ||
          formatVietnamTimestampDate(detailModalDoc.receivedAt || detailModalDoc.createdAt);

        return (
          <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
            <div className="bg-white rounded-xl max-w-2xl w-full p-4 sm:p-6 shadow-xl max-h-[92dvh] overflow-y-auto space-y-5 animate-scaleUp">
              <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 flex-wrap min-w-0">
                  <h3 className="text-base font-bold text-slate-900">Chi tiết hồ sơ văn bản</h3>
                  <span className={`font-mono text-sm font-bold px-2.5 py-0.5 rounded border break-all ${categoryStyleDetail.badge}`}>
                    {detailModalDoc.fullDocNumber}
                  </span>
                  <span className={`text-[11px] font-bold px-2 py-0.5 rounded border ${recordStatus.className}`}>
                    {recordStatus.label}
                  </span>
                  <span className="font-mono text-[10px] text-slate-500 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                    Mã hồ sơ: {detailModalDoc.id}
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => setDetailModalDoc(null)}
                  className="text-slate-400 hover:text-slate-700 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center shrink-0"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs text-slate-700">
                <div>
                  <span className="text-slate-400 block">Sổ văn bản:</span>
                  <span className="font-semibold text-slate-900">
                    {bookObj?.name || detailModalDoc.bookName || 'Đoàn trường'}
                  </span>
                  {(bookObj?.code || detailModalDoc.bookCode) && (
                    <span className="ml-1.5 font-mono text-[10px] px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 text-slate-600">
                      {bookObj?.code || detailModalDoc.bookCode}
                    </span>
                  )}
                </div>
                <div>
                  <span className="text-slate-400 block">Năm / Loại văn bản:</span>
                  <span className="font-semibold text-slate-900">
                    {detailModalDoc.year} ·{' '}
                    {isIncomingDetail
                      ? 'Văn bản đến (PDF)'
                      : `${detailModalDoc.typeName} (${detailModalDoc.typeCode}) · Số TT: ${detailModalDoc.sequenceNumber}`}
                  </span>
                </div>

                <div>
                  <span className="text-slate-400 block">Ngày ban hành:</span>
                  <span className="font-semibold font-mono text-slate-900">{issueDateDisplay}</span>
                </div>
                <div>
                  <span className="text-slate-400 block">
                    {isIncomingDetail ? 'Ngày tiếp nhận:' : 'Ngày cấp vào sổ:'}
                  </span>
                  <span className="font-semibold font-mono text-slate-900">{receivedDateDisplay}</span>
                </div>

                <div>
                  <span className="text-slate-400 block">Ngày ký:</span>
                  <span className="font-semibold font-mono text-slate-900">{signDateDisplay}</span>
                </div>

                <div>
                  <span className="text-slate-400 block">Cơ quan ban hành / Đơn vị:</span>
                  <span className="font-semibold text-slate-900">
                    {detailModalDoc.docCategory === 'incoming'
                      ? detailModalDoc.issuingAgency || 'Chưa xác định'
                      : bookObj?.unitName || settings?.unitName || 'Đoàn Trường THPT Nguyễn Chí Thanh'}
                  </span>
                </div>

                <div>
                  <span className="text-slate-400 block">
                    {detailModalDoc.docCategory === 'incoming' ? 'Người nhập:' : 'Người lấy số:'}
                  </span>
                  <span className="font-semibold text-slate-900">
                    {detailModalDoc.issuedBy} ({detailModalDoc.issuedByUsername})
                  </span>
                </div>

                <div className="sm:col-span-2">
                  <span className="text-slate-400 block">Người ký &amp; Chức vụ:</span>
                  <span className="font-semibold text-slate-900">
                    {detailModalDoc.signer} ({detailModalDoc.signerRole})
                  </span>
                </div>

                <div className="sm:col-span-2">
                  <span className="text-slate-400 block">Tên văn bản:</span>
                  <p className="font-bold text-slate-900 text-sm mt-0.5 bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                    {detailModalDoc.docName || detailModalDoc.title}
                  </p>
                </div>

                <div className="sm:col-span-2">
                  <span className="text-slate-400 block">Trích yếu / Nội dung:</span>
                  <p className="font-medium text-slate-700 text-xs mt-0.5 bg-slate-50 p-2.5 rounded-lg border border-slate-100 leading-relaxed">
                    {detailModalDoc.title || <span className="italic text-slate-400">Không có</span>}
                  </p>
                </div>

                {detailModalDoc.notes && (
                  <div className="sm:col-span-2">
                    <span className="text-slate-400 block">Ghi chú:</span>
                    <span className="text-slate-700">{detailModalDoc.notes}</span>
                  </div>
                )}

                {/* Gợi ý tên file */}
                {(() => {
                  const suggestedFileName = generateSuggestedFileName(
                    detailModalDoc.fullDocNumber,
                    detailModalDoc.docName || detailModalDoc.title
                  );
                  return (
                    <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1.5">
                      <span className="text-[11px] font-bold text-slate-600 uppercase">Tên tệp chuẩn hóa:</span>
                      <div className="flex items-center gap-2">
                        <div className="flex-1 px-3 py-1.5 bg-white border border-slate-300 rounded font-mono text-xs font-bold text-slate-900 overflow-x-auto whitespace-nowrap">
                          {detailModalDoc.driveFileName || suggestedFileName}
                        </div>
                        <button
                          type="button"
                          onClick={() => {
                            navigator.clipboard?.writeText(detailModalDoc.driveFileName || suggestedFileName).catch(() => {});
                            setCopiedDetailFileName(true);
                            setTimeout(() => setCopiedDetailFileName(false), 2000);
                          }}
                          className="px-3 py-1.5 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded shrink-0"
                        >
                          {copiedDetailFileName ? 'Đã chép' : 'Sao chép'}
                        </button>
                      </div>
                    </div>
                  );
                })()}

                {/* Tệp đính kèm theo nhóm & Phiên bản (Hỗ trợ nhiều tệp + Bản chính thức) */}
                {(() => {
                  const effectiveAttachments = getEffectiveAttachments(detailModalDoc);
                  const groups: AttachmentGroup[] = ['word_draft', 'signed_pdf', 'appendix'];
                  const canManageFiles =
                    detailModalDoc.status === 'active' &&
                    (canRetryOrAttachDoc(detailModalDoc) || canReplaceDocFile(detailModalDoc));

                  return (
                    <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-3">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <span className="text-xs font-bold text-slate-800 uppercase tracking-wide flex items-center gap-1.5">
                          <Paperclip className="w-4 h-4 text-sky-700" />
                          <span>Tệp đính kèm &amp; Phiên bản ({effectiveAttachments.length})</span>
                        </span>

                        <div className="flex flex-wrap items-center gap-2">
                          <span
                            className={`text-[11px] font-semibold px-2 py-0.5 rounded border ${
                              detailModalDoc.driveStatus === 'saved'
                                ? 'text-emerald-700 bg-emerald-50 border-emerald-200'
                                : detailModalDoc.driveStatus === 'failed'
                                ? 'text-rose-700 bg-rose-50 border-rose-200'
                                : detailModalDoc.driveStatus === 'uploading'
                                ? 'text-sky-700 bg-sky-50 border-sky-200'
                                : 'text-slate-500 bg-white border-slate-200'
                            }`}
                          >
                            {getDriveStatusLabel(detailModalDoc)}
                          </span>

                          {detailModalDoc.wordGenerationStatus === 'failed' && detailModalDoc.templateId && canRetryOrAttachDoc(detailModalDoc) && (
                            <button
                              type="button"
                              disabled={retryingWordTemplate}
                              onClick={() => handleRetryWordTemplateInDetail(detailModalDoc)}
                              className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold bg-rose-700 hover:bg-rose-800 text-white rounded-md disabled:opacity-50"
                            >
                              <RefreshCw className={`w-3 h-3 ${retryingWordTemplate ? 'animate-spin' : ''}`} />
                              <span>{retryingWordTemplate ? 'Đang tạo lại...' : 'Thử tạo & lưu lại Word'}</span>
                            </button>
                          )}

                          {canManageFiles && (
                            <button
                              type="button"
                              onClick={() => {
                                setAttTargetId(undefined);
                                setAttGroup(detailModalDoc.docCategory === 'incoming' ? 'signed_pdf' : 'word_draft');
                                setAttLabel('');
                                setAttFile(null);
                                setAttError(null);
                                setAttSetOfficial(effectiveAttachments.length === 0);
                                setShowAddAttachmentForm((prev) => !prev);
                              }}
                              className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-md"
                            >
                              <Plus className="w-3 h-3" />
                              <span>Thêm tệp đính kèm</span>
                            </button>
                          )}
                        </div>
                      </div>

                      {attError && (
                        <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg text-xs">
                          {attError}
                        </div>
                      )}

                      {showAddAttachmentForm && canManageFiles && (
                        <form
                          onSubmit={handleUploadMultiAttachment}
                          className="bg-white border border-sky-200 rounded-lg p-3 space-y-2.5 text-xs"
                        >
                          <div className="flex items-center justify-between gap-2">
                            <strong className="text-slate-900">
                              {attTargetId ? 'Tải phiên bản mới cho tệp đã chọn' : 'Thêm tệp đính kèm mới'}
                            </strong>
                            <button
                              type="button"
                              onClick={() => {
                                setShowAddAttachmentForm(false);
                                setAttTargetId(undefined);
                              }}
                              className="text-slate-400 hover:text-slate-700"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            <div className="space-y-1">
                              <label className="text-[11px] font-semibold text-slate-600">Nhóm tệp</label>
                              <select
                                value={attGroup}
                                disabled={Boolean(attTargetId)}
                                onChange={(e) => setAttGroup(e.target.value as AttachmentGroup)}
                                className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-md text-xs"
                              >
                                <option value="word_draft">Word soạn thảo</option>
                                <option value="signed_pdf">PDF đã ký</option>
                                <option value="appendix">Phụ lục</option>
                              </select>
                            </div>
                            <div className="space-y-1">
                              <label className="text-[11px] font-semibold text-slate-600">Nhãn mô tả (tùy chọn)</label>
                              <input
                                type="text"
                                value={attLabel}
                                onChange={(e) => setAttLabel(e.target.value)}
                                placeholder="Ví dụ: Bản ký số, Phụ lục 1..."
                                className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-md text-xs"
                              />
                            </div>
                          </div>

                          <input
                            type="file"
                            accept=".doc,.docx,.pdf,.xls,.xlsx,.png,.jpg,.jpeg"
                            onChange={(e) => setAttFile(e.target.files?.[0] || null)}
                            className="block w-full text-xs text-slate-700 file:mr-3 file:py-1.5 file:px-3 file:rounded-md file:border file:border-slate-300 file:text-xs file:font-semibold file:bg-slate-50 hover:file:bg-slate-100"
                          />

                          <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                            <label className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer">
                              <input
                                type="checkbox"
                                checked={attSetOfficial}
                                onChange={(e) => setAttSetOfficial(e.target.checked)}
                                className="rounded border-slate-300 text-sky-700"
                              />
                              <span>Đặt làm bản chính thức (ưu tiên khi bấm Mở tệp)</span>
                            </label>

                            <button
                              type="submit"
                              disabled={attUploading || !attFile}
                              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-md"
                            >
                              <Upload className="w-3.5 h-3.5" />
                              <span>{attUploading ? 'Đang tải lên Drive...' : 'Lưu tệp lên Drive'}</span>
                            </button>
                          </div>
                        </form>
                      )}

                      {effectiveAttachments.length === 0 ? (
                        <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                          <span className="text-[11px] text-slate-500">Chưa có tệp đính kèm trên Google Drive.</span>
                          {canRetryOrAttachDoc(detailModalDoc) && (
                            <button
                              type="button"
                              onClick={() => openUploadModal(detailModalDoc, 'retry')}
                              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg"
                            >
                              <FileUp className="w-3.5 h-3.5" />
                              <span>Đính kèm tệp chính</span>
                            </button>
                          )}
                        </div>
                      ) : (
                        <div className="space-y-3">
                          {groups.map((grp) => {
                            const items = effectiveAttachments.filter((a) => a.group === grp);
                            if (items.length === 0) return null;
                            return (
                              <div key={grp} className="space-y-1.5">
                                <h5 className="text-[11px] font-bold text-slate-600 uppercase tracking-wide">
                                  {ATTACHMENT_GROUP_LABELS[grp]} ({items.length})
                                </h5>
                                <div className="space-y-2">
                                  {items.map((att) => {
                                    const isSaved = att.driveStatus === 'saved' && Boolean(att.driveFileUrl);
                                    const isLegacyVirtual = att.id.startsWith('legacy-');
                                    return (
                                      <div
                                        key={att.id}
                                        className="bg-white border border-slate-200 rounded-lg p-3 space-y-2"
                                      >
                                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                                          <div className="min-w-0 space-y-0.5">
                                            <div className="flex flex-wrap items-center gap-1.5">
                                              <span className="font-mono font-bold text-slate-900 break-all">
                                                {att.fileName}
                                              </span>
                                              <span className="text-[10px] font-mono font-bold px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200">
                                                v{att.currentVersion}
                                              </span>
                                              {att.isOfficial && (
                                                <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-900 border border-emerald-300">
                                                  Bản chính thức
                                                </span>
                                              )}
                                              {att.label && (
                                                <span className="text-[10px] px-1.5 py-0.5 rounded bg-sky-50 text-sky-800 border border-sky-200">
                                                  {att.label}
                                                </span>
                                              )}
                                            </div>
                                            <p className="text-[11px] text-slate-500">
                                              {att.driveStatus === 'saved'
                                                ? 'Đã lưu trên Drive'
                                                : att.driveStatus === 'failed'
                                                ? `Lưu thất bại${att.driveError ? `: ${att.driveError}` : ''}`
                                                : 'Đang lưu'}
                                              {att.fileSize ? ` · ${formatFileSize(att.fileSize)}` : ''}
                                              {att.uploadedAt ? ` · ${new Date(att.uploadedAt).toLocaleString('vi-VN')}` : ''}
                                            </p>
                                          </div>

                                          <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                                            {isSaved && (
                                              <button
                                                type="button"
                                                onClick={() => handleOpenDriveFile(detailModalDoc, isLegacyVirtual ? undefined : att.id)}
                                                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-md"
                                              >
                                                <ExternalLink className="w-3 h-3" />
                                                <span>Mở tệp</span>
                                              </button>
                                            )}

                                            {!att.isOfficial && isSaved && !isLegacyVirtual && canManageFiles && (
                                              <button
                                                type="button"
                                                disabled={settingOfficialId === att.id}
                                                onClick={() => handleSetOfficialAttachment(detailModalDoc, att.id)}
                                                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-semibold bg-white hover:bg-emerald-50 border border-emerald-300 text-emerald-800 rounded-md disabled:opacity-50"
                                              >
                                                <CheckCircle2 className="w-3 h-3" />
                                                <span>{settingOfficialId === att.id ? 'Đang đặt...' : 'Đặt làm bản chính thức'}</span>
                                              </button>
                                            )}

                                            {canManageFiles && (
                                              <button
                                                type="button"
                                                onClick={() => {
                                                  if (isLegacyVirtual && canReplaceDocFile(detailModalDoc)) {
                                                    openUploadModal(detailModalDoc, isSaved ? 'replace' : 'retry');
                                                  } else {
                                                    setAttTargetId(isLegacyVirtual ? undefined : att.id);
                                                    setAttGroup(att.group);
                                                    setAttLabel(att.label || '');
                                                    setAttSetOfficial(att.isOfficial);
                                                    setAttFile(null);
                                                    setAttError(null);
                                                    setShowAddAttachmentForm(true);
                                                  }
                                                }}
                                                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 rounded-md"
                                              >
                                                <Upload className="w-3 h-3 text-sky-700" />
                                                <span>Thay phiên bản mới</span>
                                              </button>
                                            )}
                                          </div>
                                        </div>

                                        {att.versions && att.versions.length > 0 && (
                                          <details className="text-[11px] text-slate-600 pt-1 border-t border-slate-100">
                                            <summary className="cursor-pointer font-semibold text-slate-700 hover:text-sky-800">
                                              Lịch sử phiên bản ({att.versions.length})
                                            </summary>
                                            <div className="mt-1.5 space-y-1 pl-2">
                                              {att.versions.map((ver) => (
                                                <div key={ver.version} className="flex flex-wrap items-center justify-between gap-2 py-0.5">
                                                  <span>
                                                    <strong className="font-mono text-slate-800">v{ver.version}</strong> · {ver.fileName} ·{' '}
                                                    {ver.uploadedBy} ({new Date(ver.uploadedAt).toLocaleString('vi-VN')})
                                                    {ver.driveStatus === 'failed' && (
                                                      <span className="ml-1 text-rose-700 font-semibold">(Lỗi)</span>
                                                    )}
                                                  </span>
                                                  {ver.driveFileUrl && (
                                                    <a
                                                      href={ver.driveFileUrl}
                                                      target="_blank"
                                                      rel="noopener noreferrer"
                                                      className="text-sky-700 hover:underline font-semibold"
                                                    >
                                                      Mở v{ver.version}
                                                    </a>
                                                  )}
                                                </div>
                                              ))}
                                            </div>
                                          </details>
                                        )}
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })()}

                {detailModalDoc.status === 'cancelled' && (
                  <div className="sm:col-span-2 p-3 bg-rose-50 border border-rose-200 rounded-lg text-rose-800">
                    <strong className="block text-xs font-bold text-rose-900">Thông tin hủy số:</strong>
                    <p className="mt-1">Lý do: {detailModalDoc.cancelReason}</p>
                    <p className="text-[11px] text-rose-600 mt-0.5">
                      Thực hiện bởi: {detailModalDoc.cancelledBy} vào lúc{' '}
                      {detailModalDoc.cancelledAt ? new Date(detailModalDoc.cancelledAt).toLocaleString('vi-VN') : ''}
                    </p>
                  </div>
                )}

                {detailModalDoc.status === 'deleted' && (() => {
                  const delSyncMeta = getExcelDeleteSyncMeta(detailModalDoc);
                  return (
                    <div className="sm:col-span-2 p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-rose-900 space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <strong className="text-xs font-bold text-rose-950">Thông tin xóa văn bản (VB đã xóa):</strong>
                        <span className={`text-[11px] font-semibold px-2 py-0.5 rounded border ${delSyncMeta.className}`}>
                          {delSyncMeta.label}
                        </span>
                      </div>
                      <p className="text-xs">
                        <strong>Lý do xóa:</strong> {detailModalDoc.deleteReason || '—'}
                      </p>
                      <p className="text-[11px] text-rose-700">
                        Thực hiện bởi: <strong>{detailModalDoc.deletedBy || '—'}</strong> vào lúc{' '}
                        <span className="font-mono">
                          {detailModalDoc.deletedAtVietnam ||
                            (detailModalDoc.deletedAt ? new Date(detailModalDoc.deletedAt).toLocaleString('vi-VN') : '')}
                        </span>
                      </p>
                      {detailModalDoc.excelSyncError && delSyncMeta.needsRetry && (
                        <p className="text-[11px] text-amber-900 bg-amber-50 border border-amber-200 rounded px-2.5 py-1.5">
                          Lỗi đồng bộ Google Trang tính: {detailModalDoc.excelSyncError}
                        </p>
                      )}
                      {delSyncMeta.needsRetry && canRetryExcelSync(detailModalDoc) && (
                        <button
                          type="button"
                          onClick={() => handleRetryExcelSync(detailModalDoc)}
                          disabled={retryingSyncDocId === detailModalDoc.id}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-amber-600 hover:bg-amber-700 disabled:opacity-60 text-white rounded-lg"
                        >
                          <RotateCw className={`w-3.5 h-3.5 ${retryingSyncDocId === detailModalDoc.id ? 'animate-spin' : ''}`} />
                          <span>{retryingSyncDocId === detailModalDoc.id ? 'Đang đồng bộ lại...' : 'Thử đồng bộ lại Google Trang tính'}</span>
                        </button>
                      )}
                    </div>
                  );
                })()}
              </div>

              {/* Việc cần làm & Thời hạn liên quan hồ sơ */}
              {(() => {
                const confirmedSugIds = new Set(
                  detailTasks.map((t) => t.suggestionId).filter(Boolean) as string[]
                );
                const unconfirmedSuggestions = (detailModalDoc.pdfTaskSuggestions || []).filter(
                  (s) => !confirmedSugIds.has(s.id)
                );
                if (!detailTasksLoading && detailTasks.length === 0 && unconfirmedSuggestions.length === 0) {
                  return null;
                }
                return (
                  <div className="border-t border-slate-100 pt-4 space-y-2.5">
                    <div className="flex items-center justify-between gap-2">
                      <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide flex items-center gap-1.5">
                        <ListChecks className="w-4 h-4 text-sky-700" />
                        <span>Việc cần làm &amp; Thời hạn từ văn bản</span>
                      </h4>
                    </div>

                    {detailTaskError && (
                      <p className="text-xs text-rose-700 font-medium">{detailTaskError}</p>
                    )}

                    {detailTasksLoading ? (
                      <p className="text-xs text-slate-500">Đang tải đầu việc...</p>
                    ) : (
                      <div className="space-y-2 text-xs">
                        {detailTasks.map((task) => (
                          <div
                            key={task.id}
                            className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-2"
                          >
                            <div className="space-y-0.5 min-w-0">
                              <div className="flex flex-wrap items-center gap-1.5">
                                <strong className="text-slate-900">{task.title}</strong>
                                {task.dueDate ? (
                                  <span className="inline-flex items-center gap-1 font-mono text-[11px] px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-800">
                                    <CalendarClock className="w-3 h-3 text-sky-700" />
                                    {task.dueDate}
                                  </span>
                                ) : (
                                  <span className="text-[11px] px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 font-semibold">
                                    Cần xác nhận
                                  </span>
                                )}
                              </div>
                              <p className="text-[11px] text-slate-500">
                                Phụ trách: <strong>{task.assigneeName}</strong>
                                {task.recipient ? ` · Nơi nhận: ${task.recipient}` : ''}
                              </p>
                            </div>
                            <div className="flex items-center gap-1 shrink-0">
                              {(['pending', 'in_progress', 'done'] as DocumentTaskStatus[]).map((st) => (
                                <button
                                  key={st}
                                  type="button"
                                  disabled={updatingDetailTaskId === task.id}
                                  onClick={() => handleUpdateDetailTaskStatus(task, st)}
                                  className={`px-2 py-1 text-[11px] font-semibold rounded border ${
                                    task.status === st
                                      ? 'bg-sky-700 text-white border-sky-700'
                                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                                  }`}
                                >
                                  {st === 'pending' ? 'Chờ' : st === 'in_progress' ? 'Đang làm' : 'Xong'}
                                </button>
                              ))}
                            </div>
                          </div>
                        ))}

                        {unconfirmedSuggestions.map((sug) => (
                          <div
                            key={sug.id}
                            className="p-2.5 bg-amber-50/70 border border-amber-200 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-2"
                          >
                            <div className="space-y-0.5 min-w-0">
                              <div className="flex flex-wrap items-center gap-1.5">
                                <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
                                  Đề xuất chưa giao
                                </span>
                                <strong className="text-slate-900">{sug.taskTitle}</strong>
                                <span className="text-[11px] font-mono text-slate-700">
                                  {sug.dueDate || sug.rawDeadlineText || 'Cần xác nhận'}
                                </span>
                              </div>
                              <p className="text-[11px] text-slate-600">
                                Trang {sug.page}: {sug.sourceExcerpt}
                              </p>
                            </div>
                            <div className="flex items-center gap-1.5 shrink-0">
                              <select
                                aria-label="Chọn người nhận việc"
                                value={sugAssigneeMap[sug.id] || currentUser.id}
                                onChange={(e) =>
                                  setSugAssigneeMap((prev) => ({ ...prev, [sug.id]: e.target.value }))
                                }
                                className="min-h-[32px] px-2 py-1 bg-white border border-slate-300 rounded text-[11px]"
                              >
                                {detailAssignees.length > 0 ? (
                                  detailAssignees.map((a) => (
                                    <option key={a.id} value={a.id}>
                                      {a.displayName}
                                    </option>
                                  ))
                                ) : (
                                  <option value={currentUser.id}>{currentUser.displayName || currentUser.fullName}</option>
                                )}
                              </select>
                              <button
                                type="button"
                                disabled={confirmingSugId === sug.id}
                                onClick={() => handleConfirmSuggestionInDetail(detailModalDoc, sug)}
                                className="px-2.5 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded text-[11px] shrink-0 disabled:opacity-50"
                              >
                                {confirmingSugId === sug.id ? 'Đang giao...' : 'Xác nhận giao việc'}
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* Lịch sử thao tác */}
              <div className="border-t border-slate-100 pt-4">
                <div className="flex items-center gap-2 mb-3">
                  <History className="w-4 h-4 text-sky-700" />
                  <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wide">
                    Lịch sử thao tác (Lưu vết bất biến)
                  </h4>
                </div>

                <div className="space-y-2 text-xs">
                  {detailModalDoc.auditLogs && detailModalDoc.auditLogs.length > 0 ? (
                    detailModalDoc.auditLogs.map((log, idx) => {
                      const actionMeta = getAuditActionMeta(log.action);
                      return (
                        <div key={log.id || idx} className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg space-y-1">
                          <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-500">
                            <div className="flex items-center gap-2 flex-wrap">
                              <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded border ${actionMeta.className}`}>
                                {actionMeta.label}
                              </span>
                              <span className="font-semibold text-slate-800">
                                {log.userName}
                                {log.username ? ` (${log.username})` : ''}
                              </span>
                            </div>
                            <span className="font-mono">
                              {log.timestampVietnam || new Date(log.timestamp).toLocaleString('vi-VN')}
                            </span>
                          </div>
                          {log.reason && (
                            <p className="text-[11px] text-slate-700">
                              <strong>Lý do / Ghi nhận:</strong> {log.reason}
                            </p>
                          )}
                          {log.changes && Object.keys(log.changes).length > 0 && (
                            <div className="space-y-0.5 pl-2">
                              {Object.entries(log.changes).map(([field, delta]: any, fIdx) => (
                                <p key={fIdx} className="text-[11px] text-slate-600">
                                  • {formatAuditFieldLabel(field)}:{' '}
                                  {delta.old !== null && delta.old !== undefined ? (
                                    <>
                                      <del className="text-rose-600">{delta.old || '(trống)'}</del> →{' '}
                                    </>
                                  ) : null}
                                  <ins className="text-emerald-700 no-underline font-medium">
                                    {delta.new || '(trống)'}
                                  </ins>
                                </p>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })
                  ) : (
                    <p className="text-slate-400 italic">Không có lịch sử thao tác.</p>
                  )}
                </div>
              </div>

              <div className="flex justify-end pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setDetailModalDoc(null)}
                  className="px-4 py-2 text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg"
                >
                  Đóng
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* MODAL 2: CHỈNH SỬA THÔNG TIN */}
      {editModalDoc && (() => {
        const docBookId = getDocBookId(editModalDoc);
        const canManageThisBook =
          isSystemAdmin(currentUser) || hasBookPermission(currentUser, docBookId, 'manage_book');
        const bookObj = bookMap.get(docBookId);

        return (
          <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
            <div className="bg-white rounded-xl max-w-lg w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3 gap-2">
                <div className="min-w-0">
                  <h3 className="text-base font-bold text-slate-900">Sửa thông tin văn bản</h3>
                  <span className="font-mono text-xs text-sky-800 font-semibold break-all">{editModalDoc.fullDocNumber}</span>
                </div>
                <button
                  type="button"
                  onClick={() => setEditModalDoc(null)}
                  className="text-slate-400 hover:text-slate-700 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center shrink-0"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {editModalDoc.docCategory === 'incoming' ? (
                <div className="p-3 bg-sky-50 border border-sky-200 rounded-lg text-sky-950 text-xs space-y-2">
                  <label className="font-bold text-slate-700 block">Số/Ký hiệu gốc (Văn bản đến)</label>
                  <input
                    type="text"
                    value={editFullDocNumber}
                    onChange={(e) => setEditFullDocNumber(e.target.value)}
                    className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded font-mono font-bold text-base sm:text-xs"
                  />
                </div>
              ) : canManageThisBook ? (
                <div className="p-3 bg-sky-50 border border-sky-200 rounded-lg text-sky-950 text-xs space-y-2">
                  <strong className="block font-semibold">Quyền Quản trị sổ: Có thể điều chỉnh số văn bản</strong>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                    <div>
                      <label className="font-bold text-slate-700 block mb-1">Số thứ tự (STT)</label>
                      <input
                        type="number"
                        min={1}
                        value={editSequenceNumber}
                        onChange={(e) => {
                          const val = parseInt(e.target.value, 10) || 1;
                          setEditSequenceNumber(val);
                          const seqStr = val < 10 ? `0${val}` : `${val}`;
                          setEditFullDocNumber(`${seqStr}-${editModalDoc.typeCode}/${bookObj?.unitCode || settings?.unitCode || 'ĐTN'}`);
                        }}
                        className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded font-mono font-bold text-base sm:text-xs"
                      />
                    </div>
                    <div>
                      <label className="font-bold text-slate-700 block mb-1">Số/Ký hiệu đầy đủ</label>
                      <input
                        type="text"
                        value={editFullDocNumber}
                        onChange={(e) => setEditFullDocNumber(e.target.value)}
                        className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded font-mono font-bold text-base sm:text-xs"
                      />
                    </div>
                  </div>
                </div>
              ) : null}

              <form onSubmit={handleSaveEdit} className="space-y-4 text-xs">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Tên văn bản *</label>
                  <input
                    type="text"
                    value={editDocName}
                    onChange={(e) => setEditDocName(e.target.value)}
                    required
                    className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

                {editModalDoc.docCategory === 'incoming' && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div className="space-y-1">
                      <label className="font-bold text-slate-700">Cơ quan ban hành</label>
                      <input
                        type="text"
                        value={editIssuingAgency}
                        onChange={(e) => setEditIssuingAgency(e.target.value)}
                        className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="font-bold text-slate-700">Ngày ban hành</label>
                      <input
                        type="text"
                        value={editIssueDate}
                        onChange={(e) => setEditIssueDate(e.target.value)}
                        onBlur={(e) => setEditIssueDate(normalizeDateOnlyDdMmYyyy(e.target.value))}
                        placeholder="dd/mm/yyyy"
                        className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="font-bold text-slate-700">Ngày ký</label>
                      <input
                        type="text"
                        value={editSignDate}
                        onChange={(e) => setEditSignDate(e.target.value)}
                        onBlur={(e) => setEditSignDate(normalizeDateOnlyDdMmYyyy(e.target.value))}
                        placeholder="dd/mm/yyyy"
                        className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                      />
                    </div>
                  </div>
                )}

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Trích yếu / Nội dung văn bản</label>
                  <textarea
                    value={editTitle}
                    onChange={(e) => setEditTitle(e.target.value)}
                    rows={3}
                    className="w-full p-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">Người ký *</label>
                    <input
                      type="text"
                      value={editSigner}
                      onChange={(e) => setEditSigner(e.target.value)}
                      required
                      className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">Chức vụ *</label>
                    <input
                      type="text"
                      value={editSignerRole}
                      onChange={(e) => setEditSignerRole(e.target.value)}
                      required
                      className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                    />
                  </div>
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Ghi chú</label>
                  <input
                    type="text"
                    value={editNotes}
                    onChange={(e) => setEditNotes(e.target.value)}
                    className="w-full min-h-[42px] sm:min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  />
                </div>

                {editError && (
                  <div className="p-2.5 bg-rose-50 text-rose-700 rounded-lg text-xs">{editError}</div>
                )}

                <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setEditModalDoc(null)}
                    className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                  >
                    Hủy
                  </button>
                  <button
                    type="submit"
                    disabled={editSubmitting}
                    className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs"
                  >
                    {editSubmitting ? 'Đang lưu...' : 'Lưu thay đổi'}
                  </button>
                </div>
              </form>
            </div>
          </div>
        );
      })()}

      {/* MODAL 3: HỦY SỐ VĂN BẢN */}
      {cancelModalDoc && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 gap-2">
              <div className="min-w-0">
                <h3 className="text-base font-bold text-rose-700 flex items-center gap-2">
                  <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0" />
                  <span>Xác nhận hủy số văn bản</span>
                </h3>
                <span className="font-mono text-xs font-bold text-slate-900 break-all">{cancelModalDoc.fullDocNumber}</span>
              </div>
              <button
                type="button"
                onClick={() => setCancelModalDoc(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 text-xs leading-relaxed">
              Văn bản bị hủy vẫn được giữ lại trong sổ theo dõi cùng lịch sử thao tác. Số văn bản{' '}
              <strong className="font-mono">{cancelModalDoc.fullDocNumber}</strong> sẽ <strong>không được cấp lại</strong>.
            </div>

            <form onSubmit={handleSaveCancel} className="space-y-4 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-800">
                  Lý do hủy số văn bản <span className="text-rose-600">*</span>
                </label>
                <textarea
                  value={cancelReason}
                  onChange={(e) => setCancelReason(e.target.value)}
                  placeholder="Nhập lý do chi tiết..."
                  rows={3}
                  required
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-rose-600"
                />
              </div>

              {cancelError && <div className="p-2.5 bg-rose-50 text-rose-700 rounded-lg text-xs">{cancelError}</div>}

              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setCancelModalDoc(null)}
                  className="px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Đóng
                </button>
                <button
                  type="submit"
                  disabled={cancelSubmitting}
                  className="px-4 py-2 text-xs font-semibold bg-rose-700 hover:bg-rose-800 text-white rounded-lg shadow-xs"
                >
                  {cancelSubmitting ? 'Đang xử lý...' : 'Xác nhận hủy số'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 4: XÓA VĂN BẢN (CHUYỂN SANG VB ĐÃ XÓA + BẮT BUỘC NHẬP LÝ DO & XÁC NHẬN) */}
      {deleteModalDoc && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-base font-bold text-rose-700 flex items-center gap-2">
                <Trash2 className="w-5 h-5 text-rose-600 shrink-0" />
                <span>Xác nhận xóa văn bản</span>
              </h3>
              <button
                type="button"
                onClick={() => !deleteSubmitting && setDeleteModalDoc(null)}
                className="text-slate-400 hover:text-slate-700 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center shrink-0"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-900 text-xs leading-relaxed space-y-1.5">
              <p>
                Văn bản số{' '}
                <strong className="font-mono text-rose-950 font-bold break-all">{deleteModalDoc.fullDocNumber}</strong>{' '}
                sẽ được chuyển khỏi danh sách hiệu lực sang mục <strong>VB đã xóa</strong> (và sheet <strong>VB đã xóa</strong> trong Excel).
              </p>
              <p className="text-[11px] text-rose-800 break-words">
                <strong>Tên văn bản:</strong> {deleteModalDoc.docName || deleteModalDoc.title}
              </p>
              <p className="text-[11px] text-rose-800">
                Số đã cấp được giữ nguyên trong dãy, tệp đính kèm trên Google Drive được bảo toàn và lịch sử xóa sẽ được ghi nhận bất biến theo tài khoản đang đăng nhập.
              </p>
            </div>

            <form onSubmit={handleConfirmDelete} className="space-y-3.5 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-800 block">
                  Lý do xóa văn bản <span className="text-rose-600">*</span>
                </label>
                <textarea
                  value={deleteReason}
                  onChange={(e) => setDeleteReason(e.target.value)}
                  placeholder="Nhập lý do xóa (ví dụ: Nhập trùng hồ sơ, văn bản thu hồi không ban hành...)"
                  rows={3}
                  required
                  disabled={deleteSubmitting}
                  className="w-full p-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-rose-600"
                />
              </div>

              <label className="flex items-start gap-2.5 p-2.5 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={deleteConfirmed}
                  onChange={(e) => setDeleteConfirmed(e.target.checked)}
                  disabled={deleteSubmitting}
                  className="mt-0.5 w-4 h-4 text-rose-700 border-slate-300 rounded focus:ring-rose-600 shrink-0"
                />
                <span className="text-xs text-slate-700 leading-snug">
                  Tôi xác nhận xóa văn bản này khỏi sổ hiệu lực và chuyển sang <strong>VB đã xóa</strong>.
                </span>
              </label>

              {deleteError && <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-700 rounded-lg text-xs">{deleteError}</div>}

              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  disabled={deleteSubmitting}
                  onClick={() => setDeleteModalDoc(null)}
                  className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  disabled={deleteSubmitting || !deleteReason.trim() || !deleteConfirmed}
                  className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white rounded-lg shadow-xs inline-flex items-center gap-1.5"
                >
                  {deleteSubmitting ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>Đang xóa &amp; đồng bộ...</span>
                    </>
                  ) : (
                    <span>Xác nhận xóa văn bản</span>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 5: ĐÍNH KÈM / THỬ LƯU LẠI / THAY TỆP TRÊN GOOGLE DRIVE */}
      {uploadModalDoc && (() => {
        const isIncomingPdf = uploadModalDoc.docCategory === 'incoming';
        const isReplace = uploadModalMode === 'replace';

        return (
          <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
            <div className="bg-white rounded-xl max-w-lg w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div>
                  <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                    <FileUp className="w-5 h-5 text-sky-700" />
                    <span>
                      {isReplace
                        ? isIncomingPdf
                          ? 'Thay tệp PDF trên Google Drive'
                          : 'Thay tệp Word trên Google Drive'
                        : uploadModalDoc.driveStatus === 'failed'
                        ? isIncomingPdf
                          ? 'Thử lưu lại bản PDF lên Google Drive'
                          : 'Thử lưu lại tệp Word lên Google Drive'
                        : isIncomingPdf
                        ? 'Đính kèm bản PDF vào hồ sơ'
                        : 'Đính kèm tệp Word vào số văn bản'}
                    </span>
                  </h3>
                  <span className="font-mono text-xs font-bold text-sky-800">
                    {uploadModalDoc.fullDocNumber}
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => !uploadModalSubmitting && setUploadModalDoc(null)}
                  className="text-slate-400 hover:text-slate-700 p-1"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-1">
                <div className="text-slate-500">Tên văn bản:</div>
                <div className="font-bold text-slate-900">{uploadModalDoc.docName || uploadModalDoc.title}</div>
                {isReplace && uploadModalDoc.driveFileName && (
                  <div className="text-sky-800 pt-1">
                    <strong>Tệp hiện tại:</strong> <span className="font-mono">{uploadModalDoc.driveFileName}</span>
                  </div>
                )}
                {uploadModalDoc.driveStatus === 'failed' && uploadModalDoc.driveError && (
                  <div className="text-rose-700 pt-1">
                    <strong>Lý do lỗi trước đó:</strong> {uploadModalDoc.driveError}
                  </div>
                )}
              </div>

              <form onSubmit={handleUploadAttachmentSubmit} className="space-y-4 text-xs">
                <div className="space-y-2">
                  <label className="font-bold text-slate-700 block">
                    {isIncomingPdf
                      ? `Chọn tệp PDF (.pdf, tối đa ${maxWordFileSizeMB} MB)`
                      : `Chọn tệp Word (.doc hoặc .docx, tối đa ${maxWordFileSizeMB} MB)`}
                  </label>
                  <label className="flex flex-col items-center justify-center gap-2 p-4 border-2 border-dashed border-slate-300 hover:border-sky-400 rounded-lg bg-slate-50/70 cursor-pointer transition-colors">
                    <Upload className="w-6 h-6 text-sky-700" />
                    <span className="font-semibold text-slate-700">
                      {uploadModalFile
                        ? uploadModalFile.name
                        : isIncomingPdf
                        ? 'Bấm để chọn tệp PDF từ máy tính / điện thoại'
                        : 'Bấm để chọn tệp Word từ máy tính / điện thoại'}
                    </span>
                    {uploadModalFile && (
                      <span className="text-[11px] text-slate-500">
                        Dung lượng: {formatFileSize(uploadModalFile.size)}
                      </span>
                    )}
                    <input
                      type="file"
                      accept={isIncomingPdf ? '.pdf,application/pdf' : '.doc,.docx'}
                      disabled={uploadModalSubmitting}
                      className="hidden"
                      onChange={(e) => {
                        const f = e.target.files?.[0] || null;
                        setUploadModalError(null);
                        if (!f) return;
                        const err = validateAttachmentFile(f, isIncomingPdf);
                        if (err) {
                          setUploadModalError(err);
                          setUploadModalFile(null);
                          return;
                        }
                        setUploadModalFile(f);
                      }}
                    />
                  </label>
                </div>

                {uploadModalFile && (
                  <div className="p-3 bg-sky-50 border border-sky-200 rounded-lg space-y-1">
                    <span className="text-[11px] font-semibold text-sky-900 block">
                      Tên tệp sẽ lưu trên Google Drive:
                    </span>
                    <div className="font-mono font-bold text-sky-950 break-all">
                      {isReplace
                        ? buildStandardDriveFileName(
                            uploadModalDoc.fullDocNumber,
                            uploadModalDoc.docName || uploadModalDoc.title,
                            uploadModalFile.name
                          )
                        : uploadModalDoc.driveFileName ||
                          buildStandardDriveFileName(
                            uploadModalDoc.fullDocNumber,
                            uploadModalDoc.docName || uploadModalDoc.title,
                            uploadModalFile.name
                          )}
                    </div>
                  </div>
                )}

                {uploadModalError && (
                  <div className="p-3 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg text-xs">
                    {uploadModalError}
                  </div>
                )}

                <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    disabled={uploadModalSubmitting}
                    onClick={() => setUploadModalDoc(null)}
                    className="px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                  >
                    Đóng
                  </button>
                  <button
                    type="submit"
                    disabled={uploadModalSubmitting || !uploadModalFile}
                    className="px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white rounded-lg shadow-xs inline-flex items-center gap-1.5"
                  >
                    {uploadModalSubmitting ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        <span>Đang lưu tệp...</span>
                      </>
                    ) : (
                      <>
                        <FileUp className="w-3.5 h-3.5" />
                        <span>
                          {isReplace
                            ? 'Xác nhận Thay tệp'
                            : uploadModalDoc.driveStatus === 'failed'
                            ? 'Thử lưu lại'
                            : 'Tải lên Google Drive'}
                        </span>
                      </>
                    )}
                  </button>
                </div>
              </form>
            </div>
          </div>
        );
      })()}

      {/* MODAL 6: XÁC NHẬN ĐỒNG BỘ HÀNG LOẠT GOOGLE TRANG TÍNH */}
      {bulkSyncConfirmMode && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 gap-2">
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <RotateCw className="w-5 h-5 text-sky-700 shrink-0" />
                <span>Xác nhận đồng bộ Google Trang tính</span>
              </h3>
              <button
                type="button"
                disabled={bulkSyncSubmitting}
                onClick={() => setBulkSyncConfirmMode(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-3.5 bg-sky-50 border border-sky-200 rounded-lg text-xs text-sky-950 space-y-1.5">
              <p>
                <strong>Phạm vi:</strong> {currentScopeBookName}
              </p>
              <p>
                <strong>Số lượng sẽ xử lý:</strong>{' '}
                <span className="font-mono font-bold text-sky-900">
                  {bulkSyncConfirmMode === 'selected' ? selectedDeletedIds.length : pendingExcelSyncCount}
                </span>{' '}
                văn bản đã xóa
                {bulkSyncConfirmMode === 'pending'
                  ? ' (toàn bộ mục đang chờ / lỗi trong sổ, bao gồm các trang khác).'
                  : ' đã chọn.'}
              </p>
              <p className="text-[11px] text-sky-800">
                Hệ thống sẽ đồng bộ theo mã hồ sơ cố định sang sheet <strong>VB đã xóa</strong> trên Google Trang tính và chỉ xóa khỏi sheet theo dõi chính sau khi xác nhận thành công.
              </p>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                disabled={bulkSyncSubmitting}
                onClick={() => setBulkSyncConfirmMode(null)}
                className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
              >
                Hủy
              </button>
              <button
                type="button"
                disabled={bulkSyncSubmitting}
                onClick={handleConfirmBulkSync}
                className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 disabled:opacity-60 text-white rounded-lg inline-flex items-center gap-1.5"
              >
                {bulkSyncSubmitting ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    <span>Đang đồng bộ...</span>
                  </>
                ) : (
                  <>
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    <span>Xác nhận đồng bộ</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 7: DỌN KHỎI DANH SÁCH VB ĐÃ XÓA (CHỈ ADMIN TOÀN QUYỀN) */}
      {archiveModalOpen && isAdmin && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 gap-2">
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Archive className="w-5 h-5 text-slate-700 shrink-0" />
                <span>Dọn khỏi danh sách VB đã xóa</span>
              </h3>
              <button
                type="button"
                disabled={archiveSubmitting}
                onClick={() => setArchiveModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-2.5 text-xs">
              <p className="text-slate-600 leading-relaxed">
                Thao tác này chỉ gắn trạng thái <strong>Đã lưu trữ</strong> và ẩn khỏi danh sách mặc định ở <strong>VB đã xóa</strong>. Hồ sơ văn bản, số đã cấp trong dãy, liên kết Google Drive và dữ liệu trong Google Trang tính <strong>vẫn được bảo toàn nguyên vẹn</strong>.
              </p>

              <div className="space-y-2">
                {selectedDeletedIds.length > 0 && (
                  <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                    <input
                      type="radio"
                      name="archiveMode"
                      checked={archiveMode === 'selected'}
                      onChange={() => setArchiveMode('selected')}
                      className="mt-0.5 text-sky-700"
                    />
                    <span>
                      <strong>Chỉ dọn các mục đã chọn đã đồng bộ:</strong>{' '}
                      <span className="font-mono font-bold text-emerald-700">{selectedSyncedCount}</span> mục hợp lệ
                      {selectedUnsyncedCount > 0 && (
                        <span className="block text-[11px] text-amber-800 mt-0.5">
                          ({selectedUnsyncedCount} mục đang chọn chưa đồng bộ Google Trang tính sẽ được giữ lại)
                        </span>
                      )}
                    </span>
                  </label>
                )}

                <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                  <input
                    type="radio"
                    name="archiveMode"
                    checked={archiveMode === 'all_synced'}
                    onChange={() => setArchiveMode('all_synced')}
                    className="mt-0.5 text-sky-700"
                  />
                  <span>
                    <strong>Tất cả mục đã đồng bộ trong {currentScopeBookName}:</strong>{' '}
                    <span className="font-mono font-bold text-emerald-700">{syncedUnarchivedInBookCount}</span> mục
                    {pendingExcelSyncCount > 0 && (
                      <span className="block text-[11px] text-amber-800 mt-0.5">
                        ({pendingExcelSyncCount} mục chờ/lỗi đồng bộ sẽ tiếp tục hiển thị để xử lý)
                      </span>
                    )}
                  </span>
                </label>
              </div>

              {archiveError && (
                <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg">
                  {archiveError}
                </div>
              )}
            </div>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
              <button
                type="button"
                disabled={archiveSubmitting}
                onClick={() => setArchiveModalOpen(false)}
                className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
              >
                Hủy
              </button>
              <button
                type="button"
                disabled={
                  archiveSubmitting ||
                  (archiveMode === 'selected' ? selectedSyncedCount === 0 : syncedUnarchivedInBookCount === 0)
                }
                onClick={handleConfirmArchiveDeleted}
                className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-slate-900 hover:bg-slate-800 disabled:opacity-50 text-white rounded-lg inline-flex items-center gap-1.5"
              >
                {archiveSubmitting ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    <span>Đang lưu trữ...</span>
                  </>
                ) : (
                  <>
                    <Archive className="w-3.5 h-3.5" />
                    <span>Xác nhận dọn danh sách</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 8: DỌN LỊCH SỬ THAO TÁC THEO THỜI GIAN (CHỈ ADMIN TOÀN QUYỀN HỆ THỐNG) */}
      {cleanupModalOpen && isAdmin && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
          <div className="bg-white rounded-xl max-w-lg w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[92dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 gap-2">
              <h3 className="text-base font-bold text-rose-800 flex items-center gap-2">
                <Eraser className="w-5 h-5 text-rose-700 shrink-0" />
                <span>Dọn lịch sử thao tác theo thời gian</span>
              </h3>
              <button
                type="button"
                disabled={cleanupSubmitting}
                onClick={() => setCleanupModalOpen(false)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleConfirmHistoryCleanup} className="space-y-4 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-700 block">Sổ văn bản áp dụng *</label>
                <select
                  value={cleanupBookId}
                  onChange={(e) => setCleanupBookId(e.target.value)}
                  disabled={cleanupSubmitting}
                  className="w-full min-h-[40px] px-2.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-rose-600"
                >
                  {books.map((b) => (
                    <option key={b.id} value={b.id}>
                      {b.name} ({b.code})
                    </option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <label className="font-bold text-slate-700 block">Chọn phạm vi thời gian (giờ Việt Nam)</label>

                <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                  <input
                    type="radio"
                    name="cleanupScope"
                    checked={cleanupScope === 'today'}
                    onChange={() => setCleanupScope('today')}
                    className="mt-0.5 text-rose-700"
                  />
                  <span>
                    <strong className="text-slate-900">Xóa lịch sử hôm nay</strong>
                    <span className="block text-[11px] text-slate-500">Chỉ xóa các bản ghi phát sinh trong ngày hôm nay theo giờ Việt Nam.</span>
                  </span>
                </label>

                <label className="flex items-start gap-2.5 p-2.5 rounded-lg border border-slate-200 cursor-pointer hover:bg-slate-50">
                  <input
                    type="radio"
                    name="cleanupScope"
                    checked={cleanupScope === 'last_10_days'}
                    onChange={() => setCleanupScope('last_10_days')}
                    className="mt-0.5 text-rose-700"
                  />
                  <span>
                    <strong className="text-slate-900">Xóa lịch sử 10 ngày gần nhất</strong>
                    <span className="block text-[11px] text-slate-500">Bao gồm hôm nay và 9 ngày trước đó (theo lịch Việt Nam).</span>
                  </span>
                </label>

                <div className="p-2.5 rounded-lg border border-slate-200 space-y-2 hover:bg-slate-50">
                  <label className="flex items-start gap-2.5 cursor-pointer">
                    <input
                      type="radio"
                      name="cleanupScope"
                      checked={cleanupScope === 'before_date'}
                      onChange={() => setCleanupScope('before_date')}
                      className="mt-0.5 text-rose-700"
                    />
                    <span>
                      <strong className="text-slate-900">Xóa lịch sử trước ngày…</strong>
                      <span className="block text-[11px] text-slate-500">Chỉ xóa các bản ghi trước 00:00 của ngày được chọn.</span>
                    </span>
                  </label>
                  {cleanupScope === 'before_date' && (
                    <input
                      type="date"
                      aria-label="Chọn mốc ngày"
                      value={cleanupBeforeDate}
                      onChange={(e) => setCleanupBeforeDate(e.target.value)}
                      required
                      className="w-full min-h-[38px] px-2.5 bg-white border border-slate-300 rounded-lg text-xs"
                    />
                  )}
                </div>

                <div className="p-2.5 rounded-lg border border-rose-200 bg-rose-50/40 space-y-2">
                  <label className="flex items-start gap-2.5 cursor-pointer">
                    <input
                      type="radio"
                      name="cleanupScope"
                      checked={cleanupScope === 'all'}
                      onChange={() => setCleanupScope('all')}
                      className="mt-0.5 text-rose-700"
                    />
                    <span>
                      <strong className="text-rose-900">Xóa toàn bộ lịch sử của sổ này</strong>
                      <span className="block text-[11px] text-rose-700">Yêu cầu nhập chính xác cụm từ xác nhận bên dưới.</span>
                    </span>
                  </label>
                  {cleanupScope === 'all' && (
                    <div className="space-y-1 pl-6">
                      <label className="block text-[11px] font-semibold text-rose-900">
                        Nhập <code className="px-1 py-0.5 bg-white border border-rose-300 rounded font-mono">XÓA LỊCH SỬ</code> để xác nhận:
                      </label>
                      <input
                        type="text"
                        value={cleanupConfirmText}
                        onChange={(e) => setCleanupConfirmText(e.target.value)}
                        placeholder="XÓA LỊCH SỬ"
                        className="w-full min-h-[38px] px-2.5 bg-white border border-rose-300 rounded-lg font-mono font-bold text-xs text-rose-950 focus:outline-hidden focus:ring-2 focus:ring-rose-600"
                      />
                    </div>
                  )}
                </div>
              </div>

              {/* Preview Box */}
              <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                {cleanupPreviewLoading ? (
                  <div className="flex items-center gap-2 text-slate-500">
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-sky-700" />
                    <span>Đang đếm số bản ghi lịch sử khớp phạm vi...</span>
                  </div>
                ) : cleanupPreview ? (
                  <>
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-slate-500">Sổ áp dụng:</span>
                      <strong className="text-slate-900">{cleanupPreview.bookName} ({cleanupPreview.bookCode})</strong>
                    </div>
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-slate-500">Khoảng thời gian:</span>
                      <strong className="text-slate-900 text-right">{cleanupPreview.rangeLabel}</strong>
                    </div>
                    <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-200">
                      <span className="font-semibold text-slate-700">Số bản ghi lịch sử sẽ xóa:</span>
                      <span className="font-mono text-sm font-extrabold text-rose-700">
                        {cleanupPreview.matchCount} bản ghi
                      </span>
                    </div>
                  </>
                ) : null}
                <p className="text-[11px] text-slate-500 pt-1">
                  Hồ sơ văn bản, số thứ tự đã cấp chống trùng và nhật ký quản trị đợt dọn lịch sử vẫn được giữ nguyên.
                </p>
              </div>

              {cleanupPreviewError && (
                <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg">
                  {cleanupPreviewError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  disabled={cleanupSubmitting}
                  onClick={() => setCleanupModalOpen(false)}
                  className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={
                    cleanupSubmitting ||
                    cleanupPreviewLoading ||
                    !cleanupPreview ||
                    cleanupPreview.matchCount === 0 ||
                    (cleanupScope === 'all' && cleanupConfirmText.trim() !== 'XÓA LỊCH SỬ')
                  }
                  className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white rounded-lg inline-flex items-center gap-1.5"
                >
                  {cleanupSubmitting ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>Đang dọn lịch sử...</span>
                    </>
                  ) : (
                    <>
                      <Eraser className="w-3.5 h-3.5" />
                      <span>Xác nhận dọn lịch sử</span>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

