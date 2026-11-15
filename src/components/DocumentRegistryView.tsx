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
} from 'lucide-react';
import { api } from '../api';
import type {
  DocumentAuditAction,
  DocumentHistoryEntry,
  DocumentRecord,
  DocumentTypeConfig,
  RegistryBook,
  SystemSettings,
  User,
} from '../types';
import { getTypeColorStyle } from '../utils/documentColors';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  formatFileSize,
  generateSuggestedFileName,
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
  onRefresh: () => void;
  onDocumentUpdated: (doc: DocumentRecord) => void;
  initialOpenDocId?: string | null;
  onClearInitialOpenDocId?: () => void;
}

function getRecordStatusMeta(doc: DocumentRecord): { label: string; className: string } {
  if (doc.status === 'deleted') {
    return {
      label: 'Đã xóa',
      className: 'text-rose-900 bg-rose-100 border-rose-300',
    };
  }
  if (doc.status === 'cancelled') {
    return {
      label: 'Đã hủy',
      className: 'text-rose-800 bg-rose-50 border-rose-200',
    };
  }
  if (doc.docCategory === 'incoming') {
    return {
      label: 'Đã ghi sổ văn bản đến',
      className: 'text-sky-800 bg-sky-50 border-sky-200',
    };
  }
  return {
    label: 'Đã cấp số',
    className: 'text-emerald-800 bg-emerald-50 border-emerald-200',
  };
}

function getDriveStatusLabel(doc: DocumentRecord): string {
  switch (doc.driveStatus) {
    case 'saved':
      return 'Đã lưu lên Drive';
    case 'failed':
      return doc.docCategory === 'incoming' ? 'Lưu tệp thất bại' : 'Đã cấp số · Lưu tệp thất bại';
    case 'uploading':
      return 'Đang lưu';
    default:
      return 'Không đính kèm';
  }
}

function getExcelDeleteSyncMeta(doc: DocumentRecord): {
  label: string;
  className: string;
  needsRetry: boolean;
} {
  if (doc.excelDeleteSyncStatus === 'synced') {
    return {
      label: 'Đã chuyển sang sheet VB đã xóa',
      className: 'text-emerald-800 bg-emerald-50 border-emerald-200',
      needsRetry: false,
    };
  }
  if (doc.excelDeleteSyncStatus === 'failed' || doc.excelDeleteSyncStatus === 'pending') {
    return {
      label: 'Đã xóa trên hệ thống · Chờ đồng bộ Excel',
      className: 'text-amber-900 bg-amber-50 border-amber-300',
      needsRetry: true,
    };
  }
  return {
    label: 'Đã chuyển VB đã xóa (Chưa bật Excel trực tuyến)',
    className: 'text-slate-600 bg-slate-100 border-slate-200',
    needsRetry: false,
  };
}

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
      return { label: 'Đồng bộ lại Excel', className: 'bg-amber-50 text-amber-900 border-amber-200' };
    case 'request_number':
      return { label: 'Xin số', className: 'bg-slate-100 text-slate-700 border-slate-200' };
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
  syncStatus: 'Trạng thái đồng bộ Excel',
  needsReview: 'Cờ cần kiểm tra',
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
  onSelectBookId,
  currentUser,
  settings,
  serverYear,
  onRefresh,
  onDocumentUpdated,
  initialOpenDocId,
  onClearInitialOpenDocId,
}) => {
  const [registrySection, setRegistrySection] = useState<'active' | 'deleted' | 'history'>('active');
  // Search & Filter State
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'outgoing' | 'incoming'>('all');
  const [selectedYear, setSelectedYear] = useState<string>('all');
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [selectedType, setSelectedType] = useState<string>('all');
  const [selectedSigner, setSelectedSigner] = useState<string>('all');
  const [selectedStatus, setSelectedStatus] = useState<string>('all');
  const [selectedDriveStatus, setSelectedDriveStatus] = useState<string>('all');
  const [sortBy, setSortBy] = useState<'newest' | 'oldest' | 'number_asc' | 'number_desc'>('newest');

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

  // Deleted documents tab filters & retry Excel state
  const [deletedSearch, setDeletedSearch] = useState('');
  const [deletedByFilter, setDeletedByFilter] = useState('all');
  const [deletedFromDate, setDeletedFromDate] = useState('');
  const [deletedToDate, setDeletedToDate] = useState('');
  const [deletedSyncFilter, setDeletedSyncFilter] = useState<'all' | 'synced' | 'pending_sync'>('all');
  const [retryingSyncDocId, setRetryingSyncDocId] = useState<string | null>(null);
  const [syncRetryNotice, setSyncRetryNotice] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

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

  // Attachment Upload / Retry / Replace modal state
  const [uploadModalDoc, setUploadModalDoc] = useState<DocumentRecord | null>(null);
  const [uploadModalMode, setUploadModalMode] = useState<'retry' | 'replace'>('retry');
  const [uploadModalFile, setUploadModalFile] = useState<File | null>(null);
  const [uploadModalSubmitting, setUploadModalSubmitting] = useState(false);
  const [uploadModalError, setUploadModalError] = useState<string | null>(null);

  // Export Excel loading state
  const [isExporting, setIsExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

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

  // Auto-open specific document when navigated from "Mở hồ sơ đã có"
  useEffect(() => {
    if (!initialOpenDocId) return;
    const target =
      documents.find((d) => d.id === initialOpenDocId) ||
      deletedDocuments.find((d) => d.id === initialOpenDocId);
    if (target) {
      setDetailModalDoc(target);
      onClearInitialOpenDocId?.();
      return;
    }
    api
      .getDocument(initialOpenDocId)
      .then((res) => {
        if (res.document) {
          setDetailModalDoc(res.document);
        }
      })
      .catch(() => {})
      .finally(() => {
        onClearInitialOpenDocId?.();
      });
  }, [initialOpenDocId, documents, deletedDocuments, onClearInitialOpenDocId]);

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

  const handleOpenDriveFile = async (doc: DocumentRecord) => {
    setFileActionError(null);
    setOpeningFileId(doc.id);
    try {
      const res = await api.openDocumentFile(doc.id);
      if (res.url) {
        window.open(res.url, '_blank', 'noopener,noreferrer');
      }
    } catch (err: any) {
      setFileActionError(err.message || 'Không thể mở liên kết tệp Google Drive.');
    } finally {
      setOpeningFileId(null);
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

  // Derive unique years and signers
  const availableYears = useMemo(() => {
    const years = Array.from(new Set(documents.map((d) => d.year)));
    if (years.length === 0) {
      years.push(
        serverYear ||
          Number(
            new Intl.DateTimeFormat('en-CA', {
              timeZone: 'Asia/Ho_Chi_Minh',
              year: 'numeric',
            }).format(new Date())
          )
      );
    }
    return years.sort((a, b) => b - a);
  }, [documents, serverYear]);

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

    if (selectedYear !== 'all') {
      result = result.filter((d) => d.year === parseInt(selectedYear, 10));
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
        message: res.message || `Đã đồng bộ lại văn bản ${doc.fullDocNumber} sang sheet VB đã xóa.`,
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
        message: err.message || 'Đồng bộ lại Excel chưa thành công. Vui lòng thử lại sau.',
      });
    } finally {
      setRetryingSyncDocId(null);
    }
  };

  // Deleted documents derived lists
  const pendingExcelSyncCount = useMemo(
    () =>
      deletedDocuments.filter(
        (d) => d.excelDeleteSyncStatus === 'failed' || d.excelDeleteSyncStatus === 'pending'
      ).length,
    [deletedDocuments]
  );

  const availableDeletedActors = useMemo(() => {
    const map = new Map<string, string>();
    deletedDocuments.forEach((d) => {
      const key = d.deletedByUsername || d.deletedByUserId || '';
      if (key && !map.has(key)) {
        map.set(key, d.deletedBy || key);
      }
    });
    return Array.from(map.entries()).map(([username, label]) => ({ username, label }));
  }, [deletedDocuments]);

  const filteredDeletedDocuments = useMemo(() => {
    let list = [...deletedDocuments];
    if (effectiveBookScope && effectiveBookScope !== 'all') {
      list = list.filter((d) => getDocBookId(d) === effectiveBookScope);
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
    deletedDocuments,
    effectiveBookScope,
    deletedByFilter,
    deletedSyncFilter,
    deletedFromDate,
    deletedToDate,
    deletedSearch,
  ]);

  // Export Excel preserving ALL active filters + sort order + permissions
  const handleExportExcel = async () => {
    setIsExporting(true);
    setExportError(null);
    try {
      await api.downloadExcel({
        bookId: activeBookId && activeBookId !== 'all' ? activeBookId : undefined,
        year: selectedYear !== 'all' ? selectedYear : 'all',
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
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
            Sổ Theo Dõi Văn Bản
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
            Tra cứu nhanh văn bản đi và văn bản đến, theo dõi trạng thái tệp trên Google Drive và xuất báo cáo Excel chuẩn.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0 w-full lg:w-auto">
          {viewableBooks.length > 0 && (
            <div className="flex items-center gap-2 bg-white border border-slate-300 rounded-lg px-3 py-2 min-h-[42px] shadow-2xs text-xs w-full sm:w-auto min-w-0">
              <BookOpen className="w-4 h-4 text-sky-700 shrink-0" />
              <select
                aria-label="Lọc theo sổ văn bản"
                value={viewableBooks.length === 1 ? viewableBooks[0].id : activeBookId}
                onChange={(e) => onSelectBookId(e.target.value)}
                className="flex-1 sm:flex-initial min-w-0 font-semibold text-base sm:text-xs text-slate-800 bg-transparent focus:outline-hidden cursor-pointer"
              >
                {viewableBooks.length > 1 && (
                  <option value="all">Tất cả sổ có quyền ({viewableBooks.length})</option>
                )}
                {viewableBooks.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.name} ({b.unitCode})
                  </option>
                ))}
              </select>
            </div>
          )}

          <button
            type="button"
            onClick={onRefresh}
            className="flex-1 sm:flex-initial justify-center flex items-center gap-1.5 min-h-[42px] px-3.5 py-2 text-xs font-semibold bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-lg shadow-2xs transition-colors"
          >
            <RefreshCw className="w-4 h-4 text-sky-700 shrink-0" />
            <span>Làm mới dữ liệu</span>
          </button>

          {canExportScope && (
            <button
              type="button"
              onClick={handleExportExcel}
              disabled={isExporting}
              className="flex-1 sm:flex-initial justify-center flex items-center gap-2 min-h-[42px] px-4 py-2 text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 disabled:opacity-60 text-white rounded-lg shadow-xs transition-colors"
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
            {deletedDocuments.length}
          </span>
          {pendingExcelSyncCount > 0 && (
            <span
              className="inline-flex items-center gap-1 text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300"
              title="Có văn bản chờ đồng bộ sang sheet VB đã xóa"
            >
              Chờ Excel: {pendingExcelSyncCount}
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
      {/* Filter & Search Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-12 gap-2.5">
          <div className="sm:col-span-2 lg:col-span-2 xl:col-span-3 relative min-w-0">
            <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Tìm số VB, tên VB, cơ quan, tên tệp..."
              className="w-full min-h-[44px] h-11 sm:h-10 pl-9 pr-3 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            />
          </div>

          <div className="lg:col-span-1 xl:col-span-2 min-w-0">
            <select
              aria-label="Lọc nhóm văn bản"
              value={selectedCategory}
              onChange={(e) => setSelectedCategory(e.target.value as any)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600 font-semibold text-slate-700"
            >
              <option value="all">Tất cả nhóm VB</option>
              <option value="outgoing">Văn bản đi</option>
              <option value="incoming">Văn bản đến (PDF)</option>
            </select>
          </div>

          <div className="lg:col-span-1 xl:col-span-1 min-w-0">
            <select
              aria-label="Lọc theo năm"
              value={selectedYear}
              onChange={(e) => setSelectedYear(e.target.value)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            >
              <option value="all">Mọi năm</option>
              {availableYears.map((yr) => (
                <option key={yr} value={yr}>
                  {yr}
                </option>
              ))}
            </select>
          </div>

          <div className="lg:col-span-1 xl:col-span-1 min-w-0">
            <select
              aria-label="Lọc theo tháng"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            >
              <option value="all">Mọi tháng</option>
              {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
                <option key={m} value={m}>
                  Tháng {m}
                </option>
              ))}
            </select>
          </div>

          <div className="lg:col-span-1 xl:col-span-2 min-w-0">
            <select
              aria-label="Lọc theo loại văn bản"
              value={selectedType}
              onChange={(e) => setSelectedType(e.target.value)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
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

          <div className="lg:col-span-1 xl:col-span-1 min-w-0">
            <select
              aria-label="Lọc trạng thái hồ sơ"
              value={selectedStatus}
              onChange={(e) => setSelectedStatus(e.target.value)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            >
              <option value="all">Hồ sơ</option>
              <option value="active">Hiệu lực</option>
              <option value="cancelled">Đã hủy</option>
            </select>
          </div>

          <div className="lg:col-span-1 xl:col-span-2 min-w-0">
            <select
              aria-label="Lọc trạng thái tệp Drive"
              value={selectedDriveStatus}
              onChange={(e) => setSelectedDriveStatus(e.target.value)}
              className="w-full min-h-[44px] h-11 sm:h-10 px-2 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            >
              <option value="all">Tất cả trạng thái tệp</option>
              <option value="saved">Đã lưu lên Drive</option>
              <option value="failed">Lưu tệp thất bại</option>
              <option value="none">Không đính kèm</option>
            </select>
          </div>
        </div>

        {/* Sub-bar: Sort & Signer + Counts + Mobile switcher */}
        <div className="pt-2.5 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-xs text-slate-500">
          <div className="flex items-center gap-2.5 flex-wrap min-w-0">
            <div className="flex items-center gap-1.5">
              <span>Hiển thị:</span>
              <strong className="text-slate-800 font-mono">{filteredDocuments.length}</strong>
              <span>/</span>
              <strong className="text-slate-800 font-mono">{documents.length}</strong>
              <span>hồ sơ</span>
            </div>

            {availableSigners.length > 0 && (
              <select
                aria-label="Lọc theo người ký"
                value={selectedSigner}
                onChange={(e) => setSelectedSigner(e.target.value)}
                className="min-h-[36px] h-9 sm:h-8 max-w-full min-w-0 px-2 text-sm sm:text-xs bg-slate-50 border border-slate-200 rounded-md text-slate-700 focus:outline-hidden"
              >
                <option value="all">Mọi người ký</option>
                {availableSigners.map((s) => (
                  <option key={s} value={s}>
                    Ký: {s}
                  </option>
                ))}
              </select>
            )}
          </div>

          <div className="flex items-center justify-between sm:justify-end gap-2.5 flex-wrap">
            <div className="flex md:hidden items-center bg-slate-100 p-0.5 rounded-lg border border-slate-200">
              <button
                type="button"
                onClick={() => setMobileViewMode('cards')}
                className={`min-h-[34px] px-2.5 py-1 rounded text-[11px] font-semibold transition-colors flex items-center gap-1 ${
                  mobileViewMode === 'cards' ? 'bg-white text-sky-800 shadow-2xs' : 'text-slate-600'
                }`}
              >
                <LayoutList className="w-3.5 h-3.5" />
                <span>Thẻ</span>
              </button>
              <button
                type="button"
                onClick={() => setMobileViewMode('table')}
                className={`min-h-[34px] px-2.5 py-1 rounded text-[11px] font-semibold transition-colors flex items-center gap-1 ${
                  mobileViewMode === 'table' ? 'bg-white text-sky-800 shadow-2xs' : 'text-slate-600'
                }`}
              >
                <Table className="w-3.5 h-3.5" />
                <span>Bảng</span>
              </button>
            </div>

            <div className="flex items-center gap-1.5">
              <span>Sắp xếp:</span>
              <select
                aria-label="Sắp xếp danh sách"
                value={sortBy}
                onChange={(e: any) => setSortBy(e.target.value)}
                className="min-h-[36px] h-9 px-2 text-sm sm:text-xs bg-slate-50 sm:bg-transparent border border-slate-200 sm:border-0 rounded-md font-medium text-slate-700 cursor-pointer focus:outline-hidden"
              >
                <option value="newest">Mới nhất trước</option>
                <option value="oldest">Cũ nhất trước</option>
                <option value="number_asc">Số TT tăng dần</option>
                <option value="number_desc">Số TT giảm dần</option>
              </select>
            </div>
          </div>
        </div>
      </div>

      {/* Mobile Default Card View (No horizontal scroll) */}
      {mobileViewMode === 'cards' && (
        <div className="md:hidden space-y-3">
          {filteredDocuments.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-8 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">Chưa có văn bản nào phù hợp</p>
              <span className="text-xs text-slate-400 mt-1 block">Hãy đổi tiêu chí tìm kiếm hoặc cấp số mới.</span>
            </div>
          ) : (
            filteredDocuments.map((doc) => {
              const isCancelled = doc.status === 'cancelled';
              const isIncoming = doc.docCategory === 'incoming';
              const docTypeConfig = documentTypes.find((t) => t.code === doc.typeCode);
              const typeColor = getTypeColorStyle(doc.typeCode, docTypeConfig?.color);
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
                      <span className={`text-[11px] font-bold px-2 py-0.5 rounded ${typeColor.badgeBg} ${typeColor.badgeText} border ${typeColor.badgeBorder}`}>
                        {isIncoming ? 'Văn bản đến' : doc.typeName || doc.typeCode}
                      </span>
                      {doc.needsReview && (
                        <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
                          Cần kiểm tra
                        </span>
                      )}
                      <span className="font-mono font-bold text-slate-900 text-sm break-all">
                        {doc.fullDocNumber}
                      </span>
                      <button
                        type="button"
                        onClick={() => handleCopy(doc.id, doc.fullDocNumber)}
                        className="p-1.5 min-w-[32px] min-h-[32px] flex items-center justify-center text-slate-400 hover:text-slate-800 rounded"
                        title="Sao chép số văn bản"
                      >
                        {copiedId === doc.id ? (
                          <Check className="w-3.5 h-3.5 text-emerald-600" />
                        ) : (
                          <Copy className="w-3.5 h-3.5" />
                        )}
                      </button>
                    </div>

                    <span className={`text-[10px] font-bold px-2 py-0.5 rounded border shrink-0 ${recordStatus.className}`}>
                      {recordStatus.label}
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

                  {/* Separated Drive File Status & Actions */}
                  <div className="flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg border border-slate-200 bg-slate-50/70 text-[11px] flex-wrap">
                    <span
                      className={`font-semibold ${
                        doc.driveStatus === 'saved'
                          ? 'text-emerald-700'
                          : doc.driveStatus === 'failed'
                          ? 'text-rose-700'
                          : doc.driveStatus === 'uploading'
                          ? 'text-sky-700'
                          : 'text-slate-500'
                      }`}
                    >
                      {getDriveStatusLabel(doc)}
                    </span>

                    <div className="flex items-center gap-1.5 flex-wrap">
                      {doc.driveStatus === 'saved' && doc.driveFileUrl && (
                        <button
                          type="button"
                          onClick={() => handleOpenDriveFile(doc)}
                          disabled={openingFileId === doc.id}
                          className="inline-flex items-center gap-1 min-h-[34px] font-semibold text-emerald-700 bg-white px-2.5 py-1 rounded border border-emerald-200"
                        >
                          <ExternalLink className="w-3 h-3" />
                          <span>Mở tệp</span>
                        </button>
                      )}
                      {doc.driveStatus === 'saved' && canReplaceDocFile(doc) && (
                        <button
                          type="button"
                          onClick={() => openUploadModal(doc, 'replace')}
                          className="min-h-[34px] font-semibold text-sky-700 bg-white px-2.5 py-1 rounded border border-sky-200"
                        >
                          Thay tệp
                        </button>
                      )}
                      {(doc.driveStatus === 'failed' || doc.driveStatus === 'uploading') && canRetryOrAttachDoc(doc) && (
                        <button
                          type="button"
                          onClick={() => openUploadModal(doc, 'retry')}
                          className="min-h-[34px] font-semibold text-rose-700 bg-white px-2.5 py-1 rounded border border-rose-200"
                        >
                          Thử lưu lại
                        </button>
                      )}
                      {(!doc.driveStatus || doc.driveStatus === 'none') && canRetryOrAttachDoc(doc) && (
                        <button
                          type="button"
                          onClick={() => openUploadModal(doc, 'retry')}
                          className="inline-flex items-center gap-1 min-h-[34px] font-semibold text-sky-700 bg-white px-2.5 py-1 rounded border border-sky-200"
                        >
                          <FileUp className="w-3 h-3" />
                          <span>Đính kèm</span>
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Footer Actions */}
                  <div className="pt-2 border-t border-slate-100 flex items-center justify-between gap-2 text-xs">
                    <span className="text-[11px] text-slate-500 truncate min-w-0">
                      {bookObj?.name || doc.bookName}
                    </span>

                    <div className="flex items-center gap-1 bg-slate-50 p-1 rounded-lg border border-slate-200/80 shrink-0">
                      <button
                        type="button"
                        onClick={() => setDetailModalDoc(doc)}
                        className="p-2 min-w-[36px] min-h-[36px] flex items-center justify-center text-slate-600 hover:text-sky-700 hover:bg-white rounded"
                        title="Xem chi tiết"
                      >
                        <Eye className="w-4 h-4" />
                      </button>
                      {canEditDoc(doc) && (
                        <button
                          type="button"
                          onClick={() => openEditModal(doc)}
                          className="p-2 min-w-[36px] min-h-[36px] flex items-center justify-center text-slate-600 hover:text-amber-600 hover:bg-white rounded"
                          title="Sửa văn bản"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
                      )}
                      {canCancelDoc(doc) && (
                        <button
                          type="button"
                          onClick={() => openCancelModal(doc)}
                          className="p-2 min-w-[36px] min-h-[36px] flex items-center justify-center text-slate-600 hover:text-amber-700 hover:bg-white rounded"
                          title="Hủy số văn bản"
                        >
                          <Ban className="w-4 h-4" />
                        </button>
                      )}
                      {canDeleteDoc(doc) && (
                        <button
                          type="button"
                          onClick={() => openDeleteModal(doc)}
                          className="p-2 min-w-[36px] min-h-[36px] flex items-center justify-center text-slate-600 hover:text-rose-600 hover:bg-white rounded"
                          title="Xóa văn bản (chuyển sang VB đã xóa)"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}

      {/* Streamlined Desktop Table: 6 Core Columns + Compact Thao tác */}
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
                <th className="py-3 px-4 min-w-[260px]">Tên văn bản</th>
                <th className="py-3 px-3 w-32 text-center">Ngày ban hành</th>
                <th className="py-3 px-3 w-48">Cơ quan / Đơn vị</th>
                <th className="py-3 px-3 w-40 text-center">Trạng thái</th>
                <th className="py-3 px-3 w-48 text-center">Tệp</th>
                <th className="py-3 px-3 w-28 text-center">Thao tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
              {filteredDocuments.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-12 text-center text-slate-400">
                    <p className="font-medium text-slate-500">Chưa có văn bản nào phù hợp với bộ lọc</p>
                    <span className="text-xs">Hãy thử đổi tiêu chí tìm kiếm hoặc tạo cấp số mới.</span>
                  </td>
                </tr>
              ) : (
                filteredDocuments.map((doc) => {
                  const isCancelled = doc.status === 'cancelled';
                  const isIncoming = doc.docCategory === 'incoming';
                  const docTypeConfig = documentTypes.find((t) => t.code === doc.typeCode);
                  const typeColor = getTypeColorStyle(doc.typeCode, docTypeConfig?.color);
                  const recordStatus = getRecordStatusMeta(doc);
                  const displayIssueDate = normalizeDateOnlyDdMmYyyy(doc.issueDate) || '—';
                  const bookObj = bookMap.get(getDocBookId(doc));
                  const agencyOrUnit = isIncoming
                    ? doc.issuingAgency || 'Chưa xác định'
                    : bookObj?.unitName || settings?.unitName || 'Đoàn trường';

                  return (
                    <tr
                      key={doc.id}
                      className={`hover:bg-sky-50/50 transition-colors ${
                        isCancelled ? 'bg-rose-50/30' : ''
                      }`}
                    >
                      {/* 1. Số / Ký hiệu */}
                      <td className="py-3 px-3.5 align-top">
                        <div className="space-y-1">
                          <div className={`inline-flex items-center gap-1.5 font-mono font-bold ${typeColor.badgeText} ${typeColor.badgeBg} px-2.5 py-0.5 rounded border ${typeColor.badgeBorder} shadow-2xs`}>
                            <span className={`w-1.5 h-1.5 rounded-full ${typeColor.dotColor}`} />
                            <span>{doc.fullDocNumber}</span>
                            <button
                              type="button"
                              onClick={() => handleCopy(doc.id, doc.fullDocNumber)}
                              className="p-0.5 text-slate-400 hover:text-slate-800 ml-0.5"
                              title="Sao chép số văn bản"
                            >
                              {copiedId === doc.id ? (
                                <Check className="w-3.5 h-3.5 text-emerald-600" />
                              ) : (
                                <Copy className="w-3.5 h-3.5" />
                              )}
                            </button>
                          </div>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className="text-[11px] font-medium text-slate-500">
                              {isIncoming ? 'Văn bản đến' : doc.typeName}
                            </span>
                            {doc.needsReview && (
                              <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
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
                          className={`text-left font-semibold text-slate-900 hover:text-sky-800 transition-colors leading-snug ${
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
                      <td className="py-3 px-3 text-center font-mono whitespace-nowrap text-slate-800 align-top">
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

                      {/* 6. Tệp (Trạng thái Drive + Hành động tệp) */}
                      <td className="py-3 px-3 text-center align-top">
                        {doc.driveStatus === 'saved' && doc.driveFileUrl ? (
                          <div className="inline-flex flex-col items-center gap-1">
                            <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                              Đã lưu lên Drive
                            </span>
                            <div className="flex items-center gap-1.5">
                              <button
                                type="button"
                                onClick={() => handleOpenDriveFile(doc)}
                                disabled={openingFileId === doc.id}
                                className="text-[11px] font-semibold text-sky-700 hover:underline inline-flex items-center gap-0.5"
                                title={doc.driveFileName || 'Mở tệp'}
                              >
                                <ExternalLink className="w-3 h-3" />
                                <span>Mở tệp</span>
                              </button>
                              {canReplaceDocFile(doc) && (
                                <button
                                  type="button"
                                  onClick={() => openUploadModal(doc, 'replace')}
                                  className="text-[11px] font-semibold text-slate-600 hover:text-amber-700 hover:underline"
                                >
                                  · Thay tệp
                                </button>
                              )}
                            </div>
                          </div>
                        ) : doc.driveStatus === 'failed' || doc.driveStatus === 'uploading' ? (
                          <div className="inline-flex flex-col items-center gap-1">
                            <span
                              className={`text-[11px] font-semibold px-2 py-0.5 rounded border ${
                                doc.driveStatus === 'uploading'
                                  ? 'text-sky-700 bg-sky-50 border-sky-200'
                                  : 'text-rose-700 bg-rose-50 border-rose-200'
                              }`}
                              title={doc.driveError || ''}
                            >
                              {getDriveStatusLabel(doc)}
                            </span>
                            {canRetryOrAttachDoc(doc) && (
                              <button
                                type="button"
                                onClick={() => openUploadModal(doc, 'retry')}
                                className="text-[11px] font-semibold text-sky-700 hover:text-sky-900 underline"
                              >
                                Thử lưu lại
                              </button>
                            )}
                          </div>
                        ) : (
                          <div className="inline-flex items-center gap-1.5">
                            <span className="text-[11px] text-slate-400">Không đính kèm</span>
                            {canRetryOrAttachDoc(doc) && (
                              <button
                                type="button"
                                onClick={() => openUploadModal(doc, 'retry')}
                                className="p-1 text-sky-700 hover:bg-sky-50 rounded border border-sky-200"
                                title="Đính kèm tệp lên Drive"
                              >
                                <FileUp className="w-3 h-3" />
                              </button>
                            )}
                          </div>
                        )}
                      </td>

                      {/* 7. Thao tác */}
                      <td className="py-3 px-3 text-center whitespace-nowrap align-top">
                        <div className="flex items-center justify-center gap-1">
                          <button
                            type="button"
                            onClick={() => setDetailModalDoc(doc)}
                            className="p-1 text-slate-500 hover:text-sky-700 hover:bg-slate-100 rounded"
                            title="Xem chi tiết, ngày ký, người ký & lịch sử thao tác"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>

                          {canEditDoc(doc) && (
                            <button
                              type="button"
                              onClick={() => openEditModal(doc)}
                              className="p-1 text-slate-500 hover:text-amber-600 hover:bg-slate-100 rounded"
                              title="Sửa thông tin văn bản"
                            >
                              <Edit2 className="w-3.5 h-3.5" />
                            </button>
                          )}

                          {canCancelDoc(doc) && (
                            <button
                              type="button"
                              onClick={() => openCancelModal(doc)}
                              className="p-1 text-slate-500 hover:text-amber-700 hover:bg-slate-100 rounded"
                              title="Hủy số văn bản"
                            >
                              <Ban className="w-3.5 h-3.5" />
                            </button>
                          )}

                          {canDeleteDoc(doc) && (
                            <button
                              type="button"
                              onClick={() => openDeleteModal(doc)}
                              className="p-1 text-slate-500 hover:text-rose-600 hover:bg-slate-100 rounded"
                              title="Xóa văn bản (chuyển sang VB đã xóa)"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
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

          <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-2.5">
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
                aria-label="Lọc trạng thái đồng bộ Excel"
                value={deletedSyncFilter}
                onChange={(e) => setDeletedSyncFilter(e.target.value as any)}
                className="min-h-[42px] h-10 px-2.5 text-base sm:text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-hidden"
              >
                <option value="all">Mọi trạng thái Excel</option>
                <option value="pending_sync">Chờ đồng bộ Excel</option>
                <option value="synced">Đã đồng bộ VB đã xóa</option>
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

            <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
              <span>
                Hiển thị <strong className="font-mono text-slate-800">{filteredDeletedDocuments.length}</strong> văn bản đã xóa (số văn bản trong dãy và tệp trên Google Drive vẫn được bảo toàn).
              </span>
              {(deletedSearch || deletedByFilter !== 'all' || deletedSyncFilter !== 'all' || deletedFromDate || deletedToDate) && (
                <button
                  type="button"
                  onClick={() => {
                    setDeletedSearch('');
                    setDeletedByFilter('all');
                    setDeletedSyncFilter('all');
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

          {filteredDeletedDocuments.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">Không có văn bản đã xóa nào</p>
              <span className="text-xs text-slate-400 mt-1 block">
                Khi xóa văn bản khỏi danh sách hiệu lực, hồ sơ sẽ lưu tại đây và đồng bộ sang sheet &ldquo;VB đã xóa&rdquo;.
              </span>
            </div>
          ) : (
            <div className="space-y-3">
              {filteredDeletedDocuments.map((doc) => {
                const syncMeta = getExcelDeleteSyncMeta(doc);
                const bookObj = bookMap.get(getDocBookId(doc));
                return (
                  <div
                    key={doc.id}
                    className="bg-white border border-rose-200 rounded-xl p-4 shadow-2xs space-y-3"
                  >
                    <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-2">
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

                    {doc.excelSyncError && syncMeta.needsRetry && (
                      <div className="text-[11px] text-amber-900 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
                        <strong>Chi tiết lỗi đồng bộ Excel:</strong> {doc.excelSyncError}
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
        </div>
      )}

      {/* SECTION 3: LỊCH SỬ THAO TÁC (CHỈ ĐỌC, KHÔNG SỬA/XÓA) */}
      {registrySection === 'history' && (
        <div className="space-y-4">
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
                <option value="retry_excel_sync">Đồng bộ lại Excel</option>
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
              <div className="flex items-center gap-3">
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
          ) : historyLogs.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-slate-400">
              <p className="font-semibold text-slate-600 text-sm">Chưa có bản ghi lịch sử nào phù hợp</p>
              <span className="text-xs text-slate-400 mt-1 block">
                Mọi thao tác cấp số, sửa thông tin, thay tệp, hủy số và xóa văn bản đều được lưu vết tự động theo tài khoản.
              </span>
            </div>
          ) : (
            <div className="space-y-2.5">
              {historyLogs.map((entry) => {
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
        </div>
      )}

      {/* MODAL 1: CHI TIẾT HỒ SƠ & LỊCH SỬ THAO TÁC */}
      {detailModalDoc && (() => {
        const recordStatus = getRecordStatusMeta(detailModalDoc);
        const bookObj = bookMap.get(getDocBookId(detailModalDoc));
        const issueDateDisplay = normalizeDateOnlyDdMmYyyy(detailModalDoc.issueDate) || '—';
        const signDateDisplay = normalizeDateOnlyDdMmYyyy(detailModalDoc.signDate) || '—';

        return (
          <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-2.5 sm:p-4">
            <div className="bg-white rounded-xl max-w-2xl w-full p-4 sm:p-6 shadow-xl max-h-[92dvh] overflow-y-auto space-y-5 animate-scaleUp">
              <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2 flex-wrap min-w-0">
                  <h3 className="text-base font-bold text-slate-900">Chi tiết hồ sơ văn bản</h3>
                  <span className="font-mono text-sm font-bold text-sky-800 bg-sky-50 px-2.5 py-0.5 rounded border border-sky-200 break-all">
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
                    {detailModalDoc.docCategory === 'incoming'
                      ? 'Văn bản đến (PDF)'
                      : `${detailModalDoc.typeName} (${detailModalDoc.typeCode}) · Số TT: ${detailModalDoc.sequenceNumber}`}
                  </span>
                </div>

                <div>
                  <span className="text-slate-400 block">Ngày ban hành:</span>
                  <span className="font-semibold font-mono text-slate-900">{issueDateDisplay}</span>
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

                {/* Tệp & Liên kết Drive */}
                <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold text-slate-800 uppercase tracking-wide flex items-center gap-1.5">
                      <FolderCheck className="w-4 h-4 text-sky-700" />
                      <span>Tệp &amp; Liên kết Google Drive</span>
                    </span>
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
                  </div>

                  {detailModalDoc.driveStatus === 'saved' && (
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pt-1">
                      <div className="min-w-0 space-y-0.5">
                        <p className="font-mono font-bold text-slate-900 truncate">
                          Tên tệp: {detailModalDoc.driveFileName}
                        </p>
                        {detailModalDoc.driveFolderPath && (
                          <p className="text-[11px] text-slate-600">
                            Thư mục Drive: <strong className="text-slate-800">{detailModalDoc.driveFolderPath}</strong>
                          </p>
                        )}
                        {detailModalDoc.driveFileUrl && (
                          <p className="text-[11px] text-sky-800 truncate font-mono">
                            Liên kết Drive: {detailModalDoc.driveFileUrl}
                          </p>
                        )}
                        <p className="text-[11px] text-slate-500">
                          {detailModalDoc.driveFileSize ? `Dung lượng: ${formatFileSize(detailModalDoc.driveFileSize)}` : ''}
                          {detailModalDoc.driveUploadedAt
                            ? ` · Lưu lúc ${new Date(detailModalDoc.driveUploadedAt).toLocaleString('vi-VN')}`
                            : ''}
                          {detailModalDoc.driveFileId ? ` · ID: ${detailModalDoc.driveFileId}` : ''}
                        </p>
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        {detailModalDoc.driveFileUrl && (
                          <button
                            type="button"
                            onClick={() => handleOpenDriveFile(detailModalDoc)}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg shadow-2xs"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                            <span>Mở tệp</span>
                          </button>
                        )}
                        {canReplaceDocFile(detailModalDoc) && (
                          <button
                            type="button"
                            onClick={() => openUploadModal(detailModalDoc, 'replace')}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-white hover:bg-amber-50 border border-slate-300 text-slate-800 rounded-lg"
                          >
                            <Upload className="w-3.5 h-3.5 text-sky-700" />
                            <span>Thay tệp</span>
                          </button>
                        )}
                      </div>
                    </div>
                  )}

                  {(detailModalDoc.driveStatus === 'failed' || detailModalDoc.driveStatus === 'uploading') && (
                    <div className="space-y-2 pt-1">
                      <p className="text-rose-700 font-semibold">
                        {detailModalDoc.docCategory === 'incoming'
                          ? `Đã ghi sổ văn bản đến (${detailModalDoc.fullDocNumber}) · Lưu tệp thất bại`
                          : `Đã cấp số · Lưu tệp thất bại (${detailModalDoc.fullDocNumber})`}
                      </p>
                      {detailModalDoc.driveError && (
                        <p className="text-[11px] text-rose-600">{detailModalDoc.driveError}</p>
                      )}
                      {canRetryOrAttachDoc(detailModalDoc) && (
                        <button
                          type="button"
                          onClick={() => openUploadModal(detailModalDoc, 'retry')}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-rose-700 hover:bg-rose-800 text-white rounded-lg"
                        >
                          <RefreshCw className="w-3.5 h-3.5" />
                          <span>Thử lưu lại</span>
                        </button>
                      )}
                    </div>
                  )}

                  {(!detailModalDoc.driveStatus || detailModalDoc.driveStatus === 'none') && (
                    <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                      <span className="text-[11px] text-slate-500">Không đính kèm tệp trên Google Drive.</span>
                      {canRetryOrAttachDoc(detailModalDoc) && (
                        <button
                          type="button"
                          onClick={() => openUploadModal(detailModalDoc, 'retry')}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg"
                        >
                          <FileUp className="w-3.5 h-3.5" />
                          <span>Đính kèm tệp</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>

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
                          Lỗi đồng bộ Excel: {detailModalDoc.excelSyncError}
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
                          <span>{retryingSyncDocId === detailModalDoc.id ? 'Đang đồng bộ lại...' : 'Thử đồng bộ lại Excel'}</span>
                        </button>
                      )}
                    </div>
                  );
                })()}
              </div>

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
    </div>
  );
};
