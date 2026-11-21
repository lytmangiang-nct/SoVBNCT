import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  BookOpen,
  Check,
  CheckCircle2,
  Copy,
  Crosshair,
  ExternalLink,
  Eye,
  FileCheck2,
  FileText,
  FileUp,
  FolderCheck,
  Link as LinkIcon,
  Loader2,
  RefreshCw,
  RotateCcw,
  Send,
  Upload,
} from 'lucide-react';
import { api } from '../api';
import type {
  DocumentRecord,
  ExtractableFieldKey,
  ExtractedDateSource,
  PdfAnalyzeResponse,
  RegistryBook,
  SystemSettings,
  User,
} from '../types';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  formatFileSize,
  normalizeDateOnlyDdMmYyyy,
  sanitizeCustomDriveFileName,
} from '../utils/fileNaming';
import { DEFAULT_BOOK_ID, getUserDisplayName, hasBookPermission, isSystemAdmin } from '../utils/permissions';
import { SEMANTIC_STATUS_STYLES } from '../utils/documentColors';
import { PdfPreviewPane, type PdfHighlightTarget } from './PdfPreviewPane';

interface ImportPdfViewProps {
  currentUser: User;
  activeBook: RegistryBook;
  accessibleBooks: RegistryBook[];
  onSelectBook: (bookId: string) => void;
  settings: SystemSettings;
  serverYear?: number;
  onDocumentImported: (doc: DocumentRecord) => void;
  onDocumentUpdated: (doc: DocumentRecord) => void;
  onOpenExistingDocument: (docId: string) => void;
  onNavigateToIncomingRegistry?: () => void;
}

export const ImportPdfView: React.FC<ImportPdfViewProps> = ({
  currentUser,
  activeBook,
  accessibleBooks,
  onSelectBook: _onSelectBook,
  settings,
  serverYear,
  onDocumentImported,
  onDocumentUpdated,
  onOpenExistingDocument,
  onNavigateToIncomingRegistry,
}) => {
  const [inputMode, setInputMode] = useState<'url' | 'file'>('url');
  const [pdfUrlInput, setPdfUrlInput] = useState('');
  const [selectedPdfFile, setSelectedPdfFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const retryPdfFileInputRef = useRef<HTMLInputElement>(null);
  const restagePdfFileInputRef = useRef<HTMLInputElement>(null);

  const [analyzing, setAnalyzing] = useState(false);
  const [analyzeError, setAnalyzeError] = useState<string | null>(null);
  const [analysis, setAnalysis] = useState<PdfAnalyzeResponse | null>(null);
  const [localFilePreviewUrl, setLocalFilePreviewUrl] = useState<string | null>(null);
  const [jumpToPage, setJumpToPage] = useState<number | null>(null);
  const [highlightBox, setHighlightBox] = useState<PdfHighlightTarget | null>(null);

  useEffect(() => {
    if (analysis?.sourceType === 'pdf_upload' && selectedPdfFile) {
      const objectUrl = URL.createObjectURL(selectedPdfFile);
      setLocalFilePreviewUrl(objectUrl);
      return () => {
        URL.revokeObjectURL(objectUrl);
      };
    }
    setLocalFilePreviewUrl(null);
    return undefined;
  }, [analysis?.stageToken, analysis?.sourceType, selectedPdfFile]);

  // Mobile responsive view toggle while preserving form edits
  const [activeMobilePane, setActiveMobilePane] = useState<'form' | 'pdf'>('form');

  // Editable extracted fields
  const [fullDocNumber, setFullDocNumber] = useState('');
  const [docNumberConfirmed, setDocNumberConfirmed] = useState(true);
  const [docName, setDocName] = useState('');
  const [title, setTitle] = useState('');
  const [documentDate, setDocumentDate] = useState('');
  const [parentAgency, setParentAgency] = useState('');
  const [issuingAgency, setIssuingAgency] = useState('');
  const [signer, setSigner] = useState('');
  const [signerRole, setSignerRole] = useState('');
  const [notes, setNotes] = useState('');
  const [customFileName, setCustomFileName] = useState('');
  const [isEditingFileName, setIsEditingFileName] = useState(false);
  const [copiedFileName, setCopiedFileName] = useState(false);

  // Save & Conflict / Duplicate states
  const generateKey = () => `pdf-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const [idempotencyKey, setIdempotencyKey] = useState<string>(generateKey);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [filenameConflict, setFilenameConflict] = useState<{
    conflictingFileName: string;
    suggestedSuffixName: string;
    existingDocument: DocumentRecord;
  } | null>(null);
  const [ignoreDuplicatePdf, setIgnoreDuplicatePdf] = useState(false);
  const [ignoreSuspectedDuplicate, setIgnoreSuspectedDuplicate] = useState(false);
  const [liveSuspectedDuplicate, setLiveSuspectedDuplicate] = useState<DocumentRecord | null>(null);

  // Expired staged PDF recovery state (preserves all edits & idempotencyKey)
  const [tempPdfExpiredError, setTempPdfExpiredError] = useState<string | null>(null);
  const [restagingPdf, setRestagingPdf] = useState(false);
  const [restagePickFile, setRestagePickFile] = useState<File | null>(null);
  const [restageShaStatus, setRestageShaStatus] = useState<'idle' | 'matched' | 'mismatch'>('idle');
  const [allowDifferentRestageFile, setAllowDifferentRestageFile] = useState(false);

  // Result state
  const [savedDoc, setSavedDoc] = useState<DocumentRecord | null>(null);
  const [savedStageToken, setSavedStageToken] = useState<string | null>(null);
  const [retryingDrive, setRetryingDrive] = useState(false);
  const [retryDriveError, setRetryDriveError] = useState<string | null>(null);

  const importableBooks = useMemo(
    () => accessibleBooks.filter((b) => hasBookPermission(currentUser, b.id, 'import_incoming')),
    [accessibleBooks, currentUser]
  );

  const canImportActiveBook = hasBookPermission(currentUser, activeBook.id, 'import_incoming');
  const canViewActiveBook = hasBookPermission(currentUser, activeBook.id, 'view');
  const canManageActiveBook =
    isSystemAdmin(currentUser) || hasBookPermission(currentUser, activeBook.id, 'manage_book');

  const driveConfig =
    activeBook.id !== DEFAULT_BOOK_ID
      ? activeBook.googleDriveConfig
      : activeBook.googleDriveConfig || settings.googleDriveConfig;
  const driveConfigured = Boolean(driveConfig?.isConfigured);
  const maxFileSizeMB = settings.googleDriveConfig?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;
  const maxFileSizeBytes = settings.googleDriveConfig?.maxFileSizeBytes || maxFileSizeMB * 1024 * 1024;

  const isIncompleteDocNumber = (val: string): boolean => {
    const trimmed = val.trim();
    if (!trimmed) return true;
    if (/^[/\\-]/.test(trimmed)) return true;
    if (trimmed.includes('/')) {
      const beforeSlash = trimmed.split('/')[0].trim();
      if (!/\d/.test(beforeSlash)) return true;
    }
    return false;
  };

  const docNumberIncomplete = isIncompleteDocNumber(fullDocNumber);
  const requiresDocNumberReview = Boolean(
    analysis?.extraction.docNumberNeedsConfirmation || docNumberIncomplete
  );

  // Unified live file name ("Tên tệp sẽ lưu" — toàn bộ phần tên trước phần mở rộng viết IN HOA giữ dấu tiếng Việt)
  const computedDriveFileName = useMemo(() => {
    if (isEditingFileName && customFileName.trim()) {
      return sanitizeCustomDriveFileName(customFileName, '.pdf');
    }
    return buildStandardDriveFileName(fullDocNumber, docName || title, '.pdf');
  }, [fullDocNumber, docName, title, isEditingFileName, customFileName]);

  const computePdfSha256Hex = async (file: File): Promise<string | null> => {
    try {
      if (!window.crypto?.subtle) return null;
      const buf = await file.arrayBuffer();
      const digest = await window.crypto.subtle.digest('SHA-256', buf);
      return Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    } catch {
      return null;
    }
  };

  // Live check suspected duplicate when user edits fullDocNumber / issuingAgency / documentDate
  useEffect(() => {
    if (!analysis || savedDoc) return;
    if (!fullDocNumber.trim()) {
      setLiveSuspectedDuplicate(null);
      return;
    }

    const timer = setTimeout(async () => {
      try {
        const refDate =
          normalizeDateOnlyDdMmYyyy(documentDate) ||
          analysis.extraction.topRightDate?.date ||
          analysis.extraction.signatureDate?.date ||
          '';
        const yearMatch = refDate.match(/(\d{4})$/);
        const parsedYear = yearMatch ? parseInt(yearMatch[1], 10) : undefined;
        const res = await api.checkDuplicate({
          bookId: activeBook.id,
          fullDocNumber: fullDocNumber.trim(),
          issuingAgency: issuingAgency.trim(),
          year: parsedYear,
          sha256: analysis.sha256,
          excludeDocId: analysis.duplicateDocument?.id,
        });
        setLiveSuspectedDuplicate(res.suspectedDuplicate);
      } catch {
        // ignore background check error
      }
    }, 350);

    return () => clearTimeout(timer);
  }, [analysis, savedDoc, activeBook.id, fullDocNumber, issuingAgency, documentDate]);

  const applyAnalysisResult = (res: PdfAnalyzeResponse) => {
    setAnalysis(res);
    const extractedNum = res.extraction.docNumber || '';
    setFullDocNumber(extractedNum);
    setDocNumberConfirmed(
      !res.extraction.docNumberNeedsConfirmation && !isIncompleteDocNumber(extractedNum)
    );
    setDocName(res.extraction.docName || '');
    setTitle(res.extraction.title || '');
    const compStatus = res.extraction.dateComparisonStatus;
    const initialDate =
      compStatus === 'mismatch'
        ? ''
        : normalizeDateOnlyDdMmYyyy(
            res.extraction.documentDate ||
              res.extraction.issueDate ||
              res.extraction.signDate ||
              ''
          );
    setDocumentDate(initialDate);
    setParentAgency(res.extraction.parentAgency || '');
    setIssuingAgency(res.extraction.issuingAgency || '');
    setSigner(res.extraction.signer || '');
    setSignerRole(res.extraction.signerRole || '');
    setNotes('');
    setIsEditingFileName(false);
    setCustomFileName('');
    setFilenameConflict(null);
    setSubmitError(null);
    setTempPdfExpiredError(null);
    setRestagePickFile(null);
    setRestageShaStatus('idle');
    setAllowDifferentRestageFile(false);
    setIgnoreDuplicatePdf(false);
    setIgnoreSuspectedDuplicate(false);
    setLiveSuspectedDuplicate(res.suspectedDuplicateDocument || null);
    setSavedDoc(null);
    setSavedStageToken(null);
    setRetryDriveError(null);
    setHighlightBox(null);
    setActiveMobilePane('form');
    setJumpToPage(res.extraction.relevantPages?.[0] || 1);
    setIdempotencyKey(generateKey());
  };

  const handleLocateField = (fieldKey: ExtractableFieldKey, label: string) => {
    const loc = analysis?.extraction.fieldLocations?.[fieldKey];
    if (!loc) return;
    setJumpToPage(loc.page);
    setHighlightBox({
      page: loc.page,
      bbox: loc.bbox,
      label,
      estimated: Boolean(loc.estimated || !loc.bbox),
      nonce: Date.now(),
    });
    setActiveMobilePane('pdf');
  };

  const handleLocateDateSource = (src: ExtractedDateSource, label: string) => {
    setJumpToPage(src.page);
    setHighlightBox({
      page: src.page,
      bbox: src.bbox,
      label,
      estimated: Boolean(src.estimated || !src.bbox),
      nonce: Date.now(),
    });
    setActiveMobilePane('pdf');
  };

  const renderLocateButton = (fieldKey: ExtractableFieldKey, label: string) => {
    const loc = analysis?.extraction.fieldLocations?.[fieldKey];
    if (!loc) return null;
    const isEstimated = Boolean(loc.estimated || !loc.bbox);
    return (
      <div className="inline-flex items-center gap-1 flex-wrap">
        {isEstimated && (
          <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
            Cần kiểm tra
          </span>
        )}
        <button
          type="button"
          onClick={() => handleLocateField(fieldKey, label)}
          className="inline-flex items-center gap-1 px-2 py-1 min-h-[30px] text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded transition-colors"
          title={`Xem vị trí ${label} trên trang ${loc.page}`}
        >
          <Crosshair className="w-3 h-3 shrink-0" />
          <span>Xem vị trí (Tr.{loc.page})</span>
        </button>
      </div>
    );
  };

  const runAnalyzeUrl = async () => {
    if (!pdfUrlInput.trim() || analyzing) return;
    setAnalyzing(true);
    setAnalyzeError(null);
    try {
      const res = await api.analyzePdfUrl(pdfUrlInput.trim(), activeBook.id);
      applyAnalysisResult(res);
    } catch (err: any) {
      const msg =
        err?.code === 'PDF_READ_FAILED' || (typeof err?.status === 'number' && err.status >= 500)
          ? 'Chưa đọc được PDF. Vui lòng thử lại.'
          : err?.message || 'Chưa đọc được PDF. Vui lòng thử lại.';
      setAnalyzeError(msg);
    } finally {
      setAnalyzing(false);
    }
  };

  const handleAnalyzeUrl = async (e: React.FormEvent) => {
    e.preventDefault();
    await runAnalyzeUrl();
  };

  const handleAnalyzeFile = async (file: File) => {
    setAnalyzeError(null);
    setSelectedPdfFile(file);
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      setAnalyzeError('Chỉ hỗ trợ tệp định dạng .pdf.');
      return;
    }
    if (file.size <= 0) {
      setAnalyzeError('Tệp PDF được chọn đang bị rỗng (0 bytes).');
      return;
    }
    if (file.size > maxFileSizeBytes) {
      setAnalyzeError(
        `Dung lượng tệp (${formatFileSize(file.size)}) vượt quá giới hạn tối đa ${maxFileSizeMB} MB của hệ thống.`
      );
      return;
    }

    setAnalyzing(true);
    try {
      const res = await api.analyzePdfFile(file, activeBook.id);
      applyAnalysisResult(res);
    } catch (err: any) {
      const msg =
        err?.code === 'PDF_READ_FAILED' || (typeof err?.status === 'number' && err.status >= 500)
          ? 'Chưa đọc được PDF. Vui lòng thử lại.'
          : err?.message || 'Chưa đọc được PDF. Vui lòng thử lại.';
      setAnalyzeError(msg);
    } finally {
      setAnalyzing(false);
    }
  };

  const handleCopyFileName = async () => {
    if (!computedDriveFileName) return;
    try {
      await navigator.clipboard.writeText(computedDriveFileName);
      setCopiedFileName(true);
      setTimeout(() => setCopiedFileName(false), 2000);
    } catch {
      // ignore
    }
  };

  const submitImport = async (options?: {
    conflictAction?: 'suffix' | 'overwrite';
    overrideCustomName?: string;
    forceIgnoreDuplicate?: boolean;
    forceIgnoreSuspected?: boolean;
    overrideStageToken?: string;
    overrideExpectedSha256?: string;
    allowDifferentFile?: boolean;
  }) => {
    if (!analysis || submitting) return;
    setSubmitError(null);
    setFilenameConflict(null);

    if (!fullDocNumber.trim() || isIncompleteDocNumber(fullDocNumber)) {
      setSubmitError(
        'Số/ký hiệu văn bản chưa đầy đủ (thiếu phần số ở đầu). Vui lòng đối chiếu bản PDF gốc và nhập đầy đủ trước khi lưu.'
      );
      return;
    }
    if (requiresDocNumberReview && !docNumberConfirmed) {
      setSubmitError(
        'Vui lòng kiểm tra số văn bản với bản PDF gốc và đánh dấu xác nhận ở khung "Cần kiểm tra số văn bản" trước khi lưu.'
      );
      return;
    }
    if (!docName.trim() && !title.trim()) {
      setSubmitError('Vui lòng nhập Tên văn bản hoặc Trích yếu nội dung văn bản.');
      return;
    }

    const compStatus = analysis.extraction.dateComparisonStatus;
    const isMismatch = compStatus === 'mismatch';
    const rawDateTrimmed = documentDate.trim();
    const cleanDocDate = normalizeDateOnlyDdMmYyyy(rawDateTrimmed);
    if (rawDateTrimmed && !cleanDocDate) {
      setSubmitError('Ngày văn bản không hợp lệ (định dạng chuẩn dd/mm/yyyy).');
      return;
    }

    const effectiveUnifiedDate = isMismatch
      ? canManageActiveBook
        ? cleanDocDate
        : ''
      : cleanDocDate;

    const mismatchSourcesPayload = isMismatch
      ? {
          topRight: analysis.extraction.topRightDate || undefined,
          signature: analysis.extraction.signatureDate || undefined,
        }
      : undefined;

    const stageTokenToUse = options?.overrideStageToken || analysis.stageToken;
    const expectedShaToUse = options?.overrideExpectedSha256 || analysis.sha256;

    setSubmitting(true);
    try {
      const res = await api.importPdfDocument({
        bookId: activeBook.id,
        stageToken: stageTokenToUse,
        expectedSha256: expectedShaToUse,
        allowDifferentFile: Boolean(options?.allowDifferentFile || allowDifferentRestageFile),
        fullDocNumber: fullDocNumber.trim(),
        docName: docName.trim(),
        title: title.trim(),
        documentDate: effectiveUnifiedDate,
        issueDate: effectiveUnifiedDate,
        signDate: effectiveUnifiedDate,
        dateComparisonStatus: compStatus,
        dateMismatchSources: mismatchSourcesPayload,
        confirmUnifiedDate: Boolean(isMismatch && canManageActiveBook && effectiveUnifiedDate),
        parentAgency: parentAgency.trim() || undefined,
        issuingAgency: issuingAgency.trim(),
        signer: signer.trim(),
        signerRole: signerRole.trim(),
        notes: notes.trim(),
        expectedFileName: computedDriveFileName,
        conflictAction: options?.conflictAction,
        customFileName:
          options?.overrideCustomName || (isEditingFileName ? computedDriveFileName : undefined),
        ignoreDuplicatePdf: options?.forceIgnoreDuplicate ?? ignoreDuplicatePdf,
        ignoreSuspectedDuplicate: options?.forceIgnoreSuspected ?? ignoreSuspectedDuplicate,
        needsReview: analysis.extraction.needsReview,
        reviewReasons: analysis.extraction.reviewReasons,
        relevantPages: analysis.extraction.relevantPages,
        idempotencyKey,
      });

      setTempPdfExpiredError(null);
      setRestagePickFile(null);
      setRestageShaStatus('idle');
      setAllowDifferentRestageFile(false);
      setSavedDoc(res.document);
      setSavedStageToken(res.stageToken || stageTokenToUse);
      onDocumentImported(res.document);
      if (!res.driveSaved && res.uploadError) {
        setRetryDriveError(res.uploadError);
      }
    } catch (err: any) {
      if (err.code === 'TEMP_PDF_EXPIRED' || err.status === 410) {
        // Giữ nguyên toàn bộ thông tin đã sửa và idempotencyKey để nạp lại tệp không tạo trùng hồ sơ
        setTempPdfExpiredError(
          err.message ||
            'Bản PDF tạm trên máy chủ đã hết hạn lưu trữ. Toàn bộ thông tin bạn đã chỉnh sửa vẫn được giữ nguyên.'
        );
        setSubmitError(null);
      } else if (err.code === 'FILENAME_CONFLICT' && err.existingDocument) {
        setFilenameConflict({
          conflictingFileName: err.conflictingFileName || computedDriveFileName,
          suggestedSuffixName: err.suggestedSuffixName || computedDriveFileName,
          existingDocument: err.existingDocument,
        });
      } else if (err.code === 'PDF_DUPLICATE' && err.existingDocument) {
        setSubmitError(err.message);
      } else if (err.code === 'PDF_SUSPECTED_DUPLICATE' && err.existingDocument) {
        setLiveSuspectedDuplicate(err.existingDocument);
        setSubmitError(err.message);
      } else {
        setSubmitError(err.message || 'Lỗi khi lưu văn bản đến.');
      }
    } finally {
      setSubmitting(false);
    }
  };

  const handleRestageAndRetryImport = async (fileOverride?: File, forceAllowDifferent = false) => {
    if (!analysis || restagingPdf || submitting) return;
    const fileToUse = fileOverride || restagePickFile || selectedPdfFile;
    if (!fileToUse) {
      restagePdfFileInputRef.current?.click();
      return;
    }

    if (!fileToUse.name.toLowerCase().endsWith('.pdf')) {
      setSubmitError('Chỉ hỗ trợ tệp định dạng .pdf.');
      return;
    }
    if (fileToUse.size <= 0) {
      setSubmitError('Tệp PDF được chọn đang bị rỗng (0 bytes).');
      return;
    }
    if (fileToUse.size > maxFileSizeBytes) {
      setSubmitError(
        `Dung lượng tệp (${formatFileSize(fileToUse.size)}) vượt quá giới hạn tối đa ${maxFileSizeMB} MB.`
      );
      return;
    }

    const allowDiff = Boolean(forceAllowDifferent || allowDifferentRestageFile);
    setRestagingPdf(true);
    setSubmitError(null);
    try {
      const restaged = await api.restagePdfFile(fileToUse, {
        bookId: activeBook.id,
        expectedSha256: analysis.sha256,
        allowDifferentFile: allowDiff,
      });

      setSelectedPdfFile(fileToUse);
      setAnalysis((prev) =>
        prev
          ? {
              ...prev,
              stageToken: restaged.stageToken,
              sha256: restaged.sha256,
              fileSize: restaged.fileSize,
              originalFileName: restaged.originalFileName,
            }
          : prev
      );
      setRestageShaStatus(restaged.matchedOriginal ? 'matched' : 'mismatch');
      setRestagingPdf(false);

      await submitImport({
        overrideStageToken: restaged.stageToken,
        overrideExpectedSha256: restaged.sha256,
        allowDifferentFile: allowDiff,
      });
    } catch (err: any) {
      setRestagingPdf(false);
      if (err?.code === 'PDF_FILE_MISMATCH' || err?.status === 409) {
        setRestagePickFile(fileToUse);
        setRestageShaStatus('mismatch');
        setAllowDifferentRestageFile(false);
        setTempPdfExpiredError(
          err?.message ||
            'Tệp PDF vừa chọn khác với tệp đã phân tích ban đầu. Vui lòng chọn đúng tệp gốc hoặc xác nhận lưu bằng tệp mới này.'
        );
      } else {
        setSubmitError(err?.message || 'Không thể nạp lại tệp PDF để lưu. Vui lòng thử lại.');
      }
    }
  };

  const handleRestageFilePicked = async (file: File) => {
    if (!analysis) return;
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      setSubmitError('Chỉ hỗ trợ tệp định dạng .pdf.');
      return;
    }
    setSubmitError(null);
    setRestagePickFile(file);
    setAllowDifferentRestageFile(false);

    const sha = await computePdfSha256Hex(file);
    if (sha && analysis.sha256 && sha.toLowerCase() !== analysis.sha256.toLowerCase()) {
      setRestageShaStatus('mismatch');
      return;
    }

    setRestageShaStatus('matched');
    await handleRestageAndRetryImport(file, false);
  };

  const handleRetryDriveSave = async (overrideFile?: File) => {
    if (!savedDoc || retryingDrive) return;
    setRetryingDrive(true);
    setRetryDriveError(null);
    try {
      const res = await api.uploadPdfToDocument(savedDoc.id, {
        file: overrideFile,
        stageToken: !overrideFile ? savedStageToken || undefined : undefined,
        fileName: savedDoc.driveFileName,
        mode: 'retry',
      });
      setSavedDoc(res.document);
      onDocumentUpdated(res.document);
    } catch (err: any) {
      if (err.document) {
        setSavedDoc(err.document);
        onDocumentUpdated(err.document);
      }
      setRetryDriveError(err.message || 'Lưu tệp thất bại.');
    } finally {
      setRetryingDrive(false);
    }
  };

  const handleOpenSavedDriveFile = async (doc: DocumentRecord) => {
    try {
      const res = await api.openDocumentFile(doc.id);
      if (res.url) window.open(res.url, '_blank', 'noopener,noreferrer');
    } catch {
      if (doc.driveFileUrl) window.open(doc.driveFileUrl, '_blank', 'noopener,noreferrer');
    }
  };

  const handleReset = () => {
    setAnalysis(null);
    setAnalyzeError(null);
    setSelectedPdfFile(null);
    setPdfUrlInput('');
    setSavedDoc(null);
    setSavedStageToken(null);
    setFilenameConflict(null);
    setSubmitError(null);
    setTempPdfExpiredError(null);
    setRestagePickFile(null);
    setRestageShaStatus('idle');
    setAllowDifferentRestageFile(false);
    setDocNumberConfirmed(true);
    setRetryDriveError(null);
    setHighlightBox(null);
    setLiveSuspectedDuplicate(null);
    setIgnoreDuplicatePdf(false);
    setIgnoreSuspectedDuplicate(false);
    setParentAgency('');
    setDocumentDate('');
    setActiveMobilePane('form');
    setIdempotencyKey(generateKey());
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const effectiveSuspectedDup =
    liveSuspectedDuplicate ||
    (analysis?.suspectedDuplicateDocument &&
    analysis.suspectedDuplicateDocument.id !== analysis.duplicateDocument?.id
      ? analysis.suspectedDuplicateDocument
      : null);

  return (
    <div className="max-w-7xl mx-auto py-5 sm:py-8 px-3 sm:px-6 lg:px-8 space-y-5 sm:space-y-6">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <div className="flex flex-wrap items-center gap-2">
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
              Nhập Văn Bản PDF
            </h1>
            <span className={`inline-flex items-center gap-1.5 text-xs font-semibold px-2.5 py-0.5 rounded-md border ${SEMANTIC_STATUS_STYLES.incoming.badge}`}>
              <span className={`w-2 h-2 rounded-full ${SEMANTIC_STATUS_STYLES.incoming.dot}`} />
              <span>Văn bản đến</span>
            </span>
          </div>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            Tiếp nhận công văn, kế hoạch PDF từ Sở GD&amp;ĐT hoặc nguồn khác. Đọc và đối chiếu bản gốc bên cạnh, đặt tên chuẩn và lưu vào Google Drive mà{' '}
            <strong className="text-purple-800">không làm tăng dãy số văn bản đi</strong>.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0 w-full md:w-auto">
          <span className="inline-flex items-center gap-1.5 px-3 py-2 min-h-[42px] bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700">
            <BookOpen className="w-3.5 h-3.5 text-purple-700 shrink-0" />
            <span>{activeBook.name} ({activeBook.unitCode})</span>
          </span>

          {onNavigateToIncomingRegistry && canViewActiveBook && (
            <button
              type="button"
              onClick={onNavigateToIncomingRegistry}
              className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 min-h-[42px] text-xs font-semibold bg-purple-50 hover:bg-purple-100 border border-purple-200 text-purple-900 rounded-lg shadow-2xs transition-colors"
            >
              <Eye className="w-3.5 h-3.5 shrink-0" />
              <span>Tra cứu văn bản đến</span>
            </button>
          )}

          {(analysis || savedDoc) && (
            <button
              type="button"
              onClick={handleReset}
              className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 min-h-[42px] text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 rounded-lg shadow-2xs w-full sm:w-auto"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Nhập văn bản PDF khác</span>
            </button>
          )}
        </div>
      </div>

      {!canImportActiveBook && (
        <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl text-xs text-amber-900 flex items-center gap-2.5">
          <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
          <span>
            Tài khoản của bạn chưa được cấp quyền <strong>Nhập văn bản đến</strong> trong{' '}
            <strong>{activeBook.name}</strong>. Vui lòng chọn sổ khác hoặc liên hệ Quản trị viên.
          </span>
        </div>
      )}

      {/* Saved Result Banner (Tách biệt Trạng thái hồ sơ vs Trạng thái tệp) */}
      {savedDoc && (
        <div
          className={`p-4 sm:p-5 rounded-xl border shadow-xs space-y-4 ${
            savedDoc.driveStatus === 'saved'
              ? 'bg-emerald-50 border-emerald-200'
              : 'bg-amber-50 border-amber-200'
          }`}
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-start gap-3">
              <div
                className={`w-10 h-10 rounded-full flex items-center justify-center shrink-0 text-white ${
                  savedDoc.driveStatus === 'saved' ? 'bg-emerald-600' : 'bg-amber-600'
                }`}
              >
                {savedDoc.driveStatus === 'saved' ? (
                  <CheckCircle2 className="w-5 h-5" />
                ) : (
                  <AlertTriangle className="w-5 h-5" />
                )}
              </div>
              <div className="space-y-1">
                <div className="flex flex-wrap items-center gap-2">
                  <span className={`text-xs font-bold px-2.5 py-0.5 rounded border ${SEMANTIC_STATUS_STYLES.incoming.badge}`}>
                    Đã ghi sổ văn bản đến
                  </span>
                  <span
                    className={`text-xs font-bold px-2.5 py-0.5 rounded border ${
                      savedDoc.driveStatus === 'saved'
                        ? SEMANTIC_STATUS_STYLES.saved.badge
                        : SEMANTIC_STATUS_STYLES.danger.badge
                    }`}
                  >
                    {savedDoc.driveStatus === 'saved' ? 'Đã lưu lên Drive' : 'Lưu tệp thất bại'}
                  </span>
                  <span className="text-[11px] font-medium text-slate-600">
                    {savedDoc.bookName || activeBook.name}
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <span className="text-xl font-extrabold font-mono text-slate-900">
                    {savedDoc.fullDocNumber}
                  </span>
                  <span className="text-xs font-semibold px-2 py-0.5 rounded bg-white border border-slate-200 text-slate-700">
                    {savedDoc.issuingAgency || 'Cơ quan ngoài'}
                  </span>
                </div>
                <p className="text-xs text-slate-700 font-medium">
                  {savedDoc.docName || savedDoc.title}
                </p>
                <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-600">
                  <span>
                    Ngày văn bản:{' '}
                    {savedDoc.pendingDateConfirmation ? (
                      <strong className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 font-semibold">
                        Chờ quản trị viên xác nhận ngày
                      </strong>
                    ) : (
                      <strong className="font-mono text-slate-800 tabular-nums">
                        {savedDoc.issueDate || 'Chưa xác định'}
                      </strong>
                    )}
                  </span>
                  <span>·</span>
                  <span>
                    Ngày tiếp nhận: <strong className="font-mono text-slate-800 tabular-nums">{savedDoc.receivedDate || 'Hôm nay'}</strong>
                  </span>
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2">
              {savedDoc.driveStatus === 'saved' && savedDoc.driveFileUrl && (
                <button
                  type="button"
                  onClick={() => handleOpenSavedDriveFile(savedDoc)}
                  className="inline-flex items-center justify-center gap-1.5 px-4 py-2 min-h-[42px] text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg shadow-xs"
                >
                  <ExternalLink className="w-4 h-4" />
                  <span>Mở văn bản gốc</span>
                </button>
              )}
              {canViewActiveBook && (
                <button
                  type="button"
                  onClick={() => onOpenExistingDocument(savedDoc.id)}
                  className="inline-flex items-center justify-center gap-1.5 px-4 py-2 min-h-[42px] text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 rounded-lg"
                >
                  <FileCheck2 className="w-4 h-4 text-purple-700" />
                  <span>Xem hồ sơ vừa nhập</span>
                </button>
              )}
            </div>
          </div>

          {savedDoc.driveStatus === 'saved' ? (
            <div className="bg-white/90 border border-emerald-200 rounded-lg px-3.5 py-2.5 text-xs flex items-center justify-between gap-2">
              <div className="min-w-0">
                <span className="text-[11px] text-emerald-700 font-semibold block">
                  Đã lưu lên Drive với tên:
                </span>
                <span className="font-mono font-bold text-slate-900 break-all">
                  {savedDoc.driveFileName}
                </span>
              </div>
              <FolderCheck className="w-5 h-5 text-emerald-600 shrink-0" />
            </div>
          ) : (
            <div className="bg-white border border-rose-200 rounded-lg p-3.5 space-y-2.5 text-xs">
              <div className="flex items-start gap-2 text-rose-900">
                <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                <div className="space-y-0.5">
                  <strong className="font-bold block">
                    Đã ghi sổ văn bản đến ({savedDoc.fullDocNumber}) · Lưu tệp thất bại
                  </strong>
                  <p className="text-rose-700">
                    {retryDriveError ||
                      savedDoc.driveError ||
                      'Không thể tải bản PDF lên thư mục Google Drive.'}
                  </p>
                  <p className="text-[11px] text-slate-600">
                    Tên tệp dự kiến: <span className="font-mono font-bold">{savedDoc.driveFileName}</span>
                  </p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <input
                  ref={retryPdfFileInputRef}
                  type="file"
                  accept=".pdf,application/pdf"
                  className="hidden"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) handleRetryDriveSave(f);
                  }}
                />
                <button
                  type="button"
                  disabled={retryingDrive}
                  onClick={() => handleRetryDriveSave()}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white rounded-md"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${retryingDrive ? 'animate-spin' : ''}`} />
                  <span>{retryingDrive ? 'Đang thử lưu lại...' : 'Thử lưu lại'}</span>
                </button>
                <button
                  type="button"
                  disabled={retryingDrive}
                  onClick={() => retryPdfFileInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-white hover:bg-slate-50 border border-rose-300 text-rose-800 rounded-md"
                >
                  <FileUp className="w-3.5 h-3.5" />
                  <span>Chọn lại tệp PDF &amp; Thử lưu lại</span>
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Input Selector Panel */}
      {!savedDoc && canImportActiveBook && (
        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2.5 sm:gap-3">
            <div className="grid grid-cols-2 sm:inline-flex w-full sm:w-auto bg-slate-100 p-1 rounded-lg border border-slate-200">
              <button
                type="button"
                onClick={() => {
                  setInputMode('url');
                  setAnalyzeError(null);
                }}
                className={`flex items-center justify-center gap-1.5 px-3 py-2 min-h-[42px] text-xs font-semibold rounded-md transition-colors ${
                  inputMode === 'url'
                    ? 'bg-white text-sky-800 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <LinkIcon className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Dán liên kết</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setInputMode('file');
                  setAnalyzeError(null);
                }}
                className={`flex items-center justify-center gap-1.5 px-3 py-2 min-h-[42px] text-xs font-semibold rounded-md transition-colors ${
                  inputMode === 'file'
                    ? 'bg-white text-sky-800 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Upload className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Tải PDF từ máy</span>
              </button>
            </div>

            <span className="text-[11px] font-medium text-slate-600 bg-slate-50 px-2.5 py-1 rounded border border-slate-200 max-w-full break-words">
              Sổ nhập: <strong>{activeBook.name}</strong> · Định dạng .pdf · Tối đa {maxFileSizeMB} MB
            </span>
          </div>

          {inputMode === 'url' ? (
            <form onSubmit={handleAnalyzeUrl} className="flex flex-col sm:flex-row gap-2.5">
              <input
                type="url"
                value={pdfUrlInput}
                onChange={(e) => setPdfUrlInput(e.target.value)}
                placeholder="Dán liên kết tệp PDF trực tiếp hoặc link Google Drive (https://...)"
                required
                disabled={analyzing}
                className="flex-1 min-w-0 min-h-[44px] h-11 px-3.5 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
              />
              <button
                type="submit"
                disabled={analyzing || !pdfUrlInput.trim()}
                className="min-h-[44px] h-11 px-5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg shadow-xs inline-flex items-center justify-center gap-2 shrink-0 w-full sm:w-auto"
              >
                {analyzing ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Đang tải &amp; đọc PDF...</span>
                  </>
                ) : (
                  <>
                    <FileText className="w-4 h-4" />
                    <span>Tải và đọc PDF</span>
                  </>
                )}
              </button>
            </form>
          ) : (
            <div className="space-y-2">
              <input
                ref={fileInputRef}
                type="file"
                accept=".pdf,application/pdf"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  e.target.value = '';
                  if (f) handleAnalyzeFile(f);
                }}
              />
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 p-3.5 sm:p-4 border-2 border-dashed border-slate-300 hover:border-sky-400 rounded-xl bg-slate-50/70">
                <div className="space-y-0.5 text-xs min-w-0">
                  <p className="font-bold text-slate-800 break-words">
                    {selectedPdfFile
                      ? `${selectedPdfFile.name} (${formatFileSize(selectedPdfFile.size)})`
                      : 'Chọn tệp PDF văn bản từ máy tính hoặc điện thoại'}
                  </p>
                  <p className="text-[11px] text-slate-500">
                    Hệ thống tự động trích xuất văn bản, nhận dạng tiếng Việt nếu là bản scan và hiển thị PDF gốc bên cạnh để đối chiếu.
                  </p>
                </div>
                <button
                  type="button"
                  disabled={analyzing}
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 min-h-[44px] text-xs font-semibold bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white rounded-lg shadow-xs shrink-0 w-full sm:w-auto"
                >
                  {analyzing ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Đang phân tích PDF...</span>
                    </>
                  ) : (
                    <>
                      <FileUp className="w-4 h-4" />
                      <span>{selectedPdfFile ? 'Chọn tệp PDF khác' : 'Chọn tệp PDF'}</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          )}

          {analyzeError && (
            <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-900 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-start gap-2">
                <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                <span>{analyzeError}</span>
              </div>
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {inputMode === 'file' && selectedPdfFile && (
                  <button
                    type="button"
                    disabled={analyzing}
                    onClick={() => handleAnalyzeFile(selectedPdfFile)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-md"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Thử lại</span>
                  </button>
                )}
                {inputMode === 'url' && pdfUrlInput.trim() && (
                  <button
                    type="button"
                    disabled={analyzing}
                    onClick={() => runAnalyzeUrl()}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-md"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    <span>Thử lại</span>
                  </button>
                )}
                {inputMode === 'url' && (
                  <button
                    type="button"
                    onClick={() => {
                      setInputMode('file');
                      setAnalyzeError(null);
                    }}
                    className="px-3 py-1.5 bg-white hover:bg-rose-100 border border-rose-300 text-rose-800 font-semibold rounded-md"
                  >
                    Chuyển sang Tải PDF từ máy
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Side-by-Side PDF Preview + Metadata Review Form */}
      {analysis && !savedDoc && (
        <div className="space-y-4">
          {/* 1. Duplicate PDF SHA-256 Warning */}
          {analysis.duplicateDocument && (
            <div className="p-4 bg-amber-50 border border-amber-300 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-amber-950">
              <div className="flex items-start gap-2.5">
                <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <strong className="text-sm font-bold block">
                    Tệp PDF trùng hoàn toàn (SHA-256) với hồ sơ đã có: {analysis.duplicateDocument.fullDocNumber}
                  </strong>
                  <p className="text-amber-900">
                    Văn bản đã lưu: <strong>{analysis.duplicateDocument.docName || analysis.duplicateDocument.title}</strong>
                    {analysis.duplicateDocument.issuingAgency ? ` · ${analysis.duplicateDocument.issuingAgency}` : ''}
                    {analysis.duplicateDocument.issueDate ? ` · Ngày BH: ${analysis.duplicateDocument.issueDate}` : ''}.
                  </p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {canViewActiveBook ? (
                  <button
                    type="button"
                    onClick={() => onOpenExistingDocument(analysis.duplicateDocument!.id)}
                    className="px-3.5 py-2 bg-amber-700 hover:bg-amber-800 text-white font-semibold rounded-lg shadow-2xs"
                  >
                    Mở hồ sơ đã có
                  </button>
                ) : (
                  <span className="text-[11px] text-amber-800 italic">Cần quyền Xem để mở hồ sơ</span>
                )}
                <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white border border-amber-300 rounded-lg cursor-pointer font-medium">
                  <input
                    type="checkbox"
                    checked={ignoreDuplicatePdf}
                    onChange={(e) => setIgnoreDuplicatePdf(e.target.checked)}
                    className="rounded border-amber-400 text-amber-700"
                  />
                  <span>Vẫn nhập mới</span>
                </label>
              </div>
            </div>
          )}

          {/* 2. Suspected Duplicate Warning (Agency + DocNumber + Year) */}
          {effectiveSuspectedDup && (
            <div className="p-4 bg-orange-50 border border-orange-300 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-orange-950">
              <div className="flex items-start gap-2.5">
                <AlertTriangle className="w-5 h-5 text-orange-600 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <strong className="text-sm font-bold block">
                    Nghi trùng văn bản cùng Số/ký hiệu &amp; Cơ quan ban hành ({effectiveSuspectedDup.fullDocNumber})
                  </strong>
                  <p className="text-orange-900">
                    Hồ sơ đã có trong <strong>{activeBook.name}</strong>:{' '}
                    <strong>{effectiveSuspectedDup.docName || effectiveSuspectedDup.title}</strong> ·{' '}
                    {effectiveSuspectedDup.issuingAgency || 'Cùng cơ quan'} · Năm {effectiveSuspectedDup.year}.
                  </p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {canViewActiveBook ? (
                  <button
                    type="button"
                    onClick={() => onOpenExistingDocument(effectiveSuspectedDup.id)}
                    className="px-3.5 py-2 bg-orange-700 hover:bg-orange-800 text-white font-semibold rounded-lg shadow-2xs"
                  >
                    Mở hồ sơ đã có
                  </button>
                ) : (
                  <span className="text-[11px] text-orange-800 italic">Cần quyền Xem để mở hồ sơ</span>
                )}
                <label className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white border border-orange-300 rounded-lg cursor-pointer font-medium">
                  <input
                    type="checkbox"
                    checked={ignoreSuspectedDuplicate}
                    onChange={(e) => setIgnoreSuspectedDuplicate(e.target.checked)}
                    className="rounded border-orange-400 text-orange-700"
                  />
                  <span>Xác nhận nhập mới</span>
                </label>
              </div>
            </div>
          )}

          {/* Cần kiểm tra / OCR Confidence Notice */}
          {analysis.extraction.needsReview && (
            <div className="p-4 bg-amber-50/90 border border-amber-200 rounded-xl space-y-2 text-xs text-amber-950">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 rounded bg-amber-600 text-white font-bold uppercase text-[11px]">
                    Cần kiểm tra
                  </span>
                  <span className="font-bold">
                    Bấm "Xem vị trí" ở từng trường để đối chiếu trực tiếp với bản PDF gốc
                  </span>
                </div>
                {analysis.extraction.usedOcr && (
                  <span className="text-[11px] font-mono bg-white px-2 py-0.5 rounded border border-amber-200">
                    OCR tiếng Việt
                    {analysis.extraction.ocrConfidence !== undefined
                      ? ` · Độ tin cậy ~${analysis.extraction.ocrConfidence}%`
                      : ''}
                  </span>
                )}
              </div>
              <ul className="list-disc list-inside space-y-0.5 text-amber-900">
                {analysis.extraction.reviewReasons.map((reason, idx) => (
                  <li key={idx}>{reason}</li>
                ))}
              </ul>
            </div>
          )}

          {/* Mobile Screen Switcher: Thông tin văn bản / Xem PDF (giữ nguyên mọi nội dung đang sửa) */}
          <div className="flex lg:hidden items-center justify-between gap-1.5 bg-white border border-slate-200 rounded-xl p-1.5 shadow-2xs">
            <button
              type="button"
              onClick={() => setActiveMobilePane('form')}
              className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 px-2 min-h-[44px] text-xs font-bold rounded-lg transition-colors ${
                activeMobilePane === 'form'
                  ? 'bg-sky-700 text-white shadow-2xs'
                  : 'text-slate-600 hover:bg-slate-50'
              }`}
            >
              <FileText className="w-3.5 h-3.5 shrink-0" />
              <span>Thông tin văn bản</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveMobilePane('pdf')}
              className={`flex-1 flex items-center justify-center gap-1.5 py-2.5 px-2 min-h-[44px] text-xs font-bold rounded-lg transition-colors ${
                activeMobilePane === 'pdf'
                  ? 'bg-sky-700 text-white shadow-2xs'
                  : 'text-slate-600 hover:bg-slate-50'
              }`}
            >
              <Eye className="w-3.5 h-3.5 shrink-0" />
              <span>Xem PDF</span>
            </button>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 lg:gap-6 items-start">
            {/* Left: Side-by-side PDF Viewer (always mounted to preserve page/zoom) */}
            <div className={`lg:col-span-6 min-w-0 ${activeMobilePane === 'pdf' ? 'block' : 'hidden lg:block'}`}>
              {activeMobilePane === 'pdf' && (
                <div className="lg:hidden mb-2.5 flex flex-wrap items-center justify-between gap-2 bg-sky-50 border border-sky-200 rounded-lg px-3 py-2 text-xs">
                  <span className="text-sky-950 font-semibold min-w-0 truncate">
                    {highlightBox?.label
                      ? `Đang đối chiếu: ${highlightBox.label} (Tr.${highlightBox.page})`
                      : 'Đang xem bản PDF gốc'}
                  </span>
                  <button
                    type="button"
                    onClick={() => setActiveMobilePane('form')}
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 min-h-[34px] bg-white hover:bg-sky-100 border border-sky-300 text-sky-800 font-bold rounded-md shrink-0"
                  >
                    <FileText className="w-3.5 h-3.5" />
                    <span>Quay lại Thông tin văn bản</span>
                  </button>
                </div>
              )}
              <PdfPreviewPane
                pdfUrl={localFilePreviewUrl || api.getStagedPdfUrl(analysis.stageToken)}
                fileName={analysis.originalFileName}
                relevantPages={analysis.extraction.relevantPages}
                jumpToPage={jumpToPage}
                highlightBox={highlightBox}
              />
            </div>

            {/* Right: Editable Metadata Form (always mounted to preserve edits when switching on mobile) */}
            <form
              onSubmit={(e) => {
                e.preventDefault();
                submitImport();
              }}
              className={`lg:col-span-6 min-w-0 bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4 ${
                activeMobilePane === 'form' ? 'block' : 'hidden lg:block'
              }`}
            >
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-3">
                <div className="min-w-0">
                  <h2 className="text-base font-bold text-slate-900">
                    Thông tin văn bản đến
                  </h2>
                  <p className="text-[11px] text-slate-500">
                    Kiểm tra và chỉnh sửa nếu cần. Thay đổi Số/ký hiệu hoặc Tên văn bản sẽ cập nhật ngay tên tệp lưu trên Drive.
                  </p>
                </div>
                <span className={`inline-flex items-center gap-1.5 text-[11px] font-bold px-2.5 py-1 rounded border shrink-0 ${SEMANTIC_STATUS_STYLES.incoming.badge}`}>
                  <BookOpen className="w-3.5 h-3.5 shrink-0" />
                  <span>{activeBook.name} ({activeBook.unitCode})</span>
                </span>
              </div>

              <div className="space-y-1.5 min-w-0">
                <div className="flex items-center justify-between gap-1 flex-wrap">
                  <div className="flex items-center gap-1.5 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Số / Ký hiệu gốc <span className="text-rose-600">*</span>
                    </label>
                    {requiresDocNumberReview && (
                      <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
                        Cần kiểm tra số văn bản
                      </span>
                    )}
                  </div>
                  {renderLocateButton('docNumber', 'Số/ký hiệu')}
                </div>
                <input
                  type="text"
                  value={fullDocNumber}
                  onChange={(e) => {
                    const nextVal = e.target.value;
                    setFullDocNumber(nextVal);
                    setIsEditingFileName(false);
                    if (isIncompleteDocNumber(nextVal)) {
                      setDocNumberConfirmed(false);
                    }
                  }}
                  placeholder="Ví dụ: 1234/SGDĐT-GDTrH"
                  required
                  className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg font-mono font-bold text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 ${
                    requiresDocNumberReview && !docNumberConfirmed
                      ? 'border-amber-400 bg-amber-50/40'
                      : 'border-slate-300'
                  }`}
                />
                {requiresDocNumberReview && (
                  <div
                    data-testid="doc-number-review-alert"
                    className="p-3 bg-amber-50 border border-amber-300 rounded-lg space-y-2 text-xs text-amber-950"
                  >
                    <div className="flex items-start gap-2">
                      <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                      <div className="space-y-1 min-w-0">
                        <strong className="font-bold block">Cần kiểm tra số văn bản</strong>
                        <p className="text-amber-900">
                          {docNumberIncomplete
                            ? 'Số/ký hiệu đang thiếu phần số ở đầu (ví dụ dạng /QĐ-SGDĐT). Vui lòng xem góc trên bên trái trang 1 bản PDF gốc và điền đầy đủ.'
                            : analysis.extraction.docNumberReviewReason ||
                              'Vui lòng đối chiếu số/ký hiệu vừa trích xuất với góc trên bên trái trang 1 bản PDF gốc trước khi lưu.'}
                        </p>
                        {analysis.extraction.filenameDocNumberHint && (
                          <p className="text-[11px] text-amber-900 bg-white/80 border border-amber-200 rounded px-2 py-1">
                            Gợi ý đối chiếu từ tên tệp gốc (chỉ tham khảo, không tự điền):{' '}
                            <strong className="font-mono">{analysis.extraction.filenameDocNumberHint}</strong>
                          </p>
                        )}
                      </div>
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                      <label
                        className={`inline-flex items-center gap-2 px-2.5 py-1.5 bg-white border border-amber-300 rounded-md font-semibold text-amber-950 ${
                          docNumberIncomplete ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'
                        }`}
                      >
                        <input
                          type="checkbox"
                          checked={docNumberConfirmed && !docNumberIncomplete}
                          disabled={docNumberIncomplete}
                          onChange={(e) => setDocNumberConfirmed(e.target.checked)}
                          className="rounded border-amber-400 text-amber-700"
                        />
                        <span>Tôi đã đối chiếu số/ký hiệu với bản PDF gốc</span>
                      </label>
                      {analysis.extraction.fieldLocations?.docNumber && (
                        <button
                          type="button"
                          onClick={() => handleLocateField('docNumber', 'Số/ký hiệu')}
                          className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white hover:bg-sky-50 border border-sky-300 text-sky-800 font-semibold rounded-md transition-colors"
                        >
                          <Crosshair className="w-3.5 h-3.5 shrink-0" />
                          <span>Xem vị trí số văn bản (Tr.{analysis.extraction.fieldLocations.docNumber.page})</span>
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Cơ quan ban hành
                    </label>
                    {renderLocateButton('issuingAgency', 'Cơ quan ban hành')}
                  </div>
                  <input
                    type="text"
                    value={issuingAgency}
                    onChange={(e) => setIssuingAgency(e.target.value)}
                    placeholder="Ví dụ: SỞ GIÁO DỤC VÀ ĐÀO TẠO..."
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Cơ quan chủ quản <span className="text-slate-400 font-normal lowercase">(nếu có)</span>
                    </label>
                    {renderLocateButton('parentAgency', 'Cơ quan chủ quản')}
                  </div>
                  <input
                    type="text"
                    value={parentAgency}
                    onChange={(e) => setParentAgency(e.target.value)}
                    placeholder="Ví dụ: UBND TỈNH AN GIANG..."
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>
              </div>

              <div className="space-y-1 min-w-0">
                <div className="flex items-center justify-between gap-1 flex-wrap">
                  <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                    Tên văn bản <span className="text-rose-600">*</span>
                  </label>
                  {renderLocateButton('docName', 'Tên văn bản')}
                </div>
                <input
                  type="text"
                  value={docName}
                  onChange={(e) => {
                    setDocName(e.target.value);
                    setIsEditingFileName(false);
                  }}
                  placeholder="Tên văn bản ngắn gọn (tự động cập nhật tên tệp lưu trên Drive)..."
                  className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
                <p className="text-[11px] text-slate-500">
                  Tên rút gọn tạo từ trích yếu gốc để đặt tên tệp lưu trữ; chỉnh sửa tại đây sẽ cập nhật ngay tên tệp Drive viết hoa.
                </p>
              </div>

              <div className="space-y-1 min-w-0">
                <div className="flex items-center justify-between gap-1 flex-wrap">
                  <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                    Trích yếu / Nội dung
                  </label>
                  {renderLocateButton('title', 'Trích yếu')}
                </div>
                <textarea
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  rows={3}
                  placeholder="Nguyên văn tiêu đề và trích yếu đọc được từ trang đầu để đối chiếu..."
                  className="w-full p-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              {/* Ngày văn bản + Đối chiếu 2 vị trí */}
              {(() => {
                const compStatus = analysis.extraction.dateComparisonStatus;
                const topDate = analysis.extraction.topRightDate;
                const sigDate = analysis.extraction.signatureDate;
                const dateInvalid =
                  documentDate.trim().length > 0 && !normalizeDateOnlyDdMmYyyy(documentDate.trim());

                if (compStatus === 'mismatch' && (topDate || sigDate)) {
                  return (
                    <div className="p-3.5 bg-amber-50 border border-amber-300 rounded-xl space-y-3 text-xs">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="flex items-center gap-2 text-amber-950 font-bold">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                          <span>Ngày trên văn bản chưa thống nhất</span>
                        </div>
                        <span className="text-[11px] font-bold px-2 py-0.5 rounded bg-amber-200 text-amber-950 border border-amber-300">
                          Chờ quản trị viên xác nhận ngày
                        </span>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        {topDate && (
                          <div className="bg-white border border-amber-200 rounded-lg p-2.5 space-y-1.5">
                            <div className="flex items-center justify-between gap-1">
                              <span className="text-[11px] font-semibold text-slate-600">
                                {topDate.positionLabel || 'Đầu trang 1 (góc phải)'}
                              </span>
                              <button
                                type="button"
                                onClick={() => handleLocateDateSource(topDate, 'Ngày ở đầu văn bản')}
                                className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded"
                              >
                                <Crosshair className="w-3 h-3" />
                                <span>Xem vị trí (Tr.{topDate.page})</span>
                              </button>
                            </div>
                            <div className="flex items-center justify-between gap-2">
                              <span className="font-mono font-bold text-sm text-slate-900 tabular-nums">
                                {topDate.date}
                              </span>
                              {canManageActiveBook && (
                                <button
                                  type="button"
                                  onClick={() => setDocumentDate(topDate.date)}
                                  className={`px-2.5 py-1 rounded text-[11px] font-semibold border transition-colors ${
                                    documentDate === topDate.date
                                      ? 'bg-sky-700 text-white border-sky-700'
                                      : 'bg-slate-50 hover:bg-sky-50 text-slate-800 border-slate-300'
                                  }`}
                                >
                                  {documentDate === topDate.date ? 'Đã chọn' : 'Chọn ngày này'}
                                </button>
                              )}
                            </div>
                          </div>
                        )}

                        {sigDate && (
                          <div className="bg-white border border-amber-200 rounded-lg p-2.5 space-y-1.5">
                            <div className="flex items-center justify-between gap-1">
                              <span className="text-[11px] font-semibold text-slate-600">
                                {sigDate.positionLabel || 'Khối chữ ký'}
                              </span>
                              <button
                                type="button"
                                onClick={() => handleLocateDateSource(sigDate, 'Ngày ở khối chữ ký')}
                                className="inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-semibold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded"
                              >
                                <Crosshair className="w-3 h-3" />
                                <span>Xem vị trí (Tr.{sigDate.page})</span>
                              </button>
                            </div>
                            <div className="flex items-center justify-between gap-2">
                              <span className="font-mono font-bold text-sm text-slate-900 tabular-nums">
                                {sigDate.date}
                              </span>
                              {canManageActiveBook && (
                                <button
                                  type="button"
                                  onClick={() => setDocumentDate(sigDate.date)}
                                  className={`px-2.5 py-1 rounded text-[11px] font-semibold border transition-colors ${
                                    documentDate === sigDate.date
                                      ? 'bg-sky-700 text-white border-sky-700'
                                      : 'bg-slate-50 hover:bg-sky-50 text-slate-800 border-slate-300'
                                  }`}
                                >
                                  {documentDate === sigDate.date ? 'Đã chọn' : 'Chọn ngày này'}
                                </button>
                              )}
                            </div>
                          </div>
                        )}
                      </div>

                      {canManageActiveBook ? (
                        <div className="space-y-1">
                          <label className="block text-[11px] font-bold text-amber-950 uppercase tracking-wide">
                            Ngày văn bản thống nhất (dd/mm/yyyy)
                          </label>
                          <div className="flex items-center gap-2">
                            <input
                              type="text"
                              value={documentDate}
                              onChange={(e) => setDocumentDate(e.target.value)}
                              onBlur={(e) => {
                                const norm = normalizeDateOnlyDdMmYyyy(e.target.value);
                                if (norm) setDocumentDate(norm);
                              }}
                              placeholder="Chọn một ngày ở trên hoặc nhập dd/mm/yyyy (để trống nếu muốn xác nhận sau)"
                              className="flex-1 min-h-[42px] h-10 px-3 bg-white border border-amber-300 rounded-lg font-mono text-sm text-slate-900 tabular-nums focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                            />
                            {documentDate && (
                              <button
                                type="button"
                                onClick={() => setDocumentDate('')}
                                className="px-2.5 py-2 text-[11px] font-semibold text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50"
                              >
                                Xác nhận sau
                              </button>
                            )}
                          </div>
                          {dateInvalid && (
                            <p className="text-[11px] text-rose-700 font-semibold">
                              Ngày không hợp lệ. Vui lòng nhập đúng định dạng dd/mm/yyyy.
                            </p>
                          )}
                        </div>
                      ) : (
                        <p className="text-[11px] text-amber-900 bg-white/80 border border-amber-200 rounded-lg px-2.5 py-2">
                          Hệ thống không tự chọn ngày chính thức khi 2 vị trí khác nhau. Hồ sơ sẽ lưu ở trạng thái{' '}
                          <strong>Chờ quản trị viên xác nhận ngày</strong> trong Sổ văn bản.
                        </p>
                      )}
                    </div>
                  );
                }

                return (
                  <div className="space-y-1 min-w-0">
                    <div className="flex items-center justify-between gap-1 flex-wrap">
                      <div className="flex items-center gap-2 flex-wrap">
                        <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                          Ngày văn bản
                        </label>
                        {compStatus === 'matched' && (
                          <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-200">
                            Khớp 2 vị trí
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {topDate && (
                          <button
                            type="button"
                            onClick={() => handleLocateDateSource(topDate, 'Ngày ở đầu văn bản')}
                            className="inline-flex items-center gap-1 px-2 py-1 min-h-[28px] text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded"
                          >
                            <Crosshair className="w-3 h-3" />
                            <span>Đầu VB (Tr.{topDate.page})</span>
                          </button>
                        )}
                        {sigDate && (
                          <button
                            type="button"
                            onClick={() => handleLocateDateSource(sigDate, 'Ngày ở khối chữ ký')}
                            className="inline-flex items-center gap-1 px-2 py-1 min-h-[28px] text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded"
                          >
                            <Crosshair className="w-3 h-3" />
                            <span>Khối chữ ký (Tr.{sigDate.page})</span>
                          </button>
                        )}
                        {!topDate && !sigDate && renderLocateButton('documentDate', 'Ngày văn bản')}
                      </div>
                    </div>
                    <input
                      type="text"
                      value={documentDate}
                      onChange={(e) => setDocumentDate(e.target.value)}
                      onBlur={(e) => {
                        const norm = normalizeDateOnlyDdMmYyyy(e.target.value);
                        if (norm) setDocumentDate(norm);
                      }}
                      placeholder="dd/mm/yyyy (để trống nếu chưa phát hiện)"
                      className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-sm text-slate-900 tabular-nums focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                    />
                    {dateInvalid && (
                      <p className="text-[11px] text-rose-700 font-semibold">
                        Ngày văn bản không hợp lệ. Vui lòng nhập đúng định dạng dd/mm/yyyy.
                      </p>
                    )}
                    {compStatus === 'missing' && !documentDate.trim() && (
                      <p
                        data-testid="missing-printed-date-notice"
                        className="text-[11px] text-amber-900 bg-amber-50 border border-amber-200 rounded-md px-2.5 py-1.5 flex items-start gap-1.5"
                      >
                        <AlertTriangle className="w-3.5 h-3.5 text-amber-600 shrink-0 mt-0.5" />
                        <span>
                          {analysis.extraction.ignoredDigitalSignatureDate
                            ? 'Chỉ phát hiện ngày/giờ trên chữ ký số điện tử (đã bỏ qua). Vui lòng đối chiếu bản PDF gốc và nhập ngày ban hành nếu có.'
                            : 'Chưa phát hiện ngày ban hành in trên văn bản. Vui lòng đối chiếu bản PDF gốc và nhập thủ công nếu có.'}
                        </span>
                      </p>
                    )}
                  </div>
                );
              })()}

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Họ tên người ký
                    </label>
                    {renderLocateButton('signer', 'Người ký')}
                  </div>
                  <input
                    type="text"
                    value={signer}
                    onChange={(e) => setSigner(e.target.value)}
                    placeholder="Họ và tên người ký..."
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                  {analysis.extraction.signerCandidates &&
                    analysis.extraction.signerCandidates.length > 1 && (
                      <div className="flex flex-wrap gap-1.5 pt-1">
                        {analysis.extraction.signerCandidates.map((cand) => (
                          <button
                            key={cand}
                            type="button"
                            onClick={() => setSigner(cand)}
                            className={`text-[11px] px-2.5 py-1 min-h-[32px] rounded border ${
                              signer === cand
                                ? 'bg-sky-700 text-white border-sky-700'
                                : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-sky-50'
                            }`}
                          >
                            {cand}
                          </button>
                        ))}
                      </div>
                    )}
                </div>

                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Chức vụ người ký
                    </label>
                    {renderLocateButton('signerRole', 'Chức vụ')}
                  </div>
                  <input
                    type="text"
                    value={signerRole}
                    onChange={(e) => setSignerRole(e.target.value)}
                    placeholder="Ví dụ: Hiệu trưởng, Phó Giám đốc..."
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>
              </div>

              <div className="space-y-1 min-w-0">
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Ghi chú <span className="text-slate-400 font-normal lowercase">(không bắt buộc)</span>
                </label>
                <input
                  type="text"
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  placeholder="Ghi chú phân công, hạn xử lý..."
                  className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              {/* Tên tệp sẽ lưu */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2 min-w-0">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <span className="text-xs font-bold text-slate-800 uppercase tracking-wide flex items-center gap-1.5">
                    <FileText className="w-4 h-4 text-sky-700 shrink-0" />
                    <span>Tên tệp sẽ lưu</span>
                  </span>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => {
                        if (!isEditingFileName) {
                          setCustomFileName(computedDriveFileName);
                          setIsEditingFileName(true);
                        } else {
                          setIsEditingFileName(false);
                          setCustomFileName('');
                        }
                      }}
                      className="text-xs font-semibold text-sky-700 hover:underline py-1"
                    >
                      {isEditingFileName ? 'Dùng tên tự động' : 'Sửa tên tệp'}
                    </button>
                  </div>
                </div>

                {isEditingFileName ? (
                  <div className="space-y-1.5">
                    <input
                      type="text"
                      value={customFileName}
                      onChange={(e) => setCustomFileName(e.target.value)}
                      className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-sky-400 rounded-lg font-mono text-base sm:text-sm font-bold text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                    />
                    <div className="flex items-center justify-between gap-2 bg-white border border-slate-200 rounded-md px-2.5 py-1.5 text-[11px]">
                      <span className="text-slate-500 shrink-0">Tên chuẩn sẽ lưu:</span>
                      <span className="font-mono font-bold text-sky-900 truncate flex-1">{computedDriveFileName}</span>
                      <button
                        type="button"
                        onClick={handleCopyFileName}
                        className="text-sky-700 hover:underline font-semibold shrink-0"
                      >
                        {copiedFileName ? 'Đã chép' : 'Chép'}
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="flex flex-col sm:flex-row sm:items-center gap-2 min-w-0">
                    <div
                      data-testid="pdf-computed-filename"
                      className="flex-1 min-w-0 px-3 py-2 bg-white border border-slate-300 rounded-lg font-mono text-xs sm:text-sm font-bold text-sky-950 overflow-x-auto whitespace-nowrap"
                    >
                      {computedDriveFileName}
                    </div>
                    <button
                      type="button"
                      onClick={handleCopyFileName}
                      className="px-3 py-2 min-h-[40px] text-xs font-semibold bg-white hover:bg-sky-50 border border-slate-300 rounded-lg text-slate-700 shrink-0 inline-flex items-center justify-center gap-1 w-full sm:w-auto"
                    >
                      {copiedFileName ? (
                        <>
                          <Check className="w-3.5 h-3.5 text-emerald-600" />
                          <span className="text-emerald-700">Đã chép</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Chép</span>
                        </>
                      )}
                    </button>
                  </div>
                )}
                <p className="text-[11px] text-slate-500 break-words">
                  Lưu vào Google Drive tại thư mục:{' '}
                  <strong className="text-slate-700">
                    {activeBook.name} → {(normalizeDateOnlyDdMmYyyy(documentDate) || analysis.extraction.topRightDate?.date || analysis.extraction.signatureDate?.date || '').match(/(\d{4})$/)?.[1] || serverYear || Number(new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric' }).format(new Date()))} → Văn bản đến
                  </strong>{' '}
                  bằng đúng tên ở ô này (giữ nguyên số gốc{' '}
                  <strong className="font-mono text-slate-700 break-all">{fullDocNumber || '...'}</strong> trong sổ).
                </p>
              </div>

              {!driveConfigured && (
                <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-900 flex items-start gap-2">
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                  <span>
                    Thư mục Google Drive hiện chưa cấu hình xong ({driveConfig?.lastCheckMessage || 'Chưa bật'}). Hồ sơ vẫn được ghi vào sổ với trạng thái{' '}
                    <strong>Lưu tệp thất bại</strong> để bạn thử lưu lại sau.
                  </span>
                </div>
              )}


              {/* Filename conflict resolution banner (tách biệt với cảnh báo trùng văn bản) */}
              {filenameConflict && (
                <div className="p-3.5 bg-amber-50 border border-amber-300 rounded-xl space-y-2.5 text-xs text-amber-950">
                  <div className="flex items-start gap-2">
                    <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                    <div>
                      <strong className="font-bold block">
                        Trùng tên tệp trên Drive: {filenameConflict.conflictingFileName}
                      </strong>
                      <p className="text-amber-900 mt-0.5">
                        Tên tệp này đã dùng cho hồ sơ{' '}
                        <strong className="font-mono">{filenameConflict.existingDocument.fullDocNumber}</strong>{' '}
                        ({filenameConflict.existingDocument.docName || filenameConflict.existingDocument.title}). Hãy chọn cách xử lý tên tệp:
                      </p>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={() => submitImport({ conflictAction: 'suffix' })}
                      className="px-3 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-md"
                    >
                      Thêm hậu tố ({filenameConflict.suggestedSuffixName})
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setCustomFileName(filenameConflict.suggestedSuffixName);
                        setIsEditingFileName(true);
                        setFilenameConflict(null);
                      }}
                      className="px-3 py-1.5 bg-white hover:bg-slate-50 border border-amber-300 text-amber-900 font-semibold rounded-md"
                    >
                      Tự sửa tên tệp
                    </button>
                    {canViewActiveBook && (
                      <button
                        type="button"
                        onClick={() => onOpenExistingDocument(filenameConflict.existingDocument.id)}
                        className="px-3 py-1.5 bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 font-semibold rounded-md"
                      >
                        Mở hồ sơ đã có
                      </button>
                    )}
                  </div>
                </div>
              )}

              {/* Inline Recovery Banner when temporary staged PDF expired (preserves all edits & idempotencyKey) */}
              {tempPdfExpiredError && (
                <div
                  data-testid="temp-pdf-expired-recovery"
                  className="p-3.5 bg-amber-50 border border-amber-300 rounded-xl space-y-2.5 text-xs text-amber-950"
                >
                  <div className="flex items-start gap-2">
                    <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                    <div className="space-y-1 min-w-0">
                      <strong className="font-bold block">
                        Bản PDF tạm đã hết hạn lưu trữ · Mọi nội dung bạn vừa chỉnh sửa vẫn được giữ nguyên
                      </strong>
                      <p className="text-amber-900">{tempPdfExpiredError}</p>
                    </div>
                  </div>

                  <input
                    ref={restagePdfFileInputRef}
                    type="file"
                    accept=".pdf,application/pdf"
                    className="hidden"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      e.target.value = '';
                      if (f) void handleRestageFilePicked(f);
                    }}
                  />

                  {restageShaStatus === 'mismatch' && restagePickFile && (
                    <div className="p-2.5 bg-white border border-rose-300 rounded-lg space-y-2 text-rose-900">
                      <p className="font-semibold">
                        Tệp vừa chọn (<span className="font-mono">{restagePickFile.name}</span>) khác mã kiểm tra SHA-256 với tệp PDF ban đầu.
                      </p>
                      <label className="inline-flex items-center gap-2 cursor-pointer font-medium">
                        <input
                          type="checkbox"
                          checked={allowDifferentRestageFile}
                          onChange={(e) => setAllowDifferentRestageFile(e.target.checked)}
                          className="rounded border-rose-400 text-rose-700"
                        />
                        <span>Xác nhận lưu hồ sơ này bằng tệp PDF mới vừa chọn</span>
                      </label>
                    </div>
                  )}

                  {restageShaStatus === 'matched' && restagePickFile && (
                    <p className="text-[11px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 rounded-md px-2.5 py-1.5">
                      Đã đối chiếu khớp đúng tệp PDF gốc ({restagePickFile.name}).
                    </p>
                  )}

                  <div className="flex flex-wrap items-center gap-2">
                    {selectedPdfFile && restageShaStatus !== 'mismatch' && (
                      <button
                        type="button"
                        disabled={restagingPdf || submitting}
                        onClick={() => handleRestageAndRetryImport(selectedPdfFile, false)}
                        className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-2xs"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${restagingPdf ? 'animate-spin' : ''}`} />
                        <span>{restagingPdf ? 'Đang nạp lại & lưu...' : 'Thử lưu lại'}</span>
                      </button>
                    )}

                    {restageShaStatus === 'mismatch' && restagePickFile && (
                      <button
                        type="button"
                        disabled={!allowDifferentRestageFile || restagingPdf || submitting}
                        onClick={() => handleRestageAndRetryImport(restagePickFile, true)}
                        className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-2xs"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${restagingPdf ? 'animate-spin' : ''}`} />
                        <span>{restagingPdf ? 'Đang nạp lại & lưu...' : 'Lưu bằng tệp mới này'}</span>
                      </button>
                    )}

                    <button
                      type="button"
                      disabled={restagingPdf || submitting}
                      onClick={() => restagePdfFileInputRef.current?.click()}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white hover:bg-amber-100/70 disabled:opacity-50 border border-amber-400 text-amber-950 font-semibold rounded-lg"
                    >
                      <FileUp className="w-3.5 h-3.5" />
                      <span>Chọn lại tệp PDF</span>
                    </button>
                  </div>
                </div>
              )}

              {submitError && (
                <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-start gap-2">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  <span>{submitError}</span>
                </div>
              )}

              <div className="pt-2 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <span className="text-[11px] text-slate-500">
                  Người nhập: <strong className="text-slate-700">{getUserDisplayName(currentUser)}</strong>
                </span>
                <div className="flex flex-col sm:flex-row sm:items-center gap-2 w-full sm:w-auto">
                  <span className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 bg-slate-100 border border-slate-200 rounded-lg text-xs font-bold text-slate-700">
                    <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                    <span>{activeBook.name}</span>
                  </span>
                  <button
                    type="submit"
                    disabled={
                      submitting ||
                      restagingPdf ||
                      docNumberIncomplete ||
                      (requiresDocNumberReview && !docNumberConfirmed)
                    }
                    className="min-h-[48px] h-12 sm:h-11 px-6 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-sm font-semibold rounded-lg shadow-sm inline-flex items-center justify-center gap-2 w-full sm:w-auto"
                  >
                    {submitting || restagingPdf ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Đang lưu hồ sơ &amp; tải PDF...</span>
                      </>
                    ) : (
                      <>
                        <Send className="w-4 h-4" />
                        <span>Xác nhận và lưu</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
