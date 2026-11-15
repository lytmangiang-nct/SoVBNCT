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
} from '../utils/fileNaming';
import { DEFAULT_BOOK_ID, hasBookPermission } from '../utils/permissions';
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
}

export const ImportPdfView: React.FC<ImportPdfViewProps> = ({
  currentUser,
  activeBook,
  accessibleBooks,
  onSelectBook,
  settings,
  serverYear,
  onDocumentImported,
  onDocumentUpdated,
  onOpenExistingDocument,
}) => {
  const [inputMode, setInputMode] = useState<'url' | 'file'>('url');
  const [pdfUrlInput, setPdfUrlInput] = useState('');
  const [selectedPdfFile, setSelectedPdfFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const retryPdfFileInputRef = useRef<HTMLInputElement>(null);

  const [analyzing, setAnalyzing] = useState(false);
  const [analyzeError, setAnalyzeError] = useState<string | null>(null);
  const [analysis, setAnalysis] = useState<PdfAnalyzeResponse | null>(null);
  const [jumpToPage, setJumpToPage] = useState<number | null>(null);
  const [highlightBox, setHighlightBox] = useState<PdfHighlightTarget | null>(null);

  // Mobile responsive view toggle while preserving form edits
  const [activeMobilePane, setActiveMobilePane] = useState<'form' | 'pdf'>('form');

  // Editable extracted fields
  const [fullDocNumber, setFullDocNumber] = useState('');
  const [docName, setDocName] = useState('');
  const [title, setTitle] = useState('');
  const [issueDate, setIssueDate] = useState('');
  const [signDate, setSignDate] = useState('');
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

  const driveConfig =
    activeBook.id !== DEFAULT_BOOK_ID
      ? activeBook.googleDriveConfig
      : activeBook.googleDriveConfig || settings.googleDriveConfig;
  const driveConfigured = Boolean(driveConfig?.isConfigured);
  const maxFileSizeMB = settings.googleDriveConfig?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;
  const maxFileSizeBytes = settings.googleDriveConfig?.maxFileSizeBytes || maxFileSizeMB * 1024 * 1024;

  // Unified live file name ("Tên tệp sẽ lưu")
  const computedDriveFileName = useMemo(() => {
    if (isEditingFileName && customFileName.trim()) {
      const cleaned = customFileName.trim();
      return cleaned.toLowerCase().endsWith('.pdf') ? cleaned : `${cleaned}.pdf`;
    }
    return buildStandardDriveFileName(fullDocNumber, docName || title, '.pdf');
  }, [fullDocNumber, docName, title, isEditingFileName, customFileName]);

  // Live check suspected duplicate when user edits fullDocNumber / issuingAgency / issueDate
  useEffect(() => {
    if (!analysis || savedDoc) return;
    if (!fullDocNumber.trim()) {
      setLiveSuspectedDuplicate(null);
      return;
    }

    const timer = setTimeout(async () => {
      try {
        const refDate = issueDate || signDate;
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
  }, [analysis, savedDoc, activeBook.id, fullDocNumber, issuingAgency, issueDate, signDate]);

  const applyAnalysisResult = (res: PdfAnalyzeResponse) => {
    setAnalysis(res);
    setFullDocNumber(res.extraction.docNumber || '');
    setDocName(res.extraction.docName || '');
    setTitle(res.extraction.title || '');
    setIssueDate(normalizeDateOnlyDdMmYyyy(res.extraction.issueDate || ''));
    setSignDate(normalizeDateOnlyDdMmYyyy(res.extraction.signDate || ''));
    setIssuingAgency(res.extraction.issuingAgency || '');
    setSigner(res.extraction.signer || '');
    setSignerRole(res.extraction.signerRole || '');
    setNotes('');
    setIsEditingFileName(false);
    setCustomFileName('');
    setFilenameConflict(null);
    setSubmitError(null);
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

  const handleAnalyzeUrl = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!pdfUrlInput.trim() || analyzing) return;
    setAnalyzing(true);
    setAnalyzeError(null);
    try {
      const res = await api.analyzePdfUrl(pdfUrlInput.trim(), activeBook.id);
      applyAnalysisResult(res);
    } catch (err: any) {
      setAnalyzeError(err.message || 'Không thể tải và phân tích PDF từ đường dẫn.');
    } finally {
      setAnalyzing(false);
    }
  };

  const handleAnalyzeFile = async (file: File) => {
    setAnalyzeError(null);
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

    setSelectedPdfFile(file);
    setAnalyzing(true);
    try {
      const res = await api.analyzePdfFile(file, activeBook.id);
      applyAnalysisResult(res);
    } catch (err: any) {
      setAnalyzeError(err.message || 'Không thể đọc và trích xuất thông tin từ tệp PDF.');
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
  }) => {
    if (!analysis || submitting) return;
    setSubmitError(null);
    setFilenameConflict(null);

    if (!fullDocNumber.trim()) {
      setSubmitError('Vui lòng nhập hoặc xác nhận Số/ký hiệu văn bản gốc.');
      return;
    }
    if (!docName.trim() && !title.trim()) {
      setSubmitError('Vui lòng nhập Tên văn bản hoặc Trích yếu nội dung văn bản.');
      return;
    }

    setSubmitting(true);
    try {
      const res = await api.importPdfDocument({
        bookId: activeBook.id,
        stageToken: analysis.stageToken,
        fullDocNumber: fullDocNumber.trim(),
        docName: docName.trim(),
        title: title.trim(),
        issueDate: normalizeDateOnlyDdMmYyyy(issueDate.trim()),
        signDate: normalizeDateOnlyDdMmYyyy(signDate.trim()),
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

      setSavedDoc(res.document);
      setSavedStageToken(res.stageToken || analysis.stageToken);
      onDocumentImported(res.document);
      if (!res.driveSaved && res.uploadError) {
        setRetryDriveError(res.uploadError);
      }
    } catch (err: any) {
      if (err.code === 'FILENAME_CONFLICT' && err.existingDocument) {
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
    setRetryDriveError(null);
    setHighlightBox(null);
    setLiveSuspectedDuplicate(null);
    setIgnoreDuplicatePdf(false);
    setIgnoreSuspectedDuplicate(false);
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
      {/* Header + Book selector */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
            Nhập Văn Bản PDF (Văn Bản Đến)
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            Tiếp nhận công văn, kế hoạch PDF từ Sở GD&amp;ĐT hoặc nguồn khác. Đọc và đối chiếu bản gốc bên cạnh, đặt tên chuẩn và lưu vào Google Drive mà{' '}
            <strong className="text-sky-800">không làm tăng dãy số văn bản đi</strong>.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 shrink-0 w-full md:w-auto">
          {importableBooks.length > 1 && (
            <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 min-h-[42px] shadow-2xs text-xs flex-1 sm:flex-initial min-w-0">
              <BookOpen className="w-3.5 h-3.5 text-sky-700 shrink-0" />
              <select
                value={activeBook.id}
                onChange={(e) => {
                  onSelectBook(e.target.value);
                  handleReset();
                }}
                className="font-semibold text-slate-800 bg-transparent focus:outline-hidden cursor-pointer flex-1 min-w-0 text-base sm:text-xs"
              >
                {importableBooks.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.name} ({b.unitCode})
                  </option>
                ))}
              </select>
            </div>
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
                  <span className="text-xs font-bold px-2.5 py-0.5 rounded bg-sky-100 text-sky-900 border border-sky-200">
                    Đã ghi sổ văn bản đến
                  </span>
                  <span
                    className={`text-xs font-bold px-2.5 py-0.5 rounded border ${
                      savedDoc.driveStatus === 'saved'
                        ? 'bg-emerald-100 text-emerald-900 border-emerald-300'
                        : 'bg-rose-100 text-rose-800 border-rose-300'
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
                  <FileCheck2 className="w-4 h-4 text-sky-700" />
                  <span>Xem trong Sổ theo dõi</span>
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
                <span className="truncate">Dán đường dẫn PDF</span>
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
              {inputMode === 'url' && (
                <button
                  type="button"
                  onClick={() => {
                    setInputMode('file');
                    setAnalyzeError(null);
                  }}
                  className="px-3 py-1.5 bg-white hover:bg-rose-100 border border-rose-300 text-rose-800 font-semibold rounded-md shrink-0"
                >
                  Chuyển sang Tải PDF từ máy
                </button>
              )}
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
                pdfUrl={api.getStagedPdfUrl(analysis.stageToken)}
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
                {importableBooks.length > 1 ? (
                  <div className="inline-flex items-center gap-1.5 text-xs bg-sky-50 border border-sky-200 rounded-md px-2.5 py-1.5 min-h-[38px] max-w-full">
                    <BookOpen className="w-3.5 h-3.5 text-sky-700 shrink-0" />
                    <select
                      aria-label="Sổ tiếp nhận văn bản đến"
                      value={activeBook.id}
                      onChange={(e) => onSelectBook(e.target.value)}
                      className="font-bold text-sky-900 bg-transparent focus:outline-hidden cursor-pointer min-w-0 text-base sm:text-xs"
                    >
                      {importableBooks.map((b) => (
                        <option key={b.id} value={b.id}>
                          {b.name} ({b.unitCode})
                        </option>
                      ))}
                    </select>
                  </div>
                ) : (
                  <span className="text-[11px] font-bold px-2.5 py-1 rounded bg-sky-50 text-sky-800 border border-sky-200 shrink-0">
                    {activeBook.name} ({activeBook.unitCode})
                  </span>
                )}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Số / Ký hiệu gốc <span className="text-rose-600">*</span>
                    </label>
                    {renderLocateButton('docNumber', 'Số/ký hiệu')}
                  </div>
                  <input
                    type="text"
                    value={fullDocNumber}
                    onChange={(e) => {
                      setFullDocNumber(e.target.value);
                      setIsEditingFileName(false);
                    }}
                    placeholder="Ví dụ: 1234/SGDĐT-GDTrH"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono font-bold text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

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
                    placeholder="Ví dụ: Sở GD&ĐT, Trường THPT..."
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
                  placeholder="Tên ngắn gọn của văn bản (dùng để đặt tên tệp lưu trữ)..."
                  className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              <div className="space-y-1 min-w-0">
                <div className="flex items-center justify-between gap-1 flex-wrap">
                  <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                    Trích yếu / Nội dung (V/v)
                  </label>
                  {renderLocateButton('title', 'Trích yếu')}
                </div>
                <textarea
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  rows={2}
                  placeholder="Nội dung trích yếu sau V/v..."
                  className="w-full p-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Ngày ban hành
                    </label>
                    {renderLocateButton('issueDate', 'Ngày ban hành')}
                  </div>
                  <input
                    type="text"
                    value={issueDate}
                    onChange={(e) => setIssueDate(e.target.value)}
                    onBlur={(e) => setIssueDate(normalizeDateOnlyDdMmYyyy(e.target.value))}
                    placeholder="dd/mm/yyyy (để trống nếu chưa có)"
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                      Ngày ký
                    </label>
                    {renderLocateButton('signDate', 'Ngày ký')}
                  </div>
                  <input
                    type="text"
                    value={signDate}
                    onChange={(e) => setSignDate(e.target.value)}
                    onBlur={(e) => setSignDate(normalizeDateOnlyDdMmYyyy(e.target.value))}
                    placeholder="dd/mm/yyyy (chỉ điền nếu có ngày ký riêng)"
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>
              </div>

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
                  <input
                    type="text"
                    value={customFileName}
                    onChange={(e) => setCustomFileName(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-sky-400 rounded-lg font-mono text-base sm:text-sm font-bold text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                ) : (
                  <div className="flex flex-col sm:flex-row sm:items-center gap-2 min-w-0">
                    <div className="flex-1 min-w-0 px-3 py-2 bg-white border border-slate-300 rounded-lg font-mono text-xs sm:text-sm font-bold text-sky-950 overflow-x-auto whitespace-nowrap">
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
                    {activeBook.name} → {(issueDate || signDate).match(/(\d{4})$/)?.[1] || serverYear || Number(new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric' }).format(new Date()))} → Văn bản đến
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

              {submitError && (
                <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-start gap-2">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  <span>{submitError}</span>
                </div>
              )}

              <div className="pt-2 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <span className="text-[11px] text-slate-500">
                  Người nhập: <strong className="text-slate-700">{currentUser.fullName}</strong>
                </span>
                <div className="flex flex-col sm:flex-row sm:items-center gap-2 w-full sm:w-auto">
                  <span className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 bg-slate-100 border border-slate-200 rounded-lg text-xs font-bold text-slate-700">
                    <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                    <span>{activeBook.name}</span>
                  </span>
                  <button
                    type="submit"
                    disabled={submitting}
                    className="min-h-[48px] h-12 sm:h-11 px-6 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-sm font-semibold rounded-lg shadow-sm inline-flex items-center justify-center gap-2 w-full sm:w-auto"
                  >
                    {submitting ? (
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
