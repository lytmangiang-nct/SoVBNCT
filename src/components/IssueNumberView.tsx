import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  CheckCircle2,
  Copy,
  Check,
  AlertCircle,
  AlertTriangle,
  Clock,
  Sparkles,
  ArrowRight,
  Shield,
  Send,
  RotateCcw,
  X,
  Plus,
  FileText,
  Upload,
  FileUp,
  ExternalLink,
  RefreshCw,
  FolderCheck,
  BookOpen,
  Inbox,
  CheckSquare,
  XCircle,
  Download,
  FileCheck2,
  Loader2,
} from 'lucide-react';
import { api } from '../api';
import type {
  DocumentRecord,
  DocumentTypeConfig,
  NumberRequest,
  PreIssueCheckResult,
  RegistryBook,
  SignerConfig,
  SystemSettings,
  User,
  WordTemplate,
} from '../types';
import { getTypeColorStyle, DOCUMENT_COLOR_PALETTE, SEMANTIC_STATUS_STYLES } from '../utils/documentColors';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  formatFileSize,
  generateSuggestedFileName,
  getWordFileExtension,
} from '../utils/fileNaming';
import {
  DEFAULT_BOOK_ID,
  getBookEnabledTypes,
  getUserDisplayName,
  hasBookPermission,
  isSystemAdmin,
} from '../utils/permissions';

interface IssueNumberViewProps {
  currentUser: User;
  activeBook: RegistryBook;
  accessibleBooks: RegistryBook[];
  onSelectBook: (bookId: string) => void;
  booksLoaded?: boolean;
  appSyncError?: string | null;
  onRetryAppLoad?: () => void;
  settings: SystemSettings;
  documentTypes: DocumentTypeConfig[];
  signers: SignerConfig[];
  documents?: DocumentRecord[];
  numberRequests: NumberRequest[];
  serverYear?: number;
  serverFormattedDate?: string;
  onServerTimeUpdate?: (year: number, formattedDate?: string) => void;
  onRefreshRequests: () => void;
  onDocumentIssued: (doc: DocumentRecord) => void;
  onDocumentUpdated?: (doc: DocumentRecord) => void;
  onNavigateToRegistry: () => void;
  onOpenExistingDocument?: (docId: string) => void;
  onDocumentTypeAdded?: (newType: DocumentTypeConfig, updatedBook?: RegistryBook) => void;
}

export const IssueNumberView: React.FC<IssueNumberViewProps> = ({
  currentUser,
  activeBook,
  accessibleBooks,
  onSelectBook,
  booksLoaded = true,
  appSyncError = null,
  onRetryAppLoad,
  settings,
  documentTypes,
  signers: globalSigners,
  documents = [],
  numberRequests,
  serverYear,
  serverFormattedDate,
  onServerTimeUpdate,
  onRefreshRequests,
  onDocumentIssued,
  onDocumentUpdated,
  onNavigateToRegistry,
  onOpenExistingDocument,
  onDocumentTypeAdded,
}) => {
  const canIssue = hasBookPermission(currentUser, activeBook.id, 'issue_number');
  const canRequest = hasBookPermission(currentUser, activeBook.id, 'request_number');
  const canView = hasBookPermission(currentUser, activeBook.id, 'view');
  const canManageDocTypes =
    isSystemAdmin(currentUser) || hasBookPermission(currentUser, activeBook.id, 'manage_book');

  const effectiveSigners = useMemo(
    () => (activeBook.signers && activeBook.signers.length > 0 ? activeBook.signers : globalSigners),
    [activeBook.signers, globalSigners]
  );

  // Local optimistic buffer for newly added/enabled types per book so typeCode never flickers back to KH before parent props sync
  const [recentlyAddedTypesByBook, setRecentlyAddedTypesByBook] = useState<
    Record<string, DocumentTypeConfig[]>
  >({});

  const mergedDocumentTypes = useMemo(() => {
    const extras = recentlyAddedTypesByBook[activeBook.id] || [];
    if (extras.length === 0) return documentTypes;
    const map = new Map<string, DocumentTypeConfig>();
    for (const t of documentTypes) {
      map.set(t.code.trim().toUpperCase(), t);
    }
    for (const extra of extras) {
      const key = extra.code.trim().toUpperCase();
      if (!map.has(key)) {
        map.set(key, extra);
      }
    }
    return Array.from(map.values());
  }, [documentTypes, recentlyAddedTypesByBook, activeBook.id]);

  // Form State (scoped to document types enabled in activeBook)
  const activeTypes = useMemo(() => {
    const base = getBookEnabledTypes(activeBook, mergedDocumentTypes);
    const extras = recentlyAddedTypesByBook[activeBook.id] || [];
    if (extras.length === 0) return base;
    const existingCodes = new Set(base.map((t) => t.code.trim().toUpperCase()));
    const missingExtras = extras.filter(
      (t) => t.isActive && !existingCodes.has(t.code.trim().toUpperCase())
    );
    return missingExtras.length > 0 ? [...base, ...missingExtras] : base;
  }, [activeBook, mergedDocumentTypes, recentlyAddedTypesByBook]);
  const defaultTypeCode = activeTypes.length > 0 ? activeTypes[0].code : 'KH';

  const [typeCode, setTypeCode] = useState<string>(defaultTypeCode);
  const [docName, setDocName] = useState<string>('');
  const [title, setTitle] = useState<string>('');
  const [signer, setSigner] = useState<string>(effectiveSigners.length > 0 ? effectiveSigners[0].name : '');
  const [signerRole, setSignerRole] = useState<string>(effectiveSigners.length > 0 ? effectiveSigners[0].role : '');
  const [isCustomSigner, setIsCustomSigner] = useState<boolean>(false);
  const [notes, setNotes] = useState<string>('');

  // Quick Add Type Modal State (Admin / manage_book)
  const [showAddTypeModal, setShowAddTypeModal] = useState<boolean>(false);
  const [newTypeName, setNewTypeName] = useState<string>('');
  const [newTypeCode, setNewTypeCode] = useState<string>('');
  const [newTypeColor, setNewTypeColor] = useState<string>('blue');
  const [addTypeSaving, setAddTypeSaving] = useState<boolean>(false);
  const [addTypeError, setAddTypeError] = useState<string | null>(null);
  const [existingCatalogType, setExistingCatalogType] = useState<DocumentTypeConfig | null>(null);

  // Idempotency Key
  const generateKey = () => `req-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const [idempotencyKey, setIdempotencyKey] = useState<string>(generateKey);

  // Next Document Number Preview (strictly scoped by book + year + series + request sequence)
  const [previewStatus, setPreviewStatus] = useState<'loading' | 'ready' | 'error'>('loading');
  const [nextNumber, setNextNumber] = useState<string | null>(null);
  const [nextSequenceNumber, setNextSequenceNumber] = useState<number | null>(null);
  const [previewStartSequence, setPreviewStartSequence] = useState<number | null>(null);
  const [previewYear, setPreviewYear] = useState<number | null>(null);
  const [previewSeriesKey, setPreviewSeriesKey] = useState<string | null>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [sequenceConflict, setSequenceConflict] = useState<{
    expectedNextSeq: number;
    currentNextSeq: number;
    previewNumber: string;
    message: string;
  } | null>(null);

  const previewAbortRef = useRef<AbortController | null>(null);
  const previewReqSeqRef = useRef<number>(0);
  const activeSeriesKeyRef = useRef<string>('');
  const lastServerPreviewDocIdRef = useRef<string | null>(null);

  // Submission State
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [issuedSuccessDoc, setIssuedSuccessDoc] = useState<DocumentRecord | null>(null);
  const [requestSuccess, setRequestSuccess] = useState<NumberRequest | null>(null);
  const [copied, setCopied] = useState<boolean>(false);
  const [copiedFileName, setCopiedFileName] = useState<boolean>(false);
  const [copiedSuccessFileName, setCopiedSuccessFileName] = useState<boolean>(false);

  // Queue Approval State (for issue_number users)
  const [processingReqId, setProcessingReqId] = useState<string | null>(null);
  const [queueError, setQueueError] = useState<string | null>(null);
  const [rejectingReqId, setRejectingReqId] = useState<string | null>(null);
  const [rejectReasonInput, setRejectReasonInput] = useState<string>('Chưa đủ thông tin văn bản');

  // Word File Attachment State
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [selectedFileSha256, setSelectedFileSha256] = useState<string | undefined>(undefined);
  const [fileError, setFileError] = useState<string | null>(null);
  const [uploadStep, setUploadStep] = useState<'idle' | 'issuing' | 'uploading' | 'saved' | 'failed'>('idle');
  const [pendingRetryFile, setPendingRetryFile] = useState<File | null>(null);
  const [isRetryingUpload, setIsRetryingUpload] = useState<boolean>(false);
  const [retryError, setRetryError] = useState<string | null>(null);
  const [filenameConflict, setFilenameConflict] = useState<{
    conflictingFileName: string;
    suggestedSuffixName: string;
    existingDocument: DocumentRecord;
    officialDocNumberPreview?: string;
  } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const retryFileInputRef = useRef<HTMLInputElement>(null);

  // Word Templates State (auto-generate .docx from template)
  const [wordTemplates, setWordTemplates] = useState<WordTemplate[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>('');
  const [templateDraftCheck, setTemplateDraftCheck] = useState<{
    missingFields: string[];
    replacedPlaceholders: string[];
  } | null>(null);
  const [downloadingDraft, setDownloadingDraft] = useState<boolean>(false);
  const [draftError, setDraftError] = useState<string | null>(null);
  const [isRetryingTemplateWord, setIsRetryingTemplateWord] = useState<boolean>(false);

  // Pre-Issue Confirmation Modal State
  const [checkingPreIssue, setCheckingPreIssue] = useState<boolean>(false);
  const [preIssueModalData, setPreIssueModalData] = useState<{
    check: PreIssueCheckResult;
    fileSha256?: string;
  } | null>(null);

  const driveConfig =
    activeBook.id !== DEFAULT_BOOK_ID
      ? activeBook.googleDriveConfig
      : activeBook.googleDriveConfig || settings.googleDriveConfig;
  const driveConfigured = Boolean(driveConfig?.isConfigured);
  const maxWordFileSizeMB = settings.googleDriveConfig?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;
  const maxWordFileSizeBytes = settings.googleDriveConfig?.maxFileSizeBytes || maxWordFileSizeMB * 1024 * 1024;

  const [todayVietnam, setTodayVietnam] = useState<string>(() => {
    if (serverFormattedDate) return serverFormattedDate;
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Ho_Chi_Minh',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    });
    const [y, m, d] = formatter.format(new Date()).split('-');
    return `${d}/${m}/${y}`;
  });

  const [formYear, setFormYear] = useState<number>(() => {
    if (serverYear) return serverYear;
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Ho_Chi_Minh',
      year: 'numeric',
    });
    return parseInt(formatter.format(new Date()), 10);
  });

  const [yearRolloverNotice, setYearRolloverNotice] = useState<{
    fromYear: number;
    toYear: number;
    newPreviewNumber?: string;
    confirmed: boolean;
  } | null>(null);

  const hasOpenFormDraft = Boolean(docName.trim() || title.trim() || notes.trim() || selectedFile);

  useEffect(() => {
    if (serverFormattedDate) {
      setTodayVietnam(serverFormattedDate);
    }
  }, [serverFormattedDate]);

  useEffect(() => {
    if (!serverYear || serverYear === formYear) return;
    if (hasOpenFormDraft) {
      setYearRolloverNotice((prev) => ({
        fromYear: prev?.fromYear ?? formYear,
        toYear: serverYear,
        newPreviewNumber:
          previewStatus === 'ready' && previewYear === serverYear && nextNumber
            ? nextNumber
            : prev?.newPreviewNumber,
        confirmed: prev?.toYear === serverYear ? prev.confirmed : false,
      }));
    } else {
      setFormYear(serverYear);
      setYearRolloverNotice(null);
    }
  }, [serverYear, formYear, hasOpenFormDraft, previewStatus, previewYear, nextNumber]);

  // Đồng bộ số dự kiến mới nhất vào banner chuyển năm ngay khi preview năm mới trả về
  useEffect(() => {
    if (!yearRolloverNotice) return;
    if (previewStatus === 'ready' && previewYear === yearRolloverNotice.toYear && nextNumber) {
      if (yearRolloverNotice.newPreviewNumber !== nextNumber) {
        setYearRolloverNotice((prev) =>
          prev ? { ...prev, newPreviewNumber: nextNumber } : prev
        );
      }
    }
  }, [yearRolloverNotice, previewStatus, previewYear, nextNumber]);

  // Reset signer default when book changes
  useEffect(() => {
    if (!isCustomSigner && effectiveSigners.length > 0) {
      const stillValid = effectiveSigners.some((s) => s.name === signer);
      if (!stillValid) {
        setSigner(effectiveSigners[0].name);
        setSignerRole(effectiveSigners[0].role);
      }
    }
  }, [activeBook.id, effectiveSigners]);

  // Keep typeCode valid for the activeBook
  useEffect(() => {
    if (activeTypes.length > 0 && !activeTypes.some((t) => t.code === typeCode)) {
      setTypeCode(activeTypes[0].code);
    }
  }, [activeTypes, activeBook.id, typeCode]);

  // Load Word Templates scoped to activeBook
  useEffect(() => {
    if (!canIssue) {
      setWordTemplates([]);
      return;
    }
    let cancelled = false;
    api
      .getWordTemplates(activeBook.id)
      .then((res) => {
        if (!cancelled) setWordTemplates(res.templates || []);
      })
      .catch(() => {
        if (!cancelled) setWordTemplates([]);
      });
    return () => {
      cancelled = true;
    };
  }, [activeBook.id, canIssue]);

  const matchingTemplates = useMemo(
    () =>
      wordTemplates.filter(
        (t) =>
          t.bookId === activeBook.id &&
          t.typeCode.trim().toUpperCase() === typeCode.trim().toUpperCase()
      ),
    [wordTemplates, activeBook.id, typeCode]
  );

  useEffect(() => {
    if (selectedFile) {
      if (selectedTemplateId) setSelectedTemplateId('');
      return;
    }
    if (selectedTemplateId && !matchingTemplates.some((t) => t.id === selectedTemplateId)) {
      setSelectedTemplateId('');
    }
  }, [matchingTemplates, selectedTemplateId, selectedFile]);

  useEffect(() => {
    if (!selectedTemplateId || selectedFile) {
      setTemplateDraftCheck(null);
      setDraftError(null);
      return;
    }
    const timer = setTimeout(() => {
      api
        .checkWordTemplateDraft(selectedTemplateId, {
          docName: docName.trim(),
          title: title.trim(),
          signer: signer.trim(),
          signerRole: signerRole.trim(),
          issueDate: todayVietnam,
        })
        .then((res) => {
          setTemplateDraftCheck({
            missingFields: res.missingFields || [],
            replacedPlaceholders: res.replacedPlaceholders || [],
          });
          setDraftError(null);
        })
        .catch(() => {
          setTemplateDraftCheck(null);
        });
    }, 280);
    return () => clearTimeout(timer);
  }, [selectedTemplateId, selectedFile, docName, title, signer, signerRole, todayVietnam]);

  const effectiveSeriesTypeCode = activeBook.numberMode === 'shared' ? 'SHARED' : typeCode;
  const currentSeriesKey = `${activeBook.id}::${effectiveSeriesTypeCode}::${activeBook.numberMode}`;
  activeSeriesKeyRef.current = currentSeriesKey;

  const isTypeValidInActiveBook = useMemo(
    () => activeTypes.some((t) => t.code === typeCode),
    [activeTypes, typeCode]
  );

  const fetchNextNumber = React.useCallback(
    async (options?: { silent?: boolean }) => {
      const silent = Boolean(options?.silent);
      if (!booksLoaded) {
        if (!silent) {
          setPreviewStatus(appSyncError ? 'error' : 'loading');
          setNextNumber(null);
          setNextSequenceNumber(null);
          setPreviewStartSequence(null);
          setPreviewYear(null);
          setPreviewSeriesKey(null);
          setPreviewError(appSyncError || null);
        }
        return;
      }

      if (!typeCode || (!canIssue && !canRequest && !canView)) {
        setPreviewStatus('error');
        setNextNumber(null);
        setNextSequenceNumber(null);
        setPreviewStartSequence(null);
        setPreviewYear(null);
        setPreviewSeriesKey(null);
        setPreviewError('Tài khoản chưa có quyền đọc số trên sổ này.');
        return;
      }

      if (!isTypeValidInActiveBook && activeTypes.length > 0) {
        setPreviewStatus('loading');
        setNextNumber(null);
        setNextSequenceNumber(null);
        setPreviewStartSequence(null);
        setPreviewYear(null);
        setPreviewSeriesKey(null);
        return;
      }

      if (previewAbortRef.current) {
        previewAbortRef.current.abort();
      }
      const controller = new AbortController();
      previewAbortRef.current = controller;

      const reqId = ++previewReqSeqRef.current;
      const targetBookId = activeBook.id;
      const targetTypeCode = typeCode;
      const targetMode = activeBook.numberMode;
      const targetSeriesKey = `${targetBookId}::${
        targetMode === 'shared' ? 'SHARED' : targetTypeCode
      }::${targetMode}`;

      if (!silent) {
        setPreviewStatus('loading');
        setNextNumber(null);
        setNextSequenceNumber(null);
        setPreviewStartSequence(null);
        setPreviewYear(null);
        setPreviewSeriesKey(null);
        setPreviewError(null);
        setSequenceConflict(null);
      }

      try {
        const res = await api.previewNumber(
          targetTypeCode,
          targetBookId,
          undefined,
          controller.signal
        );
        if (
          reqId !== previewReqSeqRef.current ||
          activeSeriesKeyRef.current !== targetSeriesKey
        ) {
          return;
        }

        const expectedRespTypeCode =
          targetMode === 'shared' ? 'SHARED' : targetTypeCode.trim().toUpperCase();
        const actualRespTypeCode = String(res.typeCode || '').trim().toUpperCase();
        const isScopeValid =
          res.bookId === targetBookId &&
          actualRespTypeCode === expectedRespTypeCode &&
          Number.isFinite(Number(res.year)) &&
          typeof res.sequenceNumber === 'number' &&
          res.sequenceNumber >= 1 &&
          Boolean(res.previewNumber);

        if (!isScopeValid) {
          setNextNumber(null);
          setNextSequenceNumber(null);
          setPreviewStartSequence(null);
          setPreviewYear(null);
          setPreviewSeriesKey(null);
          setPreviewStatus('error');
          setPreviewError('Dữ liệu cấu hình số từ máy chủ không khớp sổ hoặc loại văn bản đang chọn.');
          return;
        }

        setNextNumber(res.previewNumber);
        setNextSequenceNumber(res.sequenceNumber);
        setPreviewStartSequence(
          typeof res.customStartSequence === 'number' && res.customStartSequence >= 1
            ? res.customStartSequence
            : null
        );
        setPreviewYear(Number(res.year));
        setPreviewSeriesKey(targetSeriesKey);
        setPreviewStatus('ready');
        setPreviewError(null);

        if (res.formattedDate) {
          setTodayVietnam(res.formattedDate);
        }
        const latestSrvYear = res.serverYear || res.year;
        if (latestSrvYear) {
          onServerTimeUpdate?.(latestSrvYear, res.formattedDate);
          if (latestSrvYear !== formYear) {
            if (docName.trim() || title.trim() || notes.trim() || selectedFile) {
              setYearRolloverNotice({
                fromYear: formYear,
                toYear: latestSrvYear,
                newPreviewNumber: res.previewNumber,
                confirmed: false,
              });
            } else {
              setFormYear(latestSrvYear);
              setYearRolloverNotice(null);
            }
          }
        }
      } catch (err: any) {
        if (err?.name === 'AbortError') return;
        if (
          reqId !== previewReqSeqRef.current ||
          activeSeriesKeyRef.current !== targetSeriesKey
        ) {
          return;
        }
        if (!silent || previewSeriesKey !== targetSeriesKey) {
          setNextNumber(null);
          setNextSequenceNumber(null);
          setPreviewStartSequence(null);
          setPreviewYear(null);
          setPreviewSeriesKey(null);
          setPreviewStatus('error');
          setPreviewError(err?.message || 'Chưa tải được cấu hình số từ máy chủ.');
        }
      }
    },
    [
      booksLoaded,
      appSyncError,
      typeCode,
      activeBook.id,
      activeBook.numberMode,
      canIssue,
      canRequest,
      canView,
      isTypeValidInActiveBook,
      activeTypes.length,
      formYear,
      docName,
      title,
      notes,
      selectedFile,
      onServerTimeUpdate,
      previewSeriesKey,
    ]
  );

  // Immediately clear stale preview & fetch fresh number when switching book, mode, typeCode, or user
  useEffect(() => {
    if (!booksLoaded) {
      if (previewAbortRef.current) {
        previewAbortRef.current.abort();
      }
      previewReqSeqRef.current += 1;
      setPreviewStatus(appSyncError ? 'error' : 'loading');
      setNextNumber(null);
      setNextSequenceNumber(null);
      setPreviewStartSequence(null);
      setPreviewYear(null);
      setPreviewSeriesKey(null);
      setPreviewError(appSyncError || null);
      setSequenceConflict(null);
      return;
    }

    fetchNextNumber({ silent: false });

    return () => {
      if (previewAbortRef.current) {
        previewAbortRef.current.abort();
      }
    };
  }, [
    currentUser.id,
    booksLoaded,
    appSyncError,
    activeBook.id,
    activeBook.numberMode,
    typeCode,
    isTypeValidInActiveBook,
    canIssue,
    canRequest,
    canView,
    serverYear,
  ]);

  // Background recheck when configVersion changes, documents update from other tabs/devices, VN midnight rolls over, or tab regains focus
  useEffect(() => {
    if (!booksLoaded) return;

    const getLocalVnYear = () => {
      try {
        const y = Number(
          new Intl.DateTimeFormat('en-CA', {
            timeZone: 'Asia/Ho_Chi_Minh',
            year: 'numeric',
          }).format(new Date())
        );
        return Number.isFinite(y) && y > 2000 ? y : formYear;
      } catch {
        return formYear;
      }
    };

    const handleRecheck = () => {
      if (document.visibilityState === 'visible') {
        fetchNextNumber({ silent: previewStatus === 'ready' && previewSeriesKey === currentSeriesKey });
      }
    };
    window.addEventListener('focus', handleRecheck);
    document.addEventListener('visibilitychange', handleRecheck);

    const midnightTimer = window.setInterval(() => {
      const localVnYear = getLocalVnYear();
      if (localVnYear !== formYear && !yearRolloverNotice) {
        fetchNextNumber({ silent: true });
      }
    }, 10000);

    return () => {
      window.removeEventListener('focus', handleRecheck);
      document.removeEventListener('visibilitychange', handleRecheck);
      window.clearInterval(midnightTimer);
    };
  }, [booksLoaded, fetchNextNumber, previewStatus, previewSeriesKey, currentSeriesKey, formYear, yearRolloverNotice]);

  const lastDocSignatureRef = useRef<string>('');
  useEffect(() => {
    if (!booksLoaded) return;
    const topDocId = documents[0]?.id || '';
    const sig = `${activeBook.configVersion || 1}::${activeBook.numberFormat}::${documents.length}::${topDocId}`;
    if (!lastDocSignatureRef.current) {
      lastDocSignatureRef.current = sig;
      return;
    }
    if (lastDocSignatureRef.current === sig) return;
    lastDocSignatureRef.current = sig;

    // If this document change came from our own just-completed issueDocument() call, we already have exact nextPreview
    if (topDocId && lastServerPreviewDocIdRef.current === topDocId) {
      lastServerPreviewDocIdRef.current = null;
      return;
    }

    fetchNextNumber({ silent: previewStatus === 'ready' && previewSeriesKey === currentSeriesKey });
  }, [
    booksLoaded,
    activeBook.configVersion,
    activeBook.numberFormat,
    documents.length,
    documents[0]?.id,
  ]);

  const handleSignerSelect = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value;
    if (val === '__custom__') {
      setIsCustomSigner(true);
      setSigner('');
      setSignerRole('');
    } else {
      setIsCustomSigner(false);
      const found = effectiveSigners.find((s) => s.name === val);
      if (found) {
        setSigner(found.name);
        setSignerRole(found.role);
      } else {
        setSigner(val);
      }
    }
  };

  const validateWordFile = (file: File): string | null => {
    const ext = getWordFileExtension(file.name);
    if (!ext) {
      return 'Định dạng tệp không hợp lệ. Vui lòng chỉ chọn tệp Word có đuôi .doc hoặc .docx.';
    }
    if (file.size <= 0) {
      return 'Tệp Word đang bị rỗng (0 bytes). Vui lòng chọn tệp khác.';
    }
    if (file.size > maxWordFileSizeBytes) {
      return `Dung lượng tệp (${formatFileSize(file.size)}) vượt quá giới hạn tối đa ${maxWordFileSizeMB} MB.`;
    }
    return null;
  };

  const computeFileSha256Hex = async (file: File): Promise<string | undefined> => {
    try {
      if (!window.crypto?.subtle) return undefined;
      const buf = await file.arrayBuffer();
      const hashBuffer = await window.crypto.subtle.digest('SHA-256', buf);
      return Array.from(new Uint8Array(hashBuffer))
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    } catch {
      return undefined;
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setFileError(null);
    setSubmitError(null);
    if (!file) return;

    const err = validateWordFile(file);
    if (err) {
      setFileError(err);
      setSelectedFile(null);
      setSelectedFileSha256(undefined);
      if (fileInputRef.current) fileInputRef.current.value = '';
      return;
    }
    setSelectedFile(file);
    setSelectedTemplateId('');
    void computeFileSha256Hex(file).then((hex) => setSelectedFileSha256(hex));
  };

  const handleClearSelectedFile = () => {
    setSelectedFile(null);
    setSelectedFileSha256(undefined);
    setFileError(null);
    setSubmitError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDownloadTemplateDraft = async () => {
    if (!selectedTemplateId || downloadingDraft) return;
    setDownloadingDraft(true);
    setDraftError(null);
    try {
      await api.downloadWordTemplateDraft(selectedTemplateId, {
        docName: docName.trim(),
        title: title.trim(),
        signer: signer.trim(),
        signerRole: signerRole.trim(),
        issueDate: todayVietnam,
      });
    } catch (err: any) {
      setDraftError(err?.message || 'Không thể tải bản nháp Word.');
    } finally {
      setDownloadingDraft(false);
    }
  };

  const performIssueSubmit = async (
    conflictAction?: 'suffix' | 'overwrite',
    confirmUpdatedSequence = false,
    overrideSha256?: string,
    overrideExpected?: { expectedYear?: number; expectedNextSeq?: number }
  ) => {
    if (isSubmitting) return;
    setSubmitError(null);
    setFilenameConflict(null);
    setRequestSuccess(null);

    if (!typeCode.trim()) {
      setSubmitError('Vui lòng chọn loại văn bản');
      return;
    }
    if (!docName.trim()) {
      setSubmitError('Vui lòng nhập tên văn bản (thông tin bắt buộc)');
      return;
    }
    if (!signer.trim()) {
      setSubmitError('Vui lòng chọn hoặc nhập người ký văn bản');
      return;
    }
    if (!signerRole.trim()) {
      setSubmitError('Vui lòng nhập chức vụ của người ký');
      return;
    }

    // Luồng 1: Người dùng chỉ có quyền "Xin số" (request_number) -> gửi yêu cầu chờ duyệt, KHÔNG được gọi cấp số trực tiếp
    if (!canIssue) {
      if (!canRequest) {
        setSubmitError(`Bạn chưa có quyền "Xin số" hoặc "Cấp số" trong ${activeBook.name}.`);
        return;
      }
      setIsSubmitting(true);
      try {
        const res = await api.createNumberRequest({
          bookId: activeBook.id,
          typeCode: typeCode.trim().toUpperCase(),
          docName: docName.trim(),
          title: title.trim(),
          signer: signer.trim(),
          signerRole: signerRole.trim(),
          notes: notes.trim(),
        });
        setRequestSuccess(res.request);
        onRefreshRequests();
        setDocName('');
        setTitle('');
        setNotes('');
      } catch (err: any) {
        setSubmitError(err.message || 'Không thể gửi yêu cầu xin số.');
      } finally {
        setIsSubmitting(false);
      }
      return;
    }

    // Chặn cấp số nếu chưa có dữ liệu số tiếp theo hợp lệ từ máy chủ cho đúng dãy đang chọn
    if (
      !booksLoaded ||
      previewStatus !== 'ready' ||
      previewSeriesKey !== currentSeriesKey ||
      nextSequenceNumber === null
    ) {
      setSubmitError('Đang kiểm tra số tiếp theo từ máy chủ hoặc chưa lấy được số mới nhất. Vui lòng thử lại.');
      return;
    }

    // Luồng 2: Người dùng có quyền "Cấp số" (issue_number) -> cấp số nguyên tử + tùy chọn tải Word lên Drive
    if (selectedFile) {
      const validationErr = validateWordFile(selectedFile);
      if (validationErr) {
        setFileError(validationErr);
        setSubmitError(validationErr);
        return;
      }
      if (!driveConfigured) {
        setSubmitError(
          `${driveConfig?.lastCheckMessage || 'Thư mục Google Drive chưa được cấu hình hợp lệ.'} Vui lòng bỏ chọn tệp để tiếp tục cấp số hoặc kiểm tra cấu hình Drive của sổ.`
        );
        return;
      }
    }

    if (yearRolloverNotice && !yearRolloverNotice.confirmed) {
      setSubmitError(
        `Hệ thống vừa chuyển sang năm mới ${yearRolloverNotice.toYear} (giờ Việt Nam). Vui lòng xác nhận chuyển năm ở khung thông báo phía trên trước khi cấp số.`
      );
      return;
    }

    setIsSubmitting(true);
    setUploadStep(selectedFile ? 'issuing' : 'idle');
    setRetryError(null);
    setSequenceConflict(null);

    let issuedDoc: DocumentRecord | null = null;
    const fileToUpload = selectedFile;
    const submittingSeriesKey = currentSeriesKey;
    const submittingTypeCode = typeCode.trim().toUpperCase();
    const submittingBookId = activeBook.id;

    try {
      const effectiveSha = overrideSha256 || selectedFileSha256;
      const res = await api.issueDocument({
        bookId: submittingBookId,
        typeCode: submittingTypeCode,
        docName: docName.trim(),
        title: title.trim(),
        signer: signer.trim(),
        signerRole: signerRole.trim(),
        notes: notes.trim(),
        idempotencyKey,
        hasAttachment: Boolean(fileToUpload),
        attachmentName: fileToUpload?.name,
        attachmentSize: fileToUpload?.size,
        fileSha256: effectiveSha,
        conflictAction,
        expectedYear:
          overrideExpected?.expectedYear ??
          (yearRolloverNotice?.confirmed ? yearRolloverNotice.toYear : formYear),
        expectedNextSeq: overrideExpected?.expectedNextSeq ?? nextSequenceNumber,
        confirmUpdatedSequence,
        templateId: !fileToUpload && selectedTemplateId ? selectedTemplateId : undefined,
      });

      issuedDoc = res.document;
      lastServerPreviewDocIdRef.current = issuedDoc.id;

      // Hủy mọi preview cũ đang chạy dở để không bao giờ ghi đè số tiếp theo mới nhất từ giao dịch vừa cấp
      if (previewAbortRef.current) {
        previewAbortRef.current.abort();
      }
      previewReqSeqRef.current += 1;

      setYearRolloverNotice(null);
      setSequenceConflict(null);
      setFormYear(issuedDoc.year);
      setIssuedSuccessDoc(issuedDoc);
      setIdempotencyKey(generateKey());
      if (issuedDoc.templateId && !fileToUpload) {
        setUploadStep(
          issuedDoc.wordGenerationStatus === 'generated' && issuedDoc.driveStatus === 'saved'
            ? 'saved'
            : issuedDoc.wordGenerationStatus === 'failed' || issuedDoc.driveStatus === 'failed'
            ? 'failed'
            : 'idle'
        );
      }

      if (res.nextPreview && activeSeriesKeyRef.current === submittingSeriesKey) {
        setNextNumber(res.nextPreview.previewNumber);
        setNextSequenceNumber(res.nextPreview.sequenceNumber);
        setPreviewStartSequence(
          typeof res.nextPreview.customStartSequence === 'number' &&
            res.nextPreview.customStartSequence >= 1
            ? res.nextPreview.customStartSequence
            : previewStartSequence
        );
        if (Number.isFinite(Number(res.nextPreview.year))) {
          setPreviewYear(Number(res.nextPreview.year));
        }
        setPreviewSeriesKey(submittingSeriesKey);
        setPreviewStatus('ready');
        setPreviewError(null);
        if (res.nextPreview.formattedDate) setTodayVietnam(res.nextPreview.formattedDate);
        if (res.nextPreview.serverYear) {
          onServerTimeUpdate?.(res.nextPreview.serverYear, res.nextPreview.formattedDate);
        }
      } else {
        fetchNextNumber({ silent: false });
      }

      onDocumentIssued(issuedDoc);

      if (fileToUpload) {
        setUploadStep('uploading');
        try {
          const uploadRes = await api.uploadWordToDocument(
            issuedDoc.id,
            fileToUpload,
            issuedDoc.driveFileName,
            'retry'
          );
          setIssuedSuccessDoc(uploadRes.document);
          onDocumentUpdated?.(uploadRes.document);
          setUploadStep('saved');
          setSelectedFile(null);
          setPendingRetryFile(null);
          if (fileInputRef.current) fileInputRef.current.value = '';
        } catch (uploadErr: any) {
          const failedDoc: DocumentRecord = uploadErr.document || {
            ...issuedDoc,
            driveStatus: 'failed',
            driveOriginalFileName: fileToUpload.name,
            driveFileSize: fileToUpload.size,
            driveError: uploadErr.message || 'Lỗi kết nối Google Drive',
          };
          setIssuedSuccessDoc(failedDoc);
          onDocumentUpdated?.(failedDoc);
          setUploadStep('failed');
          setPendingRetryFile(fileToUpload);
          setRetryError(uploadErr.message || `Đã cấp số ${issuedDoc.fullDocNumber} · Chưa lưu tệp lên Drive`);
          setSelectedFile(null);
          if (fileInputRef.current) fileInputRef.current.value = '';
        }
      }
    } catch (err: any) {
      setUploadStep('idle');
      if (err.code === 'SEQUENCE_CHANGED' && typeof err.currentNextSeq === 'number' && err.previewNumber) {
        if (previewAbortRef.current) {
          previewAbortRef.current.abort();
        }
        previewReqSeqRef.current += 1;
        setNextNumber(err.previewNumber);
        setNextSequenceNumber(err.currentNextSeq);
        setPreviewSeriesKey(submittingSeriesKey);
        setPreviewStatus('ready');
        setPreviewError(null);
        setSequenceConflict({
          expectedNextSeq: err.expectedNextSeq ?? nextSequenceNumber,
          currentNextSeq: err.currentNextSeq,
          previewNumber: err.previewNumber,
          message:
            err.message ||
            `Số dự kiến đã thay đổi thành ${err.previewNumber}. Vui lòng xác nhận để cấp theo số mới nhất.`,
        });
        setSubmitError(null);
      } else if (err.code === 'YEAR_CHANGED' && err.currentYear) {
        if (err.previewNumber) setNextNumber(err.previewNumber);
        if (typeof err.sequenceNumber === 'number') setNextSequenceNumber(err.sequenceNumber);
        if (err.formattedDate) setTodayVietnam(err.formattedDate);
        onServerTimeUpdate?.(err.currentYear, err.formattedDate);
        setYearRolloverNotice({
          fromYear: err.previousYear || formYear,
          toYear: err.currentYear,
          newPreviewNumber: err.previewNumber,
          confirmed: false,
        });
        setSubmitError(null);
      } else if (err.code === 'FILENAME_CONFLICT' && err.existingDocument) {
        setFilenameConflict({
          conflictingFileName: err.conflictingFileName || '',
          suggestedSuffixName: err.suggestedSuffixName || '',
          existingDocument: err.existingDocument,
          officialDocNumberPreview: err.officialDocNumberPreview,
        });
      } else {
        setSubmitError(err.message || 'Lỗi kết nối máy chủ khi cấp số. Dữ liệu đã nhập được giữ nguyên.');
      }
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isSubmitting || checkingPreIssue) return;

    if (!canIssue) {
      await performIssueSubmit();
      return;
    }

    setSubmitError(null);
    setFilenameConflict(null);
    setRequestSuccess(null);

    if (!typeCode.trim()) {
      setSubmitError('Vui lòng chọn loại văn bản');
      return;
    }
    if (!docName.trim()) {
      setSubmitError('Vui lòng nhập tên văn bản (thông tin bắt buộc)');
      return;
    }
    if (!signer.trim()) {
      setSubmitError('Vui lòng chọn hoặc nhập người ký văn bản');
      return;
    }
    if (!signerRole.trim()) {
      setSubmitError('Vui lòng nhập chức vụ của người ký');
      return;
    }
    if (
      !booksLoaded ||
      previewStatus !== 'ready' ||
      previewSeriesKey !== currentSeriesKey ||
      nextSequenceNumber === null
    ) {
      setSubmitError('Đang kiểm tra số tiếp theo từ máy chủ hoặc chưa lấy được số mới nhất. Vui lòng thử lại.');
      return;
    }
    if (selectedFile) {
      const validationErr = validateWordFile(selectedFile);
      if (validationErr) {
        setFileError(validationErr);
        setSubmitError(validationErr);
        return;
      }
      if (!driveConfigured) {
        setSubmitError(
          `${driveConfig?.lastCheckMessage || 'Thư mục Google Drive chưa được cấu hình hợp lệ.'} Vui lòng bỏ chọn tệp để tiếp tục cấp số hoặc kiểm tra cấu hình Drive của sổ.`
        );
        return;
      }
    }
    if (yearRolloverNotice && !yearRolloverNotice.confirmed) {
      setSubmitError(
        `Hệ thống vừa chuyển sang năm mới ${yearRolloverNotice.toYear} (giờ Việt Nam). Vui lòng xác nhận chuyển năm ở khung thông báo phía trên trước khi cấp số.`
      );
      return;
    }

    setCheckingPreIssue(true);
    try {
      const sha = selectedFile ? selectedFileSha256 || (await computeFileSha256Hex(selectedFile)) : undefined;
      if (sha && sha !== selectedFileSha256) {
        setSelectedFileSha256(sha);
      }
      const checkRes = await api.preIssueCheck({
        bookId: activeBook.id,
        typeCode: typeCode.trim().toUpperCase(),
        docName: docName.trim(),
        title: title.trim(),
        fileSha256: sha,
      });
      if (checkRes.previewNumber) {
        setNextNumber(checkRes.previewNumber);
        setNextSequenceNumber(checkRes.sequenceNumber);
      }
      setPreIssueModalData({ check: checkRes, fileSha256: sha });
    } catch (err: any) {
      setSubmitError(err?.message || 'Không thể kiểm tra đối chiếu trước khi cấp số. Vui lòng thử lại.');
    } finally {
      setCheckingPreIssue(false);
    }
  };

  const handleRetryTemplateWord = async () => {
    if (!issuedSuccessDoc || isRetryingTemplateWord) return;
    setIsRetryingTemplateWord(true);
    setRetryError(null);
    try {
      const res = await api.retryWordDrive(issuedSuccessDoc.id, issuedSuccessDoc.templateId);
      setIssuedSuccessDoc(res.document);
      onDocumentUpdated?.(res.document);
      setUploadStep('saved');
    } catch (err: any) {
      if (err.document) {
        setIssuedSuccessDoc(err.document);
        onDocumentUpdated?.(err.document);
      }
      setUploadStep('failed');
      setRetryError(err?.message || 'Chưa tạo/lưu lại được tệp Word từ mẫu lên Google Drive.');
    } finally {
      setIsRetryingTemplateWord(false);
    }
  };

  const handleRetryUpload = async (overrideFile?: File) => {
    if (!issuedSuccessDoc || isRetryingUpload) return;
    const fileToUse = overrideFile || pendingRetryFile;
    if (!fileToUse) {
      setRetryError('Vui lòng chọn lại tệp Word (.doc hoặc .docx) để thử lưu lại.');
      return;
    }

    const validationErr = validateWordFile(fileToUse);
    if (validationErr) {
      setRetryError(validationErr);
      return;
    }

    setIsRetryingUpload(true);
    setRetryError(null);
    setUploadStep('uploading');
    try {
      const res = await api.uploadWordToDocument(
        issuedSuccessDoc.id,
        fileToUse,
        issuedSuccessDoc.driveFileName,
        'retry'
      );
      setIssuedSuccessDoc(res.document);
      onDocumentUpdated?.(res.document);
      setUploadStep('saved');
      setPendingRetryFile(null);
    } catch (err: any) {
      const failedDoc: DocumentRecord = err.document || {
        ...issuedSuccessDoc,
        driveStatus: 'failed',
        driveError: err.message,
      };
      setIssuedSuccessDoc(failedDoc);
      onDocumentUpdated?.(failedDoc);
      setUploadStep('failed');
      setRetryError(err.message || `Đã cấp số ${issuedSuccessDoc.fullDocNumber} · Chưa lưu tệp lên Drive`);
    } finally {
      setIsRetryingUpload(false);
    }
  };

  const handleApproveRequest = async (reqId: string) => {
    if (processingReqId) return;
    setProcessingReqId(reqId);
    setQueueError(null);
    try {
      const res = await api.approveNumberRequest(reqId);
      onDocumentIssued(res.document);
      setIssuedSuccessDoc(res.document);
      setUploadStep('idle');
      onRefreshRequests();
    } catch (err: any) {
      setQueueError(err.message || 'Không thể duyệt và cấp số cho yêu cầu này.');
    } finally {
      setProcessingReqId(null);
    }
  };

  const handleRejectRequest = async (reqId: string, reason: string) => {
    if (processingReqId) return;
    setProcessingReqId(reqId);
    setQueueError(null);
    try {
      await api.rejectNumberRequest(reqId, reason);
      setRejectingReqId(null);
      setRejectReasonInput('Chưa đủ thông tin văn bản');
      onRefreshRequests();
    } catch (err: any) {
      setQueueError(err.message || 'Không thể từ chối yêu cầu này.');
    } finally {
      setProcessingReqId(null);
    }
  };

  const handleIssueNext = () => {
    setIssuedSuccessDoc(null);
    setRequestSuccess(null);
    setDocName('');
    setTitle('');
    setNotes('');
    setSubmitError(null);
    setSequenceConflict(null);
    setCopied(false);
    setSelectedFile(null);
    setSelectedFileSha256(undefined);
    setFileError(null);
    setPendingRetryFile(null);
    setRetryError(null);
    setUploadStep('idle');
    setYearRolloverNotice(null);
    if (serverYear) setFormYear(serverYear);
    if (fileInputRef.current) fileInputRef.current.value = '';
    setIdempotencyKey(generateKey());

    if (typeCode) {
      fetchNextNumber({ silent: previewStatus === 'ready' && previewSeriesKey === currentSeriesKey });
    }
  };

  const handleCopyDocNumber = async (docNumber: string) => {
    try {
      await navigator.clipboard.writeText(docNumber);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = docNumber;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    }
  };

  const handleCopyFileName = async (fileName: string, isSuccessCard = false) => {
    if (!fileName) return;
    try {
      await navigator.clipboard.writeText(fileName);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = fileName;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
    if (isSuccessCard) {
      setCopiedSuccessFileName(true);
      setTimeout(() => setCopiedSuccessFileName(false), 2500);
    } else {
      setCopiedFileName(true);
      setTimeout(() => setCopiedFileName(false), 2500);
    }
  };

  const handleOpenSavedWord = async (doc: DocumentRecord) => {
    try {
      const res = await api.openDocumentFile(doc.id);
      if (res.url) window.open(res.url, '_blank', 'noopener,noreferrer');
    } catch {
      if (doc.driveFileUrl) window.open(doc.driveFileUrl, '_blank', 'noopener,noreferrer');
    }
  };

  const openAddTypeDialog = () => {
    setNewTypeName('');
    setNewTypeCode('');
    setNewTypeColor('blue');
    setAddTypeError(null);
    setExistingCatalogType(null);
    setShowAddTypeModal(true);
  };

  const applyAddedOrEnabledDocType = (docType: DocumentTypeConfig, updatedBook?: RegistryBook) => {
    const targetBookId = updatedBook?.id || activeBook.id;
    setRecentlyAddedTypesByBook((prev) => {
      const current = prev[targetBookId] || [];
      if (current.some((t) => t.code.toUpperCase() === docType.code.toUpperCase())) {
        return prev;
      }
      return {
        ...prev,
        [targetBookId]: [...current, docType],
      };
    });
    if (onDocumentTypeAdded) {
      onDocumentTypeAdded(docType, updatedBook);
    }
    setTypeCode(docType.code);
    setShowAddTypeModal(false);
    setNewTypeName('');
    setNewTypeCode('');
    setNewTypeColor('blue');
    setAddTypeError(null);
    setExistingCatalogType(null);
  };

  const handleQuickAddType = async (e: React.FormEvent) => {
    e.preventDefault();
    const normalizedName = newTypeName.trim().replace(/\s+/g, ' ');
    const normalizedCode = newTypeCode.trim().replace(/\s+/g, '').toUpperCase();
    if (!normalizedName || !normalizedCode) {
      setAddTypeError('Vui lòng nhập đầy đủ tên loại văn bản và ký hiệu viết tắt.');
      return;
    }

    setAddTypeSaving(true);
    setAddTypeError(null);
    setExistingCatalogType(null);
    try {
      const res = await api.createDocumentType({
        name: normalizedName,
        code: normalizedCode,
        color: newTypeColor,
        bookId: activeBook.id,
      });
      applyAddedOrEnabledDocType(res.documentType, res.book);
    } catch (err: any) {
      if (err?.code === 'TYPE_EXISTS_IN_CATALOG_NOT_ENABLED' && err.documentType) {
        setExistingCatalogType(err.documentType);
      }
      setAddTypeError(err.message || 'Lỗi khi tạo loại văn bản mới');
    } finally {
      setAddTypeSaving(false);
    }
  };

  const handleEnableExistingCatalogType = async () => {
    if (!existingCatalogType || addTypeSaving) return;
    setAddTypeSaving(true);
    setAddTypeError(null);
    try {
      const res = await api.createDocumentType({
        name: existingCatalogType.name,
        code: existingCatalogType.code,
        color: existingCatalogType.color,
        bookId: activeBook.id,
        enableExistingForBook: true,
      });
      applyAddedOrEnabledDocType(res.documentType, res.book);
    } catch (err: any) {
      setAddTypeError(err.message || 'Không thể bật loại văn bản này cho sổ hiện tại.');
    } finally {
      setAddTypeSaving(false);
    }
  };

  const bookRequests = useMemo(
    () => numberRequests.filter((r) => r.bookId === activeBook.id),
    [numberRequests, activeBook.id]
  );
  const pendingBookRequests = useMemo(
    () => bookRequests.filter((r) => r.status === 'pending'),
    [bookRequests]
  );
  const myRequests = useMemo(
    () => bookRequests.filter((r) => r.requestedByUserId === currentUser.id),
    [bookRequests, currentUser.id]
  );

  const issueAccessibleBooks = useMemo(
    () =>
      accessibleBooks.filter(
        (b) => hasBookPermission(currentUser, b.id, 'issue_number') || hasBookPermission(currentUser, b.id, 'request_number')
      ),
    [accessibleBooks, currentUser]
  );

  return (
    <div className="max-w-4xl mx-auto py-5 sm:py-8 px-3 sm:px-6 space-y-5 sm:space-y-6">
      {/* Banner Intro */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <div className="flex flex-wrap items-center gap-2">
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
              {canIssue ? 'Cấp Số Văn Bản Tự Động' : 'Xin Số Văn Bản'}
            </h1>
            <span className={`inline-flex items-center gap-1.5 text-xs font-semibold px-2.5 py-0.5 rounded-md border ${SEMANTIC_STATUS_STYLES.outgoing.badge}`}>
              <span className={`w-2 h-2 rounded-full ${SEMANTIC_STATUS_STYLES.outgoing.dot}`} />
              <span>Văn bản đi</span>
            </span>
          </div>
          <p className="text-xs sm:text-sm text-slate-600 mt-1 leading-relaxed">
            {canIssue
              ? `Cấp số chính thức vào ${activeBook.name} (${activeBook.unitCode}) theo thứ tự nguyên tử và chống trùng số.`
              : `Gửi yêu cầu xin số vào ${activeBook.name} (${activeBook.unitCode}) để người có quyền Cấp số duyệt và cấp số chính thức.`}
          </p>
        </div>

        <span className="inline-flex items-center gap-1.5 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 shrink-0 self-start sm:self-center">
          <BookOpen className="w-4 h-4 text-blue-700 shrink-0" />
          <span>{activeBook.name} ({activeBook.unitCode})</span>
        </span>
      </div>

      {/* Permission Guard if user has neither request_number nor issue_number on activeBook */}
      {!canIssue && !canRequest && (
        <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl text-xs text-amber-900 flex items-center gap-2.5">
          <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0" />
          <span>
            Tài khoản của bạn chưa được phân quyền <strong>Xin số</strong> hoặc <strong>Cấp số</strong> trong{' '}
            <strong>{activeBook.name}</strong>. Vui lòng chọn sổ khác hoặc liên hệ Quản trị viên.
          </span>
        </div>
      )}

      {/* Role clarity notice for request_number-only users */}
      {!canIssue && canRequest && (
        <div className="p-4 bg-sky-50 border border-sky-200 rounded-xl text-xs text-sky-950 flex items-start gap-2.5">
          <Inbox className="w-5 h-5 text-sky-700 shrink-0 mt-0.5" />
          <div className="space-y-1">
            <strong className="font-bold block">Chế độ tài khoản: Xin số (Tách biệt với Cấp số)</strong>
            <p className="text-sky-900">
              Bạn có quyền <strong>Xin số</strong> trong <strong>{activeBook.name}</strong>. Sau khi gửi biểu mẫu, yêu cầu sẽ chuyển sang người có quyền <strong>Cấp số</strong> để cấp số chính thức (số hiển thị dưới đây là dự kiến tham khảo tại thời điểm hiện tại).
            </p>
          </div>
        </div>
      )}

      {/* Request Submitted Confirmation */}
      {requestSuccess && (
        <div className="p-5 bg-emerald-50 border border-emerald-200 rounded-xl shadow-xs space-y-2 text-xs text-emerald-950">
          <div className="flex items-center gap-2 font-bold text-sm text-emerald-900">
            <CheckCircle2 className="w-5 h-5 text-emerald-600" />
            <span>Đã gửi yêu cầu xin số thành công!</span>
          </div>
          <p>
            Văn bản: <strong>{requestSuccess.docName}</strong> ({requestSuccess.typeName}) · Sổ:{' '}
            <strong>{requestSuccess.bookName || activeBook.name}</strong> · Trạng thái: <strong>Chờ duyệt cấp số</strong>.
          </p>
        </div>
      )}

      {/* Pending Requests Approval Queue (Visible for users with issue_number permission) */}
      {canIssue && pendingBookRequests.length > 0 && (
        <div className="bg-amber-50/70 border border-amber-200 rounded-xl p-5 space-y-3">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-sm font-bold text-amber-950">
              <Inbox className="w-4 h-4 text-amber-700" />
              <span>Yêu cầu xin số chờ duyệt cấp số ({pendingBookRequests.length})</span>
            </div>
            <button
              type="button"
              onClick={onRefreshRequests}
              className="text-xs font-semibold text-amber-900 hover:underline flex items-center gap-1"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Làm mới</span>
            </button>
          </div>

          {queueError && (
            <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg text-xs">
              {queueError}
            </div>
          )}

          <div className="space-y-2.5">
            {pendingBookRequests.map((reqItem) => (
              <div
                key={reqItem.id}
                className="bg-white border border-amber-200 rounded-lg p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs"
              >
                <div className="space-y-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="px-2 py-0.5 rounded bg-sky-100 text-sky-900 font-bold font-mono text-[11px]">
                      {reqItem.typeName} ({reqItem.typeCode})
                    </span>
                    <strong className="text-slate-900 text-sm">{reqItem.docName}</strong>
                  </div>
                  <div className="text-slate-600 flex flex-wrap items-center gap-2 text-[11px]">
                    <span>Người xin số: <strong>{reqItem.requestedBy}</strong></span>
                    <span>·</span>
                    <span>Người ký: <strong>{reqItem.signer}</strong> ({reqItem.signerRole})</span>
                    <span>·</span>
                    <span className="font-mono">{new Date(reqItem.requestedAt).toLocaleString('vi-VN')}</span>
                  </div>
                  {reqItem.title && <p className="text-slate-500 text-[11px]">Trích yếu: {reqItem.title}</p>}
                </div>

                <div className="flex flex-col items-stretch sm:items-end gap-2 shrink-0">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      disabled={processingReqId === reqItem.id}
                      onClick={() => handleApproveRequest(reqItem.id)}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-semibold rounded-lg"
                    >
                      <CheckSquare className="w-3.5 h-3.5" />
                      <span>{processingReqId === reqItem.id ? 'Đang cấp số...' : 'Duyệt & Cấp số'}</span>
                    </button>
                    <button
                      type="button"
                      disabled={processingReqId === reqItem.id}
                      onClick={() => {
                        setRejectingReqId(rejectingReqId === reqItem.id ? null : reqItem.id);
                        setRejectReasonInput('Chưa đủ thông tin văn bản');
                      }}
                      className="inline-flex items-center gap-1 px-3 py-2 bg-white hover:bg-rose-50 border border-rose-300 text-rose-700 font-semibold rounded-lg"
                    >
                      <XCircle className="w-3.5 h-3.5" />
                      <span>Từ chối</span>
                    </button>
                  </div>
                  {rejectingReqId === reqItem.id && (
                    <div className="flex flex-wrap items-center gap-1.5 bg-rose-50 border border-rose-200 rounded-lg p-2">
                      <input
                        type="text"
                        value={rejectReasonInput}
                        onChange={(e) => setRejectReasonInput(e.target.value)}
                        placeholder="Lý do từ chối..."
                        className="px-2.5 py-1 bg-white border border-rose-300 rounded text-xs text-slate-900 min-w-[180px]"
                      />
                      <button
                        type="button"
                        disabled={processingReqId === reqItem.id}
                        onClick={() => handleRejectRequest(reqItem.id, rejectReasonInput)}
                        className="px-2.5 py-1 bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded text-xs"
                      >
                        Xác nhận
                      </button>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Issued Success Banner */}
      {issuedSuccessDoc && (() => {
        const docTypeConfig = documentTypes.find((t) => t.code === issuedSuccessDoc.typeCode);
        const successStyle = getTypeColorStyle(issuedSuccessDoc.typeCode, docTypeConfig?.color);
        const driveFailed = issuedSuccessDoc.driveStatus === 'failed';
        const driveSaved = issuedSuccessDoc.driveStatus === 'saved';

        return (
          <div className={`p-4 sm:p-6 ${successStyle.previewBg} border ${successStyle.previewBorder} rounded-xl shadow-xs transition-all animate-fadeIn`}>
            <div className="flex flex-col sm:flex-row items-start gap-3 sm:gap-4">
              <div className={`w-10 h-10 sm:w-12 sm:h-12 rounded-full ${successStyle.dotColor} flex items-center justify-center text-white shrink-0 shadow-xs`}>
                <CheckCircle2 className="w-6 h-6 sm:w-7 sm:h-7" />
              </div>

              <div className="flex-1 w-full min-w-0">
                <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-1.5 sm:gap-2 mb-1">
                      {driveFailed ? (
                        <span className="text-xs font-bold px-2.5 py-0.5 rounded bg-rose-100 text-rose-900 border border-rose-300">
                          Đã cấp số · Lưu tệp thất bại
                        </span>
                      ) : (
                        <>
                          <span className="text-xs font-bold px-2.5 py-0.5 rounded bg-emerald-100 text-emerald-900 border border-emerald-300">
                            Đã cấp số
                          </span>
                          <span
                            className={`text-xs font-bold px-2.5 py-0.5 rounded border ${
                              driveSaved
                                ? SEMANTIC_STATUS_STYLES.saved.badge
                                : issuedSuccessDoc.driveStatus === 'uploading'
                                ? SEMANTIC_STATUS_STYLES.pending.badge
                                : 'bg-slate-100 text-slate-600 border-slate-200'
                            }`}
                          >
                            {driveSaved
                              ? 'Đã lưu lên Drive'
                              : issuedSuccessDoc.driveStatus === 'uploading'
                              ? 'Đang lưu'
                              : 'Không đính kèm'}
                          </span>
                        </>
                      )}
                      <span
                        className={`text-xs font-bold px-2.5 py-0.5 rounded border ${
                          issuedSuccessDoc.syncStatus === 'synced'
                            ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                            : issuedSuccessDoc.syncStatus === 'error'
                            ? 'bg-rose-50 text-rose-800 border-rose-200'
                            : issuedSuccessDoc.syncStatus === 'pending'
                            ? 'bg-amber-50 text-amber-800 border-amber-200'
                            : 'bg-slate-100 text-slate-600 border-slate-200'
                        }`}
                      >
                        {issuedSuccessDoc.syncStatus === 'synced'
                          ? 'Đã đồng bộ Trang tính'
                          : issuedSuccessDoc.syncStatus === 'error'
                          ? 'Lỗi Trang tính'
                          : issuedSuccessDoc.syncStatus === 'pending'
                          ? 'Đang đồng bộ Trang tính'
                          : 'Trang tính chưa bật'}
                      </span>
                      <span className="text-[11px] font-medium text-slate-600">
                        {issuedSuccessDoc.bookName || activeBook.name}
                      </span>
                    </div>

                    <div className={`text-2xl sm:text-3xl font-extrabold ${successStyle.previewText} font-mono tracking-tight mt-1 flex items-center gap-2.5 sm:gap-3 flex-wrap`}>
                      <span className="break-all">{issuedSuccessDoc.fullDocNumber}</span>
                      <button
                        type="button"
                        onClick={() => handleCopyDocNumber(issuedSuccessDoc.fullDocNumber)}
                        className={`inline-flex items-center gap-1.5 min-h-[36px] px-3 py-1.5 text-xs font-semibold rounded-md bg-white border ${successStyle.badgeBorder} ${successStyle.badgeText} hover:bg-slate-50 transition-colors shadow-2xs`}
                        title="Sao chép số văn bản"
                      >
                        {copied ? (
                          <>
                            <Check className="w-3.5 h-3.5 text-emerald-600" />
                            <span className="text-emerald-700">Đã sao chép!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-3.5 h-3.5" />
                            <span>Sao chép số</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={handleIssueNext}
                      className="flex-1 sm:flex-initial justify-center flex items-center gap-2 min-h-[42px] px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs transition-colors"
                    >
                      <RotateCcw className="w-4 h-4 shrink-0" />
                      <span>Cấp số tiếp theo</span>
                    </button>

                    {canView && (
                      <button
                        type="button"
                        onClick={onNavigateToRegistry}
                        className="flex-1 sm:flex-initial justify-center flex items-center gap-2 min-h-[42px] px-4 py-2 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 rounded-lg shadow-2xs transition-colors"
                      >
                        <span>Xem sổ theo dõi</span>
                        <ArrowRight className="w-4 h-4 text-slate-500 shrink-0" />
                      </button>
                    )}
                  </div>
                </div>

                {/* Record Summary */}
                <div className={`mt-4 pt-4 border-t ${successStyle.previewBorder} grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 text-xs text-slate-700`}>
                  <div className="min-w-0">
                    <span className="text-slate-500 block">Tên văn bản:</span>
                    <p className="font-bold text-slate-900 mt-0.5 break-words">{issuedSuccessDoc.docName || issuedSuccessDoc.title}</p>
                  </div>
                  <div className="min-w-0">
                    <span className="text-slate-500 block">Trích yếu:</span>
                    <p className="font-medium text-slate-700 mt-0.5 line-clamp-2 break-words">
                      {issuedSuccessDoc.title || <span className="italic text-slate-400">Không có</span>}
                    </p>
                  </div>
                  <div className="min-w-0">
                    <span className="text-slate-500 block">Người ký:</span>
                    <p className="font-medium text-slate-900 mt-0.5 break-words">
                      {issuedSuccessDoc.signer} ({issuedSuccessDoc.signerRole})
                    </p>
                  </div>
                  <div>
                    <span className="text-slate-500 block">Ngày ban hành / cấp:</span>
                    <p className="font-medium text-slate-900 mt-0.5 font-mono">{issuedSuccessDoc.issueDate}</p>
                  </div>
                </div>

                {/* Tên tệp chuẩn từ số chính thức thực tế */}
                {(() => {
                  const officialFileName =
                    issuedSuccessDoc.driveFileName ||
                    generateSuggestedFileName(
                      issuedSuccessDoc.fullDocNumber,
                      issuedSuccessDoc.docName || issuedSuccessDoc.title
                    );
                  return (
                    <div className="mt-4 pt-3.5 border-t border-slate-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white/95 p-3.5 rounded-lg border border-slate-200/90 shadow-2xs">
                      <div className="space-y-0.5 min-w-0">
                        <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wide flex items-center gap-1.5">
                          <FileText className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                          <span>Tên tệp sẽ lưu ({issuedSuccessDoc.fullDocNumber}):</span>
                        </span>
                        <p className="font-mono text-xs sm:text-sm font-bold text-slate-900 break-all sm:break-words">
                          {officialFileName}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleCopyFileName(officialFileName, true)}
                        className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-3.5 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-md shadow-2xs transition-colors shrink-0 w-full sm:w-auto"
                      >
                        {copiedSuccessFileName ? (
                          <>
                            <Check className="w-3.5 h-3.5" />
                            <span>Đã sao chép!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-3.5 h-3.5" />
                            <span>Sao chép tên file</span>
                          </>
                        )}
                      </button>
                    </div>
                  );
                })()}

                {/* Trạng thái lưu tệp Word trên Google Drive */}
                {(uploadStep !== 'idle' || driveSaved || driveFailed || issuedSuccessDoc.driveStatus === 'uploading') && (
                  <div className="mt-3 space-y-2.5">
                    {driveSaved && (
                      <div className="p-3.5 bg-emerald-50 border border-emerald-200 rounded-lg flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div className="min-w-0 space-y-0.5">
                          <div className="text-xs font-bold text-emerald-900 flex items-center gap-1.5">
                            <FolderCheck className="w-4 h-4 text-emerald-600 shrink-0" />
                            <span>Đã đổi tên và lưu tệp lên Google Drive</span>
                          </div>
                          <p className="font-mono text-xs font-semibold text-emerald-800 break-all">
                            {issuedSuccessDoc.driveFileName}
                          </p>
                          {issuedSuccessDoc.driveFileSize && (
                            <span className="text-[11px] text-emerald-700 block">
                              Dung lượng: {formatFileSize(issuedSuccessDoc.driveFileSize)}
                            </span>
                          )}
                        </div>
                        {issuedSuccessDoc.driveFileUrl && (
                          <button
                            type="button"
                            onClick={() => handleOpenSavedWord(issuedSuccessDoc)}
                            className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 text-xs font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg shadow-2xs transition-colors shrink-0"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                            <span>Mở tệp Word</span>
                          </button>
                        )}
                      </div>
                    )}

                    {driveFailed && !isRetryingUpload && (
                      <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-lg space-y-2.5 text-xs">
                        <div className="flex items-start gap-2 text-rose-900">
                          <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                          <div className="space-y-0.5">
                            <strong className="block font-bold">
                              Đã cấp số · Lưu tệp thất bại ({issuedSuccessDoc.fullDocNumber})
                            </strong>
                            <p className="text-rose-700">
                              {retryError || issuedSuccessDoc.driveError || 'Lỗi kết nối hoặc phân quyền thư mục Google Drive.'}
                            </p>
                            {pendingRetryFile && (
                              <p className="text-[11px] text-rose-800">
                                Tệp đang chờ: <strong className="font-mono">{pendingRetryFile.name}</strong> ({formatFileSize(pendingRetryFile.size)}) &rarr; Tên mới:{' '}
                                <strong className="font-mono">
                                  {buildStandardDriveFileName(
                                    issuedSuccessDoc.fullDocNumber,
                                    issuedSuccessDoc.docName || issuedSuccessDoc.title,
                                    pendingRetryFile.name
                                  )}
                                </strong>
                              </p>
                            )}
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-2 pt-1">
                          {issuedSuccessDoc.templateId && !pendingRetryFile ? (
                            <button
                              type="button"
                              onClick={handleRetryTemplateWord}
                              disabled={isRetryingTemplateWord}
                              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-md shadow-2xs transition-colors"
                            >
                              <RefreshCw className={`w-3.5 h-3.5 ${isRetryingTemplateWord ? 'animate-spin' : ''}`} />
                              <span>{isRetryingTemplateWord ? 'Đang tạo & lưu lại...' : 'Thử tạo & lưu lại Word'}</span>
                            </button>
                          ) : (
                            <>
                              <input
                                ref={retryFileInputRef}
                                type="file"
                                accept=".doc,.docx"
                                className="hidden"
                                onChange={(e) => {
                                  const f = e.target.files?.[0];
                                  if (!f) return;
                                  const err = validateWordFile(f);
                                  if (err) {
                                    setRetryError(err);
                                    return;
                                  }
                                  setRetryError(null);
                                  setPendingRetryFile(f);
                                }}
                              />
                              <button
                                type="button"
                                onClick={() => retryFileInputRef.current?.click()}
                                className="px-3 py-1.5 bg-white hover:bg-slate-50 border border-rose-300 text-rose-800 font-semibold rounded-md transition-colors"
                              >
                                Chọn lại tệp khác
                              </button>
                              <button
                                type="button"
                                onClick={() => handleRetryUpload()}
                                disabled={isRetryingUpload || !pendingRetryFile}
                                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-md shadow-2xs transition-colors"
                              >
                                <RefreshCw className="w-3.5 h-3.5" />
                                <span>Thử lưu lại</span>
                              </button>
                            </>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        );
      })()}

      {/* Main Form (Issue or Request depending on permission) */}
      {(canIssue || canRequest) && (
        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-8 shadow-xs">
          <form onSubmit={handleSubmit} className="space-y-5 sm:space-y-6">
            {/* Top row: Loại văn bản & Số dự kiến */}
            {(() => {
              const selectedDocType = activeTypes.find((t) => t.code === typeCode);
              const currentStyle = getTypeColorStyle(typeCode, selectedDocType?.color);
              const isPreviewReady =
                booksLoaded &&
                previewStatus === 'ready' &&
                previewSeriesKey === currentSeriesKey &&
                nextSequenceNumber !== null &&
                Boolean(nextNumber);
              const effectivePreviewStatus: 'loading' | 'ready' | 'error' = !booksLoaded
                ? appSyncError
                  ? 'error'
                  : 'loading'
                : isPreviewReady
                ? 'ready'
                : previewStatus === 'error'
                ? 'error'
                : 'loading';

              return (
                <div className="grid grid-cols-1 md:grid-cols-12 gap-4 sm:gap-6 items-start">
                  <div className="md:col-span-7 space-y-2 min-w-0">
                    <div className="flex items-center justify-between gap-2 flex-wrap">
                      <label htmlFor="doc-type" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                        Loại văn bản <span className="text-rose-600">*</span>
                      </label>
                      {canManageDocTypes && (
                        <button
                          type="button"
                          onClick={openAddTypeDialog}
                          className="text-xs font-semibold text-sky-700 hover:text-sky-900 transition-colors flex items-center gap-1 min-h-[32px]"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Thêm loại văn bản</span>
                        </button>
                      )}
                    </div>

                    <select
                      id="doc-type"
                      value={typeCode}
                      onChange={(e) => {
                        const val = e.target.value;
                        if (val === '__add_new_type__') {
                          openAddTypeDialog();
                          return;
                        }
                        setTypeCode(val);
                      }}
                      className="w-full min-h-[44px] h-11 px-3 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors"
                      disabled={isSubmitting}
                    >
                      {activeTypes.map((type) => (
                        <option key={type.code} value={type.code}>
                          {type.name} ({type.code})
                        </option>
                      ))}
                      {canManageDocTypes && (
                        <option value="__add_new_type__">+ Thêm loại văn bản...</option>
                      )}
                    </select>
                    <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                      <span className="text-xs text-slate-500 break-words">
                        Sổ đang chọn: <strong className="text-slate-800">{activeBook.name}</strong> · Ký hiệu:{' '}
                        <strong className="text-slate-800">{activeBook.unitCode}</strong>
                      </span>
                      <span className={`inline-flex items-center gap-1.5 text-[11px] font-semibold px-2 py-0.5 rounded border ${SEMANTIC_STATUS_STYLES.outgoing.badge}`}>
                        Văn bản đi
                      </span>
                    </div>
                  </div>

                  {/* Số dự kiến tiếp theo */}
                  <div className={`md:col-span-5 ${currentStyle.previewBg} border ${currentStyle.previewBorder} rounded-xl p-4 flex flex-col justify-between shadow-xs transition-all min-w-0`}>
                    <div className="flex items-center justify-between text-xs mb-1.5 gap-1.5 flex-wrap">
                      <span className={`font-bold uppercase tracking-wider ${currentStyle.previewText} flex items-center gap-1.5`}>
                        <span className={`w-2 h-2 rounded-full shrink-0 ${currentStyle.dotColor}`} />
                        <span>{canIssue ? 'Số dự kiến tiếp theo' : 'Số tham khảo hiện tại'}</span>
                      </span>
                      <span className={`text-[11px] font-semibold px-2 py-0.5 rounded-md border ${currentStyle.previewTag}`}>
                        {selectedDocType?.name || typeCode}
                      </span>
                    </div>
                    <div className={`text-2xl sm:text-3xl font-black font-mono ${currentStyle.previewText} tracking-tight flex items-center gap-2 py-1 break-all min-h-[44px]`}>
                      {effectivePreviewStatus === 'loading' ? (
                        <span className="text-xs font-sans font-semibold text-slate-600 inline-flex items-center gap-2">
                          <RefreshCw className="w-4 h-4 animate-spin text-sky-700 shrink-0" />
                          <span>Đang tải mốc số đã cài… · Đang kiểm tra số tiếp theo…</span>
                        </span>
                      ) : effectivePreviewStatus === 'error' ? (
                        <div className="flex flex-wrap items-center justify-between gap-2 w-full">
                          <span className="text-xs font-sans font-bold text-rose-700 inline-flex items-center gap-1.5">
                            <AlertCircle className="w-4 h-4 shrink-0" />
                            <span>Chưa tải được cấu hình số</span>
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              if (!booksLoaded && onRetryAppLoad) {
                                onRetryAppLoad();
                              } else {
                                fetchNextNumber({ silent: false });
                              }
                            }}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-sans font-semibold bg-white hover:bg-rose-50 text-rose-800 border border-rose-300 rounded-md shadow-2xs transition-colors"
                          >
                            <RefreshCw className="w-3.5 h-3.5" />
                            <span>Thử lại</span>
                          </button>
                        </div>
                      ) : (
                        <span>{nextNumber}</span>
                      )}
                    </div>
                    {effectivePreviewStatus === 'error' && previewError && (
                      <p className="text-[11px] text-rose-700 mb-1 break-words">{previewError}</p>
                    )}
                    <div className="flex items-center justify-between mt-1 text-[11px] text-slate-600 gap-1.5 flex-wrap">
                      <span className="font-semibold text-slate-700">
                        {activeBook.name} · {activeBook.numberMode === 'shared' ? 'Dãy số chung' : `Dãy riêng ${typeCode}`} · Năm {previewYear || (yearRolloverNotice ? yearRolloverNotice.toYear : formYear)}
                      </span>
                      {effectivePreviewStatus === 'ready' ? (
                        previewStartSequence !== null ? (
                          <span className="font-semibold text-sky-800 bg-white/90 px-1.5 py-0.5 rounded border border-sky-200">
                            Số bắt đầu đã cài: {String(previewStartSequence).padStart(2, '0')}
                            {nextSequenceNumber !== null && nextSequenceNumber > previewStartSequence
                              ? ` · Đang nối tiếp`
                              : ''}
                          </span>
                        ) : (
                          <span className="font-medium text-slate-700 bg-white/90 px-1.5 py-0.5 rounded border border-slate-200">
                            Chưa thiết lập · Mặc định 01
                          </span>
                        )
                      ) : null}
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Tên văn bản */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between gap-2">
                <label htmlFor="doc-name" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Tên văn bản <span className="text-rose-600">*</span>
                </label>
                <span className="text-[11px] text-rose-600 font-semibold bg-rose-50 px-2 py-0.5 rounded border border-rose-200 shrink-0">
                  Bắt buộc
                </span>
              </div>
              <input
                id="doc-name"
                type="text"
                value={docName}
                onChange={(e) => setDocName(e.target.value)}
                placeholder="Ví dụ: Kế hoạch tổ chức hoạt động chào mừng ngày 20-11..."
                required
                className="w-full min-h-[44px] h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors placeholder:text-slate-400"
                disabled={isSubmitting}
              />

              {/* Gợi ý tên file chuẩn */}
              {(() => {
                const previewDocNumber =
                  booksLoaded &&
                  previewStatus === 'ready' &&
                  previewSeriesKey === currentSeriesKey &&
                  nextNumber
                    ? nextNumber
                    : '';
                const suggestedFileName = generateSuggestedFileName(previewDocNumber, docName);

                return (
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 sm:p-4 space-y-2 mt-2">
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800 uppercase tracking-wide">
                        <FileText className="w-4 h-4 text-sky-600 shrink-0" />
                        <span>Tên tệp dự kiến (Số ký hiệu - Tên văn bản)</span>
                      </div>
                      <span className="text-[11px] font-medium text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200">
                        Đổi <strong className="text-sky-700">/</strong> thành <strong className="text-sky-700">.</strong> · Giữ dấu tiếng Việt
                      </span>
                    </div>

                    <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 min-w-0">
                      <div className="flex-1 min-w-0 px-3 py-2 bg-white border border-slate-300 rounded-lg font-mono text-xs sm:text-sm font-bold text-slate-900 select-all break-all sm:break-words sm:overflow-x-auto shadow-2xs flex items-center min-h-[42px]">
                        {suggestedFileName ? (
                          <span className="text-sky-950 font-bold">{suggestedFileName}</span>
                        ) : (
                          <span className="text-slate-400 font-normal italic font-sans text-xs">
                            {previewDocNumber ? `${previewDocNumber.replace(/\//g, '.')} - [Tên văn bản]` : 'Nhập tên văn bản để xem tên tệp dự kiến...'}
                          </span>
                        )}
                      </div>

                      <button
                        type="button"
                        onClick={() => handleCopyFileName(suggestedFileName)}
                        disabled={!suggestedFileName}
                        className="inline-flex items-center justify-center gap-1.5 min-h-[42px] px-4 py-2 text-xs font-semibold bg-white hover:bg-sky-50 text-slate-700 hover:text-sky-700 border border-slate-300 hover:border-sky-300 rounded-lg shadow-2xs transition-colors shrink-0 disabled:opacity-50 w-full sm:w-auto"
                      >
                        {copiedFileName ? (
                          <>
                            <Check className="w-4 h-4 text-emerald-600" />
                            <span className="text-emerald-700 font-bold">Đã sao chép!</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-4 h-4" />
                            <span>Sao chép tên file</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                );
              })()}
            </div>

            {/* Trích yếu */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between gap-2">
                <label htmlFor="doc-title" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Trích yếu / Nội dung văn bản
                </label>
                <span className="text-[11px] text-slate-400 font-normal shrink-0">Không bắt buộc</span>
              </div>
              <textarea
                id="doc-title"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="Tóm tắt mục đích, nội dung chính của văn bản (nếu có)..."
                rows={3}
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors placeholder:text-slate-400"
                disabled={isSubmitting}
              />
            </div>

            {/* Signer & Role */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 sm:gap-6">
              <div className="space-y-1.5 min-w-0">
                <label htmlFor="signer-select" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Người ký <span className="text-rose-600">*</span>
                </label>

                {!isCustomSigner ? (
                  <select
                    id="signer-select"
                    value={signer}
                    onChange={handleSignerSelect}
                    className="w-full min-h-[44px] h-11 px-3 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors"
                    disabled={isSubmitting}
                  >
                    {effectiveSigners.map((s) => (
                      <option key={s.id} value={s.name}>
                        {s.name} - {s.role}
                      </option>
                    ))}
                    <option value="__custom__">+ Nhập tên người ký khác...</option>
                  </select>
                ) : (
                  <div className="space-y-2">
                    <input
                      type="text"
                      value={signer}
                      onChange={(e) => setSigner(e.target.value)}
                      placeholder="Họ và tên người ký"
                      required
                      className="w-full min-h-[44px] h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                      disabled={isSubmitting}
                    />
                    <button
                      type="button"
                      onClick={() => {
                        setIsCustomSigner(false);
                        if (effectiveSigners.length > 0) {
                          setSigner(effectiveSigners[0].name);
                          setSignerRole(effectiveSigners[0].role);
                        }
                      }}
                      className="text-xs text-sky-700 hover:underline font-medium min-h-[32px]"
                    >
                      Chọn lại từ danh sách có sẵn
                    </button>
                  </div>
                )}
              </div>

              <div className="space-y-1.5 min-w-0">
                <label htmlFor="signer-role" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Chức vụ người ký <span className="text-rose-600">*</span>
                </label>
                <input
                  id="signer-role"
                  type="text"
                  value={signerRole}
                  onChange={(e) => setSignerRole(e.target.value)}
                  placeholder="Ví dụ: Bí thư Đoàn trường / Phó Bí thư..."
                  required
                  className="w-full min-h-[44px] h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  disabled={isSubmitting}
                />
              </div>
            </div>

            {/* Automatic Issuer & Date Info */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 sm:gap-6 p-3.5 sm:p-4 bg-slate-50 border border-slate-200 rounded-lg">
              <div className="flex items-center gap-3 min-w-0">
                <div className="w-9 h-9 rounded-md bg-slate-200 flex items-center justify-center text-slate-700 shrink-0">
                  <Shield className="w-5 h-5 text-slate-600" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-xs text-slate-500 font-medium">
                    {canIssue ? 'Người lấy số (Tự động)' : 'Người xin số (Tự động)'}
                  </span>
                  <span className="text-sm font-semibold text-slate-900 break-words">
                    {getUserDisplayName(currentUser)} <span className="text-slate-400 font-normal">({currentUser.username})</span>
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-3 min-w-0">
                <div className="w-9 h-9 rounded-md bg-slate-200 flex items-center justify-center text-slate-700 shrink-0">
                  <Clock className="w-5 h-5 text-slate-600" />
                </div>
                <div className="flex flex-col min-w-0">
                  <span className="text-xs text-slate-500 font-medium">Ngày cấp số (Múi giờ Việt Nam)</span>
                  <span className="text-sm font-semibold font-mono text-slate-900">
                    {todayVietnam} <span className="text-xs font-normal text-slate-500">(dd/mm/yyyy)</span>
                  </span>
                </div>
              </div>
            </div>

            {/* Notes */}
            <div className="space-y-1.5">
              <label htmlFor="doc-notes" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                Ghi chú <span className="text-slate-400 font-normal lowercase">(không bắt buộc)</span>
              </label>
              <input
                id="doc-notes"
                type="text"
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
                placeholder="Ghi chú thêm (nơi nhận, số lượng bản phát hành, lưu ý...)"
                className="w-full min-h-[44px] h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 placeholder:text-slate-400"
                disabled={isSubmitting}
              />
            </div>

            {/* Word File Attachment (Optional — available for issue_number users) */}
            {canIssue && (() => {
              const previewDocNumber =
                booksLoaded &&
                previewStatus === 'ready' &&
                previewSeriesKey === currentSeriesKey &&
                nextNumber
                  ? nextNumber
                  : '';
              const expectedDriveName = selectedFile
                ? buildStandardDriveFileName(previewDocNumber, docName, selectedFile.name)
                : '';

              return (
                <div className="space-y-2.5 p-3.5 sm:p-4 bg-slate-50/80 border border-slate-200 rounded-xl">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <FileUp className="w-4 h-4 text-sky-700 shrink-0" />
                      <span className="text-xs font-bold text-slate-800 uppercase tracking-wide">
                        Tệp Word đính kèm
                      </span>
                      <span className="text-[11px] text-slate-500 font-normal">(không bắt buộc)</span>
                    </div>
                    <span className="text-[11px] font-medium text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200">
                      .doc, .docx · Tối đa {maxWordFileSizeMB} MB
                    </span>
                  </div>

                  <input
                    ref={fileInputRef}
                    id="word-attachment-input"
                    type="file"
                    accept=".doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    onChange={handleFileChange}
                    disabled={isSubmitting}
                    className="hidden"
                  />

                  {!selectedFile ? (
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white border border-dashed border-slate-300 rounded-lg p-3.5">
                      <div className="text-xs text-slate-600 space-y-0.5 min-w-0">
                        <p className="font-medium text-slate-800 leading-relaxed">
                          Đính kèm tệp Word để tự động đổi tên theo số chính thức và lưu vào Google Drive của {activeBook.name}
                        </p>
                        <p className="text-[11px] text-slate-500">
                          Nếu chưa có tệp, bạn vẫn có thể bấm <strong>Cấp số và lưu</strong> bình thường (kể cả khi chưa cấu hình Drive).
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        disabled={isSubmitting}
                        className="inline-flex items-center justify-center gap-1.5 min-h-[44px] px-4 py-2 text-xs font-semibold bg-white hover:bg-sky-50 text-sky-800 border border-sky-300 rounded-lg shadow-2xs transition-colors shrink-0 disabled:opacity-50 w-full sm:w-auto"
                      >
                        <Upload className="w-4 h-4" />
                        <span>Chọn tệp Word</span>
                      </button>
                    </div>
                  ) : (
                    <div className="bg-white border border-sky-200 rounded-lg p-3.5 space-y-2.5">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
                        <div className="flex items-center gap-2.5 min-w-0">
                          <div className="w-9 h-9 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center text-sky-700 shrink-0">
                            <FileText className="w-4 h-4" />
                          </div>
                          <div className="min-w-0">
                            <p className="text-xs font-bold text-slate-900 break-all sm:truncate">{selectedFile.name}</p>
                            <p className="text-[11px] text-slate-500">
                              Dung lượng: <strong className="text-slate-700">{formatFileSize(selectedFile.size)}</strong>
                            </p>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0">
                          <button
                            type="button"
                            onClick={() => fileInputRef.current?.click()}
                            disabled={isSubmitting}
                            className="flex-1 sm:flex-initial min-h-[38px] px-3 py-1.5 text-xs font-semibold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded-md disabled:opacity-50"
                          >
                            Chọn lại
                          </button>
                          <button
                            type="button"
                            onClick={handleClearSelectedFile}
                            disabled={isSubmitting}
                            className="flex-1 sm:flex-initial min-h-[38px] px-3 py-1.5 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-md disabled:opacity-50"
                          >
                            Bỏ chọn
                          </button>
                        </div>
                      </div>

                      <div className="bg-slate-50 border border-slate-200 rounded-md px-3 py-2 text-xs space-y-0.5 min-w-0">
                        <span className="text-[11px] font-semibold text-slate-500 block">
                          Tên tệp sẽ lưu (dự kiến — máy chủ tự tính lại theo số chính thức):
                        </span>
                        <span className="font-mono font-bold text-sky-900 break-all block">{expectedDriveName}</span>
                      </div>

                      {!driveConfigured && (
                        <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-amber-900 text-xs flex items-start gap-2">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                          <div className="space-y-1 min-w-0">
                            <strong className="block font-bold">Thư mục Google Drive chưa được cấu hình hợp lệ</strong>
                            <p className="text-amber-800 break-words">
                              {driveConfig?.lastCheckMessage || 'Chưa thiết lập thư mục Google Drive cho sổ này.'}
                            </p>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {fileError && (
                    <div className="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 text-xs flex items-center gap-2">
                      <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                      <span className="break-words min-w-0">{fileError}</span>
                    </div>
                  )}

                  {/* Mẫu Word tự động (khi không tải tệp Word thủ công) */}
                  {!selectedFile && matchingTemplates.length > 0 && (
                    <div className="mt-2 pt-3 border-t border-slate-200 space-y-2.5">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                          <FileCheck2 className="w-4 h-4 text-sky-700 shrink-0" />
                          <span>Hoặc tạo tự động từ Mẫu Word ({matchingTemplates.length} mẫu)</span>
                        </span>
                        <span className="text-[11px] text-slate-500">
                          Tự điền số chính thức &amp; lưu vào Google Drive khi cấp số
                        </span>
                      </div>

                      <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                        <select
                          aria-label="Chọn mẫu Word"
                          value={selectedTemplateId}
                          onChange={(e) => setSelectedTemplateId(e.target.value)}
                          disabled={isSubmitting}
                          className="flex-1 min-h-[40px] px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-900"
                        >
                          <option value="">— Không dùng mẫu Word —</option>
                          {matchingTemplates.map((tpl) => (
                            <option key={tpl.id} value={tpl.id}>
                              {tpl.name} ({tpl.originalFileName})
                            </option>
                          ))}
                        </select>

                        {selectedTemplateId && (
                          <button
                            type="button"
                            onClick={handleDownloadTemplateDraft}
                            disabled={downloadingDraft}
                            className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-3.5 py-2 text-xs font-semibold bg-white hover:bg-sky-50 text-sky-800 border border-sky-300 rounded-lg shrink-0 disabled:opacity-50"
                          >
                            <Download className="w-3.5 h-3.5 shrink-0" />
                            <span>{downloadingDraft ? 'Đang tạo nháp...' : 'Tải bản nháp (DỰ THẢO)'}</span>
                          </button>
                        )}
                      </div>

                      {selectedTemplateId && templateDraftCheck && (
                        <div className="text-[11px] bg-white border border-slate-200 rounded-lg p-2.5 space-y-1">
                          {templateDraftCheck.missingFields.length > 0 ? (
                            <p className="text-amber-800 font-medium">
                              Trường còn trống trong biểu mẫu:{' '}
                              <span className="font-mono">{templateDraftCheck.missingFields.join(', ')}</span>
                            </p>
                          ) : (
                            <p className="text-emerald-700 font-medium">
                              Đã đủ dữ liệu điền vào các trường của mẫu Word.
                            </p>
                          )}
                          {!driveConfigured && (
                            <p className="text-amber-700">
                              Lưu ý: Drive của sổ chưa cấu hình xong — số vẫn được cấp, trạng thái tệp sẽ là chưa lưu Drive để thử lại sau.
                            </p>
                          )}
                        </div>
                      )}

                      {draftError && (
                        <p className="text-xs text-rose-700 font-medium">{draftError}</p>
                      )}
                    </div>
                  )}
                </div>
              );
            })()}

            {/* Filename Conflict Banner */}
            {filenameConflict && (
              <div className="p-4 bg-amber-50 border border-amber-300 rounded-xl space-y-2.5 text-xs text-amber-950">
                <div className="flex items-start gap-2">
                  <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                  <div>
                    <strong className="font-bold block">
                      Trùng tên tệp dự kiến trên Drive: {filenameConflict.conflictingFileName}
                    </strong>
                    <p className="text-amber-900 mt-0.5">
                      Hồ sơ <strong className="font-mono">{filenameConflict.existingDocument.fullDocNumber}</strong> đã sử dụng tên tệp này.
                    </p>
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => performIssueSubmit('suffix')}
                    className="px-3 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-md"
                  >
                    Thêm hậu tố ({filenameConflict.suggestedSuffixName})
                  </button>
                  <button
                    type="button"
                    onClick={() => setFilenameConflict(null)}
                    className="px-3 py-1.5 bg-white hover:bg-slate-50 border border-amber-300 text-amber-900 font-semibold rounded-md"
                  >
                    Sửa lại Tên văn bản
                  </button>
                  {onOpenExistingDocument && canView && (
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

            {/* Sequence Changed Conflict Banner */}
            {sequenceConflict && (
              <div className="p-4 bg-amber-50 border-2 border-amber-400 rounded-xl space-y-2.5 text-xs text-amber-950 shadow-xs">
                <div className="flex items-start gap-2.5">
                  <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <strong className="font-bold text-sm block text-amber-950">
                      Số dự kiến vừa được cập nhật thành {sequenceConflict.previewNumber}
                    </strong>
                    <p className="text-amber-900">
                      {sequenceConflict.message} Toàn bộ nội dung văn bản và tệp Word bạn đang chọn vẫn được giữ nguyên.
                    </p>
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2 pt-1">
                  <button
                    type="button"
                    disabled={isSubmitting}
                    onClick={() => performIssueSubmit(undefined, true)}
                    className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-2xs transition-colors disabled:opacity-50"
                  >
                    Xác nhận cấp theo số mới ({sequenceConflict.previewNumber})
                  </button>
                  <button
                    type="button"
                    disabled={isSubmitting}
                    onClick={() => setSequenceConflict(null)}
                    className="px-3.5 py-2 bg-white hover:bg-slate-50 border border-amber-300 text-amber-900 font-semibold rounded-lg"
                  >
                    Để tôi kiểm tra lại
                  </button>
                </div>
              </div>
            )}

            {/* Year Rollover Confirmation Banner (00:00 01/01 Asia/Ho_Chi_Minh) */}
            {yearRolloverNotice && (
              <div className="p-4 bg-amber-50 border-2 border-amber-400 rounded-xl space-y-3 text-xs text-amber-950 shadow-xs">
                <div className="flex items-start gap-2.5">
                  <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <strong className="font-bold text-sm block text-amber-950">
                      Hệ thống vừa bước sang năm mới {yearRolloverNotice.toYear} (giờ Việt Nam)
                    </strong>
                    <p className="text-amber-900">
                      Biểu mẫu này được mở từ năm <strong>{yearRolloverNotice.fromYear}</strong>. Toàn bộ nội dung bạn đang nhập và tệp Word đã chọn vẫn được giữ nguyên. Số văn bản dự kiến theo dãy năm <strong>{yearRolloverNotice.toYear}</strong> đã được cập nhật thành{' '}
                      <strong className="font-mono text-sky-900">{yearRolloverNotice.newPreviewNumber || nextNumber || '---'}</strong>.
                    </p>
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2 pt-1">
                  {!yearRolloverNotice.confirmed ? (
                    <button
                      type="button"
                      onClick={() => {
                        setFormYear(yearRolloverNotice.toYear);
                        setYearRolloverNotice({ ...yearRolloverNotice, confirmed: true });
                        setSubmitError(null);
                      }}
                      className="px-4 py-2 bg-amber-700 hover:bg-amber-800 text-white font-semibold rounded-lg shadow-2xs transition-colors"
                    >
                      Xác nhận chuyển sang cấp số năm {yearRolloverNotice.toYear}
                    </button>
                  ) : (
                    <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-100 text-emerald-900 border border-emerald-300 rounded-lg font-semibold">
                      <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                      <span>Đã xác nhận cấp số theo năm {yearRolloverNotice.toYear}. Bấm "Cấp số và lưu" để hoàn tất.</span>
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* Submit Error Callout */}
            {submitError && (
              <div className="p-4 bg-rose-50 border border-rose-200 rounded-lg text-rose-800 text-sm flex items-start gap-3">
                <AlertCircle className="w-5 h-5 shrink-0 text-rose-600 mt-0.5" />
                <div className="flex-1">
                  <strong className="block font-semibold">Chưa thể thực hiện:</strong>
                  <span>{submitError}</span>
                </div>
              </div>
            )}

            {/* Submit Button & Active Book Confirm */}
            <div className="pt-3 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 sm:gap-4 border-t border-slate-100">
              <div className="text-xs text-slate-500 flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-sky-600 shrink-0" />
                <span>
                  {canIssue
                    ? 'Giao dịch máy chủ nguyên tử · Bảo vệ chống nhấn đúp · Lưu trực tuyến lâu dài'
                    : 'Yêu cầu xin số được ghi nhận tức thì và chuyển tới người có quyền Cấp số'}
                </span>
              </div>

              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5">
                <span className="inline-flex items-center justify-center gap-1.5 min-h-[44px] h-11 sm:h-12 px-3 bg-slate-100 border border-slate-200 rounded-lg text-xs font-bold text-slate-700">
                  <BookOpen className="w-3.5 h-3.5 text-blue-700 shrink-0" />
                  <span className="truncate">{activeBook.name} ({activeBook.unitCode})</span>
                </span>

                <button
                  type="submit"
                  disabled={
                    isSubmitting ||
                    checkingPreIssue ||
                    !booksLoaded ||
                    (canIssue &&
                      (previewStatus !== 'ready' ||
                        previewSeriesKey !== currentSeriesKey ||
                        nextSequenceNumber === null ||
                        Boolean(selectedFile && !driveConfigured))) ||
                    Boolean(yearRolloverNotice && !yearRolloverNotice.confirmed)
                  }
                  className="w-full sm:w-auto sm:min-w-[210px] min-h-[48px] h-12 px-6 bg-sky-700 hover:bg-sky-800 active:bg-sky-900 text-white font-semibold text-base sm:text-sm rounded-lg shadow-sm transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {checkingPreIssue ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Đang kiểm tra trước cấp số...</span>
                    </>
                  ) : isSubmitting ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                      <span>
                        {!canIssue
                          ? 'Đang gửi yêu cầu xin số...'
                          : uploadStep === 'uploading'
                          ? 'Đang lưu tệp lên Drive...'
                          : 'Đang cấp số tại máy chủ...'}
                      </span>
                    </>
                  ) : (
                    <>
                      <Send className="w-4 h-4" />
                      <span>{canIssue ? 'Cấp số và lưu' : 'Gửi yêu cầu xin số'}</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          </form>
        </div>
      )}

      {/* Request History for request_number-only users */}
      {!canIssue && canRequest && myRequests.length > 0 && (
        <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs space-y-3">
          <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
            <Inbox className="w-4 h-4 text-sky-700" />
            <span>Yêu cầu xin số của bạn tại {activeBook.name} ({myRequests.length})</span>
          </h3>
          <div className="divide-y divide-slate-100 text-xs">
            {myRequests.map((r) => (
              <div key={r.id} className="py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="space-y-0.5">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-semibold text-slate-900">{r.docName}</span>
                    <span className="px-1.5 py-0.5 bg-slate-100 rounded text-[11px] font-mono">
                      {r.typeCode}
                    </span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Gửi lúc {new Date(r.requestedAt).toLocaleString('vi-VN')} · Ký: {r.signer}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  {r.status === 'pending' && (
                    <span className="px-2.5 py-1 rounded bg-amber-50 text-amber-800 border border-amber-200 font-semibold">
                      Chờ duyệt cấp số
                    </span>
                  )}
                  {r.status === 'issued' && (
                    <>
                      <span className="px-2.5 py-1 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold font-mono">
                        Đã cấp: {r.issuedDocNumber}
                      </span>
                      {canView && r.issuedDocId && onOpenExistingDocument && (
                        <button
                          type="button"
                          onClick={() => onOpenExistingDocument(r.issuedDocId!)}
                          className="px-2.5 py-1 rounded bg-sky-50 text-sky-800 border border-sky-200 font-semibold hover:bg-sky-100"
                        >
                          Mở hồ sơ
                        </button>
                      )}
                    </>
                  )}
                  {r.status === 'rejected' && (
                    <span className="px-2.5 py-1 rounded bg-rose-50 text-rose-800 border border-rose-200 font-semibold">
                      Từ chối{r.rejectReason ? `: ${r.rejectReason}` : ''}
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Modal: Quick Add Type (Admin / manage_book) */}
      {showAddTypeModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-4 sm:p-6 shadow-xl space-y-4 animate-scaleUp max-h-[90dvh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Plus className="w-5 h-5 text-sky-700 shrink-0" />
                  <span>Thêm loại văn bản mới</span>
                </h3>
                <p className="text-[11px] text-slate-500 mt-0.5">
                  Áp dụng ngay cho sổ: <strong className="text-slate-700">{activeBook.name}</strong>
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowAddTypeModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleQuickAddType} className="space-y-4 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-700">Tên loại văn bản *</label>
                <input
                  type="text"
                  value={newTypeName}
                  onChange={(e) => {
                    setNewTypeName(e.target.value);
                    if (addTypeError) setAddTypeError(null);
                    if (existingCatalogType) setExistingCatalogType(null);
                  }}
                  placeholder="Ví dụ: Giấy mời, Quyết định, Hướng dẫn..."
                  required
                  className="w-full min-h-[44px] h-11 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-700">Ký hiệu *</label>
                <input
                  type="text"
                  value={newTypeCode}
                  onChange={(e) => {
                    setNewTypeCode(e.target.value);
                    if (addTypeError) setAddTypeError(null);
                    if (existingCatalogType) setExistingCatalogType(null);
                  }}
                  placeholder="Ví dụ: GM, QĐ, HD"
                  required
                  className="w-full min-h-[44px] h-11 px-3 bg-white border border-slate-300 rounded-lg uppercase font-mono text-base sm:text-xs font-bold"
                />
                <span className="text-[11px] text-slate-500 block">
                  Số văn bản sẽ hiển thị theo mẫu: 01-<strong>{newTypeCode.trim().replace(/\s+/g, '').toUpperCase() || 'KÝ_HIỆU'}</strong>/{activeBook.unitCode}
                </span>
              </div>

              <div className="space-y-1.5">
                <label className="font-bold text-slate-700 block">Màu sắc số &amp; thẻ nhận diện</label>
                <div className="grid grid-cols-5 gap-1.5">
                  {Object.values(DOCUMENT_COLOR_PALETTE).map((colorOpt) => (
                    <button
                      key={colorOpt.key}
                      type="button"
                      onClick={() => setNewTypeColor(colorOpt.key)}
                      className={`flex flex-col items-center gap-1 p-1.5 min-h-[44px] justify-center rounded-lg border text-[10px] transition-all ${
                        newTypeColor === colorOpt.key
                          ? `${colorOpt.tagBg} ${colorOpt.tagBorder} ring-2 ring-sky-600 font-bold shadow-xs`
                          : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-600'
                      }`}
                    >
                      <span className={`w-3.5 h-3.5 rounded-full ${colorOpt.dotColor}`} />
                      <span className="truncate w-full text-center">{colorOpt.label.split(' ')[0]}</span>
                    </button>
                  ))}
                </div>
              </div>

              {addTypeError && (
                <div className="p-3 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg text-xs space-y-2">
                  <p className="font-medium">{addTypeError}</p>
                  {existingCatalogType && (
                    <button
                      type="button"
                      disabled={addTypeSaving}
                      onClick={handleEnableExistingCatalogType}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-md transition-colors disabled:opacity-50"
                    >
                      <Check className="w-3.5 h-3.5" />
                      <span>
                        Bật {existingCatalogType.code} ({existingCatalogType.name}) cho sổ {activeBook.name}
                      </span>
                    </button>
                  )}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowAddTypeModal(false)}
                  className="min-h-[40px] px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={addTypeSaving}
                  className="min-h-[40px] px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs disabled:opacity-50"
                >
                  {addTypeSaving ? 'Đang lưu...' : 'Lưu và áp dụng'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: Pre-Issue Confirmation (Kiểm tra trước khi cấp số) */}
      {preIssueModalData && (() => {
        const { check, fileSha256 } = preIssueModalData;
        const selectedTypeObj = activeTypes.find((t) => t.code === typeCode);
        const selectedTplObj = selectedTemplateId
          ? matchingTemplates.find((t) => t.id === selectedTemplateId)
          : undefined;
        const expectedWillSaveName = selectedFile
          ? buildStandardDriveFileName(check.previewNumber, docName, selectedFile.name)
          : selectedTplObj
          ? buildStandardDriveFileName(check.previewNumber, docName, '.docx')
          : 'Không đính kèm';

        return (
          <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4">
            <div className="bg-white rounded-xl max-w-lg w-full p-4 sm:p-6 shadow-2xl border border-slate-200 space-y-4 max-h-[90dvh] overflow-y-auto">
              <div className="flex items-start justify-between gap-2 border-b border-slate-100 pb-3">
                <div>
                  <h3 className="text-base font-bold text-slate-900">
                    Xác nhận trước khi cấp số
                  </h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Kiểm tra lại thông tin hồ sơ trước khi chốt số chính thức
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setPreIssueModalData(null)}
                  className="text-slate-400 hover:text-slate-700 p-1.5 min-w-[36px] min-h-[36px] flex items-center justify-center"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <dl className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-50 border border-slate-200 rounded-xl p-3.5 text-xs">
                <div>
                  <dt className="text-slate-500">Sổ văn bản</dt>
                  <dd className="font-bold text-slate-900 mt-0.5">{check.bookName}</dd>
                </div>
                <div>
                  <dt className="text-slate-500">Năm · Loại văn bản</dt>
                  <dd className="font-bold text-slate-900 mt-0.5 tabular-nums">
                    {check.year} · {selectedTypeObj?.name || typeCode} ({typeCode})
                  </dd>
                </div>
                <div>
                  <dt className="text-slate-500">Số dự kiến</dt>
                  <dd className="font-mono font-extrabold text-base text-sky-800 mt-0.5 tabular-nums">
                    {check.previewNumber}
                  </dd>
                </div>
                <div>
                  <dt className="text-slate-500">Người ký</dt>
                  <dd className="font-bold text-slate-900 mt-0.5">
                    {signer} ({signerRole})
                  </dd>
                </div>
                <div className="sm:col-span-2">
                  <dt className="text-slate-500">Tên văn bản</dt>
                  <dd className="font-semibold text-slate-900 mt-0.5 break-words">{docName.trim()}</dd>
                </div>
                <div className="sm:col-span-2">
                  <dt className="text-slate-500">Tên tệp sẽ lưu</dt>
                  <dd className="font-mono font-semibold text-slate-800 mt-0.5 break-all">
                    {expectedWillSaveName}
                    {selectedTplObj && !selectedFile && (
                      <span className="ml-1.5 text-[11px] font-sans text-sky-700 font-semibold">
                        (Tạo từ mẫu: {selectedTplObj.name})
                      </span>
                    )}
                  </dd>
                </div>
              </dl>

              {check.sameFileDocs.length > 0 && (
                <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl space-y-2 text-xs text-rose-950">
                  <div className="flex items-start gap-2 font-bold">
                    <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    <span>Cảnh báo: Tệp đính kèm trùng nội dung (SHA-256) với hồ sơ đã có trong sổ</span>
                  </div>
                  <div className="space-y-1.5">
                    {check.sameFileDocs.map((d) => (
                      <div key={d.id} className="flex flex-wrap items-center justify-between gap-2 bg-white border border-rose-200 rounded-lg px-2.5 py-1.5">
                        <span className="min-w-0">
                          <strong className="font-mono">{d.fullDocNumber}</strong> · {d.docName} ({d.issueDate})
                        </span>
                        {d.canView && onOpenExistingDocument && (
                          <button
                            type="button"
                            onClick={() => {
                              setPreIssueModalData(null);
                              onOpenExistingDocument(d.id);
                            }}
                            className="text-sky-700 font-semibold hover:underline shrink-0"
                          >
                            Mở hồ sơ
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {check.nearDuplicateDocs.length > 0 && (
                <div className="p-3.5 bg-amber-50 border border-amber-200 rounded-xl space-y-2 text-xs text-amber-950">
                  <div className="flex items-start gap-2 font-bold">
                    <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                    <span>Cảnh báo: Có hồ sơ gần giống tên văn bản trong sổ (chỉ để đối chiếu)</span>
                  </div>
                  <div className="space-y-1.5">
                    {check.nearDuplicateDocs.map((d) => (
                      <div key={d.id} className="flex flex-wrap items-center justify-between gap-2 bg-white border border-amber-200 rounded-lg px-2.5 py-1.5">
                        <span className="min-w-0">
                          <strong className="font-mono">{d.fullDocNumber}</strong> · {d.docName} ({d.issueDate})
                        </span>
                        {d.canView && onOpenExistingDocument && (
                          <button
                            type="button"
                            onClick={() => {
                              setPreIssueModalData(null);
                              onOpenExistingDocument(d.id);
                            }}
                            className="text-sky-700 font-semibold hover:underline shrink-0"
                          >
                            Mở hồ sơ
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <div className="flex flex-col-reverse sm:flex-row items-stretch sm:items-center justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setPreIssueModalData(null)}
                  className="min-h-[42px] px-4 py-2 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 rounded-lg"
                >
                  Quay lại sửa
                </button>
                <button
                  type="button"
                  disabled={isSubmitting}
                  onClick={async () => {
                    const modalState = preIssueModalData;
                    setPreIssueModalData(null);
                    await performIssueSubmit(undefined, false, fileSha256, {
                      expectedYear: modalState.check.year,
                      expectedNextSeq: modalState.check.sequenceNumber,
                    });
                  }}
                  className="min-h-[42px] px-5 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs inline-flex items-center justify-center gap-1.5"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  <span>Xác nhận cấp số {check.previewNumber}</span>
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
};
