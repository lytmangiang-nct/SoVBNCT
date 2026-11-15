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
} from 'lucide-react';
import { api } from '../api';
import type {
  DocumentRecord,
  DocumentTypeConfig,
  NumberRequest,
  RegistryBook,
  SignerConfig,
  SystemSettings,
  User,
} from '../types';
import { getTypeColorStyle, DOCUMENT_COLOR_PALETTE } from '../utils/documentColors';
import {
  buildStandardDriveFileName,
  DEFAULT_MAX_WORD_MB,
  formatFileSize,
  generateSuggestedFileName,
  getWordFileExtension,
} from '../utils/fileNaming';
import { DEFAULT_BOOK_ID, getBookEnabledTypes, hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface IssueNumberViewProps {
  currentUser: User;
  activeBook: RegistryBook;
  accessibleBooks: RegistryBook[];
  onSelectBook: (bookId: string) => void;
  settings: SystemSettings;
  documentTypes: DocumentTypeConfig[];
  signers: SignerConfig[];
  documents?: DocumentRecord[];
  numberRequests: NumberRequest[];
  onRefreshRequests: () => void;
  onDocumentIssued: (doc: DocumentRecord) => void;
  onDocumentUpdated?: (doc: DocumentRecord) => void;
  onNavigateToRegistry: () => void;
  onOpenExistingDocument?: (docId: string) => void;
  onDocumentTypeAdded?: (newType: DocumentTypeConfig) => void;
}

export const IssueNumberView: React.FC<IssueNumberViewProps> = ({
  currentUser,
  activeBook,
  accessibleBooks,
  onSelectBook,
  settings,
  documentTypes,
  signers: globalSigners,
  documents = [],
  numberRequests,
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

  const effectiveSigners = useMemo(
    () => (activeBook.signers && activeBook.signers.length > 0 ? activeBook.signers : globalSigners),
    [activeBook.signers, globalSigners]
  );

  // Form State (scoped to document types enabled in activeBook)
  const activeTypes = useMemo(
    () => getBookEnabledTypes(activeBook, documentTypes),
    [activeBook, documentTypes]
  );
  const defaultTypeCode = activeTypes.length > 0 ? activeTypes[0].code : 'KH';

  const [typeCode, setTypeCode] = useState<string>(defaultTypeCode);
  const [docName, setDocName] = useState<string>('');
  const [title, setTitle] = useState<string>('');
  const [signer, setSigner] = useState<string>(effectiveSigners.length > 0 ? effectiveSigners[0].name : '');
  const [signerRole, setSignerRole] = useState<string>(effectiveSigners.length > 0 ? effectiveSigners[0].role : '');
  const [isCustomSigner, setIsCustomSigner] = useState<boolean>(false);
  const [notes, setNotes] = useState<string>('');

  // Quick Add Type Modal State (Admin only)
  const [showAddTypeModal, setShowAddTypeModal] = useState<boolean>(false);
  const [newTypeName, setNewTypeName] = useState<string>('');
  const [newTypeCode, setNewTypeCode] = useState<string>('');
  const [newTypeColor, setNewTypeColor] = useState<string>('blue');
  const [addTypeSaving, setAddTypeSaving] = useState<boolean>(false);
  const [addTypeError, setAddTypeError] = useState<string | null>(null);

  // Idempotency Key
  const generateKey = () => `req-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const [idempotencyKey, setIdempotencyKey] = useState<string>(generateKey);

  // Next Document Number Preview
  const [nextNumber, setNextNumber] = useState<string>('...');
  const [loadingNextNumber, setLoadingNextNumber] = useState<boolean>(false);

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

  // Word File Attachment State
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
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

  const driveConfig =
    activeBook.id !== DEFAULT_BOOK_ID
      ? activeBook.googleDriveConfig
      : activeBook.googleDriveConfig || settings.googleDriveConfig;
  const driveConfigured = Boolean(driveConfig?.isConfigured);
  const maxWordFileSizeMB = settings.googleDriveConfig?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;
  const maxWordFileSizeBytes = settings.googleDriveConfig?.maxFileSizeBytes || maxWordFileSizeMB * 1024 * 1024;

  const [todayVietnam] = useState<string>(() => {
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Ho_Chi_Minh',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    });
    const [y, m, d] = formatter.format(new Date()).split('-');
    return `${d}/${m}/${y}`;
  });

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

  // Fetch preview number scoped to activeBook
  useEffect(() => {
    let active = true;
    const fetchNextNumber = async () => {
      if (!typeCode || (!canIssue && !canRequest && !canView)) {
        setNextNumber('---');
        return;
      }
      setLoadingNextNumber(true);
      try {
        const res = await api.previewNumber(typeCode, activeBook.id);
        if (active) {
          setNextNumber(res.previewNumber);
        }
      } catch {
        if (active) {
          setNextNumber('---');
        }
      } finally {
        if (active) setLoadingNextNumber(false);
      }
    };

    fetchNextNumber();

    const handleRecheck = () => {
      if (document.visibilityState === 'visible') {
        fetchNextNumber();
      }
    };
    window.addEventListener('focus', handleRecheck);
    document.addEventListener('visibilitychange', handleRecheck);

    return () => {
      active = false;
      window.removeEventListener('focus', handleRecheck);
      document.removeEventListener('visibilitychange', handleRecheck);
    };
  }, [typeCode, activeBook.id, activeBook.numberMode, documents.length, documents[0]?.id, canIssue, canRequest, canView]);

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

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setFileError(null);
    setSubmitError(null);
    if (!file) return;

    const err = validateWordFile(file);
    if (err) {
      setFileError(err);
      setSelectedFile(null);
      if (fileInputRef.current) fileInputRef.current.value = '';
      return;
    }
    setSelectedFile(file);
  };

  const handleClearSelectedFile = () => {
    setSelectedFile(null);
    setFileError(null);
    setSubmitError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const performIssueSubmit = async (conflictAction?: 'suffix' | 'overwrite') => {
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

    setIsSubmitting(true);
    setUploadStep(selectedFile ? 'issuing' : 'idle');
    setRetryError(null);

    let issuedDoc: DocumentRecord | null = null;
    const fileToUpload = selectedFile;

    try {
      const res = await api.issueDocument({
        bookId: activeBook.id,
        typeCode: typeCode.trim().toUpperCase(),
        docName: docName.trim(),
        title: title.trim(),
        signer: signer.trim(),
        signerRole: signerRole.trim(),
        notes: notes.trim(),
        idempotencyKey,
        hasAttachment: Boolean(fileToUpload),
        attachmentName: fileToUpload?.name,
        attachmentSize: fileToUpload?.size,
        conflictAction,
      });

      issuedDoc = res.document;
      setIssuedSuccessDoc(issuedDoc);
      onDocumentIssued(issuedDoc);

      api.previewNumber(typeCode.trim().toUpperCase(), activeBook.id).then((preview) => {
        setNextNumber(preview.previewNumber);
      }).catch(() => {});

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
      if (err.code === 'FILENAME_CONFLICT' && err.existingDocument) {
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
    await performIssueSubmit();
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

  const handleRejectRequest = async (reqId: string) => {
    if (processingReqId) return;
    const reason = window.prompt('Nhập lý do từ chối yêu cầu xin số (tùy chọn):', 'Chưa đủ thông tin văn bản');
    if (reason === null) return;
    setProcessingReqId(reqId);
    setQueueError(null);
    try {
      await api.rejectNumberRequest(reqId, reason);
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
    setCopied(false);
    setSelectedFile(null);
    setFileError(null);
    setPendingRetryFile(null);
    setRetryError(null);
    setUploadStep('idle');
    if (fileInputRef.current) fileInputRef.current.value = '';
    setIdempotencyKey(generateKey());

    if (typeCode) {
      api.previewNumber(typeCode, activeBook.id).then((preview) => {
        setNextNumber(preview.previewNumber);
      }).catch(() => {});
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

  const handleQuickAddType = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTypeName.trim() || !newTypeCode.trim()) {
      setAddTypeError('Vui lòng nhập đầy đủ tên loại văn bản và ký hiệu viết tắt');
      return;
    }
    setAddTypeSaving(true);
    setAddTypeError(null);
    try {
      const res = await api.createDocumentType({
        name: newTypeName.trim(),
        code: newTypeCode.trim().toUpperCase(),
        color: newTypeColor,
      });
      if (onDocumentTypeAdded) {
        onDocumentTypeAdded(res.documentType);
      }
      setTypeCode(res.documentType.code);
      setShowAddTypeModal(false);
      setNewTypeName('');
      setNewTypeCode('');
      setNewTypeColor('blue');
    } catch (err: any) {
      setAddTypeError(err.message || 'Lỗi khi tạo loại văn bản mới');
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
    <div className="max-w-4xl mx-auto py-8 px-4 sm:px-6 space-y-6">
      {/* Banner Intro + Book Switcher */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            {canIssue ? 'Cấp Số Văn Bản Tự Động' : 'Xin Số Văn Bản'}
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            {canIssue
              ? `Cấp số chính thức vào ${activeBook.name} (${activeBook.unitCode}) theo thứ tự nguyên tử và chống trùng số.`
              : `Gửi yêu cầu xin số vào ${activeBook.name} (${activeBook.unitCode}) để người có quyền Cấp số duyệt và cấp số chính thức.`}
          </p>
        </div>

        {issueAccessibleBooks.length > 1 && (
          <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-3 py-2 shadow-2xs text-xs shrink-0">
            <BookOpen className="w-3.5 h-3.5 text-sky-700" />
            <select
              value={activeBook.id}
              onChange={(e) => {
                onSelectBook(e.target.value);
                setIssuedSuccessDoc(null);
                setRequestSuccess(null);
                setSubmitError(null);
              }}
              className="font-semibold text-slate-800 bg-transparent focus:outline-hidden cursor-pointer"
            >
              {issueAccessibleBooks.map((b) => (
                <option key={b.id} value={b.id}>
                  {b.name} ({b.unitCode})
                </option>
              ))}
            </select>
          </div>
        )}
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

                <div className="flex items-center gap-2 shrink-0">
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
                    onClick={() => handleRejectRequest(reqItem.id)}
                    className="inline-flex items-center gap-1 px-3 py-2 bg-white hover:bg-rose-50 border border-rose-300 text-rose-700 font-semibold rounded-lg"
                  >
                    <XCircle className="w-3.5 h-3.5" />
                    <span>Từ chối</span>
                  </button>
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
          <div className={`p-6 ${successStyle.previewBg} border ${successStyle.previewBorder} rounded-xl shadow-xs transition-all animate-fadeIn`}>
            <div className="flex items-start gap-4">
              <div className={`w-12 h-12 rounded-full ${successStyle.dotColor} flex items-center justify-center text-white shrink-0 shadow-xs`}>
                <CheckCircle2 className="w-7 h-7" />
              </div>

              <div className="flex-1 min-w-0">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <div className="flex flex-wrap items-center gap-2 mb-1">
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
                                ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                                : issuedSuccessDoc.driveStatus === 'uploading'
                                ? 'bg-sky-50 text-sky-800 border-sky-200'
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
                      <span className="text-[11px] font-medium text-slate-600">
                        {issuedSuccessDoc.bookName || activeBook.name}
                      </span>
                    </div>

                    <div className={`text-3xl font-extrabold ${successStyle.previewText} font-mono tracking-tight mt-1 flex items-center gap-3 flex-wrap`}>
                      <span>{issuedSuccessDoc.fullDocNumber}</span>
                      <button
                        type="button"
                        onClick={() => handleCopyDocNumber(issuedSuccessDoc.fullDocNumber)}
                        className={`inline-flex items-center gap-1.5 px-3 py-1 text-xs font-semibold rounded-md bg-white border ${successStyle.badgeBorder} ${successStyle.badgeText} hover:bg-slate-50 transition-colors shadow-2xs`}
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

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleIssueNext}
                      className="flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs transition-colors"
                    >
                      <RotateCcw className="w-4 h-4" />
                      <span>Cấp số tiếp theo</span>
                    </button>

                    {canView && (
                      <button
                        type="button"
                        onClick={onNavigateToRegistry}
                        className="flex items-center gap-2 px-4 py-2 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 rounded-lg shadow-2xs transition-colors"
                      >
                        <span>Xem sổ theo dõi</span>
                        <ArrowRight className="w-4 h-4 text-slate-500" />
                      </button>
                    )}
                  </div>
                </div>

                {/* Record Summary */}
                <div className={`mt-4 pt-4 border-t ${successStyle.previewBorder} grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 text-xs text-slate-700`}>
                  <div>
                    <span className="text-slate-500 block">Tên văn bản:</span>
                    <p className="font-bold text-slate-900 mt-0.5">{issuedSuccessDoc.docName || issuedSuccessDoc.title}</p>
                  </div>
                  <div>
                    <span className="text-slate-500 block">Trích yếu:</span>
                    <p className="font-medium text-slate-700 mt-0.5 line-clamp-2">
                      {issuedSuccessDoc.title || <span className="italic text-slate-400">Không có</span>}
                    </p>
                  </div>
                  <div>
                    <span className="text-slate-500 block">Người ký:</span>
                    <p className="font-medium text-slate-900 mt-0.5">
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
                          <FileText className="w-3.5 h-3.5 text-sky-600" />
                          <span>Tên tệp sẽ lưu (theo số chính thức {issuedSuccessDoc.fullDocNumber}):</span>
                        </span>
                        <p className="font-mono text-xs sm:text-sm font-bold text-slate-900 truncate">
                          {officialFileName}
                        </p>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleCopyFileName(officialFileName, true)}
                        className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-md shadow-2xs transition-colors shrink-0"
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
                          <p className="font-mono text-xs font-semibold text-emerald-800 truncate">
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
        <div className="bg-white border border-slate-200 rounded-xl p-6 sm:p-8 shadow-xs">
          <form onSubmit={handleSubmit} className="space-y-6">
            {/* Top row: Loại văn bản & Số dự kiến */}
            {(() => {
              const selectedDocType = activeTypes.find((t) => t.code === typeCode);
              const currentStyle = getTypeColorStyle(typeCode, selectedDocType?.color);

              return (
                <div className="grid grid-cols-1 md:grid-cols-12 gap-6 items-start">
                  <div className="md:col-span-7 space-y-2">
                    <div className="flex items-center justify-between">
                      <label htmlFor="doc-type" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                        Loại văn bản <span className="text-rose-600">*</span>
                      </label>
                      {isSystemAdmin(currentUser) && (
                        <button
                          type="button"
                          onClick={() => {
                            setNewTypeName('');
                            setNewTypeCode('');
                            setNewTypeColor('blue');
                            setAddTypeError(null);
                            setShowAddTypeModal(true);
                          }}
                          className="text-xs font-semibold text-sky-700 hover:text-sky-900 transition-colors flex items-center gap-1"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Thêm loại văn bản</span>
                        </button>
                      )}
                    </div>

                    <select
                      id="doc-type"
                      value={typeCode}
                      onChange={(e) => setTypeCode(e.target.value)}
                      className="w-full h-10 px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors"
                      disabled={isSubmitting}
                    >
                      {activeTypes.map((type) => (
                        <option key={type.code} value={type.code}>
                          {type.name} ({type.code})
                        </option>
                      ))}
                    </select>
                    <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                      <span className="text-xs text-slate-500">
                        Sổ: <strong className="text-slate-700">{activeBook.name}</strong> · Ký hiệu đơn vị:{' '}
                        <strong className="text-slate-700">{activeBook.unitCode}</strong> ({activeBook.unitName})
                      </span>
                      {issueAccessibleBooks.length > 1 && (
                        <div className="inline-flex items-center gap-1.5 text-xs bg-sky-50 border border-sky-200 rounded-md px-2 py-1">
                          <BookOpen className="w-3.5 h-3.5 text-sky-700 shrink-0" />
                          <select
                            aria-label="Chọn sổ cấp số"
                            value={activeBook.id}
                            onChange={(e) => {
                              onSelectBook(e.target.value);
                              setIssuedSuccessDoc(null);
                              setRequestSuccess(null);
                              setSubmitError(null);
                            }}
                            disabled={isSubmitting}
                            className="font-bold text-sky-900 bg-transparent focus:outline-hidden cursor-pointer"
                          >
                            {issueAccessibleBooks.map((b) => (
                              <option key={b.id} value={b.id}>
                                {b.name} ({b.unitCode})
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Số dự kiến tiếp theo */}
                  <div className={`md:col-span-5 ${currentStyle.previewBg} border ${currentStyle.previewBorder} rounded-xl p-4 flex flex-col justify-between shadow-xs transition-all`}>
                    <div className="flex items-center justify-between text-xs mb-1.5 gap-1.5">
                      <span className={`font-bold uppercase tracking-wider ${currentStyle.previewText} flex items-center gap-1.5`}>
                        <span className={`w-2 h-2 rounded-full ${currentStyle.dotColor}`} />
                        <span>{canIssue ? 'Số cấp tiếp theo (dự kiến)' : 'Số tham khảo hiện tại'}</span>
                      </span>
                      <span className={`text-[11px] font-semibold px-2 py-0.5 rounded-md border ${currentStyle.previewTag}`}>
                        {selectedDocType?.name || typeCode}
                      </span>
                    </div>
                    <div className={`text-2xl sm:text-3xl font-black font-mono ${currentStyle.previewText} tracking-tight flex items-center gap-2 py-1`}>
                      {loadingNextNumber ? (
                        <span className="text-xs font-normal text-slate-500">Đang cập nhật...</span>
                      ) : (
                        <span>{nextNumber}</span>
                      )}
                    </div>
                    <div className="flex items-center justify-between mt-1 text-[11px] text-slate-600 gap-1.5 flex-wrap">
                      <span className="font-semibold text-slate-700">
                        {activeBook.name} · {activeBook.numberMode === 'shared' ? 'Dãy số chung' : `Dãy riêng ${typeCode}`}
                      </span>
                      <span className="font-medium text-emerald-700 bg-white/90 px-1.5 py-0.5 rounded border border-emerald-200">
                        Tự động nối tiếp
                      </span>
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Tên văn bản */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between">
                <label htmlFor="doc-name" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Tên văn bản <span className="text-rose-600">*</span>
                </label>
                <span className="text-[11px] text-rose-600 font-semibold bg-rose-50 px-2 py-0.5 rounded border border-rose-200">
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
                className="w-full h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-sm font-medium text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors placeholder:text-slate-400"
                disabled={isSubmitting}
              />

              {/* Gợi ý tên file chuẩn */}
              {(() => {
                const previewDocNumber = nextNumber && nextNumber !== '...' && nextNumber !== '---' ? nextNumber : '';
                const suggestedFileName = generateSuggestedFileName(previewDocNumber, docName);

                return (
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 sm:p-4 space-y-2 mt-2">
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800 uppercase tracking-wide">
                        <FileText className="w-4 h-4 text-sky-600" />
                        <span>Tên tệp dự kiến (Số ký hiệu - Tên văn bản)</span>
                      </div>
                      <span className="text-[11px] font-medium text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200">
                        Đổi <strong className="text-sky-700">/</strong> thành <strong className="text-sky-700">.</strong> · Giữ dấu tiếng Việt
                      </span>
                    </div>

                    <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                      <div className="flex-1 px-3.5 py-2 bg-white border border-slate-300 rounded-lg font-mono text-xs sm:text-sm font-bold text-slate-900 select-all overflow-x-auto whitespace-nowrap shadow-2xs flex items-center min-h-10">
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
                        className="inline-flex items-center justify-center gap-1.5 px-4 py-2 text-xs font-semibold bg-white hover:bg-sky-50 text-slate-700 hover:text-sky-700 border border-slate-300 hover:border-sky-300 rounded-lg shadow-2xs transition-colors shrink-0 disabled:opacity-50"
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
              <div className="flex items-center justify-between">
                <label htmlFor="doc-title" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Trích yếu / Nội dung văn bản
                </label>
                <span className="text-[11px] text-slate-400 font-normal">Không bắt buộc</span>
              </div>
              <textarea
                id="doc-title"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="Tóm tắt mục đích, nội dung chính của văn bản (nếu có)..."
                rows={3}
                className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors placeholder:text-slate-400"
                disabled={isSubmitting}
              />
            </div>

            {/* Signer & Role */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-1.5">
                <label htmlFor="signer-select" className="block text-xs font-bold text-slate-700 uppercase tracking-wide">
                  Người ký <span className="text-rose-600">*</span>
                </label>

                {!isCustomSigner ? (
                  <select
                    id="signer-select"
                    value={signer}
                    onChange={handleSignerSelect}
                    className="w-full h-11 px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 focus:border-sky-600 transition-colors"
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
                      className="w-full h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
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
                      className="text-xs text-sky-700 hover:underline font-medium"
                    >
                      Chọn lại từ danh sách có sẵn
                    </button>
                  </div>
                )}
              </div>

              <div className="space-y-1.5">
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
                  className="w-full h-11 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  disabled={isSubmitting}
                />
              </div>
            </div>

            {/* Automatic Issuer & Date Info */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 p-4 bg-slate-50 border border-slate-200 rounded-lg">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-md bg-slate-200 flex items-center justify-center text-slate-700 shrink-0">
                  <Shield className="w-5 h-5 text-slate-600" />
                </div>
                <div className="flex flex-col">
                  <span className="text-xs text-slate-500 font-medium">
                    {canIssue ? 'Người lấy số (Tự động)' : 'Người xin số (Tự động)'}
                  </span>
                  <span className="text-sm font-semibold text-slate-900">
                    {currentUser.fullName} <span className="text-slate-400 font-normal">({currentUser.username})</span>
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-md bg-slate-200 flex items-center justify-center text-slate-700 shrink-0">
                  <Clock className="w-5 h-5 text-slate-600" />
                </div>
                <div className="flex flex-col">
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
                className="w-full h-10 px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 placeholder:text-slate-400"
                disabled={isSubmitting}
              />
            </div>

            {/* Word File Attachment (Optional — available for issue_number users) */}
            {canIssue && (() => {
              const previewDocNumber = nextNumber && nextNumber !== '...' && nextNumber !== '---' ? nextNumber : '';
              const expectedDriveName = selectedFile
                ? buildStandardDriveFileName(previewDocNumber, docName, selectedFile.name)
                : '';

              return (
                <div className="space-y-2.5 p-4 bg-slate-50/80 border border-slate-200 rounded-xl">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <FileUp className="w-4 h-4 text-sky-700" />
                      <span className="text-xs font-bold text-slate-800 uppercase tracking-wide">
                        Tệp Word đính kèm
                      </span>
                      <span className="text-[11px] text-slate-500 font-normal">(không bắt buộc)</span>
                    </div>
                    <span className="text-[11px] font-medium text-slate-600 bg-white px-2 py-0.5 rounded border border-slate-200">
                      Định dạng .doc, .docx · Tối đa {maxWordFileSizeMB} MB
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
                      <div className="text-xs text-slate-600 space-y-0.5">
                        <p className="font-medium text-slate-800">
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
                        className="inline-flex items-center justify-center gap-1.5 px-4 py-2 text-xs font-semibold bg-white hover:bg-sky-50 text-sky-800 border border-sky-300 rounded-lg shadow-2xs transition-colors shrink-0 disabled:opacity-50"
                      >
                        <Upload className="w-3.5 h-3.5" />
                        <span>Chọn tệp Word</span>
                      </button>
                    </div>
                  ) : (
                    <div className="bg-white border border-sky-200 rounded-lg p-3.5 space-y-2.5">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div className="flex items-center gap-2.5 min-w-0">
                          <div className="w-9 h-9 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center text-sky-700 shrink-0">
                            <FileText className="w-4 h-4" />
                          </div>
                          <div className="min-w-0">
                            <p className="text-xs font-bold text-slate-900 truncate">{selectedFile.name}</p>
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
                            className="px-2.5 py-1.5 text-xs font-semibold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded-md disabled:opacity-50"
                          >
                            Chọn lại
                          </button>
                          <button
                            type="button"
                            onClick={handleClearSelectedFile}
                            disabled={isSubmitting}
                            className="px-2.5 py-1.5 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-md disabled:opacity-50"
                          >
                            Bỏ chọn
                          </button>
                        </div>
                      </div>

                      <div className="bg-slate-50 border border-slate-200 rounded-md px-3 py-2 text-xs space-y-0.5">
                        <span className="text-[11px] font-semibold text-slate-500 block">
                          Tên tệp sẽ lưu (dự kiến — nếu người khác vừa cấp số trước, máy chủ sẽ tự tính lại theo số chính thức):
                        </span>
                        <span className="font-mono font-bold text-sky-900 break-all">{expectedDriveName}</span>
                      </div>

                      {!driveConfigured && (
                        <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-amber-900 text-xs flex items-start gap-2">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                          <div className="space-y-1">
                            <strong className="block font-bold">Thư mục Google Drive chưa được cấu hình hợp lệ</strong>
                            <p className="text-amber-800">
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
                      <span>{fileError}</span>
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
            <div className="pt-2 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4 border-t border-slate-100">
              <div className="text-xs text-slate-500 flex items-center gap-2">
                <Sparkles className="w-4 h-4 text-sky-600 shrink-0" />
                <span>
                  {canIssue
                    ? 'Giao dịch máy chủ nguyên tử · Bảo vệ chống nhấn đúp · Lưu trực tuyến lâu dài'
                    : 'Yêu cầu xin số được ghi nhận tức thì và chuyển tới người có quyền Cấp số'}
                </span>
              </div>

              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5">
                {issueAccessibleBooks.length > 1 ? (
                  <div className="inline-flex items-center justify-between gap-1.5 h-12 px-3 bg-sky-50 border border-sky-200 rounded-lg text-xs">
                    <span className="text-sky-700 font-medium whitespace-nowrap">Sổ:</span>
                    <select
                      aria-label="Xác nhận sổ khi cấp số"
                      value={activeBook.id}
                      onChange={(e) => {
                        onSelectBook(e.target.value);
                        setIssuedSuccessDoc(null);
                        setRequestSuccess(null);
                        setSubmitError(null);
                      }}
                      disabled={isSubmitting}
                      className="font-bold text-sky-900 bg-transparent focus:outline-hidden cursor-pointer"
                    >
                      {issueAccessibleBooks.map((b) => (
                        <option key={b.id} value={b.id}>
                          {b.name} ({b.unitCode})
                        </option>
                      ))}
                    </select>
                  </div>
                ) : (
                  <span className="inline-flex items-center justify-center gap-1.5 h-12 px-3 bg-slate-100 border border-slate-200 rounded-lg text-xs font-bold text-slate-700">
                    <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                    <span>{activeBook.name}</span>
                  </span>
                )}

                <button
                  type="submit"
                  disabled={isSubmitting || Boolean(canIssue && selectedFile && !driveConfigured)}
                  className="w-full sm:w-auto sm:min-w-[200px] h-12 px-6 bg-sky-700 hover:bg-sky-800 active:bg-sky-900 text-white font-semibold text-sm rounded-lg shadow-sm transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isSubmitting ? (
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

      {/* Modal: Quick Add Type (Admin Only) */}
      {showAddTypeModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-6 shadow-xl space-y-4 animate-scaleUp">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Plus className="w-5 h-5 text-sky-700" />
                <span>Thêm loại văn bản mới</span>
              </h3>
              <button
                type="button"
                onClick={() => setShowAddTypeModal(false)}
                className="text-slate-400 hover:text-slate-700 p-1"
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
                  onChange={(e) => setNewTypeName(e.target.value)}
                  placeholder="Ví dụ: Quyết định, Hướng dẫn, Tờ trình..."
                  required
                  className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-xs"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-700">Ký hiệu viết tắt *</label>
                <input
                  type="text"
                  value={newTypeCode}
                  onChange={(e) => setNewTypeCode(e.target.value)}
                  placeholder="Ví dụ: QĐ, HD, TTr"
                  required
                  className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg uppercase font-mono text-xs font-bold"
                />
                <span className="text-[11px] text-slate-500 block">
                  Số văn bản sẽ hiển thị theo mẫu: 01-<strong>{newTypeCode || 'KÍ_HIỆU'}</strong>/{activeBook.unitCode}
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
                      className={`flex flex-col items-center gap-1 p-1.5 rounded-lg border text-[10px] transition-all ${
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
                <div className="p-2.5 bg-rose-50 text-rose-700 rounded-lg text-xs">
                  {addTypeError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowAddTypeModal(false)}
                  className="px-3.5 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={addTypeSaving}
                  className="px-4 py-2 text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white rounded-lg shadow-xs"
                >
                  {addTypeSaving ? 'Đang tạo...' : 'Lưu và áp dụng'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
