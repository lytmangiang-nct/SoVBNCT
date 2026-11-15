import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Users,
  Shield,
  Cloud,
  Check,
  AlertTriangle,
  Plus,
  Trash2,
  Save,
  CheckCircle2,
  RotateCw,
  RefreshCw,
  Copy,
  Clock,
  UserCheck,
  UserX,
  Lock,
  ShieldCheck,
  HardDrive,
  FolderCheck,
  Upload,
  BookOpen,
  Sliders,
  ChevronDown,
  ChevronRight,
  Menu,
  X,
} from 'lucide-react';
import { api } from '../api';
import type {
  BookPermission,
  DocumentTypeConfig,
  GoogleDriveAuthMode,
  RegistryBook,
  SignerConfig,
  StartSequenceSeriesConfig,
  SystemAuditLog,
  SystemSettings,
  User,
} from '../types';
import { getTypeColorStyle, DOCUMENT_COLOR_PALETTE } from '../utils/documentColors';
import { DEFAULT_MAX_WORD_MB, extractGoogleDriveFolderId } from '../utils/fileNaming';
import {
  ALL_BOOK_PERMISSIONS,
  BOOK_PERMISSION_META,
  DEFAULT_BOOK_ID,
  getBookEnabledTypes,
  hasBookPermission,
  isSystemAdmin,
  normalizeBookPermissionList,
} from '../utils/permissions';
import {
  GOOGLE_APPS_SCRIPT_TEMPLATE,
  GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE,
  PERMISSION_PRESETS,
  SETTINGS_NAV_GROUPS,
  getBookIdentityStyle,
  type SettingsSubTab,
} from './settings/settingsConstants';
import { BooksListSection } from './settings/BooksListSection';
import { CreateBookDrawer, type CreateBookFormValues } from './settings/CreateBookDrawer';
import {
  ConfirmDeleteModal,
  PendingBookSwitchModal,
  ResetPasswordModal,
  ToggleBookLockModal,
} from './settings/SettingsModals';

interface SettingsViewProps {
  currentUser: User;
  books: RegistryBook[];
  activeBookId: string;
  onSelectBookId: (bookId: string) => void;
  settings: SystemSettings;
  documentTypes: DocumentTypeConfig[];
  signers: SignerConfig[];
  serverYear?: number;
  configLoaded?: boolean;
  syncError?: string | null;
  onSettingsUpdated: (newSettings: SystemSettings) => void;
  onBooksUpdated?: () => void;
  onRefreshData: () => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  currentUser,
  books,
  activeBookId,
  onSelectBookId,
  settings,
  documentTypes,
  signers,
  serverYear,
  configLoaded = true,
  syncError = null,
  onSettingsUpdated,
  onBooksUpdated,
  onRefreshData,
}) => {
  const isAdmin = isSystemAdmin(currentUser);
  const currentYear =
    serverYear ||
    Number(
      new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Ho_Chi_Minh',
        year: 'numeric',
      }).format(new Date())
    );
  const isConfigBlocked = !configLoaded || Boolean(syncError);

  const managedBooks = useMemo(() => {
    return books.filter((b) => isAdmin || hasBookPermission(currentUser, b.id, 'manage_book'));
  }, [books, isAdmin, currentUser]);

  const [selectedBookId, setSelectedBookId] = useState<string>(() => {
    if (managedBooks.some((b) => b.id === activeBookId)) return activeBookId;
    return managedBooks[0]?.id || books[0]?.id || DEFAULT_BOOK_ID;
  });

  useEffect(() => {
    if (!managedBooks.some((b) => b.id === selectedBookId) && managedBooks.length > 0) {
      setSelectedBookId(
        managedBooks.some((b) => b.id === activeBookId) ? activeBookId : managedBooks[0].id
      );
    }
  }, [managedBooks, selectedBookId, activeBookId]);

  const selectedBook = useMemo(() => {
    return (
      books.find((b) => b.id === selectedBookId) ||
      managedBooks[0] ||
      books[0] || {
        id: DEFAULT_BOOK_ID,
        code: 'DOAN_TRUONG',
        name: 'Đoàn trường',
        unitName: settings.unitName,
        unitCode: settings.unitCode,
        isShared: true,
        numberMode: settings.numberMode,
        numberFormat: settings.numberFormat,
        customStartSequences: settings.customStartSequences || {},
        signers,
        isActive: true,
        createdAt: '',
      }
    );
  }, [books, managedBooks, selectedBookId, settings, signers]);

  const canManageSelectedBook = isAdmin || hasBookPermission(currentUser, selectedBook.id, 'manage_book');

  // Sub-tab inside Settings
  const [activeSubTab, setActiveSubTab] = useState<SettingsSubTab>('books');
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  // Book Management State (Multi-book drawer & lock modal)
  const [createBookDrawerOpen, setCreateBookDrawerOpen] = useState(false);
  const [bookSaving, setBookSaving] = useState(false);
  const [bookMessage, setBookMessage] = useState<{ success: boolean; text: string } | null>(null);
  const [createBookServerError, setCreateBookServerError] = useState<string | null>(null);
  const [toggleLockConfirmBook, setToggleLockConfirmBook] = useState<RegistryBook | null>(null);
  const [toggleLockSaving, setToggleLockSaving] = useState(false);

  // Generic confirm & reset password modals (no browser alert/confirm/prompt)
  const [confirmDeleteRequest, setConfirmDeleteRequest] = useState<{
    title: string;
    description: string;
    confirmLabel: string;
    onConfirm: () => Promise<void>;
  } | null>(null);
  const [confirmDeleteBusy, setConfirmDeleteBusy] = useState(false);
  const [resetPassUser, setResetPassUser] = useState<User | null>(null);
  const [resetPassSaving, setResetPassSaving] = useState(false);

  // Unsaved-changes book switch guard
  const [pendingBookSwitch, setPendingBookSwitch] = useState<{
    targetBookId: string;
    targetBookName: string;
    currentBookName: string;
    nextTab?: SettingsSubTab;
  } | null>(null);
  const [switchingBookSaving, setSwitchingBookSaving] = useState(false);
  const [switchBookError, setSwitchBookError] = useState<string | null>(null);

  // General Settings Form (per selected book)
  const [bookTitleInput, setBookTitleInput] = useState(selectedBook.name);
  const [unitName, setUnitName] = useState(selectedBook.unitName);
  const [unitCode, setUnitCode] = useState(selectedBook.unitCode);
  const [numberFormat, setNumberFormat] = useState(selectedBook.numberFormat);
  const [bookEnabledTypes, setBookEnabledTypes] = useState<string[]>(() =>
    getBookEnabledTypes(selectedBook, documentTypes).map((t) => t.code)
  );
  const [generalSaving, setGeneralSaving] = useState(false);
  const [generalSuccess, setGeneralSuccess] = useState(false);
  const [generalError, setGeneralError] = useState<string | null>(null);

  // Numbering Mode (per selected book)
  const [targetMode, setTargetMode] = useState<'separate' | 'shared'>(selectedBook.numberMode);
  const [modeConfirmText, setModeConfirmText] = useState('');
  const [modeSaving, setModeSaving] = useState(false);
  const [modeError, setModeError] = useState<string | null>(null);
  const [modeSuccess, setModeSuccess] = useState(false);

  // Custom Starting Sequence (keyed strictly by bookId + year + typeCode)
  const [seqYear, setSeqYear] = useState<number>(currentYear);
  const [seqTypeCode, setSeqTypeCode] = useState<string>(() => {
    if (selectedBook.numberMode === 'shared') return 'SHARED';
    return getBookEnabledTypes(selectedBook, documentTypes)[0]?.code || 'KH';
  });
  const [savedSeriesMap, setSavedSeriesMap] = useState<Record<string, StartSequenceSeriesConfig>>({});
  const [draftStartInputBySeries, setDraftStartInputBySeries] = useState<Record<string, string>>({});
  const [seqConfigLoading, setSeqConfigLoading] = useState(false);
  const [seqConfigLoadError, setSeqConfigLoadError] = useState<string | null>(null);
  const [savingSeriesMap, setSavingSeriesMap] = useState<Record<string, boolean>>({});
  const [seqSaveErrorBySeries, setSeqSaveErrorBySeries] = useState<Record<string, string | null>>({});
  const [seqSaveSuccessNotice, setSeqSaveSuccessNotice] = useState<{
    seriesId: string;
    message: string;
  } | null>(null);

  // Doc Type Modal / Form
  const [newTypeName, setNewTypeName] = useState('');
  const [newTypeCode, setNewTypeCode] = useState('');
  const [newTypeColor, setNewTypeColor] = useState('blue');
  const [typeSaving, setTypeSaving] = useState(false);
  const [typeError, setTypeError] = useState<string | null>(null);

  // Signer Modal / Form (per selected book)
  const [newSignerName, setNewSignerName] = useState('');
  const [newSignerRole, setNewSignerRole] = useState('');
  const [signerSaving, setSignerSaving] = useState(false);
  const [signerError, setSignerError] = useState<string | null>(null);

  // Google Drive Word/PDF Storage state (strictly isolated per selected book)
  const effectiveBookDriveCfg = useMemo(() => {
    if (selectedBook.id !== DEFAULT_BOOK_ID) {
      return selectedBook.googleDriveConfig;
    }
    return selectedBook.googleDriveConfig || settings.googleDriveConfig;
  }, [selectedBook, settings.googleDriveConfig]);

  const [gdEnabled, setGdEnabled] = useState<boolean>(Boolean(effectiveBookDriveCfg?.enabled));
  const [gdAuthMode, setGdAuthMode] = useState<GoogleDriveAuthMode>(
    effectiveBookDriveCfg?.authMode || 'apps_script'
  );
  const [gdFolderUrl, setGdFolderUrl] = useState<string>(
    effectiveBookDriveCfg?.folderUrl || effectiveBookDriveCfg?.folderId || ''
  );
  const [gdAppsScriptUrl, setGdAppsScriptUrl] = useState<string>('');
  const [gdServiceAccountJson, setGdServiceAccountJson] = useState<string>('');
  const [gdSaving, setGdSaving] = useState(false);
  const [gdSaveMessage, setGdSaveMessage] = useState<{ success: boolean; message: string } | null>(null);
  const [gdTesting, setGdTesting] = useState(false);
  const [gdTestResult, setGdTestResult] = useState<{
    success: boolean;
    message: string;
    folderName?: string;
  } | null>(null);
  const [copiedDriveScript, setCopiedDriveScript] = useState(false);

  // Google Sheets state (Master + per-book optional)
  const [gsEnabled, setGsEnabled] = useState(settings.googleSheetsConfig?.enabled || false);
  const [gsWebhookUrl, setGsWebhookUrl] = useState(settings.googleSheetsConfig?.webhookUrl || '');
  const [gsTestBookId, setGsTestBookId] = useState<string>(selectedBook.id);
  const [gsSaving, setGsSaving] = useState(false);
  const [gsTesting, setGsTesting] = useState(false);
  const [gsTestResult, setGsTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [gsSyncingAll, setGsSyncingAll] = useState(false);
  const [gsSyncResult, setGsSyncResult] = useState<string | null>(null);
  const [copiedScript, setCopiedScript] = useState(false);
  const [bookGsEnabled, setBookGsEnabled] = useState<boolean>(Boolean(selectedBook.googleSheetsConfig?.enabled));
  const [bookGsWebhookUrl, setBookGsWebhookUrl] = useState<string>(selectedBook.googleSheetsConfig?.webhookUrl || '');
  const [bookGsSaving, setBookGsSaving] = useState(false);
  const [bookGsMessage, setBookGsMessage] = useState<string | null>(null);

  // Dirty-state & optimistic concurrency tracking
  const [generalDirty, setGeneralDirty] = useState(false);
  const [driveDirty, setDriveDirty] = useState(false);
  const [bookGsDirty, setBookGsDirty] = useState(false);
  const [settingsConflict, setSettingsConflict] = useState<{
    message: string;
    serverVersion?: number;
    serverBook?: RegistryBook;
    retryAction?: () => Promise<void>;
  } | null>(null);
  const loadedBookSnapshotRef = useRef<{ bookId: string; version: number } | null>(null);
  const prevServerYearRef = useRef<number>(currentYear);

  useEffect(() => {
    if (prevServerYearRef.current !== currentYear) {
      if (seqYear === prevServerYearRef.current) {
        setSeqYear(currentYear);
      }
      prevServerYearRef.current = currentYear;
    }
  }, [currentYear, seqYear]);

  const applySelectedBookToForms = useCallback(
    (book: RegistryBook, isBookSwitch = true) => {
      const enabledForBook = getBookEnabledTypes(book, documentTypes);
      setBookTitleInput(book.name);
      setUnitName(book.unitName);
      setUnitCode(book.unitCode);
      setNumberFormat(book.numberFormat);
      setBookEnabledTypes(enabledForBook.map((t) => t.code));
      setTargetMode(book.numberMode);
      // Khi chỉ làm mới dữ liệu nền trên cùng một sổ, giữ nguyên loại văn bản người dùng đang chọn.
      // Khi đổi sang sổ khác, giữ lại loại đang chọn nếu sổ mới cũng bật loại đó; nếu không thì chọn loại hợp lệ đầu tiên.
      setSeqTypeCode((prevType) => {
        if (book.numberMode === 'shared') return 'SHARED';
        const matchInNewBook = enabledForBook.find(
          (t) => t.code.trim().toUpperCase() === prevType.trim().toUpperCase()
        );
        if (!isBookSwitch && matchInNewBook) return matchInNewBook.code;
        if (matchInNewBook) return matchInNewBook.code;
        return enabledForBook[0]?.code || 'KH';
      });
      setModeConfirmText('');
      if (isBookSwitch) {
        setModeError(null);
        setModeSuccess(false);
        setGeneralSuccess(false);
        setGdSaveMessage(null);
        setGdTestResult(null);
        setBookGsMessage(null);
        setSeqSaveSuccessNotice(null);
      }

      const cfg =
        book.id !== DEFAULT_BOOK_ID
          ? book.googleDriveConfig
          : book.googleDriveConfig || settings.googleDriveConfig;
      setGdEnabled(Boolean(cfg?.enabled));
      setGdAuthMode(cfg?.authMode || 'apps_script');
      setGdFolderUrl(cfg?.folderUrl || cfg?.folderId || '');
      setGdAppsScriptUrl('');
      setGdServiceAccountJson('');

      setBookGsEnabled(Boolean(book.googleSheetsConfig?.enabled));
      setBookGsWebhookUrl(book.googleSheetsConfig?.webhookUrl || '');
      setGsTestBookId(book.id);

      setGeneralDirty(false);
      setDriveDirty(false);
      setBookGsDirty(false);
      setSettingsConflict(null);
      loadedBookSnapshotRef.current = {
        bookId: book.id,
        version: book.configVersion || 1,
      };
    },
    [documentTypes, settings.googleDriveConfig]
  );

  // Sync form states cleanly when switching selectedBook or when remote config changes while form is clean
  useEffect(() => {
    const incomingVersion = selectedBook.configVersion || 1;
    const prev = loadedBookSnapshotRef.current;

    if (!prev || prev.bookId !== selectedBook.id) {
      applySelectedBookToForms(selectedBook, true);
      return;
    }

    if (incomingVersion !== prev.version) {
      const hasUnsavedChanges =
        generalDirty ||
        driveDirty ||
        bookGsDirty ||
        targetMode !== selectedBook.numberMode;
      if (hasUnsavedChanges) {
        setSettingsConflict((existing) => ({
          message:
            'Cài đặt đã thay đổi trên thiết bị khác. Hãy tải lại cấu hình mới nhất từ máy chủ hoặc xác nhận ghi đè bằng nội dung bạn đang sửa.',
          serverVersion: incomingVersion,
          serverBook: selectedBook,
          retryAction: existing?.retryAction,
        }));
      } else {
        applySelectedBookToForms(selectedBook, false);
      }
    }
  }, [
    selectedBook,
    applySelectedBookToForms,
    generalDirty,
    driveDirty,
    bookGsDirty,
    targetMode,
  ]);

  // Canonical series identity: bookId + year + typeCode
  const enabledTypesForSelectedBook = useMemo(
    () => getBookEnabledTypes(selectedBook, documentTypes),
    [selectedBook, documentTypes]
  );

  const effectiveSeqTypeCode = useMemo(() => {
    if (selectedBook.numberMode === 'shared') return 'SHARED';
    const found = enabledTypesForSelectedBook.find(
      (t) => t.code.trim().toUpperCase() === seqTypeCode.trim().toUpperCase()
    );
    return found ? found.code : enabledTypesForSelectedBook[0]?.code || 'KH';
  }, [selectedBook.numberMode, enabledTypesForSelectedBook, seqTypeCode]);

  const activeSeriesId = `${selectedBook.id}::${seqYear}::${effectiveSeqTypeCode}`;
  const activeSeriesKeyDisplay = `${selectedBook.code}:${seqYear}:${effectiveSeqTypeCode}`;

  const activeSeriesIdRef = useRef<string>(activeSeriesId);
  activeSeriesIdRef.current = activeSeriesId;
  const activeBookIdRef = useRef<string>(selectedBook.id);
  activeBookIdRef.current = selectedBook.id;
  const draftStartInputBySeriesRef = useRef<Record<string, string>>(draftStartInputBySeries);
  draftStartInputBySeriesRef.current = draftStartInputBySeries;

  const readRequestSeqRef = useRef<number>(0);
  const latestCompletedReadReqBySeriesRef = useRef<Record<string, number>>({});
  const readAbortControllerRef = useRef<AbortController | null>(null);

  const loadStartSequenceSeries = useCallback(
    async (bookId: string, year: number, typeCode: string) => {
      if (isConfigBlocked || !bookId || !typeCode) return;

      if (readAbortControllerRef.current) {
        readAbortControllerRef.current.abort();
      }
      const controller = new AbortController();
      readAbortControllerRef.current = controller;

      const reqId = ++readRequestSeqRef.current;
      const targetSeriesId = `${bookId}::${year}::${typeCode}`;

      if (activeSeriesIdRef.current === targetSeriesId) {
        setSeqConfigLoading(true);
        setSeqConfigLoadError(null);
      }

      try {
        const res = await api.getStartSequenceConfig(bookId, year, typeCode, controller.signal);
        const prevCompleted = latestCompletedReadReqBySeriesRef.current[targetSeriesId] || 0;
        if (reqId >= prevCompleted) {
          latestCompletedReadReqBySeriesRef.current[targetSeriesId] = reqId;
          setSavedSeriesMap((prev) => ({
            ...prev,
            [targetSeriesId]: res,
          }));
        }

        if (reqId === readRequestSeqRef.current && activeSeriesIdRef.current === targetSeriesId) {
          setSeqConfigLoading(false);
          setSeqConfigLoadError(null);
        }
      } catch (err: any) {
        if (err?.name === 'AbortError') return;
        if (reqId === readRequestSeqRef.current && activeSeriesIdRef.current === targetSeriesId) {
          setSeqConfigLoading(false);
          setSeqConfigLoadError(err?.message || 'Không thể tải cấu hình dãy số từ máy chủ.');
        }
      }
    },
    [isConfigBlocked]
  );

  useEffect(() => {
    if (activeSubTab === 'numbering' && !isConfigBlocked) {
      loadStartSequenceSeries(selectedBook.id, seqYear, effectiveSeqTypeCode);
    }
    return () => {
      if (readAbortControllerRef.current) {
        readAbortControllerRef.current.abort();
      }
    };
  }, [
    activeSubTab,
    isConfigBlocked,
    selectedBook.id,
    selectedBook.configVersion,
    selectedBook.numberMode,
    selectedBook.numberFormat,
    selectedBook.unitCode,
    seqYear,
    effectiveSeqTypeCode,
    loadStartSequenceSeries,
  ]);

  const activeSavedConfig = savedSeriesMap[activeSeriesId] || null;
  const savedInputString =
    activeSavedConfig?.isConfigured && activeSavedConfig.customStartSequence != null
      ? String(activeSavedConfig.customStartSequence)
      : '';
  const hasSeriesDraft = Object.prototype.hasOwnProperty.call(
    draftStartInputBySeries,
    activeSeriesId
  );
  const currentStartInputValue = hasSeriesDraft
    ? draftStartInputBySeries[activeSeriesId]
    : savedInputString;
  const isSeriesDirty =
    hasSeriesDraft && draftStartInputBySeries[activeSeriesId].trim() !== savedInputString;
  const isActiveSeriesSaving = Boolean(savingSeriesMap[activeSeriesId]);
  const activeSeriesSaveError = seqSaveErrorBySeries[activeSeriesId] || null;
  const parsedStartInputNum = Number(currentStartInputValue.trim());
  const isStartInputValid =
    currentStartInputValue.trim() !== '' &&
    Number.isInteger(parsedStartInputNum) &&
    parsedStartInputNum >= 1;

  // Microsoft Graph
  const [msEnabled, setMsEnabled] = useState(settings.msGraphConfig.enabled);
  const [msTenantId, setMsTenantId] = useState(settings.msGraphConfig.tenantId);
  const [msClientId, setMsClientId] = useState(settings.msGraphConfig.clientId);
  const [msClientSecret, setMsClientSecret] = useState(settings.msGraphConfig.clientSecret);
  const [msDriveId, setMsDriveId] = useState(settings.msGraphConfig.driveId);
  const [msFilePath, setMsFilePath] = useState(settings.msGraphConfig.filePath);
  const [msTableName, setMsTableName] = useState(settings.msGraphConfig.tableName);
  const [msTesting, setMsTesting] = useState(false);
  const [msTestResult, setMsTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [msSaving, setMsSaving] = useState(false);
  const [msSyncingAll, setMsSyncingAll] = useState(false);
  const [msSyncResult, setMsSyncResult] = useState<string | null>(null);

  // User Management & Per-Book Permissions
  const [usersList, setUsersList] = useState<User[]>([]);
  const [loadingUsers, setLoadingUsers] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newFullName, setNewFullName] = useState('');
  const [newUserPass, setNewUserPass] = useState('');
  const [newBookPermissions, setNewBookPermissions] = useState<Record<string, BookPermission[]>>({
    [DEFAULT_BOOK_ID]: ['view', 'request_number'],
  });
  const [userSaving, setUserSaving] = useState(false);
  const [userError, setUserError] = useState<string | null>(null);

  const [editingPermUserId, setEditingPermUserId] = useState<string | null>(null);
  const [editingBookPermissions, setEditingBookPermissions] = useState<Record<string, BookPermission[]>>({});
  const [savingPermUserId, setSavingPermUserId] = useState<string | null>(null);
  const [auditLogs, setAuditLogs] = useState<SystemAuditLog[]>([]);
  const [loadingAuditLogs, setLoadingAuditLogs] = useState(false);

  const togglePermissionInMap = (
    map: Record<string, BookPermission[]>,
    bookId: string,
    perm: BookPermission
  ): Record<string, BookPermission[]> => {
    const current = new Set(map[bookId] || []);
    if (perm === 'view' && current.has('issue_number') && current.has('view')) {
      // Quyền Lấy số luôn kèm quyền Chỉ xem trên cùng sổ
      return {
        ...map,
        [bookId]: normalizeBookPermissionList(Array.from(current)),
      };
    }
    if (current.has(perm)) current.delete(perm);
    else current.add(perm);
    return {
      ...map,
      [bookId]: normalizeBookPermissionList(Array.from(current)),
    };
  };

  // Fetch users & audit logs when opening users tab
  const handleOpenUsersTab = async () => {
    setActiveSubTab('users');
    if (!isAdmin) return;
    setLoadingUsers(true);
    setLoadingAuditLogs(true);
    try {
      const [usersRes, logsRes] = await Promise.all([
        api.getUsers(),
        api.getAuditLogs().catch(() => ({ logs: [] as SystemAuditLog[] })),
      ]);
      setUsersList(usersRes.users);
      setAuditLogs(logsRes.logs);
    } catch {
      // ignore
    } finally {
      setLoadingUsers(false);
      setLoadingAuditLogs(false);
    }
  };

  const handleConflictError = (err: any, retryAction: () => Promise<void>): boolean => {
    if (err?.status === 409 && err?.code === 'SETTINGS_CONFLICT') {
      setSettingsConflict({
        message:
          err.message ||
          'Cài đặt đã thay đổi trên thiết bị khác. Vui lòng tải lại cấu hình mới hoặc xác nhận ghi đè.',
        serverVersion: err.serverVersion,
        serverBook: err.serverBook,
        retryAction,
      });
      return true;
    }
    return false;
  };

  // Create new Registry Book (Admin only, via Drawer)
  const handleCreateBookSubmit = async (values: CreateBookFormValues): Promise<boolean> => {
    if (!isAdmin || isConfigBlocked) return false;
    setBookSaving(true);
    setCreateBookServerError(null);
    setBookMessage(null);
    try {
      const res = await api.createBook({
        code: values.code || undefined,
        name: values.name,
        unitName: values.unitName,
        unitCode: values.unitCode,
        isShared: values.isShared,
        numberMode: values.numberMode,
        numberFormat: values.numberFormat || '{so}-{loai}/{don_vi}',
        enabledTypeCodes: values.enabledTypeCodes,
      });
      setCreateBookDrawerOpen(false);
      setSelectedBookId(res.book.id);
      onSelectBookId(res.book.id);
      setBookMessage({
        success: true,
        text: `Đã tạo sổ "${res.book.name}" (Mã sổ: ${res.book.code} · Ký hiệu: ${res.book.unitCode}). Bạn có thể cấu hình Drive riêng và phân quyền tài khoản cho sổ mới.`,
      });
      onBooksUpdated?.();
      onRefreshData();
      return true;
    } catch (err: any) {
      setCreateBookServerError(err.message || 'Lỗi khi tạo sổ văn bản mới');
      return false;
    } finally {
      setBookSaving(false);
    }
  };

  const handleConfirmToggleBookLock = async () => {
    const book = toggleLockConfirmBook;
    if (!book || !isAdmin || isConfigBlocked || book.id === DEFAULT_BOOK_ID) return;
    setToggleLockSaving(true);
    try {
      await api.updateBook(book.id, {
        isActive: !book.isActive,
        expectedVersion: book.configVersion,
      });
      setToggleLockConfirmBook(null);
      setBookMessage({
        success: true,
        text: book.isActive
          ? `Đã tạm khóa sổ "${book.name}".`
          : `Đã mở khóa sổ "${book.name}".`,
      });
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      setToggleLockConfirmBook(null);
      if (
        !handleConflictError(err, async () => {
          await api.updateBook(book.id, { isActive: !book.isActive, forceOverwrite: true });
          onBooksUpdated?.();
          onRefreshData();
        })
      ) {
        setBookMessage({ success: false, text: err.message || 'Lỗi khi cập nhật trạng thái sổ' });
      }
    } finally {
      setToggleLockSaving(false);
    }
  };

  const hasUnsavedChangesForSelectedBook =
    generalDirty ||
    driveDirty ||
    bookGsDirty ||
    targetMode !== selectedBook.numberMode ||
    isSeriesDirty;

  const commitBookSwitch = (targetBookId: string, nextTab?: SettingsSubTab) => {
    setSelectedBookId(targetBookId);
    onSelectBookId(targetBookId);
    if (nextTab) {
      if (nextTab === 'users') {
        handleOpenUsersTab();
      } else {
        setActiveSubTab(nextTab);
      }
    }
    setPendingBookSwitch(null);
    setSwitchBookError(null);
  };

  const requestSelectBook = (targetBookId: string, nextTab?: SettingsSubTab) => {
    if (targetBookId === selectedBook.id) {
      if (nextTab) {
        if (nextTab === 'users') handleOpenUsersTab();
        else setActiveSubTab(nextTab);
      }
      return;
    }
    if (hasUnsavedChangesForSelectedBook) {
      const target = books.find((b) => b.id === targetBookId);
      setSwitchBookError(null);
      setPendingBookSwitch({
        targetBookId,
        targetBookName: target?.name || 'Sổ mới',
        currentBookName: selectedBook.name,
        nextTab,
      });
      return;
    }
    commitBookSwitch(targetBookId, nextTab);
  };

  const handleDiscardAndSwitchBook = () => {
    if (!pendingBookSwitch) return;
    setDraftStartInputBySeries((prev) => {
      const next = { ...prev };
      for (const k of Object.keys(next)) {
        if (k.startsWith(`${selectedBook.id}::`)) {
          delete next[k];
        }
      }
      return next;
    });
    setGeneralDirty(false);
    setDriveDirty(false);
    setBookGsDirty(false);
    setTargetMode(selectedBook.numberMode);
    setGeneralError(null);
    commitBookSwitch(pendingBookSwitch.targetBookId, pendingBookSwitch.nextTab);
  };

  const handleSaveAndSwitchBook = async () => {
    if (!pendingBookSwitch || switchingBookSaving) return;
    setSwitchingBookSaving(true);
    setSwitchBookError(null);
    try {
      let currentVersion =
        loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion;

      if (generalDirty) {
        if (bookEnabledTypes.length === 0) {
          throw new Error('Vui lòng chọn ít nhất 1 loại văn bản được sử dụng trong sổ này.');
        }
        const res = await api.updateBook(selectedBook.id, {
          name: bookTitleInput.trim() || selectedBook.name,
          unitName: unitName.trim(),
          unitCode: unitCode.trim().toUpperCase(),
          numberFormat: numberFormat.trim(),
          enabledTypeCodes: bookEnabledTypes,
          expectedVersion: currentVersion,
        });
        currentVersion = res.book.configVersion || (currentVersion || 1) + 1;
        setGeneralDirty(false);
      }

      if (targetMode !== selectedBook.numberMode) {
        const res = await api.updateNumberMode(
          targetMode,
          modeConfirmText.trim(),
          selectedBook.id,
          currentVersion
        );
        currentVersion = (currentVersion || 1) + 1;
        onSettingsUpdated(res.settings);
        setModeConfirmText('');
      }

      if (isSeriesDirty) {
        const trimmed = currentStartInputValue.trim();
        const nextNum = Number(trimmed);
        if (!trimmed || !Number.isInteger(nextNum) || nextNum < 1) {
          throw new Error('Số bắt đầu phải là số nguyên dương từ 1 trở lên.');
        }
        if (activeSavedConfig && nextNum < activeSavedConfig.minAllowedStartSequence) {
          throw new Error(
            `Số bắt đầu (${nextNum}) phải >= ${activeSavedConfig.minAllowedStartSequence}.`
          );
        }
        const seqRes = await api.setStartSequence(
          seqYear,
          effectiveSeqTypeCode,
          nextNum,
          selectedBook.id
        );
        if (seqRes.book?.configVersion) {
          currentVersion = seqRes.book.configVersion;
        }
        setDraftStartInputBySeries((prev) => {
          const next = { ...prev };
          delete next[activeSeriesId];
          return next;
        });
      }

      if (driveDirty) {
        const gdRes = await api.saveGoogleDriveConfig({
          bookId: selectedBook.id,
          enabled: gdEnabled,
          authMode: gdAuthMode,
          folderUrl: gdFolderUrl.trim(),
          appsScriptUrl: gdAppsScriptUrl.trim() || undefined,
          serviceAccountJson: gdServiceAccountJson.trim() || undefined,
          scopeUnitCode: selectedBook.unitCode,
          expectedVersion: currentVersion,
        });
        currentVersion = (currentVersion || 1) + 1;
        if (selectedBook.id === DEFAULT_BOOK_ID) {
          onSettingsUpdated({ ...settings, googleDriveConfig: gdRes.googleDriveConfig });
        }
        setDriveDirty(false);
      }

      if (bookGsDirty && isAdmin) {
        const gsRes = await api.saveBookGoogleSheetsConfig(selectedBook.id, {
          enabled: bookGsEnabled,
          webhookUrl: bookGsWebhookUrl.trim(),
          expectedVersion: currentVersion,
        });
        currentVersion = gsRes.book?.configVersion || (currentVersion || 1) + 1;
        setBookGsDirty(false);
      }

      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: currentVersion || 1,
      };
      onBooksUpdated?.();
      onRefreshData();
      commitBookSwitch(pendingBookSwitch.targetBookId, pendingBookSwitch.nextTab);
    } catch (err: any) {
      setSwitchBookError(err?.message || 'Không thể lưu thay đổi hiện tại.');
    } finally {
      setSwitchingBookSaving(false);
    }
  };

  // 1. Save General Settings for selected book
  const saveGeneralRequest = async (forceOverwrite = false) => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    if (bookEnabledTypes.length === 0) {
      setSettingsConflict(null);
      setGeneralError('Vui lòng chọn ít nhất 1 loại văn bản được sử dụng trong sổ này.');
      return;
    }
    setGeneralSaving(true);
    setGeneralSuccess(false);
    setGeneralError(null);
    try {
      const res = await api.updateBook(selectedBook.id, {
        name: bookTitleInput.trim() || selectedBook.name,
        unitName: unitName.trim(),
        unitCode: unitCode.trim().toUpperCase(),
        numberFormat: numberFormat.trim(),
        enabledTypeCodes: bookEnabledTypes,
        expectedVersion: loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
        forceOverwrite,
      });
      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: res.book.configVersion || (selectedBook.configVersion || 1) + 1,
      };
      setGeneralDirty(false);
      setSettingsConflict(null);
      onBooksUpdated?.();
      onRefreshData();
      setGeneralSuccess(true);
      setTimeout(() => setGeneralSuccess(false), 3000);
    } catch (err: any) {
      if (!handleConflictError(err, () => saveGeneralRequest(true))) {
        setGeneralError(err.message || 'Lỗi khi lưu cài đặt sổ');
      }
    } finally {
      setGeneralSaving(false);
    }
  };

  const handleSaveGeneral = async (e: React.FormEvent) => {
    e.preventDefault();
    await saveGeneralRequest(false);
  };

  // 2. Change Number Mode for selected book
  const saveModeRequest = async (forceOverwrite = false) => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    setModeSaving(true);
    setModeError(null);
    setModeSuccess(false);
    try {
      const res = await api.updateNumberMode(
        targetMode,
        modeConfirmText.trim(),
        selectedBook.id,
        loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
        forceOverwrite
      );
      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: (selectedBook.configVersion || 1) + 1,
      };
      setSettingsConflict(null);
      onSettingsUpdated(res.settings);
      onBooksUpdated?.();
      onRefreshData();
      setModeSuccess(true);
      setModeConfirmText('');
      loadStartSequenceSeries(
        selectedBook.id,
        seqYear,
        targetMode === 'shared' ? 'SHARED' : effectiveSeqTypeCode
      );
      setTimeout(() => setModeSuccess(false), 3000);
    } catch (err: any) {
      if (!handleConflictError(err, () => saveModeRequest(true))) {
        setModeError(err.message || 'Lỗi khi đổi chế độ dãy số');
      }
    } finally {
      setModeSaving(false);
    }
  };

  const handleChangeMode = async (e: React.FormEvent) => {
    e.preventDefault();
    await saveModeRequest(false);
  };

  // 3. Set Start Sequence for selected (book, year, typeCode) with out-of-order save isolation
  const handleSetStartSequence = async (e: React.FormEvent) => {
    e.preventDefault();
    if (
      !canManageSelectedBook ||
      isConfigBlocked ||
      seqConfigLoading ||
      Boolean(seqConfigLoadError)
    ) {
      return;
    }

    const targetBookId = selectedBook.id;
    const targetBookName = selectedBook.name;
    const targetYear = seqYear;
    const targetTypeCode = effectiveSeqTypeCode;
    const targetSeriesId = `${targetBookId}::${targetYear}::${targetTypeCode}`;
    const rawDraft = currentStartInputValue;
    const trimmed = rawDraft.trim();
    const nextNum = Number(trimmed);

    if (!trimmed || !Number.isInteger(nextNum) || nextNum < 1) {
      setSeqSaveErrorBySeries((prev) => ({
        ...prev,
        [targetSeriesId]: 'Vui lòng nhập số bắt đầu là số nguyên dương từ 1 trở lên.',
      }));
      return;
    }

    if (activeSavedConfig && nextNum < activeSavedConfig.minAllowedStartSequence) {
      setSeqSaveErrorBySeries((prev) => ({
        ...prev,
        [targetSeriesId]: `Số bắt đầu (${nextNum}) phải lớn hơn số lớn nhất đã cấp/hủy trong dãy (${activeSavedConfig.maxIssuedSequence}). Mức tối thiểu là ${activeSavedConfig.minAllowedStartSequence}.`,
      }));
      return;
    }

    setSavingSeriesMap((prev) => ({ ...prev, [targetSeriesId]: true }));
    setSeqSaveErrorBySeries((prev) => ({ ...prev, [targetSeriesId]: null }));
    setSeqSaveSuccessNotice((prev) => (prev?.seriesId === targetSeriesId ? null : prev));

    try {
      const res = await api.setStartSequence(targetYear, targetTypeCode, nextNum, targetBookId);

      // Cập nhật riêng cấu hình đã lưu của đúng dãy vừa lưu, không ảnh hưởng dãy đang xem nếu người dùng đã chuyển loại
      const updatedConfig: StartSequenceSeriesConfig = {
        seriesKey: res.seriesKey,
        storageKey: res.storageKey,
        bookId: res.bookId,
        bookCode: res.bookCode,
        year: res.year,
        typeCode: res.typeCode,
        isConfigured: res.isConfigured,
        customStartSequence: res.customStartSequence,
        maxIssuedSequence: res.maxIssuedSequence,
        minAllowedStartSequence: res.minAllowedStartSequence,
        nextExpectedSeq: res.nextExpectedSeq,
        previewNumber: res.previewNumber,
        configVersion: res.configVersion,
        customStartSequences: res.customStartSequences,
      };
      latestCompletedReadReqBySeriesRef.current[targetSeriesId] = ++readRequestSeqRef.current;
      setSavedSeriesMap((prev) => ({
        ...prev,
        [targetSeriesId]: updatedConfig,
      }));

      // Chỉ xóa bản nháp của dãy vừa lưu nếu người dùng chưa nhập giá trị khác trong lúc chờ phản hồi
      if (draftStartInputBySeriesRef.current[targetSeriesId] === rawDraft) {
        setDraftStartInputBySeries((prev) => {
          const next = { ...prev };
          delete next[targetSeriesId];
          return next;
        });
      }

      if (activeBookIdRef.current === targetBookId && res.book?.configVersion) {
        loadedBookSnapshotRef.current = {
          bookId: targetBookId,
          version: res.book.configVersion,
        };
      }

      const isStillViewingSameSeries = activeSeriesIdRef.current === targetSeriesId;
      setSeqSaveSuccessNotice({
        seriesId: targetSeriesId,
        message: isStillViewingSameSeries
          ? `Đã lưu Số bắt đầu đã cài (${String(nextNum).padStart(2, '0')}) cho ${targetBookName} · Năm ${targetYear} · ${targetTypeCode}!`
          : `Đã lưu xong mốc ${targetBookName} · Năm ${targetYear} · ${targetTypeCode} = ${String(nextNum).padStart(2, '0')} (dãy đang chọn vẫn được giữ nguyên).`,
      });

      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      setSeqSaveErrorBySeries((prev) => ({
        ...prev,
        [targetSeriesId]: err.message || 'Lỗi khi lưu số bắt đầu',
      }));
    } finally {
      setSavingSeriesMap((prev) => {
        const next = { ...prev };
        delete next[targetSeriesId];
        return next;
      });
    }
  };

  // 4. Create Document Type (Admin / manage_book)
  const handleCreateDocType = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook || isConfigBlocked) return;
    const normalizedName = newTypeName.trim().replace(/\s+/g, ' ');
    const normalizedCode = newTypeCode.trim().replace(/\s+/g, '').toUpperCase();
    if (!normalizedName || !normalizedCode) {
      setTypeError('Vui lòng nhập đầy đủ tên loại văn bản và ký hiệu viết tắt.');
      return;
    }
    setTypeSaving(true);
    setTypeError(null);
    try {
      await api.createDocumentType({
        name: normalizedName,
        code: normalizedCode,
        color: newTypeColor,
        bookId: selectedBook.id,
      });
      setNewTypeName('');
      setNewTypeCode('');
      setNewTypeColor('blue');
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      if (err?.code === 'TYPE_EXISTS_IN_CATALOG_NOT_ENABLED') {
        try {
          await api.createDocumentType({
            name: normalizedName,
            code: normalizedCode,
            color: newTypeColor,
            bookId: selectedBook.id,
            enableExistingForBook: true,
          });
          setNewTypeName('');
          setNewTypeCode('');
          setNewTypeColor('blue');
          onBooksUpdated?.();
          onRefreshData();
          return;
        } catch (enableErr: any) {
          setTypeError(enableErr.message || 'Lỗi khi bật loại văn bản cho sổ');
          return;
        }
      }
      setTypeError(err.message || 'Lỗi khi tạo loại văn bản');
    } finally {
      setTypeSaving(false);
    }
  };

  const handleUpdateDocTypeColor = async (id: string, color: string) => {
    if (!isAdmin || isConfigBlocked) return;
    try {
      await api.updateDocumentType(id, { color });
      onRefreshData();
    } catch (err: any) {
      setTypeError(err.message || 'Lỗi khi cập nhật màu sắc loại văn bản');
    }
  };

  const handleDeleteDocType = (id: string, code: string) => {
    if (!isAdmin || isConfigBlocked) return;
    setConfirmDeleteRequest({
      title: `Xóa loại văn bản "${code}"?`,
      description: 'Loại văn bản này sẽ bị xóa khỏi danh mục hệ thống (các văn bản đã cấp số trước đó vẫn được bảo toàn).',
      confirmLabel: 'Xác nhận xóa',
      onConfirm: async () => {
        setConfirmDeleteBusy(true);
        try {
          await api.deleteDocumentType(id);
          setConfirmDeleteRequest(null);
          onRefreshData();
        } catch (err: any) {
          setTypeError(err.message || 'Lỗi khi xóa loại văn bản');
          setConfirmDeleteRequest(null);
        } finally {
          setConfirmDeleteBusy(false);
        }
      },
    });
  };

  // 5. Book Signers management (Admin or manage_book)
  const displayedBookSigners = useMemo(() => {
    if (selectedBook.signers && selectedBook.signers.length > 0) {
      return selectedBook.signers;
    }
    return signers;
  }, [selectedBook.signers, signers]);

  const saveBookSignerRequest = async (updatedList: SignerConfig[], forceOverwrite = false, resetInput = false) => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    setSignerSaving(true);
    setSignerError(null);
    try {
      const res = await api.updateBook(selectedBook.id, {
        signers: updatedList,
        expectedVersion: loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
        forceOverwrite,
      });
      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: res.book.configVersion || (selectedBook.configVersion || 1) + 1,
      };
      setSettingsConflict(null);
      if (resetInput) {
        setNewSignerName('');
        setNewSignerRole('');
      }
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      if (!handleConflictError(err, () => saveBookSignerRequest(updatedList, true, resetInput))) {
        setSignerError(err.message || 'Lỗi khi cập nhật người ký');
      }
    } finally {
      setSignerSaving(false);
    }
  };

  const handleCreateBookSigner = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook || isConfigBlocked) return;
    const nextSigner: SignerConfig = {
      id: `sig-${Date.now()}`,
      name: newSignerName.trim(),
      role: newSignerRole.trim(),
    };
    const updatedList = [...displayedBookSigners, nextSigner];
    await saveBookSignerRequest(updatedList, false, true);
    if (isAdmin && selectedBook.id === DEFAULT_BOOK_ID) {
      await api.createSigner({ name: nextSigner.name, role: nextSigner.role }).catch(() => {});
    }
  };

  const handleDeleteBookSigner = (signerId: string) => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    const target = displayedBookSigners.find((s) => s.id === signerId);
    setConfirmDeleteRequest({
      title: `Xóa người ký "${target?.name || ''}"?`,
      description: `Người ký này sẽ bị gỡ khỏi danh sách gợi ý của sổ "${selectedBook.name}".`,
      confirmLabel: 'Xóa người ký',
      onConfirm: async () => {
        setConfirmDeleteBusy(true);
        try {
          const updatedList = displayedBookSigners.filter((s) => s.id !== signerId);
          await saveBookSignerRequest(updatedList, false, false);
          if (isAdmin && selectedBook.id === DEFAULT_BOOK_ID) {
            await api.deleteSigner(signerId).catch(() => {});
          }
          setConfirmDeleteRequest(null);
        } finally {
          setConfirmDeleteBusy(false);
        }
      },
    });
  };

  // 6. Google Sheets Config & Sync Handlers
  const handleSaveGoogleSheets = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin || isConfigBlocked) return;
    setGsSaving(true);
    try {
      const res = await api.saveGoogleSheetsConfig({
        enabled: gsEnabled,
        webhookUrl: gsWebhookUrl.trim(),
      });
      onSettingsUpdated({ ...settings, googleSheetsConfig: res.googleSheetsConfig });
      setGsSyncResult('Đã lưu cấu hình Google Trang tính thành công!');
    } catch (err: any) {
      setGsSyncResult(`Lỗi: ${err.message || 'Lỗi khi lưu cấu hình Google Trang tính'}`);
    } finally {
      setGsSaving(false);
    }
  };

  const handleTestGoogleSheets = async () => {
    setGsTesting(true);
    setGsTestResult(null);
    try {
      const res = await api.testGoogleSheetsConnection(gsWebhookUrl.trim(), gsTestBookId || selectedBook.id);
      setGsTestResult(res);
    } catch (err: any) {
      setGsTestResult({ success: false, message: err.message || 'Lỗi kiểm tra Google Sheets' });
    } finally {
      setGsTesting(false);
    }
  };

  const handleSyncAllGoogleSheets = async (targetBookId?: string) => {
    setGsSyncingAll(true);
    setGsSyncResult(null);
    try {
      const res = await api.syncAllGoogleSheets(targetBookId);
      setGsSyncResult(res.message);
      onRefreshData();
    } catch (err: any) {
      setGsSyncResult('Lỗi: ' + err.message);
    } finally {
      setGsSyncingAll(false);
    }
  };

  const saveBookGoogleSheetsRequest = async (forceOverwrite = false) => {
    if (!isAdmin || isConfigBlocked) return;
    setBookGsSaving(true);
    setBookGsMessage(null);
    try {
      const res = await api.saveBookGoogleSheetsConfig(selectedBook.id, {
        enabled: bookGsEnabled,
        webhookUrl: bookGsWebhookUrl.trim(),
        expectedVersion: loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
        forceOverwrite,
      });
      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: res.book?.configVersion || (selectedBook.configVersion || 1) + 1,
      };
      setBookGsDirty(false);
      setSettingsConflict(null);
      onBooksUpdated?.();
      onRefreshData();
      setBookGsMessage(`Đã lưu webhook Google Sheets riêng cho sổ "${selectedBook.name}".`);
    } catch (err: any) {
      if (!handleConflictError(err, () => saveBookGoogleSheetsRequest(true))) {
        setBookGsMessage(`Lỗi: ${err.message || 'Không thể lưu cấu hình Google Sheets của sổ.'}`);
      }
    } finally {
      setBookGsSaving(false);
    }
  };

  const handleSaveBookGoogleSheets = async (e: React.FormEvent) => {
    e.preventDefault();
    await saveBookGoogleSheetsRequest(false);
  };

  // 6b. Google Drive Storage Handlers (per selected book)
  const saveGoogleDriveRequest = async (forceOverwrite = false) => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    setGdSaving(true);
    setGdSaveMessage(null);
    try {
      const res = await api.saveGoogleDriveConfig({
        bookId: selectedBook.id,
        enabled: gdEnabled,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim(),
        appsScriptUrl: gdAppsScriptUrl.trim() || undefined,
        serviceAccountJson: gdServiceAccountJson.trim() || undefined,
        scopeUnitCode: selectedBook.unitCode,
        expectedVersion: loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
        forceOverwrite,
      });
      loadedBookSnapshotRef.current = {
        bookId: selectedBook.id,
        version: (selectedBook.configVersion || 1) + 1,
      };
      setDriveDirty(false);
      setSettingsConflict(null);
      if (selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      onRefreshData();
      setGdAppsScriptUrl('');
      setGdServiceAccountJson('');
      setGdSaveMessage({
        success: true,
        message: `Đã lưu cấu hình thư mục Google Drive cho "${selectedBook.name}". Hãy bấm "Kiểm tra kết nối" để xác nhận quyền tải tệp.`,
      });
    } catch (err: any) {
      if (!handleConflictError(err, () => saveGoogleDriveRequest(true))) {
        setGdSaveMessage({
          success: false,
          message: err.message || 'Lỗi khi lưu cấu hình Google Drive',
        });
      }
    } finally {
      setGdSaving(false);
    }
  };

  const handleSaveGoogleDrive = async (e: React.FormEvent) => {
    e.preventDefault();
    await saveGoogleDriveRequest(false);
  };

  const handleClearDriveCredential = async (type: 'apps_script' | 'service_account') => {
    if (!canManageSelectedBook || isConfigBlocked) return;
    setGdSaving(true);
    setGdSaveMessage(null);
    try {
      const res = await api.saveGoogleDriveConfig({
        bookId: selectedBook.id,
        enabled: gdEnabled,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim(),
        clearAppsScriptUrl: type === 'apps_script',
        clearServiceAccountJson: type === 'service_account',
        scopeUnitCode: selectedBook.unitCode,
        expectedVersion: loadedBookSnapshotRef.current?.version ?? selectedBook.configVersion,
      });
      if (selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      if (type === 'apps_script') setGdAppsScriptUrl('');
      else setGdServiceAccountJson('');
      setGdSaveMessage({
        success: true,
        message: 'Đã xóa thông tin xác thực cũ.',
      });
    } catch (err: any) {
      setGdSaveMessage({
        success: false,
        message: err.message || 'Không thể xóa thông tin xác thực',
      });
    } finally {
      setGdSaving(false);
    }
  };

  const handleTestGoogleDrive = async () => {
    if (!canManageSelectedBook) return;
    setGdTesting(true);
    setGdTestResult(null);
    try {
      const res = await api.testGoogleDriveConnection({
        bookId: selectedBook.id,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim() || undefined,
        appsScriptUrl: gdAppsScriptUrl.trim() || undefined,
        serviceAccountJson: gdServiceAccountJson.trim() || undefined,
      });
      if (res.googleDriveConfig && selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      setGdTestResult({
        success: res.success,
        folderName: res.folderName,
        message: res.message,
      });
    } catch (err: any) {
      setGdTestResult({
        success: false,
        message: err.message || 'Kiểm tra kết nối thư mục Google Drive thất bại.',
      });
    } finally {
      setGdTesting(false);
    }
  };

  const handleServiceAccountFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      setGdServiceAccountJson(text);
      setDriveDirty(true);
      setGdSaveMessage(null);
    } catch {
      setGdSaveMessage({
        success: false,
        message: 'Không đọc được nội dung tệp JSON đã chọn.',
      });
    }
  };

  const [msSaveMessage, setMsSaveMessage] = useState<{ success: boolean; message: string } | null>(null);
  const [userFeedback, setUserFeedback] = useState<{ success: boolean; message: string } | null>(null);

  // 7. Save Microsoft Graph Config
  const handleSaveMsGraph = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setMsSaving(true);
    setMsSaveMessage(null);
    try {
      const res = await api.saveSyncConfig({
        enabled: msEnabled,
        tenantId: msTenantId,
        clientId: msClientId,
        clientSecret: msClientSecret,
        driveId: msDriveId,
        filePath: msFilePath,
        tableName: msTableName,
      });
      onSettingsUpdated({ ...settings, msGraphConfig: res.msGraphConfig });
      setMsSaveMessage({ success: true, message: 'Đã lưu cấu hình OneDrive / SharePoint!' });
    } catch (err: any) {
      setMsSaveMessage({ success: false, message: err.message || 'Lỗi khi lưu cấu hình' });
    } finally {
      setMsSaving(false);
    }
  };

  const handleTestMsConnection = async () => {
    setMsTesting(true);
    setMsTestResult(null);
    try {
      const res = await api.testSyncConnection();
      setMsTestResult(res);
    } catch (err: any) {
      setMsTestResult({ success: false, message: err.message || 'Lỗi kết nối' });
    } finally {
      setMsTesting(false);
    }
  };

  const handleSyncAll = async () => {
    setMsSyncingAll(true);
    setMsSyncResult(null);
    try {
      const res = await api.syncAllPending();
      setMsSyncResult(res.message);
      onRefreshData();
    } catch (err: any) {
      setMsSyncResult('Lỗi: ' + err.message);
    } finally {
      setMsSyncingAll(false);
    }
  };

  // 8. User Management Handlers (Admin Only)
  const handleCreateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    if (newUserPass.trim().length < 6) {
      setUserError('Mật khẩu khởi tạo phải có tối thiểu 6 ký tự');
      return;
    }
    setUserSaving(true);
    setUserError(null);
    setUserFeedback(null);
    try {
      await api.createUser({
        username: newUsername.trim(),
        fullName: newFullName.trim(),
        password: newUserPass.trim(),
        role: 'issuer',
        bookPermissions: newBookPermissions,
      });
      const createdUsername = newUsername.trim();
      setNewUsername('');
      setNewFullName('');
      setNewUserPass('');
      setNewBookPermissions({ [DEFAULT_BOOK_ID]: ['view', 'request_number'] });
      setUserFeedback({
        success: true,
        message: `Đã tạo tài khoản @${createdUsername} và áp dụng phân quyền theo sổ.`,
      });
      handleOpenUsersTab();
    } catch (err: any) {
      setUserError(err.message || 'Lỗi khi tạo tài khoản');
    } finally {
      setUserSaving(false);
    }
  };

  const handleStartEditPermissions = (u: User) => {
    if (editingPermUserId === u.id) {
      setEditingPermUserId(null);
      return;
    }
    setEditingPermUserId(u.id);
    setEditingBookPermissions({ ...(u.bookPermissions || {}) });
  };

  const handleSaveUserPermissions = async (u: User) => {
    if (!isAdmin) return;
    setSavingPermUserId(u.id);
    setUserFeedback(null);
    try {
      await api.updateUser(u.id, { bookPermissions: editingBookPermissions });
      setEditingPermUserId(null);
      setUserFeedback({
        success: true,
        message: `Đã cập nhật phân quyền theo sổ cho tài khoản @${u.username}.`,
      });
      await handleOpenUsersTab();
    } catch (err: any) {
      setUserFeedback({
        success: false,
        message: err.message || 'Lỗi khi lưu phân quyền tài khoản',
      });
    } finally {
      setSavingPermUserId(null);
    }
  };

  const handleToggleUserStatus = async (u: User) => {
    if (u.username === 'admin') {
      setUserFeedback({
        success: false,
        message: 'Không thể khóa tài khoản Quản trị viên cố định',
      });
      return;
    }
    setUserFeedback(null);
    try {
      await api.updateUser(u.id, { isActive: !u.isActive });
      setUserFeedback({
        success: true,
        message: u.isActive
          ? `Đã tạm khóa tài khoản @${u.username}.`
          : `Đã mở khóa tài khoản @${u.username}.`,
      });
      handleOpenUsersTab();
    } catch (err: any) {
      setUserFeedback({
        success: false,
        message: err.message || 'Lỗi khi cập nhật trạng thái tài khoản',
      });
    }
  };

  const handleDeleteUser = (u: User) => {
    if (u.username === 'admin') {
      setUserFeedback({
        success: false,
        message: 'Không thể xóa tài khoản Quản trị viên cố định',
      });
      return;
    }
    setConfirmDeleteRequest({
      title: `Xóa tài khoản @${u.username}?`,
      description: `Tài khoản "${u.fullName}" sẽ bị xóa khỏi hệ thống. Lịch sử văn bản đã tạo vẫn được giữ nguyên.`,
      confirmLabel: 'Xác nhận xóa tài khoản',
      onConfirm: async () => {
        setConfirmDeleteBusy(true);
        try {
          await api.deleteUser(u.id);
          setConfirmDeleteRequest(null);
          setUserFeedback({
            success: true,
            message: `Đã xóa tài khoản @${u.username}.`,
          });
          await handleOpenUsersTab();
        } catch (err: any) {
          setConfirmDeleteRequest(null);
          setUserFeedback({
            success: false,
            message: err.message || 'Lỗi khi xóa tài khoản',
          });
        } finally {
          setConfirmDeleteBusy(false);
        }
      },
    });
  };

  const handleResetUserPass = (u: User) => {
    setResetPassUser(u);
  };

  const handleConfirmResetPassword = async (newPassword: string) => {
    if (!resetPassUser) return;
    setResetPassSaving(true);
    setUserFeedback(null);
    try {
      await api.updateUser(resetPassUser.id, { password: newPassword });
      const uname = resetPassUser.username;
      setResetPassUser(null);
      setUserFeedback({
        success: true,
        message: `Đã đặt lại mật khẩu cho tài khoản @${uname} thành công.`,
      });
    } catch (err: any) {
      setUserFeedback({
        success: false,
        message: err.message || 'Lỗi khi đổi mật khẩu',
      });
    } finally {
      setResetPassSaving(false);
    }
  };

  const renderPermissionMatrix = (
    permMap: Record<string, BookPermission[]>,
    onChangeMap: (next: Record<string, BookPermission[]>) => void
  ) => {
    const targetBooks = books.filter((b) => b.isActive);
    return (
      <div className="space-y-3">
        {targetBooks.map((book) => {
          const currentPerms = normalizeBookPermissionList(permMap[book.id] || []);
          const hasIssuePerm = currentPerms.includes('issue_number');
          return (
            <div key={book.id} className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                  <span className="font-bold text-slate-900 text-xs">{book.name}</span>
                  <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-600">
                    {book.unitCode}
                  </span>
                  <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-800">
                    {book.code}
                  </span>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  {PERMISSION_PRESETS.map((preset) => (
                    <button
                      key={preset.label}
                      type="button"
                      onClick={() =>
                        onChangeMap({
                          ...permMap,
                          [book.id]: normalizeBookPermissionList([...preset.perms]),
                        })
                      }
                      className="px-2.5 py-1 min-h-[32px] text-[10px] font-semibold bg-white hover:bg-sky-50 text-slate-700 hover:text-sky-800 border border-slate-200 rounded transition-colors"
                    >
                      {preset.label}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                {ALL_BOOK_PERMISSIONS.map((perm) => {
                  const meta = BOOK_PERMISSION_META[perm];
                  const checked = currentPerms.includes(perm);
                  const lockedViewByIssue = perm === 'view' && hasIssuePerm;
                  return (
                    <label
                      key={perm}
                      title={
                        lockedViewByIssue
                          ? 'Quyền Lấy số tự động bao gồm quyền Chỉ xem trên sổ này'
                          : meta.description
                      }
                      className={`flex items-start gap-2.5 p-2.5 min-h-[42px] rounded-md border text-[11px] transition-colors ${
                        lockedViewByIssue ? 'cursor-default' : 'cursor-pointer'
                      } ${
                        checked
                          ? 'bg-sky-50/80 border-sky-300 text-sky-950 font-semibold'
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={lockedViewByIssue}
                        onChange={() => onChangeMap(togglePermissionInMap(permMap, book.id, perm))}
                        className="mt-0.5 rounded text-sky-700 focus:ring-sky-700 disabled:opacity-75"
                      />
                      <span className="leading-tight">
                        {meta.shortLabel}
                        {lockedViewByIssue && (
                          <span className="block text-[9px] font-normal text-sky-700">
                            Kèm theo quyền Lấy số
                          </span>
                        )}
                        {(perm === 'request_number' || perm === 'issue_number') && (
                          <span className="block text-[9px] font-normal text-slate-400">
                            {perm === 'request_number' ? 'Chỉ gửi xin số' : 'Tự động kèm Chỉ xem'}
                          </span>
                        )}
                      </span>
                    </label>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  const visibleNavGroups = useMemo(() => {
    return SETTINGS_NAV_GROUPS.filter((group) => !group.adminOnly || isAdmin).map((group) => ({
      ...group,
      items: group.items.filter((item) => !item.adminOnly || isAdmin),
    }));
  }, [isAdmin]);

  const activeNavMeta = useMemo(() => {
    for (const group of visibleNavGroups) {
      const item = group.items.find((i) => i.id === activeSubTab);
      if (item) {
        return { groupTitle: group.title, itemLabel: item.label };
      }
    }
    return { groupTitle: 'Quản lý sổ', itemLabel: 'Danh sách sổ' };
  }, [visibleNavGroups, activeSubTab]);

  const handleSelectNavItem = (tabId: SettingsSubTab) => {
    setMobileNavOpen(false);
    if (tabId === 'users') {
      handleOpenUsersTab();
    } else {
      setActiveSubTab(tabId);
    }
  };

  const getNavItemBadge = (tabId: SettingsSubTab): number | null => {
    if (tabId === 'books') return books.length;
    if (tabId === 'signers') return displayedBookSigners.length;
    if (tabId === 'docTypes') return documentTypes.length;
    return null;
  };

  const selectedBookIdentity = getBookIdentityStyle(selectedBook);

  return (
    <div className="w-full max-w-7xl mx-auto py-4 sm:py-6 px-3 sm:px-6 lg:px-8 space-y-4 sm:space-y-5 overflow-x-hidden">
      {/* Page Header */}
      <div className="min-w-0">
        <h1 className="text-lg sm:text-2xl font-bold tracking-tight text-slate-900">
          Cài Đặt Hệ Thống &amp; Quản Trị Sổ
        </h1>
        <p className="text-xs sm:text-sm text-slate-600 mt-1">
          Quản lý sổ văn bản, quy tắc cấp số, danh mục, lưu trữ đồng bộ và phân quyền theo từng sổ.
        </p>
      </div>

      {/* Mobile: "Chọn mục cài đặt" button (replaces horizontal scrolling tabs) */}
      <div className="lg:hidden">
        <button
          type="button"
          onClick={() => setMobileNavOpen(true)}
          aria-expanded={mobileNavOpen}
          className="w-full min-h-[46px] px-3.5 py-2.5 bg-white hover:bg-slate-50 border border-slate-300 rounded-xl shadow-2xs flex items-center justify-between gap-3 text-left transition-colors"
        >
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-8 h-8 rounded-lg bg-sky-50 text-sky-700 border border-sky-200 flex items-center justify-center shrink-0">
              <Menu className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <span className="text-[11px] font-semibold text-slate-500 block leading-tight">
                Chọn mục cài đặt · {activeNavMeta.groupTitle}
              </span>
              <span className="text-xs font-bold text-sky-900 block truncate mt-0.5">
                {activeNavMeta.itemLabel}
              </span>
            </div>
          </div>
          <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
        </button>
      </div>

      {/* Mobile Grouped Settings Picker Drawer/Modal */}
      {mobileNavOpen && (
        <div
          className="fixed inset-0 z-50 lg:hidden flex flex-col justify-end sm:justify-center sm:items-center bg-slate-900/50 backdrop-blur-[1px] sm:p-4"
          role="dialog"
          aria-modal="true"
          aria-label="Chọn mục cài đặt"
        >
          <button
            type="button"
            aria-label="Đóng danh sách mục cài đặt"
            onClick={() => setMobileNavOpen(false)}
            className="flex-1 w-full cursor-default sm:hidden"
          />
          <div className="w-full sm:max-w-md bg-white rounded-t-2xl sm:rounded-2xl border-t sm:border border-slate-200 shadow-2xl max-h-[85vh] flex flex-col overflow-hidden">
            <div className="px-4 py-3.5 border-b border-slate-200 flex items-center justify-between gap-2 shrink-0">
              <span className="text-sm font-bold text-slate-900">Chọn mục cài đặt</span>
              <button
                type="button"
                onClick={() => setMobileNavOpen(false)}
                className="p-2 min-w-[36px] min-h-[36px] inline-flex items-center justify-center rounded-lg text-slate-500 hover:bg-slate-100"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="overflow-y-auto p-3 space-y-4">
              {visibleNavGroups.map((group) => (
                <div key={group.key} className="space-y-1">
                  <div className="px-2.5 py-1 text-[11px] font-bold uppercase tracking-wider text-slate-400">
                    {group.title}
                  </div>
                  <div className="space-y-1">
                    {group.items.map((item) => {
                      const active = activeSubTab === item.id;
                      const badge = getNavItemBadge(item.id);
                      return (
                        <button
                          key={item.id}
                          type="button"
                          onClick={() => handleSelectNavItem(item.id)}
                          className={`w-full min-h-[42px] px-3 py-2 rounded-lg text-left text-xs font-semibold flex items-center justify-between gap-2 border transition-colors ${
                            active
                              ? 'bg-sky-50 text-sky-900 border-sky-200'
                              : 'bg-white text-slate-700 border-transparent hover:bg-slate-50'
                          }`}
                        >
                          <span className="break-words">{item.label}</span>
                          <div className="flex items-center gap-1.5 shrink-0">
                            {badge !== null && (
                              <span
                                className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                  active ? 'bg-sky-100 text-sky-800' : 'bg-slate-100 text-slate-600'
                                }`}
                              >
                                {badge}
                              </span>
                            )}
                            <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
                          </div>
                        </button>
                      );
                    })}
                    {group.key === 'books_management' && isAdmin && (
                      <button
                        type="button"
                        disabled={isConfigBlocked}
                        onClick={() => {
                          setMobileNavOpen(false);
                          setCreateBookServerError(null);
                          setCreateBookDrawerOpen(true);
                        }}
                        className="w-full min-h-[42px] px-3 py-2 rounded-lg text-left text-xs font-semibold flex items-center gap-2 text-sky-700 hover:bg-sky-50 border border-dashed border-sky-200 disabled:opacity-50"
                      >
                        <Plus className="w-3.5 h-3.5 shrink-0" />
                        <span>Tạo sổ mới</span>
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Main 2-column layout: Left Vertical Grouped Sidebar (lg+) + Right Content Area */}
      <div className="lg:flex lg:items-start lg:gap-6 min-w-0">
        {/* Desktop Left Navigation */}
        <aside
          aria-label="Nhóm mục cài đặt"
          className="hidden lg:block w-64 shrink-0 sticky top-20"
        >
          <nav className="bg-white border border-slate-200 rounded-xl p-3 shadow-xs space-y-4">
            {visibleNavGroups.map((group) => (
              <div key={group.key} className="space-y-1">
                <h2 className="px-2.5 py-1 text-[11px] font-bold uppercase tracking-wider text-slate-400">
                  {group.title}
                </h2>
                <div className="space-y-1">
                  {group.items.map((item) => {
                    const active = activeSubTab === item.id;
                    const badge = getNavItemBadge(item.id);
                    return (
                      <button
                        key={item.id}
                        type="button"
                        onClick={() => handleSelectNavItem(item.id)}
                        className={`w-full min-h-[38px] px-3 py-2 rounded-lg text-left text-xs font-semibold flex items-center justify-between gap-2 border transition-colors ${
                          active
                            ? 'bg-sky-50 text-sky-900 border-sky-200'
                            : 'bg-transparent text-slate-600 border-transparent hover:text-slate-900 hover:bg-slate-50'
                        }`}
                      >
                        <span className="leading-snug">{item.label}</span>
                        {badge !== null && (
                          <span
                            className={`px-1.5 py-0.5 rounded text-[10px] font-bold shrink-0 ${
                              active ? 'bg-sky-100 text-sky-800' : 'bg-slate-100 text-slate-600'
                            }`}
                          >
                            {badge}
                          </span>
                        )}
                      </button>
                    );
                  })}

                  {group.key === 'books_management' && isAdmin && (
                    <button
                      type="button"
                      disabled={isConfigBlocked}
                      onClick={() => {
                        setCreateBookServerError(null);
                        setCreateBookDrawerOpen(true);
                      }}
                      className="w-full min-h-[38px] px-3 py-2 rounded-lg text-left text-xs font-semibold flex items-center gap-2 text-sky-700 hover:bg-sky-50 border border-dashed border-sky-200 transition-colors disabled:opacity-50"
                    >
                      <Plus className="w-3.5 h-3.5 shrink-0" />
                      <span>Tạo sổ mới</span>
                    </button>
                  )}
                </div>
              </div>
            ))}
          </nav>
        </aside>

        {/* Right Content Area */}
        <div className="flex-1 min-w-0 space-y-5">
          {/* Pinned Top Bar: Sổ đang cấu hình */}
          {managedBooks.length > 0 && (
            <div className="bg-white border border-slate-200 rounded-xl p-3 sm:px-4 sm:py-3 shadow-xs flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 min-w-0">
              <div className="flex items-center gap-2.5 min-w-0 flex-1">
                <div
                  className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border ${selectedBookIdentity.iconBg} ${selectedBookIdentity.iconText} ${selectedBookIdentity.badgeBorder}`}
                >
                  <BookOpen className="w-4 h-4" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <label
                      htmlFor="settings-active-book-select"
                      className="text-[11px] font-semibold text-slate-500 whitespace-nowrap"
                    >
                      Sổ đang cấu hình
                    </label>
                    {hasUnsavedChangesForSelectedBook && (
                      <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 text-[10px] font-bold">
                        Có thay đổi chưa lưu
                      </span>
                    )}
                  </div>
                  <select
                    id="settings-active-book-select"
                    value={selectedBook.id}
                    onChange={(e) => requestSelectBook(e.target.value)}
                    className="w-full min-w-0 max-w-full mt-0.5 font-bold text-slate-900 bg-transparent border-0 p-0 focus:outline-hidden cursor-pointer text-base sm:text-sm truncate"
                  >
                    {managedBooks.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name} ({b.unitCode}){!b.isActive ? ' — Tạm khóa' : ''}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2 shrink-0 sm:pl-2">
                <span
                  className={`font-mono font-bold px-2.5 py-1 rounded-md border text-[11px] ${selectedBookIdentity.badgeBg} ${selectedBookIdentity.badgeText} ${selectedBookIdentity.badgeBorder}`}
                >
                  {selectedBook.unitCode}
                </span>
                <span className="font-mono px-2 py-1 rounded-md bg-slate-100 text-slate-600 border border-slate-200 text-[11px]">
                  {selectedBook.code}
                </span>
              </div>
            </div>
          )}

          {isConfigBlocked && (
            <div className="p-4 bg-rose-50 border border-rose-300 rounded-xl flex flex-wrap items-center justify-between gap-3 text-xs text-rose-950">
              <div className="flex items-start gap-2.5 min-w-0">
                <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                <div className="min-w-0">
                  <strong className="block font-bold">
                    Chưa cập nhật được dữ liệu cấu hình từ máy chủ
                  </strong>
                  <span className="break-words">
                    {syncError ||
                      'Đang tải cấu hình mới nhất. Các nút lưu cài đặt tạm khóa để tránh ghi đè nhầm cấu hình cũ.'}
                  </span>
                </div>
              </div>
              <button
                type="button"
                onClick={onRefreshData}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-lg transition-colors shrink-0"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>Thử lại</span>
              </button>
            </div>
          )}

          {settingsConflict && (
            <div className="p-4 bg-amber-50 border border-amber-300 rounded-xl space-y-2.5 text-xs text-amber-950">
              <div className="flex items-start gap-2.5">
                <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                <div className="space-y-1 min-w-0">
                  <strong className="block font-bold">
                    Cài đặt đã thay đổi trên thiết bị khác
                  </strong>
                  <span className="break-words">{settingsConflict.message}</span>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2 sm:pl-7">
                <button
                  type="button"
                  onClick={() => {
                    const targetBook = settingsConflict.serverBook || selectedBook;
                    applySelectedBookToForms(targetBook, true);
                    onRefreshData();
                  }}
                  className="px-3 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg transition-colors"
                >
                  Tải lại cài đặt mới từ máy chủ
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    if (settingsConflict.retryAction) {
                      await settingsConflict.retryAction();
                    } else {
                      if (settingsConflict.serverVersion) {
                        loadedBookSnapshotRef.current = {
                          bookId: selectedBook.id,
                          version: settingsConflict.serverVersion,
                        };
                      }
                      setSettingsConflict(null);
                    }
                  }}
                  className="px-3 py-1.5 bg-amber-100 hover:bg-amber-200 text-amber-950 border border-amber-300 font-semibold rounded-lg transition-colors"
                >
                  Giữ nội dung đang sửa &amp; ghi đè lên máy chủ
                </button>
              </div>
            </div>
          )}

          {!isAdmin && (
            <div className="p-3 bg-sky-50 border border-sky-200 rounded-xl text-xs text-sky-900 flex items-center gap-2">
              <Sliders className="w-4 h-4 text-sky-700 shrink-0" />
              <span className="break-words">
                Bạn đang cấu hình với quyền <strong>Quản trị sổ</strong> cho sổ{' '}
                <strong>{selectedBook.name}</strong>. Các thiết lập toàn hệ thống và phân quyền tài khoản do Quản trị viên quản lý.
              </span>
            </div>
          )}

          {/* SUB-TAB: DANH SÁCH SỔ */}
          {activeSubTab === 'books' && (
            <BooksListSection
              books={books}
              selectedBookId={selectedBook.id}
              currentUser={currentUser}
              isAdmin={isAdmin}
              isConfigBlocked={isConfigBlocked}
              documentTypes={documentTypes}
              bookFeedback={bookMessage}
              onDismissFeedback={() => setBookMessage(null)}
              onConfigureBook={(bookId) => requestSelectBook(bookId, 'general')}
              onOpenCreateDrawer={() => {
                setCreateBookServerError(null);
                setCreateBookDrawerOpen(true);
              }}
              onRequestToggleLockBook={(book) => setToggleLockConfirmBook(book)}
            />
          )}

      {/* SUB-TAB 1: THÔNG TIN SỔ & MẪU SỐ */}
      {activeSubTab === 'general' && (
        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs max-w-2xl">
          <div className="mb-4 pb-3 border-b border-slate-100 flex items-center justify-between">
            <div className="min-w-0">
              <h2 className="text-base font-bold text-slate-900">
                Thông tin đơn vị & Mẫu số: {selectedBook.name}
              </h2>
              <p className="text-xs text-slate-500">
                {selectedBook.isShared ? 'Sổ dùng chung' : 'Sổ riêng đơn vị'} · Mã sổ cố định:{' '}
                <code className="font-mono font-bold text-sky-800">{selectedBook.code}</code>
              </p>
            </div>
          </div>

          <form onSubmit={handleSaveGeneral} className="space-y-4 sm:space-y-5 text-xs">
            <div className="space-y-1">
              <label className="font-bold text-slate-700">Mã sổ nội bộ cố định (Tab Google Sheets)</label>
              <input
                type="text"
                value={selectedBook.code}
                disabled
                className="w-full min-h-[42px] h-10 sm:h-9 px-3 bg-slate-100 border border-slate-200 rounded-lg font-mono text-base sm:text-xs text-slate-600 cursor-not-allowed"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Tên hiển thị sổ văn bản</label>
              <input
                type="text"
                value={bookTitleInput}
                onChange={(e) => {
                  setBookTitleInput(e.target.value);
                  setGeneralDirty(true);
                }}
                disabled={!canManageSelectedBook || generalSaving || isConfigBlocked}
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Tên đơn vị</label>
              <input
                type="text"
                value={unitName}
                onChange={(e) => {
                  setUnitName(e.target.value);
                  setGeneralDirty(true);
                }}
                disabled={!canManageSelectedBook || generalSaving || isConfigBlocked}
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Ký hiệu đơn vị chính thức</label>
              <input
                type="text"
                value={unitCode}
                onChange={(e) => {
                  setUnitCode(e.target.value);
                  setGeneralDirty(true);
                }}
                disabled={!canManageSelectedBook || generalSaving || isConfigBlocked}
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 font-mono focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
              <span className="text-[11px] text-slate-500 block">
                Ví dụ: ĐTN (Đoàn trường), CB (Chi bộ), THPTNCT (Nhà trường) — quản trị viên có thể đổi theo đúng ký hiệu đơn vị đang dùng.
              </span>
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Mẫu định dạng số văn bản</label>
              <input
                type="text"
                value={numberFormat}
                onChange={(e) => {
                  setNumberFormat(e.target.value);
                  setGeneralDirty(true);
                }}
                disabled={!canManageSelectedBook || generalSaving || isConfigBlocked}
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 font-mono focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
              <span className="text-[11px] text-slate-500 block">
                Các biến thay thế: <code className="bg-slate-100 px-1 py-0.5 rounded">{'{so}'}</code> (số thứ tự),{' '}
                <code className="bg-slate-100 px-1 py-0.5 rounded">{'{loai}'}</code> (loại văn bản),{' '}
                <code className="bg-slate-100 px-1 py-0.5 rounded">{'{don_vi}'}</code> (ký hiệu đơn vị). Ví dụ:{' '}
                <code className="font-bold text-sky-800">01-KH/{unitCode || 'ĐTN'}</code>
              </span>
            </div>

            <div className="space-y-2">
              <label className="font-bold text-slate-700 block">
                Loại văn bản được sử dụng trong sổ "{selectedBook.name}"
              </label>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
                {documentTypes.filter((t) => t.isActive).map((t) => {
                  const checked = bookEnabledTypes.includes(t.code);
                  return (
                    <label
                      key={t.code}
                      className={`flex items-center gap-2.5 p-2.5 min-h-[42px] rounded-lg border cursor-pointer ${
                        checked ? 'bg-sky-50 border-sky-300 text-sky-950 font-semibold' : 'bg-white border-slate-200 text-slate-600'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={!canManageSelectedBook || generalSaving || isConfigBlocked}
                        onChange={() => {
                          setBookEnabledTypes((prev) =>
                            prev.includes(t.code) ? prev.filter((c) => c !== t.code) : [...prev, t.code]
                          );
                          setGeneralDirty(true);
                        }}
                        className="rounded text-sky-700"
                      />
                      <span>{t.name} ({t.code})</span>
                    </label>
                  );
                })}
              </div>
            </div>

            {generalError && (
              <div className="p-3 bg-rose-50 text-rose-700 border border-rose-200 rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                <span>{generalError}</span>
              </div>
            )}

            {generalSuccess && (
              <div className="p-3 bg-emerald-50 text-emerald-800 rounded-lg flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                <span>Đã lưu thành công cấu hình sổ "{bookTitleInput}"!</span>
              </div>
            )}

            {canManageSelectedBook && (
              <div className="pt-2">
                <button
                  type="submit"
                  disabled={generalSaving || isConfigBlocked}
                  className="w-full sm:w-auto min-h-[44px] px-5 py-2.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center justify-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  <span>{generalSaving ? 'Đang lưu...' : 'Lưu cài đặt sổ'}</span>
                </button>
              </div>
            )}
          </form>
        </div>
      )}

      {/* SUB-TAB 2: QUY TẮC DÃY SỐ & SỐ BẮT ĐẦU */}
      {activeSubTab === 'numbering' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4">
            <h2 className="text-base font-bold text-slate-900">
              Chế Độ Dãy Số Trong Năm — {selectedBook.name}
            </h2>
            <p className="text-xs text-slate-600">
              Quy tắc quản lý cách tăng số văn bản đi cho sổ <strong>{selectedBook.name}</strong>. Sang năm mới (múi giờ Asia/Ho_Chi_Minh), dãy số tự động bắt đầu lại từ 01.
            </p>

            <form onSubmit={handleChangeMode} className="space-y-4 text-xs">
              <div className="space-y-3">
                <label className="flex items-start gap-3 p-3.5 min-h-[48px] border rounded-lg cursor-pointer transition-colors hover:bg-slate-50 border-slate-300">
                  <input
                    type="radio"
                    name="numberMode"
                    value="separate"
                    checked={targetMode === 'separate'}
                    onChange={() => setTargetMode('separate')}
                    disabled={!canManageSelectedBook}
                    className="mt-0.5 text-sky-600 focus:ring-sky-600"
                  />
                  <div>
                    <strong className="block font-semibold text-slate-900">
                      Mỗi loại văn bản có một dãy số riêng trong năm (Khuyên dùng)
                    </strong>
                    <span className="text-slate-500 text-[11px] block mt-0.5">
                      Kế hoạch tăng riêng (01-KH, 02-KH...), Báo cáo tăng riêng (01-BC, 02-BC...).
                    </span>
                  </div>
                </label>

                <label className="flex items-start gap-3 p-3.5 min-h-[48px] border rounded-lg cursor-pointer transition-colors hover:bg-slate-50 border-slate-300">
                  <input
                    type="radio"
                    name="numberMode"
                    value="shared"
                    checked={targetMode === 'shared'}
                    onChange={() => setTargetMode('shared')}
                    disabled={!canManageSelectedBook}
                    className="mt-0.5 text-sky-600 focus:ring-sky-600"
                  />
                  <div>
                    <strong className="block font-semibold text-slate-900">
                      Dùng một dãy số chung cho tất cả các loại văn bản
                    </strong>
                    <span className="text-slate-500 text-[11px] block mt-0.5">
                      Tất cả văn bản dùng chung 1 thứ tự liên tục từ 01, 02, 03...
                    </span>
                  </div>
                </label>
              </div>

              {targetMode !== selectedBook.numberMode && (
                <div className="p-4 bg-rose-50 border border-rose-300 rounded-lg space-y-2">
                  <div className="flex items-start gap-2 text-rose-800">
                    <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                    <div>
                      <strong className="block font-semibold">Cảnh báo chuyển đổi chế độ cấp số:</strong>
                      <span>
                        Nếu trong năm sổ này đã phát sinh văn bản, để xác nhận vui lòng nhập{' '}
                        <strong className="text-rose-950">XÁC NHẬN CHUYỂN ĐỔI</strong>:
                      </span>
                    </div>
                  </div>

                  <input
                    type="text"
                    value={modeConfirmText}
                    onChange={(e) => setModeConfirmText(e.target.value)}
                    placeholder="Nhập: XÁC NHẬN CHUYỂN ĐỔI"
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-rose-300 rounded-lg text-base sm:text-xs font-bold text-rose-900"
                  />
                </div>
              )}

              {modeError && (
                <div className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs">{modeError}</div>
              )}

              {modeSuccess && (
                <div className="p-3 bg-emerald-50 text-emerald-800 rounded-lg text-xs flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>Đã cập nhật chế độ cấp số thành công!</span>
                </div>
              )}

              {canManageSelectedBook && targetMode !== selectedBook.numberMode && (
                <button
                  type="submit"
                  disabled={modeSaving || isConfigBlocked}
                  className="w-full sm:w-auto min-h-[44px] px-4 py-2.5 bg-rose-700 hover:bg-rose-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {modeSaving ? 'Đang cập nhật...' : 'Xác nhận áp dụng chế độ mới'}
                </button>
              )}
            </form>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-2">
              <div className="min-w-0">
                <h2 className="text-base font-bold text-slate-900">
                  Thiết Lập Số Bắt Đầu — {selectedBook.name}
                </h2>
                <p className="text-xs text-slate-600 mt-1 break-words">
                  Mỗi loại văn bản có mốc bắt đầu và dãy số riêng theo từng sổ, từng năm. Khóa dãy đang chọn:{' '}
                  <code className="font-mono font-bold text-sky-800 bg-sky-50 px-1.5 py-0.5 rounded border border-sky-200 break-all">
                    {activeSeriesKeyDisplay}
                  </code>
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                {seqConfigLoading ? (
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-[11px] font-semibold bg-sky-50 text-sky-800 border border-sky-200">
                    <RotateCw className="w-3 h-3 animate-spin" />
                    <span>Đang tải cấu hình...</span>
                  </span>
                ) : seqConfigLoadError ? (
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-semibold bg-rose-50 text-rose-800 border border-rose-200">
                    <AlertTriangle className="w-3 h-3" />
                    <span>Lỗi tải cấu hình</span>
                  </span>
                ) : isSeriesDirty ? (
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-bold bg-amber-100 text-amber-900 border border-amber-300">
                    <span>Chưa lưu</span>
                  </span>
                ) : activeSavedConfig?.isConfigured ? (
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200">
                    <CheckCircle2 className="w-3 h-3" />
                    <span>Đã thiết lập mốc riêng</span>
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200">
                    <span>Chưa thiết lập · Mặc định 01</span>
                  </span>
                )}
              </div>
            </div>

            {seqConfigLoadError && (
              <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-900 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-start gap-2 min-w-0">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  <span className="break-words">
                    Không tải được cấu hình cho dãy <strong>{activeSeriesKeyDisplay}</strong>: {seqConfigLoadError}. Nút lưu tạm khóa để tránh ghi đè sai.
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => loadStartSequenceSeries(selectedBook.id, seqYear, effectiveSeqTypeCode)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 min-h-[36px] bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-md shrink-0"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Thử tải lại</span>
                </button>
              </div>
            )}

            <form onSubmit={handleSetStartSequence} className="space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Năm áp dụng</label>
                  <select
                    aria-label="Năm áp dụng số bắt đầu"
                    value={seqYear}
                    onChange={(e) => setSeqYear(parseInt(e.target.value, 10))}
                    disabled={!canManageSelectedBook || isConfigBlocked}
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  >
                    {Array.from(
                      new Set([
                        currentYear + 1,
                        currentYear,
                        currentYear - 1,
                        ...Object.keys(selectedBook.customStartSequences || {})
                          .map((k) => parseInt(k.split('_')[0], 10))
                          .filter((y) => !Number.isNaN(y)),
                      ])
                    )
                      .sort((a, b) => b - a)
                      .map((y) => (
                        <option key={y} value={y}>
                          Năm {y}
                        </option>
                      ))}
                  </select>
                </div>

                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Loại văn bản</label>
                  <select
                    aria-label="Loại văn bản thiết lập số bắt đầu"
                    value={effectiveSeqTypeCode}
                    onChange={(e) => setSeqTypeCode(e.target.value)}
                    disabled={!canManageSelectedBook || isConfigBlocked}
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  >
                    {selectedBook.numberMode === 'shared' ? (
                      <option value="SHARED">Dãy số chung (SHARED)</option>
                    ) : (
                      enabledTypesForSelectedBook.map((t) => (
                        <option key={t.code} value={t.code}>
                          {t.name} ({t.code})
                        </option>
                      ))
                    )}
                  </select>
                </div>

                <div className="space-y-1 min-w-0">
                  <div className="flex items-center justify-between gap-2">
                    <label htmlFor="start-sequence-input" className="font-bold text-slate-700">
                      Số bắt đầu đã cài
                    </label>
                    <div className="flex items-center gap-1.5">
                      {isSeriesDirty && (
                        <>
                          <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300 text-[10px] font-bold">
                            Chưa lưu
                          </span>
                          <button
                            type="button"
                            onClick={() => {
                              setDraftStartInputBySeries((prev) => {
                                const next = { ...prev };
                                delete next[activeSeriesId];
                                return next;
                              });
                              setSeqSaveErrorBySeries((prev) => ({
                                ...prev,
                                [activeSeriesId]: null,
                              }));
                            }}
                            className="text-[11px] font-semibold text-slate-500 hover:text-slate-800 underline"
                          >
                            Hoàn tác
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                  <input
                    id="start-sequence-input"
                    type="number"
                    min={1}
                    step={1}
                    value={currentStartInputValue}
                    placeholder="Chưa thiết lập · Mặc định 01"
                    onChange={(e) => {
                      const raw = e.target.value;
                      setSeqSaveErrorBySeries((prev) => ({ ...prev, [activeSeriesId]: null }));
                      setDraftStartInputBySeries((prev) => {
                        if (raw.trim() === savedInputString) {
                          const next = { ...prev };
                          delete next[activeSeriesId];
                          return next;
                        }
                        return { ...prev, [activeSeriesId]: raw };
                      });
                    }}
                    disabled={
                      !canManageSelectedBook ||
                      isConfigBlocked ||
                      seqConfigLoading ||
                      Boolean(seqConfigLoadError)
                    }
                    className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg font-mono font-bold text-base sm:text-xs ${
                      isSeriesDirty ? 'border-amber-400 ring-1 ring-amber-300' : 'border-slate-300'
                    }`}
                  />
                </div>
              </div>

              {/* Explicit side-by-side summary: Số bắt đầu đã cài vs Số tiếp theo dự kiến */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="p-3 bg-sky-50 border border-sky-200 rounded-lg space-y-1 min-w-0">
                  <span className="text-[11px] font-semibold text-sky-800 block">
                    Số bắt đầu đã cài ({activeSeriesKeyDisplay})
                  </span>
                  <div className="font-mono text-base font-bold text-sky-950 break-words">
                    {seqConfigLoading
                      ? 'Đang tải...'
                      : seqConfigLoadError
                      ? 'Lỗi tải cấu hình'
                      : activeSavedConfig?.isConfigured && activeSavedConfig.customStartSequence != null
                      ? `${String(activeSavedConfig.customStartSequence).padStart(2, '0')} (${activeSavedConfig.customStartSequence})`
                      : 'Chưa thiết lập · Mặc định 01'}
                  </div>
                  <span className="text-[11px] text-sky-700 block">
                    {isSeriesDirty
                      ? `Bản nháp đang nhập: "${currentStartInputValue.trim() || 'trống'}" (Chưa lưu — bấm Lưu số bắt đầu để áp dụng).`
                      : 'Mốc đã lưu giữ cố định theo sổ, năm và loại văn bản; không tự thay đổi khi cấp số mới.'}
                  </span>
                </div>

                <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg space-y-1 min-w-0">
                  <span className="text-[11px] font-semibold text-emerald-800 block">
                    Số tiếp theo dự kiến ({selectedBook.name} · Năm {seqYear} · {effectiveSeqTypeCode})
                  </span>
                  <div className="font-mono text-base font-bold text-emerald-950 break-words">
                    {seqConfigLoading
                      ? 'Đang tính...'
                      : seqConfigLoadError
                      ? 'Chưa có dữ liệu dự kiến'
                      : activeSavedConfig
                      ? `${String(activeSavedConfig.nextExpectedSeq).padStart(2, '0')} → ${activeSavedConfig.previewNumber}`
                      : 'Chưa có dữ liệu dự kiến'}
                  </div>
                  <span className="text-[11px] text-emerald-800 block">
                    {activeSavedConfig
                      ? `= max(Mốc đã cài, Số lớn nhất đã cấp/hủy: ${activeSavedConfig.maxIssuedSequence} + 1). Mức đặt tối thiểu: ${activeSavedConfig.minAllowedStartSequence}.`
                      : 'Tính từ max(Mốc đã cài, Số lớn nhất đã cấp/hủy trong cùng dãy + 1).'}
                  </span>
                </div>
              </div>

              {Object.keys(selectedBook.customStartSequences || {}).length > 0 && (
                <div className="p-3 bg-sky-50/70 border border-sky-200 rounded-lg text-[11px] text-sky-900">
                  <strong className="block mb-1">
                    Các mốc Số bắt đầu đã cài trên sổ "{selectedBook.name}":
                  </strong>
                  <div className="flex flex-wrap gap-2">
                    {Object.entries(selectedBook.customStartSequences)
                      .sort(([a], [b]) => a.localeCompare(b))
                      .map(([k, val]) => {
                        const [yr, ...rest] = k.split('_');
                        const tCode = rest.join('_');
                        const isActiveRow =
                          String(seqYear) === yr && tCode === effectiveSeqTypeCode;
                        return (
                          <span
                            key={k}
                            className={`px-2 py-0.5 rounded font-mono border ${
                              isActiveRow
                                ? 'bg-sky-700 text-white border-sky-700'
                                : 'bg-white border-sky-200 text-sky-900'
                            }`}
                          >
                            {selectedBook.code}:{yr}:{tCode} = <strong>{val}</strong>
                          </span>
                        );
                      })}
                  </div>
                </div>
              )}

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-[11px] text-slate-600">
                <strong>Quy tắc bảo vệ:</strong> Số bắt đầu của từng dãy không được nhỏ hơn hoặc bằng số lớn nhất từng cấp (bao gồm số đã hủy) trong chính sổ, năm và loại văn bản đó.
              </div>

              {activeSeriesSaveError && (
                <div className="p-3 bg-rose-50 border border-rose-200 text-rose-700 rounded-lg text-xs">
                  {activeSeriesSaveError}
                </div>
              )}

              {seqSaveSuccessNotice && (
                <div className="p-3 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-lg text-xs flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>{seqSaveSuccessNotice.message}</span>
                </div>
              )}

              {canManageSelectedBook && (
                <button
                  type="submit"
                  disabled={
                    isActiveSeriesSaving ||
                    isConfigBlocked ||
                    seqConfigLoading ||
                    Boolean(seqConfigLoadError) ||
                    !isStartInputValid
                  }
                  className="w-full sm:w-auto min-h-[44px] px-4 py-2.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {isActiveSeriesSaving ? 'Đang lưu...' : 'Lưu số bắt đầu'}
                </button>
              )}
            </form>
          </div>
        </div>
      )}

      {/* SUB-TAB 3: LOẠI VĂN BẢN (GLOBAL) */}
      {activeSubTab === 'docTypes' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <h2 className="text-base font-bold text-slate-900">
                Danh Sách Loại Văn Bản Hệ Thống
              </h2>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {documentTypes.map((t) => {
                const style = getTypeColorStyle(t.code, t.color);
                return (
                  <div key={t.id} className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <span
                        className={`font-mono font-bold px-2.5 py-1 rounded border ${style.badgeText} ${style.badgeBg} ${style.badgeBorder} shadow-2xs`}
                      >
                        <span className={`inline-block w-1.5 h-1.5 rounded-full ${style.dotColor} mr-1.5`} />
                        {t.code}
                      </span>
                      <div>
                        <h4 className="font-semibold text-slate-900">{t.name}</h4>
                        <span className="text-[11px] text-slate-500">
                          Mẫu: 01-{t.code}/{selectedBook.unitCode}
                        </span>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-2 sm:gap-3">
                      {isAdmin && (
                        <div className="flex flex-wrap items-center gap-1 bg-slate-50 p-1 rounded-md border border-slate-200">
                          <span className="text-[10px] text-slate-500 font-medium px-1">Màu:</span>
                          {Object.values(DOCUMENT_COLOR_PALETTE).map((colorOpt) => (
                            <button
                              key={colorOpt.key}
                              type="button"
                              onClick={() => handleUpdateDocTypeColor(t.id, colorOpt.key)}
                              className={`w-3.5 h-3.5 rounded-full ${colorOpt.dotColor} transition-transform ${
                                (t.color || 'blue') === colorOpt.key
                                  ? 'ring-2 ring-offset-1 ring-slate-800 scale-125'
                                  : 'opacity-60 hover:opacity-100 hover:scale-110'
                              }`}
                              title={`Đổi màu sang ${colorOpt.label}`}
                            />
                          ))}
                        </div>
                      )}

                      <span className="text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-semibold whitespace-nowrap">
                        Đang sử dụng
                      </span>
                      {isAdmin && (
                        <button
                          onClick={() => handleDeleteDocType(t.id, t.code)}
                          className="p-2 min-w-[36px] min-h-[36px] inline-flex items-center justify-center text-slate-400 hover:text-rose-600 hover:bg-slate-100 rounded transition-colors"
                          title="Xóa loại văn bản này"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {canManageSelectedBook && (
            <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Plus className="w-4 h-4 text-sky-700" />
                <span>Thêm loại văn bản mới</span>
              </h3>

              <form onSubmit={handleCreateDocType} className="grid grid-cols-1 sm:grid-cols-12 gap-3 text-xs">
                <div className="sm:col-span-6 space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Tên loại văn bản *</label>
                  <input
                    type="text"
                    value={newTypeName}
                    onChange={(e) => setNewTypeName(e.target.value)}
                    placeholder="Ví dụ: Hướng dẫn / Quyết định..."
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  />
                </div>

                <div className="sm:col-span-4 space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Ký hiệu viết tắt *</label>
                  <input
                    type="text"
                    value={newTypeCode}
                    onChange={(e) => setNewTypeCode(e.target.value)}
                    placeholder="Ví dụ: HD / QĐ"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg uppercase font-mono text-base sm:text-xs"
                  />
                </div>

                <div className="sm:col-span-2 flex items-end">
                  <button
                    type="submit"
                    disabled={typeSaving || isConfigBlocked}
                    className="w-full min-h-[44px] h-11 sm:h-10 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                  >
                    {typeSaving ? '...' : 'Thêm'}
                  </button>
                </div>

                <div className="sm:col-span-12 space-y-1.5 pt-1">
                  <label className="font-bold text-slate-700 block">Chọn màu sắc nhận diện loại văn bản</label>
                  <div className="grid grid-cols-5 sm:grid-cols-10 gap-1.5">
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

                {typeError && (
                  <div className="sm:col-span-12 p-2 bg-rose-50 text-rose-700 rounded text-xs">
                    {typeError}
                  </div>
                )}
              </form>
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB 4: NGƯỜI KÝ & CHỨC VỤ THEO SỔ */}
      {activeSubTab === 'signers' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <h2 className="text-base font-bold text-slate-900">
                Danh Mục Người Ký Gợi Ý — {selectedBook.name}
              </h2>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {displayedBookSigners.map((s) => (
                <div key={s.id} className="p-4 flex items-center justify-between">
                  <div>
                    <h4 className="font-semibold text-slate-900">{s.name}</h4>
                    <span className="text-slate-500 text-[11px]">{s.role}</span>
                  </div>

                  {canManageSelectedBook && (
                    <button
                      onClick={() => handleDeleteBookSigner(s.id)}
                      className="p-2 min-w-[38px] min-h-[38px] inline-flex items-center justify-center text-slate-400 hover:text-rose-600 hover:bg-slate-100 rounded transition-colors shrink-0"
                      title="Xóa người ký này"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          {canManageSelectedBook && (
            <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Plus className="w-4 h-4 text-sky-700 shrink-0" />
                <span>Thêm người ký mới vào sổ "{selectedBook.name}"</span>
              </h3>

              <form onSubmit={handleCreateBookSigner} className="grid grid-cols-1 sm:grid-cols-12 gap-3 text-xs">
                <div className="sm:col-span-5 space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Họ và tên *</label>
                  <input
                    type="text"
                    value={newSignerName}
                    onChange={(e) => setNewSignerName(e.target.value)}
                    placeholder="Ví dụ: Võ Văn Hoàng"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  />
                </div>

                <div className="sm:col-span-5 space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Chức vụ *</label>
                  <input
                    type="text"
                    value={newSignerRole}
                    onChange={(e) => setNewSignerRole(e.target.value)}
                    placeholder="Ví dụ: Phó Bí thư Đoàn trường"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                  />
                </div>

                <div className="sm:col-span-2 flex items-end">
                  <button
                    type="submit"
                    disabled={signerSaving || isConfigBlocked}
                    className="w-full min-h-[44px] h-11 sm:h-10 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                  >
                    {signerSaving ? '...' : 'Thêm'}
                  </button>
                </div>

                {signerError && (
                  <div className="sm:col-span-12 p-2 bg-rose-50 text-rose-700 rounded text-xs">
                    {signerError}
                  </div>
                )}
              </form>
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB: LƯU TRỮ GOOGLE DRIVE (THEO TỪNG SỔ) */}
      {activeSubTab === 'googledrive' && (() => {
        const parsedPreviewFolderId = extractGoogleDriveFolderId(gdFolderUrl);
        const currentDriveCfg = effectiveBookDriveCfg;
        const maxWordMB = currentDriveCfg?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;

        return (
          <div className="space-y-6 max-w-3xl min-w-0">
            <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-5 min-w-0">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div className="min-w-0 flex-1">
                  <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                    <HardDrive className="w-5 h-5 text-sky-700 shrink-0" />
                    <span>Lưu Trữ Tệp Google Drive — {selectedBook.name}</span>
                  </h2>
                  <p className="text-xs text-slate-600 mt-1 break-words">
                    Cấu hình thư mục Google Drive riêng cho sổ <strong>{selectedBook.name} ({selectedBook.unitCode})</strong>. Tệp Word khi cấp số hoặc PDF văn bản đến sẽ tự đổi tên theo mẫu <code className="font-mono">Số.ký hiệu - Tên văn bản.đuôi</code> và lưu vào thư mục này.
                  </p>
                </div>

                <div className="flex flex-col sm:items-end gap-1 shrink-0">
                  <span
                    className={`text-xs px-2.5 py-1 rounded font-semibold ${
                      currentDriveCfg?.isConfigured
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : gdEnabled
                        ? 'bg-amber-50 text-amber-800 border border-amber-200'
                        : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {currentDriveCfg?.isConfigured
                      ? 'Đã sẵn sàng lưu tệp'
                      : gdEnabled
                      ? 'Thiếu thông tin cấu hình'
                      : 'Chưa bật'}
                  </span>
                  <span className="text-[11px] text-slate-500">
                    Giới hạn tệp máy chủ: <strong>Tối đa {maxWordMB} MB</strong>
                  </span>
                </div>
              </div>

              <form onSubmit={handleSaveGoogleDrive} className="space-y-5 text-xs">
                <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                  <input
                    type="checkbox"
                    checked={gdEnabled}
                    onChange={(e) => {
                      setGdEnabled(e.target.checked);
                      setDriveDirty(true);
                    }}
                    disabled={!canManageSelectedBook || isConfigBlocked}
                    className="rounded text-sky-700 focus:ring-sky-700"
                  />
                  <span className="font-semibold text-slate-900">
                    Bật lưu trữ tệp Word / PDF vào thư mục Google Drive cho sổ "{selectedBook.name}"
                  </span>
                </label>

                <div className="p-3 bg-sky-50/70 border border-sky-200 rounded-lg text-[11px] text-sky-950 space-y-1">
                  <strong className="block">Cấu trúc thư mục tự động trên Google Drive của sổ này (cách ly hoàn toàn với sổ khác):</strong>
                  <div className="font-mono font-semibold text-sky-900">
                    Thư mục gốc ({selectedBook.code}) → {selectedBook.name} → {currentYear} → Văn bản đi / Văn bản đến
                  </div>
                </div>

                {canManageSelectedBook && (
                  <div className="space-y-1.5 min-w-0">
                    <label className="font-bold text-slate-700 block">
                      URL hoặc ID thư mục Google Drive (Sổ: {selectedBook.unitCode}) *
                    </label>
                    <input
                      type="text"
                      value={gdFolderUrl}
                      onChange={(e) => {
                        setGdFolderUrl(e.target.value);
                        setDriveDirty(true);
                      }}
                      disabled={isConfigBlocked}
                      placeholder="https://drive.google.com/drive/folders/1AbCDef... hoặc Folder ID"
                      className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs focus:ring-2 focus:ring-sky-600"
                    />
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-500">
                      <span className="break-all">
                        ID thư mục nhận diện:{' '}
                        {parsedPreviewFolderId ? (
                          <strong className="font-mono text-sky-800">{parsedPreviewFolderId}</strong>
                        ) : (
                          <span className="text-rose-600">Chưa hợp lệ</span>
                        )}
                      </span>
                      {(gdTestResult?.folderName || currentDriveCfg?.folderName) && (
                        <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold">
                          <FolderCheck className="w-3.5 h-3.5 shrink-0" />
                          <span>Thư mục: {gdTestResult?.folderName || currentDriveCfg?.folderName}</span>
                        </span>
                      )}
                    </div>
                  </div>
                )}

                {canManageSelectedBook && (
                  <div className="space-y-2 min-w-0">
                    <label className="font-bold text-slate-700 block">
                      Phương thức kết nối Google Drive (Xác thực phía máy chủ)
                    </label>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <label
                        className={`flex items-start gap-2.5 p-3 rounded-lg border cursor-pointer transition-colors ${
                          gdAuthMode === 'apps_script'
                            ? 'bg-sky-50 border-sky-400'
                            : 'bg-white border-slate-200 hover:bg-slate-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="gdAuthMode"
                          checked={gdAuthMode === 'apps_script'}
                          onChange={() => {
                            setGdAuthMode('apps_script');
                            setDriveDirty(true);
                          }}
                          disabled={isConfigBlocked}
                          className="mt-0.5 text-sky-700"
                        />
                        <div>
                          <strong className="block text-slate-900">Google Apps Script (Khuyên dùng)</strong>
                          <span className="text-[11px] text-slate-600">
                            Phù hợp cho <strong>Drive của tôi</strong> & tài khoản Google cá nhân/trường học. Hỗ trợ cả thay tệp cũ.
                          </span>
                        </div>
                      </label>

                      <label
                        className={`flex items-start gap-2.5 p-3 rounded-lg border cursor-pointer transition-colors ${
                          gdAuthMode === 'service_account'
                            ? 'bg-sky-50 border-sky-400'
                            : 'bg-white border-slate-200 hover:bg-slate-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="gdAuthMode"
                          checked={gdAuthMode === 'service_account'}
                          onChange={() => {
                            setGdAuthMode('service_account');
                            setDriveDirty(true);
                          }}
                          disabled={isConfigBlocked}
                          className="mt-0.5 text-sky-700"
                        />
                        <div>
                          <strong className="block text-slate-900">Google Drive API (Service Account)</strong>
                          <span className="text-[11px] text-slate-600">
                            Dùng trực tiếp cho <strong>Shared Drive</strong> Google Workspace qua khóa JSON tài khoản dịch vụ.
                          </span>
                        </div>
                      </label>
                    </div>
                  </div>
                )}

                {canManageSelectedBook && gdAuthMode === 'apps_script' && (
                  <div className="space-y-4 min-w-0">
                    <div className="bg-sky-50/70 border border-sky-200 rounded-xl p-3.5 sm:p-4 space-y-3 min-w-0">
                      <h3 className="font-bold text-sky-950 flex items-center gap-2">
                        <CheckCircle2 className="w-4 h-4 text-sky-700 shrink-0" />
                        <span>Hướng dẫn triển khai Apps Script (Hỗ trợ upload, đối soát & thay tệp)</span>
                      </h3>
                      <ol className="list-decimal pl-5 space-y-1 text-slate-700">
                        <li>Truy cập <strong>script.google.com</strong> &rarr; Tạo dự án mới &rarr; Dán đoạn mã bên dưới.</li>
                        <li>Triển khai &rarr; Ứng dụng web (Execute as: <strong>Tôi</strong>, Who has access: <strong>Bất kỳ ai</strong>).</li>
                        <li>Dán <strong>URL ứng dụng web (/exec)</strong> vào ô bên dưới (nếu dùng chung credential mặc định của hệ thống có thể để trống nếu đã cấu hình ở Sổ mặc định).</li>
                      </ol>

                      <div className="rounded-lg overflow-hidden border border-slate-800 bg-slate-900 min-w-0">
                        <div className="flex items-center justify-between gap-2 px-3 py-2 bg-slate-800/80 border-b border-slate-700">
                          <span className="text-[11px] font-mono text-slate-300">Apps Script Google Drive</span>
                          <button
                            type="button"
                            onClick={() => {
                              navigator.clipboard.writeText(GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE);
                              setCopiedDriveScript(true);
                              setTimeout(() => setCopiedDriveScript(false), 2500);
                            }}
                            className="px-2.5 py-1.5 min-h-[32px] text-xs font-semibold bg-sky-700 hover:bg-sky-600 text-white rounded shadow transition-colors inline-flex items-center gap-1 shrink-0"
                          >
                            {copiedDriveScript ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                            <span>{copiedDriveScript ? 'Đã sao chép' : 'Sao chép mã'}</span>
                          </button>
                        </div>
                        <pre className="p-3 bg-slate-900 text-slate-100 text-[11px] font-mono overflow-x-auto max-h-52 leading-relaxed w-full max-w-full">
{GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE}
                        </pre>
                      </div>
                    </div>

                    <div className="space-y-1.5 min-w-0">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <label className="font-bold text-slate-700">
                          URL Web App Google Apps Script (/exec)
                        </label>
                        {currentDriveCfg?.hasAppsScriptUrl && (
                          <div className="flex flex-wrap items-center gap-2 min-w-0">
                            <span className="text-[11px] text-emerald-700 font-mono bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200 break-all">
                              Đã lưu: {currentDriveCfg.appsScriptUrlMasked}
                            </span>
                            <button
                              type="button"
                              onClick={() => handleClearDriveCredential('apps_script')}
                              className="text-[11px] text-rose-600 hover:underline py-1"
                            >
                              Xóa URL
                            </button>
                          </div>
                        )}
                      </div>
                      <input
                        type="url"
                        value={gdAppsScriptUrl}
                        onChange={(e) => {
                          setGdAppsScriptUrl(e.target.value);
                          setDriveDirty(true);
                        }}
                        disabled={isConfigBlocked}
                        placeholder={
                          currentDriveCfg?.hasAppsScriptUrl
                            ? 'Để trống nếu giữ nguyên URL đang lưu, hoặc dán URL mới để thay thế...'
                            : 'https://script.google.com/macros/s/AKfycbx.../exec'
                        }
                        className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs focus:ring-2 focus:ring-sky-600"
                      />
                    </div>
                  </div>
                )}

                {canManageSelectedBook && gdAuthMode === 'service_account' && (
                  <div className="space-y-4">
                    <div className="space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <label className="font-bold text-slate-700">
                          Khóa JSON Tài khoản dịch vụ (Service Account Key)
                        </label>
                        {currentDriveCfg?.hasServiceAccount && (
                          <div className="flex items-center gap-2">
                            <span className="text-[11px] text-emerald-700 font-mono bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                              Đã cấu hình: {currentDriveCfg.serviceAccountEmail}
                            </span>
                            <button
                              type="button"
                              onClick={() => handleClearDriveCredential('service_account')}
                              className="text-[11px] text-rose-600 hover:underline"
                            >
                              Xóa khóa
                            </button>
                          </div>
                        )}
                      </div>

                      <div className="flex items-center gap-2">
                        <label className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg font-semibold text-slate-700 cursor-pointer">
                          <Upload className="w-3.5 h-3.5 text-sky-700" />
                          <span>Chọn tệp .json từ máy tính</span>
                          <input
                            type="file"
                            accept=".json,application/json"
                            className="hidden"
                            onChange={handleServiceAccountFileUpload}
                          />
                        </label>
                        {gdServiceAccountJson && (
                          <span className="text-[11px] text-sky-700 font-semibold">
                            Đã nạp nội dung JSON mới (bấm Lưu cấu hình để áp dụng)
                          </span>
                        )}
                      </div>

                      <textarea
                        value={gdServiceAccountJson}
                        onChange={(e) => {
                          setGdServiceAccountJson(e.target.value);
                          setDriveDirty(true);
                        }}
                        disabled={isConfigBlocked}
                        rows={4}
                        placeholder={
                          currentDriveCfg?.hasServiceAccount
                            ? 'Để trống nếu giữ nguyên khóa JSON hiện tại, hoặc dán JSON mới để thay thế...'
                            : '{\n  "type": "service_account",\n  "client_email": "...",\n  "private_key": "-----BEGIN PRIVATE KEY-----..."\n}'
                        }
                        className="w-full p-2.5 bg-white border border-slate-300 rounded-lg font-mono text-[11px] focus:ring-2 focus:ring-sky-600"
                      />
                    </div>
                  </div>
                )}

                {currentDriveCfg?.lastCheckedAt && !gdTestResult && (
                  <div className="text-[11px] text-slate-500 flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-slate-400" />
                    <span>
                      Kiểm tra gần nhất ({new Date(currentDriveCfg.lastCheckedAt).toLocaleString('vi-VN')}):{' '}
                      <strong className={currentDriveCfg.lastCheckStatus === 'ok' ? 'text-emerald-700' : 'text-rose-700'}>
                        {currentDriveCfg.lastCheckMessage}
                      </strong>
                    </span>
                  </div>
                )}

                {gdSaveMessage && (
                  <div
                    className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                      gdSaveMessage.success
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : 'bg-rose-50 text-rose-800 border border-rose-200'
                    }`}
                  >
                    {gdSaveMessage.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    )}
                    <span>{gdSaveMessage.message}</span>
                  </div>
                )}

                {gdTestResult && (
                  <div
                    className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                      gdTestResult.success
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : 'bg-rose-50 text-rose-800 border border-rose-200'
                    }`}
                  >
                    {gdTestResult.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    )}
                    <span>{gdTestResult.message}</span>
                  </div>
                )}

                {canManageSelectedBook && (
                  <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2.5 sm:gap-3 pt-2">
                    <button
                      type="submit"
                      disabled={gdSaving || isConfigBlocked}
                      className="min-h-[44px] px-4 py-2.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center justify-center gap-1.5"
                    >
                      <Save className="w-3.5 h-3.5" />
                      <span>{gdSaving ? 'Đang lưu...' : 'Lưu cấu hình Google Drive'}</span>
                    </button>

                    <button
                      type="button"
                      onClick={handleTestGoogleDrive}
                      disabled={gdTesting || !parsedPreviewFolderId}
                      className="min-h-[44px] px-4 py-2.5 bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-800 font-semibold rounded-lg transition-colors flex items-center justify-center gap-1.5"
                    >
                      <RotateCw className={`w-3.5 h-3.5 ${gdTesting ? 'animate-spin' : ''}`} />
                      <span>{gdTesting ? 'Đang kiểm tra quyền ghi...' : 'Kiểm tra kết nối'}</span>
                    </button>
                  </div>
                )}
              </form>
            </div>
          </div>
        );
      })()}

      {/* SUB-TAB: GOOGLE TRANG TÍNH (ADMIN ONLY) */}
      {activeSubTab === 'googlesheets' && isAdmin && (
        <div className="space-y-6 max-w-3xl min-w-0">
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-5 min-w-0">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <div className="min-w-0 flex-1">
                <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded bg-emerald-600 text-white flex items-center justify-center text-xs font-bold shadow-2xs shrink-0">
                    GS
                  </span>
                  <span>Đồng Bộ Vào Google Trang Tính (Google Sheets)</span>
                </h2>
                <p className="text-xs text-slate-600 mt-1">
                  Tự động thêm/cập nhật từng dòng văn bản sang Google Trang tính của trường.
                </p>
              </div>

              <div className="shrink-0">
                <span
                  className={`text-xs px-2.5 py-1 rounded font-semibold ${
                    gsEnabled ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-600'
                  }`}
                >
                  {gsEnabled ? 'Đang hoạt động' : 'Chưa kích hoạt'}
                </span>
              </div>
            </div>

            <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-3.5 sm:p-4 text-xs space-y-3 min-w-0">
              <h3 className="font-bold text-emerald-950 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" />
                <span>Cách tạo liên kết đồng bộ đến Google Sheets:</span>
              </h3>
              <ol className="list-decimal pl-5 space-y-1.5 text-slate-700">
                <li>Mở Google Trang tính &rarr; Tiện ích mở rộng &rarr; Apps Script.</li>
                <li>Dán đoạn mã bên dưới, bấm Lưu và Triển khai dạng Ứng dụng web (Execute as: Tôi, Access: Bất kỳ ai).</li>
              </ol>

              <div className="rounded-lg overflow-hidden border border-slate-800 bg-slate-900 min-w-0">
                <div className="flex items-center justify-between gap-2 px-3 py-2 bg-slate-800/80 border-b border-slate-700">
                  <span className="text-[11px] font-mono text-slate-300">Apps Script Google Sheets</span>
                  <button
                    type="button"
                    onClick={() => {
                      navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE);
                      setCopiedScript(true);
                      setTimeout(() => setCopiedScript(false), 2500);
                    }}
                    className="px-2.5 py-1.5 min-h-[32px] text-xs font-semibold bg-emerald-700 hover:bg-emerald-600 text-white rounded shadow transition-colors inline-flex items-center gap-1 shrink-0"
                  >
                    {copiedScript ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                    <span>{copiedScript ? 'Đã sao chép' : 'Sao chép mã'}</span>
                  </button>
                </div>
                <pre className="p-3 bg-slate-900 text-slate-100 text-[11px] font-mono overflow-x-auto max-h-56 leading-relaxed w-full max-w-full">
{GOOGLE_APPS_SCRIPT_TEMPLATE}
                </pre>
              </div>
            </div>

            <form onSubmit={handleSaveGoogleSheets} className="space-y-4 text-xs">
              <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                <input
                  type="checkbox"
                  checked={gsEnabled}
                  onChange={(e) => setGsEnabled(e.target.checked)}
                  className="rounded text-emerald-700 focus:ring-emerald-700"
                />
                <span className="font-semibold text-slate-900">
                  Bật tự động đồng bộ sang Google Trang tính
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="space-y-1 sm:col-span-2 min-w-0">
                  <label className="font-bold text-slate-700">
                    URL Ứng Dụng Web Google Apps Script Master (/exec) *
                  </label>
                  <input
                    type="url"
                    value={gsWebhookUrl}
                    onChange={(e) => setGsWebhookUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/AKfycbx.../exec"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs focus:ring-2 focus:ring-emerald-600"
                  />
                  <span className="text-[11px] text-slate-500 block break-words">
                    Tự động chia mỗi sổ thành 1 tab theo Mã sổ cố định: {books.map((b) => b.code).join(', ')}.
                  </span>
                </div>

                <div className="space-y-1 sm:col-span-2 min-w-0">
                  <label className="font-bold text-slate-700">Sổ dùng để kiểm tra / đồng bộ theo sổ</label>
                  <select
                    value={gsTestBookId}
                    onChange={(e) => setGsTestBookId(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs font-semibold"
                  >
                    {books.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name} — Tab: {b.code} ({b.unitCode})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {gsTestResult && (
                <div
                  className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                    gsTestResult.success
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}
                >
                  <span>{gsTestResult.message}</span>
                </div>
              )}

              {gsSyncResult && (
                <div className="p-3 bg-slate-100 text-slate-800 rounded-lg text-xs">{gsSyncResult}</div>
              )}

              <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2.5 sm:gap-3 pt-2">
                <button
                  type="submit"
                  disabled={gsSaving || isConfigBlocked}
                  className="min-h-[44px] px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {gsSaving ? 'Đang lưu...' : 'Lưu cấu hình Master Sheets'}
                </button>

                <button
                  type="button"
                  onClick={handleTestGoogleSheets}
                  disabled={gsTesting || !gsWebhookUrl}
                  className="min-h-[44px] px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg transition-colors flex items-center justify-center gap-1.5"
                >
                  <RotateCw className={`w-3.5 h-3.5 ${gsTesting ? 'animate-spin' : ''}`} />
                  <span>{gsTesting ? 'Đang gửi kiểm tra...' : 'Gửi dòng thử nghiệm lên tab sổ'}</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleSyncAllGoogleSheets()}
                  disabled={gsSyncingAll || !gsEnabled || !gsWebhookUrl}
                  className="min-h-[44px] px-4 py-2.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {gsSyncingAll ? 'Đang chuyển toàn bộ...' : 'Đồng bộ toàn bộ các sổ'}
                </button>
              </div>
            </form>

            {/* Optional per-book Google Sheets webhook */}
            <div className="pt-5 border-t border-slate-200 space-y-3">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  Google Sheets riêng cho sổ "{selectedBook.name}" (tùy chọn)
                </h3>
                <p className="text-[11px] text-slate-500">
                  Dùng khi cần đẩy thêm bản sao riêng cho bảng tính nội bộ của {selectedBook.name} bên cạnh bảng tính Master của trường.
                </p>
              </div>

              <form onSubmit={handleSaveBookGoogleSheets} className="space-y-3 text-xs">
                <label className="flex items-center gap-2.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={bookGsEnabled}
                    onChange={(e) => {
                      setBookGsEnabled(e.target.checked);
                      setBookGsDirty(true);
                    }}
                    disabled={isConfigBlocked}
                    className="rounded text-emerald-700"
                  />
                  <span className="font-semibold text-slate-800">
                    Bật đồng bộ thêm sang Google Sheets riêng của sổ "{selectedBook.name}" ({selectedBook.code})
                  </span>
                </label>

                <input
                  type="url"
                  value={bookGsWebhookUrl}
                  onChange={(e) => {
                    setBookGsWebhookUrl(e.target.value);
                    setBookGsDirty(true);
                  }}
                  disabled={isConfigBlocked}
                  placeholder="https://script.google.com/macros/s/.../exec (để trống nếu chỉ dùng bảng tính Master)"
                  className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                />

                {bookGsMessage && (
                  <div className="p-2.5 rounded-lg bg-slate-100 text-slate-800 text-xs">{bookGsMessage}</div>
                )}

                <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2">
                  <button
                    type="submit"
                    disabled={bookGsSaving || isConfigBlocked}
                    className="min-h-[42px] px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white font-semibold rounded-lg"
                  >
                    {bookGsSaving ? 'Đang lưu...' : `Lưu Sheets riêng của ${selectedBook.code}`}
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSyncAllGoogleSheets(selectedBook.id)}
                    disabled={gsSyncingAll}
                    className="min-h-[42px] px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg"
                  >
                    Đồng bộ riêng sổ {selectedBook.name}
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* SUB-TAB: MICROSOFT GRAPH EXCEL (ADMIN ONLY) */}
      {activeSubTab === 'msgraph' && isAdmin && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-5">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <div className="min-w-0">
                <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Cloud className="w-5 h-5 text-sky-700 shrink-0" />
                  <span>Đồng Bộ Tự Động Vào Excel Trực Tuyến (OneDrive / SharePoint)</span>
                </h2>
              </div>
              <span
                className={`text-xs px-2.5 py-1 rounded font-semibold shrink-0 ${
                  msEnabled ? 'bg-emerald-50 text-emerald-800' : 'bg-slate-100 text-slate-600'
                }`}
              >
                {msEnabled ? 'Đã bật đồng bộ' : 'Chưa bật'}
              </span>
            </div>

            <form onSubmit={handleSaveMsGraph} className="space-y-4 text-xs">
              <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                <input
                  type="checkbox"
                  checked={msEnabled}
                  onChange={(e) => setMsEnabled(e.target.checked)}
                  className="rounded text-sky-700 focus:ring-sky-700"
                />
                <span className="font-semibold text-slate-900">
                  Bật tính năng tự động đồng bộ vào Microsoft Graph Excel
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Azure Directory (Tenant) ID</label>
                  <input
                    type="text"
                    value={msTenantId}
                    onChange={(e) => setMsTenantId(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Application (Client) ID</label>
                  <input
                    type="text"
                    value={msClientId}
                    onChange={(e) => setMsClientId(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
                <div className="space-y-1 sm:col-span-2 min-w-0">
                  <label className="font-bold text-slate-700">Client Secret</label>
                  <input
                    type="password"
                    value={msClientSecret}
                    onChange={(e) => setMsClientSecret(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Drive ID (Tùy chọn)</label>
                  <input
                    type="text"
                    value={msDriveId}
                    onChange={(e) => setMsDriveId(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Đường dẫn file Excel</label>
                  <input
                    type="text"
                    value={msFilePath}
                    onChange={(e) => setMsFilePath(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
                <div className="space-y-1 sm:col-span-2 min-w-0">
                  <label className="font-bold text-slate-700">Tên bảng dữ liệu (Table Name)</label>
                  <input
                    type="text"
                    value={msTableName}
                    onChange={(e) => setMsTableName(e.target.value)}
                    className="w-full min-h-[44px] h-11 sm:h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
                  />
                </div>
              </div>

              {msSaveMessage && (
                <div
                  className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                    msSaveMessage.success
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}
                >
                  {msSaveMessage.success ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                  ) : (
                    <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  )}
                  <span>{msSaveMessage.message}</span>
                </div>
              )}

              {msTestResult && (
                <div
                  className={`p-3 rounded-lg text-xs ${
                    msTestResult.success
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}
                >
                  {msTestResult.message}
                </div>
              )}

              {msSyncResult && (
                <div className="p-3 bg-slate-100 text-slate-800 rounded-lg text-xs">{msSyncResult}</div>
              )}

              <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-2.5 sm:gap-3 pt-2">
                <button
                  type="submit"
                  disabled={msSaving || isConfigBlocked}
                  className="min-h-[44px] px-4 py-2.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {msSaving ? 'Đang lưu...' : 'Lưu cấu hình'}
                </button>
                <button
                  type="button"
                  onClick={handleTestMsConnection}
                  disabled={msTesting || !msTenantId || !msClientId}
                  className="min-h-[44px] px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg transition-colors flex items-center justify-center gap-1.5"
                >
                  <RotateCw className={`w-3.5 h-3.5 ${msTesting ? 'animate-spin' : ''}`} />
                  <span>{msTesting ? 'Đang kiểm tra...' : 'Kiểm tra kết nối'}</span>
                </button>
                <button
                  type="button"
                  onClick={handleSyncAll}
                  disabled={msSyncingAll || !msEnabled}
                  className="min-h-[44px] px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {msSyncingAll ? 'Đang đồng bộ...' : 'Đồng bộ lại tất cả'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* SUB-TAB: QUẢN LÝ TÀI KHOẢN & PHÂN QUYỀN THEO TỪNG SỔ (ADMIN ONLY) */}
      {activeSubTab === 'users' && isAdmin && (
        <div className="space-y-6 max-w-4xl">
          {userFeedback && (
            <div
              className={`p-3.5 rounded-xl border text-xs flex items-start justify-between gap-2 ${
                userFeedback.success
                  ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                  : 'bg-rose-50 text-rose-800 border-rose-200'
              }`}
            >
              <div className="flex items-start gap-2 min-w-0">
                {userFeedback.success ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                )}
                <span className="break-words">{userFeedback.message}</span>
              </div>
              <button
                type="button"
                onClick={() => setUserFeedback(null)}
                className="text-[11px] font-semibold opacity-75 hover:opacity-100 shrink-0"
              >
                Đóng
              </button>
            </div>
          )}

          <div className="bg-sky-50 border border-sky-200 rounded-xl p-4 text-xs text-sky-900 space-y-2">
            <div className="flex items-center gap-2 font-bold text-sky-950 text-sm">
              <ShieldCheck className="w-4 h-4 text-sky-700" />
              <span>Chính sách tài khoản & Phân quyền chi tiết theo từng sổ</span>
            </div>
            <ul className="list-disc list-inside space-y-1 text-slate-700 leading-relaxed">
              <li>
                Mặc định <strong>01 tài khoản Quản trị viên (admin)</strong> cố định toàn quyền quản lý hệ thống.
              </li>
              <li>
                Quản trị viên gán <strong>Họ và tên cố định</strong> cho từng tài khoản tạo mới; người dùng không thể tự thay đổi tên này.
              </li>
              <li>
                Mỗi tài khoản có thể được phân chi tiết <strong>9 nhóm quyền theo từng sổ văn bản</strong>:{' '}
                <em>Xem, Nhập văn bản đến, Xin số, Cấp số, Sửa, Thay tệp, Hủy, Xuất sổ, Quản trị sổ</em>.
                Quyền <strong>Xin số</strong> và <strong>Cấp số</strong> được tách biệt hoàn toàn.
              </li>
            </ul>
          </div>

          {/* Account List */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100">
              <h2 className="text-base font-bold text-slate-900">
                Danh Sách Tài Khoản Hệ Thống
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Bấm <strong>Phân quyền sổ</strong> để điều chỉnh 9 quyền chi tiết trên từng sổ văn bản cho mỗi tài khoản.
              </p>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {loadingUsers ? (
                <div className="p-6 text-center text-slate-400">Đang tải danh sách tài khoản...</div>
              ) : (
                usersList.map((u) => {
                  const isMainAdmin = isSystemAdmin(u);
                  const isEditingPerms = editingPermUserId === u.id;
                  return (
                    <div key={u.id} className="p-4 space-y-3">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div className="space-y-1.5">
                          <div className="flex items-center gap-2 flex-wrap">
                            <strong className="text-slate-900 font-bold text-sm">{u.fullName}</strong>
                            <span className="font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded text-[11px]">
                              @{u.username}
                            </span>
                            {isMainAdmin ? (
                              <span className="bg-amber-100 text-amber-900 font-bold px-2 py-0.5 rounded text-[10px] flex items-center gap-1">
                                <Shield className="w-3 h-3 text-amber-700" />
                                Quản trị viên cố định (Toàn quyền)
                              </span>
                            ) : (
                              <span className="bg-sky-100 text-sky-900 font-medium px-2 py-0.5 rounded text-[10px]">
                                Tên gán cố định
                              </span>
                            )}
                          </div>

                          {!isMainAdmin && (
                            <div className="flex flex-wrap gap-1.5">
                              {books
                                .filter((b) => b.isActive)
                                .map((book) => {
                                  const perms = u.bookPermissions?.[book.id] || [];
                                  if (perms.length === 0) return null;
                                  return (
                                    <span
                                      key={book.id}
                                      className="inline-flex items-center gap-1 px-2 py-0.5 bg-slate-100 border border-slate-200 rounded text-[10px] text-slate-700"
                                    >
                                      <strong className="text-sky-900">{book.unitCode}:</strong>
                                      <span>
                                        {perms.map((p) => BOOK_PERMISSION_META[p]?.shortLabel || p).join(', ')}
                                      </span>
                                    </span>
                                  );
                                })}
                            </div>
                          )}

                          <div className="text-[11px] text-slate-500">
                            Ngày tạo: {new Date(u.createdAt).toLocaleDateString('vi-VN')}
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-2 shrink-0">
                          <span
                            className={`text-[11px] px-2.5 py-1 rounded-md font-semibold ${
                              u.isActive
                                ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                : 'bg-rose-50 text-rose-800 border border-rose-200'
                            }`}
                          >
                            {u.isActive ? 'Đang hoạt động' : 'Tạm khóa'}
                          </span>

                          {!isMainAdmin && (
                            <button
                              type="button"
                              onClick={() => handleStartEditPermissions(u)}
                              className={`px-2.5 py-1.5 min-h-[36px] rounded-md font-semibold text-[11px] border transition-colors flex items-center gap-1 ${
                                isEditingPerms
                                  ? 'bg-sky-700 text-white border-sky-700'
                                  : 'bg-sky-50 hover:bg-sky-100 text-sky-800 border-sky-200'
                              }`}
                            >
                              <Users className="w-3 h-3" />
                              <span>{isEditingPerms ? 'Đóng phân quyền' : 'Phân quyền sổ'}</span>
                            </button>
                          )}

                          <button
                            type="button"
                            onClick={() => handleResetUserPass(u)}
                            className="px-2.5 py-1.5 min-h-[36px] bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md font-medium text-[11px] transition-colors flex items-center gap-1"
                          >
                            <Lock className="w-3 h-3 text-slate-500" />
                            <span>Đổi MK</span>
                          </button>

                          {!isMainAdmin && (
                            <>
                              <button
                                type="button"
                                onClick={() => handleToggleUserStatus(u)}
                                className={`px-2.5 py-1.5 min-h-[36px] rounded-md font-medium text-[11px] transition-colors flex items-center gap-1 ${
                                  u.isActive
                                    ? 'bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200'
                                    : 'bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200'
                                }`}
                              >
                                {u.isActive ? <UserX className="w-3 h-3" /> : <UserCheck className="w-3 h-3" />}
                                <span>{u.isActive ? 'Khóa' : 'Mở'}</span>
                              </button>

                              <button
                                type="button"
                                onClick={() => handleDeleteUser(u)}
                                className="p-2 min-w-[36px] min-h-[36px] inline-flex items-center justify-center text-slate-400 hover:text-rose-700 hover:bg-rose-50 rounded-md transition-colors"
                                title="Xóa tài khoản này khỏi hệ thống"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </>
                          )}
                        </div>
                      </div>

                      {!isMainAdmin && isEditingPerms && (
                        <div className="pt-3 border-t border-slate-100 space-y-3">
                          {renderPermissionMatrix(editingBookPermissions, setEditingBookPermissions)}
                          <div className="flex flex-col sm:flex-row justify-end gap-2">
                            <button
                              type="button"
                              onClick={() => setEditingPermUserId(null)}
                              className="min-h-[42px] px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-semibold"
                            >
                              Hủy
                            </button>
                            <button
                              type="button"
                              onClick={() => handleSaveUserPermissions(u)}
                              disabled={savingPermUserId === u.id || isConfigBlocked}
                              className="min-h-[42px] px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white rounded-lg font-semibold flex items-center justify-center gap-1.5 disabled:opacity-50"
                            >
                              <Save className="w-3.5 h-3.5" />
                              <span>{savingPermUserId === u.id ? 'Đang lưu...' : 'Lưu phân quyền sổ'}</span>
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          </div>

          {/* Add New User Form */}
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs">
            <h3 className="text-sm font-bold text-slate-900 mb-1 flex items-center gap-2">
              <Plus className="w-4 h-4 text-sky-700 shrink-0" />
              <span>Thêm tài khoản mới và phân quyền theo từng sổ</span>
            </h3>
            <p className="text-xs text-slate-500 mb-4">
              Họ và tên được gán cố định cho tài khoản. Chọn rõ các quyền trên từng sổ (Xin số tách biệt với Cấp số).
            </p>

            <form onSubmit={handleCreateUser} className="space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Tên đăng nhập *</label>
                  <input
                    type="text"
                    value={newUsername}
                    onChange={(e) => setNewUsername(e.target.value)}
                    placeholder="Ví dụ: chidoan12a"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-sm"
                  />
                </div>

                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Họ và tên (Gán cố định) *</label>
                  <input
                    type="text"
                    value={newFullName}
                    onChange={(e) => setNewFullName(e.target.value)}
                    placeholder="Ví dụ: Nguyễn Văn A"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-sm"
                  />
                </div>

                <div className="space-y-1 min-w-0">
                  <label className="font-bold text-slate-700">Mật khẩu khởi tạo *</label>
                  <input
                    type="password"
                    value={newUserPass}
                    onChange={(e) => setNewUserPass(e.target.value)}
                    placeholder="Tối thiểu 6 ký tự"
                    required
                    className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-sm"
                  />
                </div>
              </div>

              <div className="space-y-2">
                <label className="font-bold text-slate-700 block">
                  Phân quyền theo từng sổ văn bản
                </label>
                {renderPermissionMatrix(newBookPermissions, setNewBookPermissions)}
              </div>

              {userError && (
                <div className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs border border-rose-200 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>{userError}</span>
                </div>
              )}

              <div className="pt-2 flex items-center justify-end">
                <button
                  type="submit"
                  disabled={userSaving || isConfigBlocked}
                  className="w-full sm:w-auto min-h-[44px] px-6 py-2.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center justify-center gap-2 disabled:opacity-50"
                >
                  <Plus className="w-4 h-4" />
                  <span>{userSaving ? 'Đang tạo tài khoản...' : 'Tạo tài khoản & áp dụng quyền'}</span>
                </button>
              </div>
            </form>
          </div>

          {/* Nhật ký thay đổi phân quyền & quản trị sổ */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  Nhật ký thay đổi phân quyền & cấu hình sổ ({auditLogs.length})
                </h3>
                <p className="text-[11px] text-slate-500">
                  Ghi nhận các thao tác gán/thu hồi quyền, khóa tài khoản và tạo/sửa sổ để phục vụ kiểm tra.
                </p>
              </div>
              <button
                type="button"
                onClick={handleOpenUsersTab}
                className="text-xs font-semibold text-sky-700 hover:underline flex items-center gap-1"
              >
                <RotateCw className="w-3.5 h-3.5" />
                <span>Làm mới</span>
              </button>
            </div>

            <div className="divide-y divide-slate-100 text-xs max-h-80 overflow-y-auto">
              {loadingAuditLogs ? (
                <div className="p-4 text-center text-slate-400">Đang tải nhật ký...</div>
              ) : auditLogs.length === 0 ? (
                <div className="p-4 text-center text-slate-400">Chưa có nhật ký thay đổi phân quyền hoặc sổ nào.</div>
              ) : (
                auditLogs.slice(0, 40).map((log) => (
                  <div key={log.id} className="p-3.5 space-y-1">
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px]">
                      <div className="flex items-center gap-2">
                        <span className="px-2 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-900 font-mono font-bold">
                          {log.action}
                        </span>
                        <strong className="text-slate-800">{log.actorName}</strong>
                        {log.bookCode && (
                          <span className="font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-700">
                            {log.bookCode}
                          </span>
                        )}
                      </div>
                      <span className="font-mono text-slate-400">
                        {new Date(log.timestamp).toLocaleString('vi-VN')}
                      </span>
                    </div>
                    <p className="text-slate-700">{log.details}</p>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}
        </div>
      </div>

      {/* Slide-over drawer for + Tạo sổ mới */}
      <CreateBookDrawer
        open={createBookDrawerOpen}
        onClose={() => !bookSaving && setCreateBookDrawerOpen(false)}
        existingBooks={books}
        documentTypes={documentTypes}
        saving={bookSaving}
        isConfigBlocked={isConfigBlocked}
        serverError={createBookServerError}
        onClearServerError={() => setCreateBookServerError(null)}
        onSubmitBook={handleCreateBookSubmit}
      />

      {/* Unsaved-changes guard when switching Sổ đang cấu hình */}
      <PendingBookSwitchModal
        pendingSwitch={pendingBookSwitch}
        saving={switchingBookSaving}
        errorMessage={switchBookError}
        onSaveAndSwitch={handleSaveAndSwitchBook}
        onDiscardAndSwitch={handleDiscardAndSwitchBook}
        onCancel={() => {
          if (!switchingBookSaving) {
            setPendingBookSwitch(null);
            setSwitchBookError(null);
          }
        }}
      />

      {/* Explicit confirmation for Tạm khóa / Mở khóa sổ */}
      <ToggleBookLockModal
        book={toggleLockConfirmBook}
        saving={toggleLockSaving}
        onConfirm={handleConfirmToggleBookLock}
        onCancel={() => !toggleLockSaving && setToggleLockConfirmBook(null)}
      />

      {/* Generic confirmation modal (delete doc type / signer / user) */}
      <ConfirmDeleteModal
        request={confirmDeleteRequest}
        busy={confirmDeleteBusy}
        onClose={() => !confirmDeleteBusy && setConfirmDeleteRequest(null)}
      />

      {/* Password reset modal */}
      <ResetPasswordModal
        user={resetPassUser}
        saving={resetPassSaving}
        onSubmit={handleConfirmResetPassword}
        onClose={() => !resetPassSaving && setResetPassUser(null)}
      />
    </div>
  );
};

