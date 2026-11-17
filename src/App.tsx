import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { Navbar, type AppTab, type RegistrySection } from './components/Navbar';
import { IssueNumberView } from './components/IssueNumberView';
import { ImportPdfView } from './components/ImportPdfView';
import { DocumentRegistryView } from './components/DocumentRegistryView';
import { TasksView } from './components/TasksView';
import { ReconciliationView } from './components/ReconciliationView';
import { StatisticsView } from './components/StatisticsView';
import { SettingsView } from './components/SettingsView';
import type { SettingsSubTab } from './components/settings/settingsConstants';
import { LoginModal } from './components/LoginModal';
import { api, buildApiUrl, clearAuth, updateStoredUser } from './api';
import type {
  DocumentRecord,
  DocumentTypeConfig,
  NumberRequest,
  PersistenceStatus,
  RegistryBook,
  SignerConfig,
  SystemSettings,
  User,
} from './types';
import {
  DEFAULT_BOOK_ID,
  getAccessibleBooks,
  getDocBookId,
  hasAnyBookPermission,
  hasBookPermission,
  isSystemAdmin,
} from './utils/permissions';
import { getDocumentFilterYear } from './utils/fileNaming';
import { AlertTriangle, Bell, BookOpen, RefreshCw } from 'lucide-react';

function getClientVietnamTimeFallback() {
  try {
    const formatter = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Ho_Chi_Minh',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    });
    const parts = formatter.format(new Date()).split('-');
    return {
      year: parseInt(parts[0], 10),
      formattedDate: `${parts[2]}/${parts[1]}/${parts[0]}`,
    };
  } catch {
    const now = new Date();
    return {
      year: now.getFullYear(),
      formattedDate: now.toLocaleDateString('vi-VN'),
    };
  }
}

export default function App() {
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [authChecking, setAuthChecking] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<AppTab>('issue');
  const [registrySection, setRegistrySection] = useState<RegistrySection>('active');
  const [registryScopeAll, setRegistryScopeAll] = useState<boolean>(false);
  const [registryPresetCategory, setRegistryPresetCategory] = useState<'all' | 'outgoing' | 'incoming' | null>(null);
  const [pendingOpenDocId, setPendingOpenDocId] = useState<string | null>(null);

  const initialVnTime = useMemo(() => getClientVietnamTimeFallback(), []);
  const [serverYear, setServerYear] = useState<number>(initialVnTime.year);
  const [serverFormattedDate, setServerFormattedDate] = useState<string>(initialVnTime.formattedDate);

  // Core Data
  const [settings, setSettings] = useState<SystemSettings>({
    unitName: 'Đoàn Trường THPT Nguyễn Chí Thanh',
    unitCode: 'ĐTN',
    numberMode: 'separate',
    numberFormat: '{so}-{loai}/{don_vi}',
    customStartSequences: {},
  });
  const [books, setBooks] = useState<RegistryBook[]>([]);
  const [activeBookId, setActiveBookId] = useState<string>(DEFAULT_BOOK_ID);
  const [numberRequests, setNumberRequests] = useState<NumberRequest[]>([]);
  const [documentTypes, setDocumentTypes] = useState<DocumentTypeConfig[]>([]);
  const [signers, setSigners] = useState<SignerConfig[]>([]);
  const [documents, setDocuments] = useState<DocumentRecord[]>([]);
  const [deletedDocuments, setDeletedDocuments] = useState<DocumentRecord[]>([]);
  const [loadingData, setLoadingData] = useState<boolean>(false);
  const [settingsLoaded, setSettingsLoaded] = useState<boolean>(false);
  const [lastSyncAt, setLastSyncAt] = useState<Date | null>(null);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [persistenceStatus, setPersistenceStatus] = useState<PersistenceStatus | null>(null);
  const [legacyReconciliationWarnings, setLegacyReconciliationWarnings] = useState<string[]>([]);
  const [requestedSettingsTab, setRequestedSettingsTab] = useState<SettingsSubTab | null>(null);
  const syncInFlightRef = React.useRef<boolean>(false);

  // Live Toast Notification
  const [liveToast, setLiveToast] = useState<{ message: string; type?: 'info' | 'success' } | null>(null);

  const showToast = (message: string, type: 'info' | 'success' = 'info') => {
    setLiveToast({ message, type });
    setTimeout(() => setLiveToast(null), 4000);
  };

  const applyServerTime = useCallback((timeVn?: { year?: number; formattedDate?: string }) => {
    if (!timeVn) return;
    if (typeof timeVn.year === 'number' && Number.isFinite(timeVn.year)) {
      setServerYear((prev) => (prev !== timeVn.year ? timeVn.year! : prev));
    }
    if (typeof timeVn.formattedDate === 'string' && timeVn.formattedDate) {
      setServerFormattedDate((prev) => (prev !== timeVn.formattedDate ? timeVn.formattedDate! : prev));
    }
  }, []);

  // 1. Initial Auth Check
  useEffect(() => {
    const checkAuth = async () => {
      try {
        const res = await api.getCurrentUser();
        setCurrentUser(res.user);
      } catch {
        clearAuth();
        setCurrentUser(null);
      } finally {
        setAuthChecking(false);
      }
    };

    checkAuth();

    const handleExpired = () => {
      setCurrentUser(null);
      setSettingsLoaded(false);
      setPersistenceStatus(null);
      setLegacyReconciliationWarnings([]);
      setBooks([]);
      setDocuments([]);
      setDeletedDocuments([]);
      setNumberRequests([]);
    };
    window.addEventListener('auth_expired', handleExpired);
    return () => window.removeEventListener('auth_expired', handleExpired);
  }, []);

  const refreshBooks = useCallback(async () => {
    if (!currentUser) return;
    try {
      const res = await api.getBooks();
      setBooks(res.books);
      applyServerTime(res.serverTimeVietnam);
    } catch (err) {
      console.error('Lỗi tải danh sách sổ:', err);
    }
  }, [currentUser, applyServerTime]);

  const refreshRequests = useCallback(async () => {
    if (!currentUser) return;
    try {
      const res = await api.getNumberRequests();
      setNumberRequests(res.requests);
    } catch {
      // ignore if user has no access
    }
  }, [currentUser]);

  // 2. Load Core Application Data (Single Server Source of Truth, In-place refresh)
  const loadAppData = useCallback(async () => {
    if (!currentUser || syncInFlightRef.current) return;
    syncInFlightRef.current = true;
    setLoadingData(true);
    try {
      const [meRes, settingsRes, docsRes, deletedDocsRes, booksRes, reqsRes] = await Promise.all([
        api.getCurrentUser(),
        api.getSettings(),
        api.getDocuments({}),
        api.getDocuments({ status: 'deleted' }).catch(() => ({ total: 0, documents: [] as DocumentRecord[] })),
        api.getBooks(),
        api.getNumberRequests().catch(() => ({ requests: [] as NumberRequest[] })),
      ]);

      if (!meRes.user.isActive) {
        clearAuth();
        setCurrentUser(null);
        setSettingsLoaded(false);
        return;
      }

      updateStoredUser(meRes.user);
      setCurrentUser((prev) => {
        if (!prev) return meRes.user;
        if (
          JSON.stringify(prev.bookPermissions || {}) !== JSON.stringify(meRes.user.bookPermissions || {}) ||
          prev.role !== meRes.user.role ||
          prev.isActive !== meRes.user.isActive ||
          prev.fullName !== meRes.user.fullName ||
          (prev.displayName || '') !== (meRes.user.displayName || '')
        ) {
          return meRes.user;
        }
        return prev;
      });

      setSettings(settingsRes.settings);
      setDocumentTypes(settingsRes.documentTypes);
      setSigners(settingsRes.signers);
      setDocuments(docsRes.documents);
      setDeletedDocuments(deletedDocsRes.documents);
      setBooks(booksRes.books);
      setNumberRequests(reqsRes.requests);
      if (settingsRes.persistenceStatus) {
        setPersistenceStatus(settingsRes.persistenceStatus);
      } else if (meRes.persistenceStatus) {
        setPersistenceStatus(meRes.persistenceStatus);
      }
      setLegacyReconciliationWarnings(settingsRes.legacyReconciliationWarnings || []);
      applyServerTime(settingsRes.serverTimeVietnam || booksRes.serverTimeVietnam);

      setSettingsLoaded(true);
      setSyncError(null);
      setLastSyncAt(new Date());
    } catch (err: any) {
      console.error('Lỗi khi tải dữ liệu ứng dụng:', err);
      if (err?.persistenceStatus) {
        setPersistenceStatus(err.persistenceStatus);
      }
      if (err?.status !== 401) {
        setSyncError(err?.message || 'Không thể đồng bộ dữ liệu mới nhất từ máy chủ.');
      }
    } finally {
      syncInFlightRef.current = false;
      setLoadingData(false);
    }
  }, [currentUser?.id, applyServerTime]);

  useEffect(() => {
    if (currentUser?.id) {
      loadAppData();
    }
  }, [currentUser?.id, loadAppData]);

  // Tự động nạp dữ liệu mới nhất khi chuyển sổ hoặc mở Sổ theo dõi / Thống kê / Cài đặt
  useEffect(() => {
    if (!currentUser?.id || !settingsLoaded) return;
    if (activeTab === 'settings' || activeTab === 'registry' || activeTab === 'statistics') {
      loadAppData();
    }
  }, [activeTab, activeBookId, currentUser?.id, settingsLoaded, loadAppData]);

  // 3. Server-Sent Events (SSE) for Real-Time Multi-Device Updates
  useEffect(() => {
    if (!currentUser) return;

    let eventSource: EventSource | null = null;
    let reconnectTimeout: ReturnType<typeof setTimeout>;

    const connectSSE = () => {
      eventSource = new EventSource(buildApiUrl('/api/events'), { withCredentials: true });

      eventSource.addEventListener('document_issued', (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data);
          const newDoc: DocumentRecord = data.document;
          setDocuments((prev) => {
            if (prev.some((d) => d.id === newDoc.id)) return prev;
            return [newDoc, ...prev];
          });
          showToast(`Văn bản mới vừa được ghi: ${newDoc.fullDocNumber}`, 'success');
        } catch (err) {
          console.error(err);
        }
      });

      eventSource.addEventListener('document_updated', (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data);
          const updatedDoc: DocumentRecord = data.document;
          if (updatedDoc.status === 'deleted') {
            setDocuments((prev) => prev.filter((d) => d.id !== updatedDoc.id));
            setDeletedDocuments((prev) => {
              if (updatedDoc.deletedListArchived) {
                return prev.filter((d) => d.id !== updatedDoc.id);
              }
              const exists = prev.some((d) => d.id === updatedDoc.id);
              return exists
                ? prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
                : [updatedDoc, ...prev];
            });
          } else {
            setDocuments((prev) =>
              prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
            );
          }
        } catch (err) {
          console.error(err);
        }
      });

      eventSource.addEventListener('document_cancelled', (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data);
          const cancelledDoc: DocumentRecord = data.document;
          setDocuments((prev) =>
            prev.map((d) => (d.id === cancelledDoc.id ? cancelledDoc : d))
          );
          showToast(`Văn bản ${cancelledDoc.fullDocNumber} đã được hủy`, 'info');
        } catch (err) {
          console.error(err);
        }
      });

      eventSource.addEventListener('document_deleted', (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data);
          if (data.id) {
            setDocuments((prev) => prev.filter((d) => d.id !== data.id));
          }
          if (data.document) {
            const deletedDoc: DocumentRecord = data.document;
            setDeletedDocuments((prev) => {
              const exists = prev.some((d) => d.id === deletedDoc.id);
              return exists
                ? prev.map((d) => (d.id === deletedDoc.id ? deletedDoc : d))
                : [deletedDoc, ...prev];
            });
            showToast(`Văn bản ${deletedDoc.fullDocNumber} đã chuyển sang VB đã xóa`, 'info');
          }
        } catch (err) {
          console.error(err);
        }
      });

      eventSource.addEventListener('books_updated', () => {
        refreshBooks();
      });

      eventSource.addEventListener('user_permissions_updated', async (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data || '{}');
          const matchesMe =
            (!data.userId && !data.targetUserId) ||
            data.userId === currentUser.id ||
            data.targetUserId === currentUser.id;
          if (matchesMe) {
            const permsUnchanged =
              data.user &&
              JSON.stringify(currentUser.bookPermissions || {}) ===
                JSON.stringify(data.user.bookPermissions || {}) &&
              currentUser.isActive === data.user.isActive &&
              currentUser.role === data.user.role;
            if (data.user && permsUnchanged && (currentUser.displayName || '') !== (data.user.displayName || '')) {
              updateStoredUser(data.user);
              setCurrentUser(data.user);
              showToast('Tên hiển thị tài khoản của bạn vừa được cập nhật.', 'info');
            } else {
              await loadAppData();
              showToast('Quyền truy cập sổ văn bản của bạn vừa được cập nhật.', 'info');
            }
          }
        } catch {
          // ignore
        }
      });

      eventSource.addEventListener('number_requests_updated', () => {
        refreshRequests();
      });

      eventSource.addEventListener('documents_deleted_archived', () => {
        loadAppData();
      });

      eventSource.addEventListener('document_history_cleaned', () => {
        loadAppData();
      });

      eventSource.addEventListener('settings_updated', (e: MessageEvent) => {
        try {
          const data = JSON.parse(e.data);
          if (data.settings) {
            if (isSystemAdmin(currentUser) || hasAnyBookPermission(currentUser, 'manage_book')) {
              api.getSettings().then((res) => {
                setSettings(res.settings);
                applyServerTime(res.serverTimeVietnam);
              }).catch(() => {});
            } else {
              setSettings(data.settings);
            }
          }
          if (data.documentTypes) setDocumentTypes(data.documentTypes);
          if (data.signers) setSigners(data.signers);
        } catch (err) {
          console.error(err);
        }
      });

      eventSource.onerror = () => {
        if (eventSource) eventSource.close();
        reconnectTimeout = setTimeout(connectSSE, 4000);
      };
    };

    connectSSE();

    // Kiểm tra lại định kỳ mỗi 30 giây khi tab đang mở + khi quay lại tab hoặc có mạng trở lại
    const syncLiveState = () => {
      if (document.visibilityState !== 'visible') return;
      loadAppData();
    };

    const pollInterval = setInterval(syncLiveState, 30000);
    const midnightCheckInterval = setInterval(() => {
      if (document.visibilityState !== 'visible') return;
      const clientVn = getClientVietnamTimeFallback();
      if (clientVn.year !== serverYear) {
        loadAppData();
      }
    }, 10000);

    const handleVisibilityOrFocus = () => {
      syncLiveState();
    };
    const handleOnline = () => {
      loadAppData();
    };
    window.addEventListener('focus', handleVisibilityOrFocus);
    window.addEventListener('online', handleOnline);
    document.addEventListener('visibilitychange', handleVisibilityOrFocus);

    return () => {
      if (eventSource) eventSource.close();
      clearTimeout(reconnectTimeout);
      clearInterval(pollInterval);
      clearInterval(midnightCheckInterval);
      window.removeEventListener('focus', handleVisibilityOrFocus);
      window.removeEventListener('online', handleOnline);
      document.removeEventListener('visibilitychange', handleVisibilityOrFocus);
    };
  }, [currentUser?.id, serverYear, loadAppData, refreshBooks, refreshRequests, applyServerTime]);

  const accessibleBooks = useMemo(
    () => getAccessibleBooks(currentUser, books),
    [currentUser, books]
  );

  // Ensure activeBookId points to an accessible book
  useEffect(() => {
    if (accessibleBooks.length === 0 || !currentUser) return;
    if (!accessibleBooks.some((b) => b.id === activeBookId)) {
      setActiveBookId(accessibleBooks[0].id);
    }
  }, [accessibleBooks, activeBookId, currentUser]);

  const activeBook: RegistryBook = useMemo(() => {
    const found =
      accessibleBooks.find((b) => b.id === activeBookId) ||
      books.find((b) => b.id === activeBookId) ||
      accessibleBooks[0] ||
      books[0];
    if (found) return found;
    return {
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
      googleDriveConfig: settings.googleDriveConfig,
      isActive: true,
      createdAt: '',
    };
  }, [accessibleBooks, books, activeBookId, settings, signers]);

  // Enforce permission-based tab fallback strictly for activeBook & currentUser
  useEffect(() => {
    if (!currentUser || loadingData) return;
    const isAdmin = isSystemAdmin(currentUser);
    const hasPerm = (perm: 'issue_number' | 'request_number' | 'import_incoming' | 'view' | 'delete' | 'manage_book') => {
      if (isAdmin) return true;
      if (activeBook?.id && books.some((b) => b.id === activeBook.id)) {
        return hasBookPermission(currentUser, activeBook.id, perm);
      }
      return hasAnyBookPermission(currentUser, perm, books);
    };

    const allowedTabs: AppTab[] = [];
    if (hasPerm('issue_number') || hasPerm('request_number')) {
      allowedTabs.push('issue');
    }
    if (hasPerm('import_incoming')) {
      allowedTabs.push('import_pdf');
    }
    if (hasPerm('view') || hasPerm('delete')) {
      allowedTabs.push('registry');
    }
    if (hasPerm('view')) {
      allowedTabs.push('statistics');
    }
    allowedTabs.push('tasks');
    if (isAdmin || hasAnyBookPermission(currentUser, 'manage_book', books)) {
      allowedTabs.push('reconciliation');
    }
    if (hasPerm('manage_book') || isAdmin) {
      allowedTabs.push('settings');
    }

    if (allowedTabs.length > 0 && !allowedTabs.includes(activeTab)) {
      if (allowedTabs.includes('registry')) {
        setRegistrySection(hasPerm('view') ? 'active' : 'deleted');
      }
      setActiveTab(allowedTabs[0]);
    } else if (
      activeTab === 'registry' &&
      !hasPerm('view') &&
      hasPerm('delete') &&
      registrySection !== 'deleted'
    ) {
      setRegistrySection('deleted');
    }
  }, [currentUser, books, activeBook?.id, activeTab, registrySection, loadingData]);

  // Handlers
  const handleLogout = async () => {
    await api.logout();
    setCurrentUser(null);
    setSettingsLoaded(false);
    setLastSyncAt(null);
    setSyncError(null);
    setPersistenceStatus(null);
    setLegacyReconciliationWarnings([]);
    setBooks([]);
    setDocuments([]);
    setDeletedDocuments([]);
    setNumberRequests([]);
    setActiveBookId(DEFAULT_BOOK_ID);
  };

  const handleDocumentIssued = (newDoc: DocumentRecord) => {
    setDocuments((prev) => {
      if (prev.some((d) => d.id === newDoc.id)) return prev;
      return [newDoc, ...prev];
    });
    setLastSyncAt(new Date());
  };

  const handleDocumentUpdated = (updatedDoc: DocumentRecord) => {
    if (updatedDoc.status === 'deleted') {
      setDocuments((prev) => prev.filter((d) => d.id !== updatedDoc.id));
      setDeletedDocuments((prev) => {
        if (updatedDoc.deletedListArchived) {
          return prev.filter((d) => d.id !== updatedDoc.id);
        }
        const exists = prev.some((d) => d.id === updatedDoc.id);
        return exists
          ? prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
          : [updatedDoc, ...prev];
      });
    } else {
      setDocuments((prev) =>
        prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
      );
    }
    setLastSyncAt(new Date());
  };

  const handleOpenExistingDocument = (docId: string) => {
    setPendingOpenDocId(docId);
    setRegistrySection('active');
    setActiveTab('registry');
  };

  const handleNavigateToIncomingRegistry = () => {
    setRegistryScopeAll(false);
    setRegistryPresetCategory('incoming');
    setRegistrySection('active');
    setActiveTab('registry');
  };

  if (authChecking) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-center space-y-3">
          <div className="w-10 h-10 border-3 border-sky-700 border-t-transparent rounded-full animate-spin mx-auto" />
          <p className="text-xs text-slate-500 font-medium">Đang kiểm tra bảo mật...</p>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return <LoginModal onLoginSuccess={(user) => setCurrentUser(user)} />;
  }

  // Quick stats summary for server Vietnam year on active book (or all accessible books on registry)
  const currentYear = serverYear;
  const effectiveRegistryBookId =
    activeTab === 'registry' && registryScopeAll && accessibleBooks.length > 1
      ? 'all'
      : activeBook.id;
  const currentYearDocs = documents.filter((d) => {
    if (getDocumentFilterYear(d) !== currentYear) return false;
    if (effectiveRegistryBookId === 'all') return true;
    return getDocBookId(d) === activeBook.id;
  });
  const totalCount = currentYearDocs.length;
  const activeBookDeletedDocs = deletedDocuments.filter((d) =>
    effectiveRegistryBookId === 'all' ? true : getDocBookId(d) === activeBook.id
  );
  const pendingExcelSyncCount = activeBookDeletedDocs.filter(
    (d) => d.excelDeleteSyncStatus === 'failed' || d.excelDeleteSyncStatus === 'pending'
  ).length;
  const noBookPermission = settingsLoaded && accessibleBooks.length === 0;

  return (
    <div className="min-h-[100dvh] bg-slate-50 text-slate-900 flex flex-col font-sans overflow-x-clip">
      {/* Top Navigation Bar */}
      <Navbar
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        registrySection={registrySection}
        onSelectRegistrySection={setRegistrySection}
        deletedCount={activeBookDeletedDocs.length}
        pendingExcelSyncCount={pendingExcelSyncCount}
        currentUser={currentUser}
        books={books}
        activeBook={activeBook}
        onLogout={handleLogout}
      />

      {/* Thanh thông tin sổ: Bộ chọn sổ, năm, tổng số văn bản và nút Làm mới dạng biểu tượng */}
      <section
        aria-label="Thông tin sổ văn bản"
        className="bg-sky-900 text-white border-b border-sky-950 py-2.5 px-3 sm:px-6 lg:px-8"
      >
        <div className="max-w-7xl mx-auto flex items-center justify-between gap-2.5 sm:gap-4">
          <div className="flex flex-wrap items-center gap-x-3 sm:gap-x-4 gap-y-2 min-w-0 flex-1">
            {/* Bộ chọn sổ duy nhất */}
            {accessibleBooks.length > 0 && (
              <div className="flex items-center gap-2 bg-sky-800/90 border border-sky-600/80 rounded-lg px-2.5 h-9 shrink-0 max-w-full min-w-0">
                <BookOpen className="w-4 h-4 text-sky-200 shrink-0" />
                <span className="text-xs text-sky-200 font-medium whitespace-nowrap">Sổ:</span>
                {accessibleBooks.length > 1 ? (
                  <select
                    aria-label="Sổ đang chọn"
                    value={effectiveRegistryBookId}
                    onChange={(e) => {
                      const val = e.target.value;
                      if (val === 'all') {
                        setRegistryScopeAll(true);
                      } else {
                        setRegistryScopeAll(false);
                        setActiveBookId(val);
                      }
                    }}
                    className="min-w-0 max-w-[210px] sm:max-w-[270px] bg-transparent text-white text-xs sm:text-sm font-bold focus:outline-hidden cursor-pointer truncate [&>option]:text-slate-900"
                  >
                    {activeTab === 'registry' && (
                      <option value="all">Tất cả sổ có quyền ({accessibleBooks.length})</option>
                    )}
                    {accessibleBooks.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name} ({b.unitCode})
                      </option>
                    ))}
                  </select>
                ) : (
                  <span className="text-xs sm:text-sm font-bold text-white truncate max-w-[220px]">
                    {activeBook.name} ({activeBook.unitCode})
                  </span>
                )}
              </div>
            )}

            {/* Năm & Tổng số văn bản */}
            <div className="flex items-center gap-2 sm:gap-3 text-xs sm:text-sm text-sky-100 whitespace-nowrap tabular-nums">
              <span className="font-bold uppercase tracking-wider text-sky-200">
                Năm {currentYear}
              </span>
              <span className="text-sky-400" aria-hidden="true">·</span>
              <span>
                Tổng: <strong className="font-mono text-white font-bold">{totalCount}</strong> văn bản
              </span>
            </div>
          </div>

          {/* Nút Làm mới duy nhất */}
          <button
            type="button"
            onClick={loadAppData}
            disabled={loadingData}
            aria-label="Làm mới"
            title={
              lastSyncAt
                ? `Làm mới dữ liệu (Cập nhật lúc ${lastSyncAt.toLocaleTimeString('vi-VN', {
                    hour: '2-digit',
                    minute: '2-digit',
                    second: '2-digit',
                  })})`
                : 'Làm mới dữ liệu'
            }
            className="inline-flex items-center gap-1.5 h-9 px-3 rounded-lg bg-sky-800/90 hover:bg-sky-700 border border-sky-600/80 text-white text-xs font-semibold transition-colors shrink-0 disabled:opacity-60 whitespace-nowrap"
          >
            <RefreshCw className={`w-3.5 h-3.5 shrink-0 ${loadingData ? 'animate-spin' : ''}`} />
            <span>Làm mới</span>
          </button>
        </div>
      </section>

      {/* Sync Failure / Storage Attention Banner */}
      {syncError && (
        <div className="bg-rose-50 border-b border-rose-200 px-3 sm:px-6 lg:px-8 py-2.5">
          <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-2 text-xs text-rose-900">
            <div className="flex items-center gap-2 min-w-0 flex-1">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
              <span className="break-words min-w-0">
                <strong>
                  {persistenceStatus && !persistenceStatus.durable
                    ? `Lưu trữ chính: ${persistenceStatus.connectionLabel}`
                    : 'Chưa cập nhật được dữ liệu'}
                </strong>{' '}
                — {syncError}
              </span>
            </div>
            <div className="flex flex-wrap items-center gap-2 shrink-0">
              {isSystemAdmin(currentUser) && persistenceStatus && !persistenceStatus.durable && (
                <button
                  type="button"
                  onClick={() => {
                    setRequestedSettingsTab('database');
                    setActiveTab('settings');
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 min-h-[36px] bg-amber-600 hover:bg-amber-700 text-white font-semibold rounded-md transition-colors"
                >
                  <span>Mở Lưu trữ &amp; Đồng bộ (Redis)</span>
                </button>
              )}
              <button
                type="button"
                onClick={loadAppData}
                disabled={loadingData}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 min-h-[36px] bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-md transition-colors disabled:opacity-50"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${loadingData ? 'animate-spin' : ''}`} />
                <span>Thử lại</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Live Toast Notification */}
      {liveToast && (
        <div className="fixed bottom-3 left-3 right-3 sm:left-auto sm:bottom-6 sm:right-6 sm:max-w-sm z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-xl flex items-center gap-3 animate-slideUp border border-slate-700 text-xs">
          <Bell className="w-4 h-4 text-amber-400 shrink-0" />
          <span className="break-words min-w-0">{liveToast.message}</span>
        </div>
      )}

      {/* Main View Area */}
      <main className="flex-1">
        {noBookPermission ? (
          <div className="max-w-2xl mx-auto my-12 px-4">
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 text-center space-y-3">
              <AlertTriangle className="w-8 h-8 text-amber-600 mx-auto" />
              <h2 className="text-base font-bold text-amber-950">
                Tài khoản chưa được phân quyền trên sổ văn bản nào
              </h2>
              <p className="text-xs text-amber-900">
                Quyền truy cập của bạn có thể vừa được điều chỉnh bởi Quản trị viên. Vui lòng liên hệ Quản trị viên hoặc bấm Làm mới dữ liệu để kiểm tra lại.
              </p>
              <button
                type="button"
                onClick={loadAppData}
                className="inline-flex items-center gap-1.5 px-4 py-2 bg-amber-700 hover:bg-amber-800 text-white text-xs font-semibold rounded-lg"
              >
                <RefreshCw className={`w-3.5 h-3.5 ${loadingData ? 'animate-spin' : ''}`} />
                <span>Làm mới dữ liệu</span>
              </button>
            </div>
          </div>
        ) : (
          <>
            {activeTab === 'issue' && (
              <IssueNumberView
                currentUser={currentUser}
                activeBook={activeBook}
                accessibleBooks={accessibleBooks.length > 0 ? accessibleBooks : [activeBook]}
                onSelectBook={setActiveBookId}
                booksLoaded={settingsLoaded}
                appSyncError={syncError}
                onRetryAppLoad={loadAppData}
                settings={settings}
                documentTypes={documentTypes}
                signers={signers}
                documents={documents}
                numberRequests={numberRequests}
                serverYear={serverYear}
                serverFormattedDate={serverFormattedDate}
                onServerTimeUpdate={(year, formattedDate) => applyServerTime({ year, formattedDate })}
                onRefreshRequests={refreshRequests}
                onDocumentIssued={handleDocumentIssued}
                onDocumentUpdated={handleDocumentUpdated}
                onNavigateToRegistry={() => {
                  setRegistrySection('active');
                  setActiveTab('registry');
                }}
                onOpenExistingDocument={handleOpenExistingDocument}
                onDocumentTypeAdded={(newType, updatedBook) => {
                  setDocumentTypes((prev) => {
                    if (prev.some((t) => t.code === newType.code)) {
                      return prev.map((t) => (t.code === newType.code ? newType : t));
                    }
                    return [...prev, newType];
                  });
                  if (updatedBook) {
                    setBooks((prev) =>
                      prev.map((b) => (b.id === updatedBook.id ? updatedBook : b))
                    );
                  }
                  showToast(
                    updatedBook && documentTypes.some((t) => t.code === newType.code)
                      ? `Đã bật loại văn bản ${newType.name} (${newType.code}) cho ${updatedBook.name}`
                      : `Đã thêm loại văn bản mới: ${newType.name} (${newType.code})`,
                    'success'
                  );
                }}
              />
            )}

            {activeTab === 'import_pdf' && (
              <ImportPdfView
                currentUser={currentUser}
                activeBook={activeBook}
                accessibleBooks={accessibleBooks.length > 0 ? accessibleBooks : [activeBook]}
                onSelectBook={setActiveBookId}
                settings={settings}
                serverYear={serverYear}
                onDocumentImported={handleDocumentIssued}
                onDocumentUpdated={handleDocumentUpdated}
                onOpenExistingDocument={handleOpenExistingDocument}
                onNavigateToIncomingRegistry={handleNavigateToIncomingRegistry}
              />
            )}

            {activeTab === 'registry' && (
              <DocumentRegistryView
                documents={documents}
                deletedDocuments={deletedDocuments}
                documentTypes={documentTypes}
                books={books.length > 0 ? books : [activeBook]}
                activeBookId={effectiveRegistryBookId}
                onSelectBookId={(val) => {
                  if (val === 'all') {
                    setRegistryScopeAll(true);
                  } else {
                    setRegistryScopeAll(false);
                    setActiveBookId(val);
                  }
                }}
                settings={settings}
                serverYear={serverYear}
                currentUser={currentUser}
                section={registrySection}
                onSectionChange={setRegistrySection}
                onRefresh={loadAppData}
                onDocumentUpdated={handleDocumentUpdated}
                initialOpenDocId={pendingOpenDocId}
                onClearInitialOpenDocId={() => setPendingOpenDocId(null)}
                presetCategory={registryPresetCategory}
                onClearPresetCategory={() => setRegistryPresetCategory(null)}
              />
            )}

            {activeTab === 'tasks' && (
              <TasksView
                currentUser={currentUser}
                accessibleBooks={accessibleBooks.length > 0 ? accessibleBooks : [activeBook]}
                activeBookId={activeBook.id}
                onOpenDocument={handleOpenExistingDocument}
              />
            )}

            {activeTab === 'reconciliation' && (
              <ReconciliationView
                currentUser={currentUser}
                books={books.length > 0 ? books : [activeBook]}
                activeBookId={activeBook.id}
                serverYear={serverYear}
                onOpenDocument={handleOpenExistingDocument}
                onRefreshApp={loadAppData}
              />
            )}

            {activeTab === 'statistics' && (
              <StatisticsView
                currentYear={currentYear}
                books={books.length > 0 ? books : [activeBook]}
                activeBookId={activeBook.id}
                onSelectBookId={setActiveBookId}
                currentUser={currentUser}
              />
            )}

            {activeTab === 'settings' && (
              <SettingsView
                currentUser={currentUser}
                books={books.length > 0 ? books : [activeBook]}
                activeBookId={activeBook.id}
                onSelectBookId={setActiveBookId}
                settings={settings}
                documentTypes={documentTypes}
                signers={signers}
                serverYear={serverYear}
                configLoaded={settingsLoaded}
                syncError={syncError}
                persistenceStatus={persistenceStatus}
                legacyReconciliationWarnings={legacyReconciliationWarnings}
                requestedSubTab={requestedSettingsTab}
                onClearRequestedSubTab={() => setRequestedSettingsTab(null)}
                onPersistenceStatusChange={(status) => setPersistenceStatus(status)}
                onSettingsUpdated={(newSettings) => setSettings(newSettings)}
                onBooksUpdated={refreshBooks}
                onRefreshData={loadAppData}
                onCurrentUserUpdated={(updatedUser) => {
                  updateStoredUser(updatedUser);
                  setCurrentUser(updatedUser);
                }}
              />
            )}
          </>
        )}
      </main>

      {/* Minimalist Human Footer */}
      <footer className="border-t border-slate-200 bg-white py-4 px-4 text-center text-xs text-slate-500">
        <span>© {currentYear} Đoàn TNCS Hồ Chí Minh - Đoàn Trường THPT Nguyễn Chí Thanh. Bản quyền thuộc về tổ chức Đoàn.</span>
      </footer>
    </div>
  );
}
