import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { Navbar, type AppTab } from './components/Navbar';
import { IssueNumberView } from './components/IssueNumberView';
import { ImportPdfView } from './components/ImportPdfView';
import { DocumentRegistryView } from './components/DocumentRegistryView';
import { StatisticsView } from './components/StatisticsView';
import { SettingsView } from './components/SettingsView';
import { LoginModal } from './components/LoginModal';
import { api, buildApiUrl, clearAuth } from './api';
import type {
  DocumentRecord,
  DocumentTypeConfig,
  NumberRequest,
  RegistryBook,
  SignerConfig,
  SystemSettings,
  User,
} from './types';
import {
  DEFAULT_BOOK_ID,
  getAccessibleBooks,
  getBookEnabledTypes,
  getDocBookId,
  hasAnyBookPermission,
  hasBookPermission,
  isSystemAdmin,
} from './utils/permissions';
import { AlertTriangle, Bell, BookOpen, Clock, FileText, RefreshCw } from 'lucide-react';

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
    msGraphConfig: {
      enabled: false,
      tenantId: '',
      clientId: '',
      clientSecret: '',
      driveId: '',
      filePath: '',
      tableName: 'SoCapSoVanBan',
    },
  });
  const [books, setBooks] = useState<RegistryBook[]>([]);
  const [activeBookId, setActiveBookId] = useState<string>(DEFAULT_BOOK_ID);
  const [numberRequests, setNumberRequests] = useState<NumberRequest[]>([]);
  const [documentTypes, setDocumentTypes] = useState<DocumentTypeConfig[]>([]);
  const [signers, setSigners] = useState<SignerConfig[]>([]);
  const [documents, setDocuments] = useState<DocumentRecord[]>([]);
  const [loadingData, setLoadingData] = useState<boolean>(false);
  const [settingsLoaded, setSettingsLoaded] = useState<boolean>(false);
  const [lastSyncAt, setLastSyncAt] = useState<Date | null>(null);
  const [syncError, setSyncError] = useState<string | null>(null);
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
      setBooks([]);
      setDocuments([]);
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
      const [meRes, settingsRes, docsRes, booksRes, reqsRes] = await Promise.all([
        api.getCurrentUser(),
        api.getSettings(),
        api.getDocuments({}),
        api.getBooks(),
        api.getNumberRequests().catch(() => ({ requests: [] as NumberRequest[] })),
      ]);

      if (!meRes.user.isActive) {
        clearAuth();
        setCurrentUser(null);
        setSettingsLoaded(false);
        return;
      }

      setCurrentUser((prev) => {
        if (!prev) return meRes.user;
        if (
          JSON.stringify(prev.bookPermissions || {}) !== JSON.stringify(meRes.user.bookPermissions || {}) ||
          prev.role !== meRes.user.role ||
          prev.isActive !== meRes.user.isActive ||
          prev.fullName !== meRes.user.fullName
        ) {
          return meRes.user;
        }
        return prev;
      });

      setSettings(settingsRes.settings);
      setDocumentTypes(settingsRes.documentTypes);
      setSigners(settingsRes.signers);
      setDocuments(docsRes.documents);
      setBooks(booksRes.books);
      setNumberRequests(reqsRes.requests);
      applyServerTime(settingsRes.serverTimeVietnam || booksRes.serverTimeVietnam);

      setSettingsLoaded(true);
      setSyncError(null);
      setLastSyncAt(new Date());
    } catch (err: any) {
      console.error('Lỗi khi tải dữ liệu ứng dụng:', err);
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
          setDocuments((prev) =>
            prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
          );
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
            await loadAppData();
            showToast('Quyền truy cập sổ văn bản của bạn vừa được cập nhật.', 'info');
          }
        } catch {
          // ignore
        }
      });

      eventSource.addEventListener('number_requests_updated', () => {
        refreshRequests();
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
      window.removeEventListener('focus', handleVisibilityOrFocus);
      window.removeEventListener('online', handleOnline);
      document.removeEventListener('visibilitychange', handleVisibilityOrFocus);
    };
  }, [currentUser?.id, loadAppData, refreshBooks, refreshRequests, applyServerTime]);

  const accessibleBooks = useMemo(
    () => getAccessibleBooks(currentUser, books),
    [currentUser, books]
  );

  // Sync activeBookId to a book authorized for the current activeTab
  useEffect(() => {
    if (accessibleBooks.length === 0 || !currentUser) return;
    const bookSupportsTab = (b: RegistryBook): boolean => {
      if (isSystemAdmin(currentUser)) return true;
      if (activeTab === 'issue') {
        return hasBookPermission(currentUser, b.id, 'issue_number') || hasBookPermission(currentUser, b.id, 'request_number');
      }
      if (activeTab === 'import_pdf') {
        return hasBookPermission(currentUser, b.id, 'import_incoming');
      }
      if (activeTab === 'registry' || activeTab === 'statistics') {
        return hasBookPermission(currentUser, b.id, 'view');
      }
      if (activeTab === 'settings') {
        return hasBookPermission(currentUser, b.id, 'manage_book');
      }
      return true;
    };

    const currentBook = accessibleBooks.find((b) => b.id === activeBookId);
    if (!currentBook || !bookSupportsTab(currentBook)) {
      const preferred = accessibleBooks.find(bookSupportsTab) || accessibleBooks[0];
      if (preferred && preferred.id !== activeBookId) {
        setActiveBookId(preferred.id);
      }
    }
  }, [accessibleBooks, activeBookId, activeTab, currentUser]);

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

  // Enforce permission-based tab fallback (chỉ mở các tab người dùng có quyền)
  useEffect(() => {
    if (!currentUser || loadingData) return;
    const isAdmin = isSystemAdmin(currentUser);
    const allowedTabs: AppTab[] = [];
    if (
      isAdmin ||
      hasAnyBookPermission(currentUser, 'issue_number', books) ||
      hasAnyBookPermission(currentUser, 'request_number', books)
    ) {
      allowedTabs.push('issue');
    }
    if (isAdmin || hasAnyBookPermission(currentUser, 'import_incoming', books)) {
      allowedTabs.push('import_pdf');
    }
    if (isAdmin || hasAnyBookPermission(currentUser, 'view', books)) {
      allowedTabs.push('registry');
      allowedTabs.push('statistics');
    }
    if (isAdmin || hasAnyBookPermission(currentUser, 'manage_book', books)) {
      allowedTabs.push('settings');
    }

    if (allowedTabs.length > 0 && !allowedTabs.includes(activeTab)) {
      setActiveTab(allowedTabs[0]);
    }
  }, [currentUser, books, activeTab, loadingData]);

  // Handlers
  const handleLogout = async () => {
    await api.logout();
    setCurrentUser(null);
    setSettingsLoaded(false);
    setLastSyncAt(null);
    setSyncError(null);
    setBooks([]);
    setDocuments([]);
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
    setDocuments((prev) =>
      prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
    );
    setLastSyncAt(new Date());
  };

  const handleOpenExistingDocument = (docId: string) => {
    setPendingOpenDocId(docId);
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

  // Quick stats summary for server Vietnam year on active book
  const currentYear = serverYear;
  const currentYearDocs = documents.filter(
    (d) => d.year === currentYear && getDocBookId(d) === activeBook.id
  );
  const totalCount = currentYearDocs.length;
  const incomingCount = currentYearDocs.filter((d) => d.docCategory === 'incoming').length;
  const noBookPermission = settingsLoaded && accessibleBooks.length === 0;

  return (
    <div className="min-h-[100dvh] bg-slate-50 text-slate-900 flex flex-col font-sans overflow-x-clip">
      {/* Top Navigation Bar */}
      <Navbar
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        currentUser={currentUser}
        books={books}
        accessibleBooks={accessibleBooks}
        activeBook={activeBook}
        onSelectBook={setActiveBookId}
        onLogout={handleLogout}
        lastSyncAt={lastSyncAt}
        syncingData={loadingData}
        onRefreshData={loadAppData}
      />

      {/* Global Dashboard Summary Banner */}
      <section className="bg-sky-900 text-white border-b border-sky-950 py-3 px-3 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 min-w-0">
            <div className="flex items-start sm:items-center gap-2.5 min-w-0">
              <div className="w-8 h-8 sm:w-9 sm:h-9 rounded-lg bg-sky-800/80 flex items-center justify-center shrink-0 border border-sky-700/60 mt-0.5 sm:mt-0">
                <FileText className="w-4 h-4 sm:w-5 sm:h-5 text-sky-200" />
              </div>
              <div className="min-w-0 space-y-0.5">
                <div className="flex items-center gap-1.5 sm:gap-2 flex-wrap">
                  <span className="text-xs uppercase font-bold tracking-wider text-sky-200">
                    Năm {currentYear}
                  </span>
                  <span className="text-[11px] text-sky-300">·</span>
                  <span className="text-xs text-sky-200">
                    Tổng: <strong className="font-mono text-white font-bold">{totalCount}</strong> văn bản
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[11px] text-sky-300">
                  <span className="truncate max-w-full">
                    {activeBook.unitName} ({activeBook.unitCode})
                  </span>
                  {lastSyncAt && (
                    <>
                      <span className="hidden sm:inline">·</span>
                      <span className="inline-flex items-center gap-1 text-sky-200 font-mono">
                        <Clock className="w-3 h-3 shrink-0" />
                        Cập nhật lần cuối:{' '}
                        {lastSyncAt.toLocaleTimeString('vi-VN', {
                          hour: '2-digit',
                          minute: '2-digit',
                          second: '2-digit',
                        })}
                      </span>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* Always-visible Book Selector on all screens */}
            {accessibleBooks.length > 0 && (
              <div className="flex items-center gap-2 bg-sky-800/90 border border-sky-600/80 rounded-lg px-2.5 py-1.5 min-h-[40px] sm:min-h-[36px] shrink-0 w-full sm:w-auto min-w-0">
                <BookOpen className="w-3.5 h-3.5 text-sky-200 shrink-0" />
                <span className="text-[11px] text-sky-200 font-medium whitespace-nowrap">Sổ:</span>
                {accessibleBooks.length > 1 ? (
                  <select
                    aria-label="Sổ đang chọn"
                    value={activeBook.id}
                    onChange={(e) => setActiveBookId(e.target.value)}
                    className="flex-1 sm:flex-initial min-w-0 bg-transparent text-white text-xs sm:text-sm font-bold focus:outline-hidden cursor-pointer [&>option]:text-slate-900"
                  >
                    {accessibleBooks.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name} ({b.unitCode})
                      </option>
                    ))}
                  </select>
                ) : (
                  <span className="text-xs font-bold text-white truncate">
                    {activeBook.name} ({activeBook.unitCode})
                  </span>
                )}
              </div>
            )}
          </div>

          {/* Quick counts for book-enabled types + Văn bản đến */}
          <div className="flex items-center gap-1.5 sm:gap-2 flex-wrap text-[11px] sm:text-xs">
            {getBookEnabledTypes(activeBook, documentTypes).map((t) => {
              const count = currentYearDocs.filter(
                (d) => d.docCategory !== 'incoming' && d.typeCode === t.code
              ).length;
              return (
                <div
                  key={t.code}
                  className="bg-sky-800/80 border border-sky-700/60 px-2 sm:px-2.5 py-1 rounded-md flex items-center gap-1.5"
                >
                  <span className="text-sky-200 font-medium">{t.name}:</span>
                  <span className="font-mono font-bold text-white">{count}</span>
                </div>
              );
            })}
            <div className="bg-indigo-800/80 border border-indigo-600/70 px-2 sm:px-2.5 py-1 rounded-md flex items-center gap-1.5">
              <span className="text-indigo-200 font-medium">Văn bản đến:</span>
              <span className="font-mono font-bold text-white">{incomingCount}</span>
            </div>
          </div>
        </div>
      </section>

      {/* Sync Failure Banner */}
      {syncError && (
        <div className="bg-rose-50 border-b border-rose-200 px-3 sm:px-6 lg:px-8 py-2.5">
          <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-2 text-xs text-rose-900">
            <div className="flex items-center gap-2 min-w-0">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
              <span className="break-words min-w-0">
                <strong>Chưa cập nhật được dữ liệu</strong> — {syncError}
              </span>
            </div>
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
                onNavigateToRegistry={() => setActiveTab('registry')}
                onOpenExistingDocument={handleOpenExistingDocument}
                onDocumentTypeAdded={(newType) => {
                  setDocumentTypes((prev) => {
                    if (prev.some((t) => t.code === newType.code)) return prev;
                    return [...prev, newType];
                  });
                  showToast(`Đã thêm loại văn bản mới: ${newType.name} (${newType.code})`, 'success');
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
              />
            )}

            {activeTab === 'registry' && (
              <DocumentRegistryView
                documents={documents}
                documentTypes={documentTypes}
                books={books.length > 0 ? books : [activeBook]}
                activeBookId={activeBook.id}
                onSelectBookId={setActiveBookId}
                settings={settings}
                serverYear={serverYear}
                currentUser={currentUser}
                onRefresh={loadAppData}
                onDocumentUpdated={handleDocumentUpdated}
                initialOpenDocId={pendingOpenDocId}
                onClearInitialOpenDocId={() => setPendingOpenDocId(null)}
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
                onSettingsUpdated={(newSettings) => setSettings(newSettings)}
                onBooksUpdated={refreshBooks}
                onRefreshData={loadAppData}
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
