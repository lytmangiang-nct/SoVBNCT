import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { Navbar, type AppTab } from './components/Navbar';
import { IssueNumberView } from './components/IssueNumberView';
import { ImportPdfView } from './components/ImportPdfView';
import { DocumentRegistryView } from './components/DocumentRegistryView';
import { StatisticsView } from './components/StatisticsView';
import { SettingsView } from './components/SettingsView';
import { LoginModal } from './components/LoginModal';
import { api, getStoredUser, clearAuth } from './api';
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
import { Bell, BookOpen, FileText } from 'lucide-react';

export default function App() {
  const [currentUser, setCurrentUser] = useState<User | null>(getStoredUser);
  const [authChecking, setAuthChecking] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<AppTab>('issue');
  const [pendingOpenDocId, setPendingOpenDocId] = useState<string | null>(null);

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

  // Live Toast Notification
  const [liveToast, setLiveToast] = useState<{ message: string; type?: 'info' | 'success' } | null>(null);

  const showToast = (message: string, type: 'info' | 'success' = 'info') => {
    setLiveToast({ message, type });
    setTimeout(() => setLiveToast(null), 4000);
  };

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
    };
    window.addEventListener('auth_expired', handleExpired);
    return () => window.removeEventListener('auth_expired', handleExpired);
  }, []);

  const refreshBooks = useCallback(async () => {
    if (!currentUser) return;
    try {
      const res = await api.getBooks();
      setBooks(res.books);
    } catch (err) {
      console.error('Lỗi tải danh sách sổ:', err);
    }
  }, [currentUser]);

  const refreshRequests = useCallback(async () => {
    if (!currentUser) return;
    try {
      const res = await api.getNumberRequests();
      setNumberRequests(res.requests);
    } catch {
      // ignore if user has no access
    }
  }, [currentUser]);

  // 2. Load Core Application Data
  const loadAppData = useCallback(async () => {
    if (!currentUser) return;
    setLoadingData(true);
    try {
      const [settingsRes, docsRes, booksRes, reqsRes] = await Promise.all([
        api.getSettings(),
        api.getDocuments({}),
        api.getBooks().catch(() => ({ books: [] as RegistryBook[] })),
        api.getNumberRequests().catch(() => ({ requests: [] as NumberRequest[] })),
      ]);

      setSettings(settingsRes.settings);
      setDocumentTypes(settingsRes.documentTypes);
      setSigners(settingsRes.signers);
      setDocuments(docsRes.documents);
      if (booksRes.books.length > 0) {
        setBooks(booksRes.books);
      }
      setNumberRequests(reqsRes.requests);
    } catch (err: any) {
      console.error('Lỗi khi tải dữ liệu ứng dụng:', err);
    } finally {
      setLoadingData(false);
    }
  }, [currentUser]);

  useEffect(() => {
    if (currentUser) {
      loadAppData();
    }
  }, [currentUser, loadAppData]);

  // 3. Server-Sent Events (SSE) for Real-Time Multi-Device Updates
  useEffect(() => {
    if (!currentUser) return;

    let eventSource: EventSource | null = null;
    let reconnectTimeout: ReturnType<typeof setTimeout>;

    const connectSSE = () => {
      eventSource = new EventSource('/api/events');

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
          if (!data.targetUserId || data.targetUserId === currentUser.id) {
            try {
              const meRes = await api.getCurrentUser();
              if (!meRes.user.isActive) {
                clearAuth();
                setCurrentUser(null);
                return;
              }
              setCurrentUser(meRes.user);
              refreshBooks();
              showToast('Quyền truy cập sổ văn bản của bạn vừa được cập nhật.', 'info');
            } catch {
              clearAuth();
              setCurrentUser(null);
            }
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
              api.getSettings().then((res) => setSettings(res.settings)).catch(() => {});
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

    // Fallback polling mỗi 4 giây khi tab đang mở để đảm bảo đồng bộ giữa điện thoại & máy tính
    const syncLiveState = () => {
      if (document.visibilityState !== 'visible') return;
      loadAppData();
      api
        .getCurrentUser()
        .then((res) => {
          setCurrentUser((prev) => {
            if (!prev) return res.user;
            if (
              JSON.stringify(prev.bookPermissions || {}) !== JSON.stringify(res.user.bookPermissions || {}) ||
              prev.role !== res.user.role ||
              prev.isActive !== res.user.isActive
            ) {
              return res.user;
            }
            return prev;
          });
        })
        .catch(() => {});
    };

    const pollInterval = setInterval(syncLiveState, 4000);

    const handleVisibilityOrFocus = () => {
      syncLiveState();
    };
    window.addEventListener('focus', handleVisibilityOrFocus);
    document.addEventListener('visibilitychange', handleVisibilityOrFocus);

    return () => {
      if (eventSource) eventSource.close();
      clearTimeout(reconnectTimeout);
      clearInterval(pollInterval);
      window.removeEventListener('focus', handleVisibilityOrFocus);
      document.removeEventListener('visibilitychange', handleVisibilityOrFocus);
    };
  }, [currentUser?.id, loadAppData, refreshBooks, refreshRequests]);

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
  };

  const handleDocumentIssued = (newDoc: DocumentRecord) => {
    setDocuments((prev) => {
      if (prev.some((d) => d.id === newDoc.id)) return prev;
      return [newDoc, ...prev];
    });
  };

  const handleDocumentUpdated = (updatedDoc: DocumentRecord) => {
    setDocuments((prev) =>
      prev.map((d) => (d.id === updatedDoc.id ? updatedDoc : d))
    );
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

  // Quick stats summary for current year on active book
  const currentYear = new Date().getFullYear();
  const currentYearDocs = documents.filter(
    (d) => d.year === currentYear && getDocBookId(d) === activeBook.id
  );
  const totalCount = currentYearDocs.length;
  const incomingCount = currentYearDocs.filter((d) => d.docCategory === 'incoming').length;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans">
      {/* Top Navigation Bar */}
      <Navbar
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        currentUser={currentUser}
        books={books}
        onLogout={handleLogout}
      />

      {/* Global Dashboard Summary Banner */}
      <section className="bg-sky-900 text-white border-b border-sky-950 py-3.5 px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-sky-800/80 flex items-center justify-center shrink-0 border border-sky-700/60">
              <FileText className="w-5 h-5 text-sky-200" />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-xs uppercase font-bold tracking-wider text-sky-200">
                  Năm {currentYear}
                </span>
                <span className="text-[11px] text-sky-300">·</span>
                <span className="text-xs text-sky-200">
                  Tổng: <strong className="font-mono text-white font-bold">{totalCount}</strong> văn bản
                </span>
                {accessibleBooks.length > 1 && (
                  <>
                    <span className="text-[11px] text-sky-300">·</span>
                    <span className="inline-flex items-center gap-1 text-[11px] bg-sky-800 border border-sky-700 px-2 py-0.5 rounded text-sky-100 font-semibold">
                      <BookOpen className="w-3 h-3" />
                      {activeBook.name}
                    </span>
                  </>
                )}
              </div>
              <p className="text-[11px] text-sky-300">
                {activeBook.unitName} ({activeBook.unitCode})
              </p>
            </div>
          </div>

          {/* Quick counts for book-enabled types + Văn bản đến */}
          <div className="flex items-center gap-2 flex-wrap text-xs">
            {getBookEnabledTypes(activeBook, documentTypes).map((t) => {
              const count = currentYearDocs.filter(
                (d) => d.docCategory !== 'incoming' && d.typeCode === t.code
              ).length;
              return (
                <div
                  key={t.code}
                  className="bg-sky-800/80 border border-sky-700/60 px-2.5 py-1 rounded-md flex items-center gap-1.5"
                >
                  <span className="text-sky-300 font-medium">{t.name}:</span>
                  <span className="font-mono font-bold text-white">{count}</span>
                </div>
              );
            })}
            <div className="bg-indigo-800/80 border border-indigo-600/70 px-2.5 py-1 rounded-md flex items-center gap-1.5">
              <span className="text-indigo-200 font-medium">Văn bản đến:</span>
              <span className="font-mono font-bold text-white">{incomingCount}</span>
            </div>
          </div>
        </div>
      </section>

      {/* Live Toast Notification */}
      {liveToast && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-xl flex items-center gap-3 animate-slideUp border border-slate-700 text-xs">
          <Bell className="w-4 h-4 text-amber-400 shrink-0" />
          <span>{liveToast.message}</span>
        </div>
      )}

      {/* Main View Area */}
      <main className="flex-1">
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
            onSettingsUpdated={(newSettings) => setSettings(newSettings)}
            onBooksUpdated={refreshBooks}
            onRefreshData={loadAppData}
          />
        )}
      </main>

      {/* Minimalist Human Footer */}
      <footer className="border-t border-slate-200 bg-white py-4 px-4 text-center text-xs text-slate-500">
        <span>© {currentYear} Đoàn TNCS Hồ Chí Minh - Đoàn Trường THPT Nguyễn Chí Thanh. Bản quyền thuộc về tổ chức Đoàn.</span>
      </footer>
    </div>
  );
}
