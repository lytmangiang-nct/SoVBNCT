import React, { useEffect, useRef, useState } from 'react';
import {
  FileText,
  FileUp,
  BookOpen,
  BarChart3,
  Settings,
  LogOut,
  ShieldCheck,
  UserCheck,
  ChevronDown,
  History,
  Trash2,
  MoreHorizontal,
  UserCircle2,
  ListChecks,
  ClipboardCheck,
} from 'lucide-react';
import type { BookPermission, RegistryBook, User } from '../types';
import {
  getUserDisplayName,
  getUserSystemRoleLabel,
  hasAnyBookPermission,
  hasBookPermission,
  isSystemAdmin,
} from '../utils/permissions';

export type AppTab =
  | 'issue'
  | 'import_pdf'
  | 'registry'
  | 'tasks'
  | 'reconciliation'
  | 'statistics'
  | 'settings';
export type RegistrySection = 'active' | 'deleted' | 'history';

interface NavbarProps {
  activeTab: AppTab;
  setActiveTab: (tab: AppTab) => void;
  registrySection: RegistrySection;
  onSelectRegistrySection: (section: RegistrySection) => void;
  deletedCount?: number;
  pendingExcelSyncCount?: number;
  currentUser: User | null;
  books: RegistryBook[];
  activeBook?: RegistryBook;
  onLogout: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  setActiveTab,
  registrySection,
  onSelectRegistrySection,
  deletedCount = 0,
  pendingExcelSyncCount = 0,
  currentUser,
  books,
  activeBook,
  onLogout,
}) => {
  const [moreMenuOpen, setMoreMenuOpen] = useState(false);
  const [accountMenuOpen, setAccountMenuOpen] = useState(false);

  const moreRef = useRef<HTMLDivElement | null>(null);
  const accountRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      const target = event.target as Node;
      if (moreRef.current && !moreRef.current.contains(target)) {
        setMoreMenuOpen(false);
      }
      if (accountRef.current && !accountRef.current.contains(target)) {
        setAccountMenuOpen(false);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setMoreMenuOpen(false);
        setAccountMenuOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handleOutsideClick);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  const isAdmin = isSystemAdmin(currentUser);

  const hasPermForActiveBook = (perm: BookPermission): boolean => {
    if (isAdmin) return true;
    if (activeBook?.id) {
      return hasBookPermission(currentUser, activeBook.id, perm);
    }
    return hasAnyBookPermission(currentUser, perm, books);
  };

  const canIssueDirect = hasPermForActiveBook('issue_number');
  const canRequestNumber = hasPermForActiveBook('request_number');
  const canIssueOrRequest = canIssueDirect || canRequestNumber;
  const canImportPdf = hasPermForActiveBook('import_incoming');
  const canViewRegistry = hasPermForActiveBook('view');
  const canViewDeleted = canViewRegistry || hasPermForActiveBook('delete');
  const canManageSettings = hasPermForActiveBook('manage_book');
  const canAccessReconciliation = isAdmin || hasAnyBookPermission(currentUser, 'manage_book', books);

  const displayName = getUserDisplayName(currentUser);
  const systemRoleLabel = getUserSystemRoleLabel(currentUser);

  type PrimaryNavItem = {
    key: 'issue' | 'import_pdf' | 'registry';
    label: string;
    visible: boolean;
    active: boolean;
    icon: React.ReactNode;
    onClick: () => void;
  };

  const allPrimaryItems: PrimaryNavItem[] = [
    {
      key: 'issue',
      label: canIssueDirect ? 'Cấp số' : 'Xin số',
      visible: canIssueOrRequest,
      active: activeTab === 'issue',
      icon: <FileText className="w-4 h-4 shrink-0" />,
      onClick: () => {
        setActiveTab('issue');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'import_pdf',
      label: 'Nhập PDF',
      visible: canImportPdf,
      active: activeTab === 'import_pdf',
      icon: <FileUp className="w-4 h-4 shrink-0" />,
      onClick: () => {
        setActiveTab('import_pdf');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'registry',
      label: 'Sổ theo dõi',
      visible: canViewRegistry,
      active: activeTab === 'registry' && registrySection === 'active',
      icon: <BookOpen className="w-4 h-4 shrink-0" />,
      onClick: () => {
        onSelectRegistrySection('active');
        setActiveTab('registry');
        setMoreMenuOpen(false);
      },
    },
  ];
  const primaryItems = allPrimaryItems.filter((item) => item.visible);

  type MoreNavItem = {
    key: string;
    label: string;
    visible: boolean;
    active: boolean;
    badge?: number;
    alertDot?: boolean;
    icon: React.ReactNode;
    onClick: () => void;
  };

  const allMoreItems: MoreNavItem[] = [
    {
      key: 'tasks',
      label: 'Việc của tôi',
      visible: Boolean(currentUser),
      active: activeTab === 'tasks',
      icon: <ListChecks className="w-4 h-4 text-sky-700 shrink-0" />,
      onClick: () => {
        setActiveTab('tasks');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'statistics',
      label: 'Thống kê',
      visible: canViewRegistry,
      active: activeTab === 'statistics',
      icon: <BarChart3 className="w-4 h-4 text-sky-700 shrink-0" />,
      onClick: () => {
        setActiveTab('statistics');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'history',
      label: 'Lịch sử thao tác',
      visible: canViewRegistry,
      active: activeTab === 'registry' && registrySection === 'history',
      icon: <History className="w-4 h-4 text-sky-700 shrink-0" />,
      onClick: () => {
        onSelectRegistrySection('history');
        setActiveTab('registry');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'deleted',
      label: 'VB đã xóa',
      visible: canViewDeleted,
      active: activeTab === 'registry' && registrySection === 'deleted',
      badge: deletedCount,
      alertDot: pendingExcelSyncCount > 0,
      icon: <Trash2 className="w-4 h-4 text-rose-700 shrink-0" />,
      onClick: () => {
        onSelectRegistrySection('deleted');
        setActiveTab('registry');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'reconciliation',
      label: 'Đối soát dữ liệu',
      visible: canAccessReconciliation,
      active: activeTab === 'reconciliation',
      icon: <ClipboardCheck className="w-4 h-4 text-sky-700 shrink-0" />,
      onClick: () => {
        setActiveTab('reconciliation');
        setMoreMenuOpen(false);
      },
    },
    {
      key: 'settings',
      label: 'Cài đặt',
      visible: canManageSettings,
      active: activeTab === 'settings',
      icon: <Settings className="w-4 h-4 text-slate-700 shrink-0" />,
      onClick: () => {
        setActiveTab('settings');
        setMoreMenuOpen(false);
      },
    },
  ];
  const moreItems = allMoreItems.filter((item) => item.visible);

  const isMoreActive = moreItems.some((item) => item.active);

  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200 shadow-xs">
      {/* Row 1: Flexbox layout — Logo + tên ngắn | Menu chính (md+) | Thêm + Tài khoản */}
      <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 h-14 sm:h-16 flex items-center justify-between gap-2 sm:gap-4">
        {/* Brand: Logo + Tên ngắn "Sổ văn bản" */}
        <div className="flex items-center gap-2.5 min-w-0 flex-1 md:flex-initial">
          <div className="w-9 h-9 rounded-lg bg-sky-700 flex items-center justify-center text-white font-bold shadow-xs shrink-0">
            <svg className="w-5 h-5 text-white" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
              <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z" />
            </svg>
          </div>
          <span className="text-sm sm:text-base font-bold tracking-tight text-sky-950 truncate min-w-0">
            Sổ văn bản
          </span>
        </div>

        {/* Center Navigation (Desktop & Tablet >= md): Cấp số, Nhập PDF, Sổ theo dõi */}
        <nav
          aria-label="Điều hướng chính"
          className="hidden md:flex items-center justify-center gap-1.5 min-w-0"
        >
          {primaryItems.map((item) => (
            <button
              key={item.key}
              type="button"
              onClick={item.onClick}
              className={`inline-flex items-center justify-center gap-2 h-9 px-3.5 text-sm font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 ${
                item.active
                  ? 'bg-sky-50 text-sky-900 border-sky-200'
                  : 'bg-transparent text-slate-600 border-transparent hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              {item.icon}
              <span>{item.label}</span>
            </button>
          ))}
        </nav>

        {/* Right Controls: Nút Thêm + Biểu tượng Tài khoản */}
        <div className="flex items-center gap-2 shrink-0">
          {moreItems.length > 0 && (
            <div ref={moreRef} className="relative">
              <button
                type="button"
                onClick={() => {
                  setMoreMenuOpen((prev) => !prev);
                  setAccountMenuOpen(false);
                }}
                aria-haspopup="menu"
                aria-expanded={moreMenuOpen}
                className={`inline-flex items-center justify-center gap-1.5 h-9 px-3 text-xs sm:text-sm font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 ${
                  isMoreActive || moreMenuOpen
                    ? 'bg-sky-50 text-sky-900 border-sky-200'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                <MoreHorizontal className="w-4 h-4 shrink-0 md:hidden" />
                <span>Thêm</span>
                {pendingExcelSyncCount > 0 && (
                  <span
                    className="w-2 h-2 rounded-full bg-amber-500 shrink-0"
                    title="Có văn bản đã xóa chờ đồng bộ Google Trang tính"
                  />
                )}
                <ChevronDown
                  className={`w-3.5 h-3.5 shrink-0 transition-transform ${
                    moreMenuOpen ? 'rotate-180' : ''
                  }`}
                />
              </button>

              {moreMenuOpen && (
                <div
                  role="menu"
                  aria-label="Menu Thêm"
                  className="absolute right-0 mt-2 w-60 max-w-[calc(100vw-1.5rem)] bg-white border border-slate-200 rounded-xl shadow-xl py-1.5 z-50"
                >
                  {moreItems.map((item) => (
                    <button
                      key={item.key}
                      type="button"
                      role="menuitem"
                      onClick={item.onClick}
                      className={`w-full flex items-center justify-between gap-2 px-3.5 py-2.5 text-left text-xs sm:text-sm font-semibold transition-colors ${
                        item.active
                          ? 'bg-sky-50 text-sky-900'
                          : 'text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      <span className="flex items-center gap-2.5 min-w-0">
                        {item.icon}
                        <span className="truncate">{item.label}</span>
                      </span>
                      {(item.badge !== undefined && item.badge > 0) || item.alertDot ? (
                        <span className="inline-flex items-center gap-1 shrink-0">
                          {item.alertDot && <span className="w-2 h-2 rounded-full bg-amber-500" />}
                          {item.badge !== undefined && item.badge > 0 && (
                            <span className="font-mono text-[11px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 tabular-nums">
                              {item.badge}
                            </span>
                          )}
                        </span>
                      ) : null}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Account Trigger Button & Menu */}
          {currentUser && (
            <div ref={accountRef} className="relative">
              <button
                type="button"
                onClick={() => {
                  setAccountMenuOpen((prev) => !prev);
                  setMoreMenuOpen(false);
                }}
                aria-haspopup="menu"
                aria-expanded={accountMenuOpen}
                aria-label={`Tài khoản: ${displayName}`}
                title={`${displayName} (@${currentUser.username} · ${systemRoleLabel})`}
                className={`inline-flex items-center justify-center w-9 h-9 md:w-auto md:px-2.5 md:gap-2 rounded-full border transition-colors shrink-0 min-w-0 max-w-full ${
                  accountMenuOpen
                    ? 'bg-sky-100 text-sky-900 border-sky-300'
                    : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200/70'
                }`}
              >
                {isAdmin ? (
                  <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0" />
                ) : (
                  <UserCircle2 className="w-5 h-5 text-sky-700 shrink-0" />
                )}
                <span className="hidden md:block truncate max-w-[140px] lg:max-w-[190px] text-xs font-semibold text-slate-800">
                  {displayName}
                </span>
              </button>

              {accountMenuOpen && (
                <div
                  role="menu"
                  aria-label="Menu Tài khoản"
                  className="absolute right-0 mt-2 w-64 max-w-[calc(100vw-1.5rem)] bg-white border border-slate-200 rounded-xl shadow-xl p-3 space-y-3 z-50"
                >
                  <div className="flex items-start gap-2.5 pb-2.5 border-b border-slate-100 min-w-0">
                    <div className="w-9 h-9 rounded-full bg-sky-50 border border-sky-200 flex items-center justify-center shrink-0">
                      {isAdmin ? (
                        <ShieldCheck className="w-4 h-4 text-amber-600" />
                      ) : (
                        <UserCheck className="w-4 h-4 text-sky-700" />
                      )}
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className="text-sm font-bold text-slate-900 break-words leading-snug">
                        {displayName}
                      </p>
                      <p className="text-xs text-slate-500 mt-0.5 break-words">
                        @{currentUser.username} · {systemRoleLabel}
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      setAccountMenuOpen(false);
                      onLogout();
                    }}
                    className="w-full flex items-center justify-center gap-2 h-9 px-3 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors"
                  >
                    <LogOut className="w-4 h-4 shrink-0" />
                    <span>Đăng xuất</span>
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Row 2 on Mobile (< md): Primary 3 tabs cleanly in CSS Grid without overlap */}
      {primaryItems.length > 0 && (
        <nav
          aria-label="Điều hướng chính trên điện thoại"
          className="md:hidden border-t border-slate-100 bg-slate-50/80 px-3 py-1.5 grid gap-1.5"
          style={{ gridTemplateColumns: `repeat(${primaryItems.length}, minmax(0, 1fr))` }}
        >
          {primaryItems.map((item) => (
            <button
              key={item.key}
              type="button"
              onClick={item.onClick}
              className={`flex items-center justify-center gap-1.5 h-9 px-2 text-xs font-semibold rounded-lg border transition-colors min-w-0 ${
                item.active
                  ? 'bg-white text-sky-900 border-sky-200 shadow-2xs'
                  : 'bg-transparent text-slate-600 border-transparent hover:bg-white/60'
              }`}
            >
              {item.icon}
              <span className="truncate">{item.label}</span>
            </button>
          ))}
        </nav>
      )}
    </header>
  );
};
