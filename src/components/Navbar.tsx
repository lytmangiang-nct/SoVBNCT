import React, { useState } from 'react';
import {
  FileText,
  FileUp,
  BookOpen,
  BarChart3,
  Settings,
  LogOut,
  ShieldCheck,
  UserCheck,
  RefreshCw,
  Menu,
  X,
} from 'lucide-react';
import type { RegistryBook, User } from '../types';
import { hasAnyBookPermission, isSystemAdmin } from '../utils/permissions';

export type AppTab = 'issue' | 'import_pdf' | 'registry' | 'statistics' | 'settings';

interface NavbarProps {
  activeTab: AppTab;
  setActiveTab: (tab: AppTab) => void;
  currentUser: User | null;
  books: RegistryBook[];
  accessibleBooks?: RegistryBook[];
  activeBook?: RegistryBook;
  onSelectBook?: (bookId: string) => void;
  onLogout: () => void;
  lastSyncAt?: Date | null;
  syncingData?: boolean;
  onRefreshData?: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  setActiveTab,
  currentUser,
  books,
  accessibleBooks = [],
  activeBook,
  onSelectBook,
  onLogout,
  lastSyncAt,
  syncingData = false,
  onRefreshData,
}) => {
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const isAdmin = isSystemAdmin(currentUser);
  const canIssueDirect = isAdmin || hasAnyBookPermission(currentUser, 'issue_number', books);
  const canRequestNumber = isAdmin || hasAnyBookPermission(currentUser, 'request_number', books);
  const canIssueOrRequest = canIssueDirect || canRequestNumber;
  const canImportPdf = isAdmin || hasAnyBookPermission(currentUser, 'import_incoming', books);
  const canViewRegistry = isAdmin || hasAnyBookPermission(currentUser, 'view', books);
  const canViewStats = canViewRegistry;
  const canManageSettings = isAdmin || hasAnyBookPermission(currentUser, 'manage_book', books);

  const issueDesktopLabel = canIssueDirect ? 'Cấp số văn bản' : 'Xin số văn bản';
  const issueMobileLabel = canIssueDirect ? 'Cấp số' : 'Xin số';
  const roleLabel = isAdmin
    ? 'Quản trị viên'
    : canIssueDirect
    ? 'Cán bộ cấp số'
    : canRequestNumber
    ? 'Cán bộ xin số'
    : 'Cán bộ sổ văn bản';

  const formattedLastSync = lastSyncAt
    ? lastSyncAt.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
    : null;

  const allNavItems: Array<{
    key: AppTab;
    label: string;
    mobileLabel: string;
    visible: boolean;
    icon: React.ReactNode;
  }> = [
    {
      key: 'issue',
      label: issueDesktopLabel,
      mobileLabel: issueMobileLabel,
      visible: canIssueOrRequest,
      icon: <FileText className="w-4 h-4 text-sky-700 shrink-0" />,
    },
    {
      key: 'import_pdf',
      label: 'Nhập văn bản PDF',
      mobileLabel: 'Nhập PDF',
      visible: canImportPdf,
      icon: <FileUp className="w-4 h-4 text-sky-700 shrink-0" />,
    },
    {
      key: 'registry',
      label: 'Sổ theo dõi',
      mobileLabel: 'Sổ theo dõi',
      visible: canViewRegistry,
      icon: <BookOpen className="w-4 h-4 text-sky-700 shrink-0" />,
    },
    {
      key: 'statistics',
      label: 'Thống kê',
      mobileLabel: 'Thống kê',
      visible: canViewStats,
      icon: <BarChart3 className="w-4 h-4 text-sky-700 shrink-0" />,
    },
    {
      key: 'settings',
      label: 'Cài đặt',
      mobileLabel: 'Cài đặt',
      visible: canManageSettings,
      icon: <Settings className="w-4 h-4 text-sky-700 shrink-0" />,
    },
  ];
  const navItems = allNavItems.filter((item) => item.visible);

  const activeNavItem = navItems.find((item) => item.key === activeTab) || navItems[0];

  const handleSelectTab = (tab: AppTab) => {
    setActiveTab(tab);
    setMobileMenuOpen(false);
  };

  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200 shadow-xs">
      <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 h-14 sm:h-16 flex items-center justify-between gap-2 sm:gap-3">
        {/* Zone 1: Brand Wordmark */}
        <div className="flex items-center gap-2 sm:gap-3 min-w-0">
          <div className="w-8 h-8 sm:w-10 sm:h-10 rounded-lg bg-sky-700 flex items-center justify-center text-white font-bold shadow-xs shrink-0">
            <svg className="w-4 h-4 sm:w-6 sm:h-6 text-white" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z" />
            </svg>
          </div>
          <div className="flex flex-col min-w-0">
            <span className="text-xs sm:text-sm font-bold tracking-tight text-sky-900 truncate">
              THPT NGUYỄN CHÍ THANH
            </span>
            <span className="text-[10px] sm:text-[11px] text-slate-500 font-medium truncate">
              Sổ Cấp Số Văn Bản
            </span>
          </div>
        </div>

        {/* Zone 2: Navigation tabs on desktop (lg+) / medium tablet (md+) */}
        <nav className="hidden lg:flex items-center gap-1 min-w-0">
          {navItems.map((item) => (
            <button
              key={item.key}
              type="button"
              onClick={() => handleSelectTab(item.key)}
              className={`flex items-center gap-2 px-3 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === item.key
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              {item.icon}
              <span>{item.label}</span>
            </button>
          ))}
        </nav>

        {/* Zone 3: Sync, User Profile, Logout (desktop) & Compact Menu Trigger (mobile/tablet) */}
        <div className="flex items-center gap-1.5 sm:gap-2.5 shrink-0">
          {onRefreshData && (
            <div className="flex items-center gap-2">
              {formattedLastSync && (
                <span
                  className="hidden xl:inline text-[11px] text-slate-500 font-mono whitespace-nowrap"
                  title={`Cập nhật lần cuối lúc ${formattedLastSync}`}
                >
                  Cập nhật lần cuối: {formattedLastSync}
                </span>
              )}
              <button
                type="button"
                onClick={onRefreshData}
                disabled={syncingData}
                aria-label="Làm mới dữ liệu"
                title={formattedLastSync ? `Cập nhật lần cuối: ${formattedLastSync}` : 'Làm mới dữ liệu từ máy chủ'}
                className="inline-flex items-center justify-center gap-1.5 min-h-[40px] min-w-[40px] sm:min-h-[36px] px-2.5 py-1.5 text-xs font-semibold text-sky-800 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded-lg transition-colors whitespace-nowrap shrink-0 disabled:opacity-60"
              >
                <RefreshCw className={`w-4 h-4 sm:w-3.5 sm:h-3.5 ${syncingData ? 'animate-spin' : ''}`} />
                <span className="hidden sm:inline">{syncingData ? 'Đang tải...' : 'Làm mới dữ liệu'}</span>
              </button>
            </div>
          )}

          {currentUser && (
            <div className="hidden sm:flex items-center gap-1.5 text-xs text-slate-700 lg:border-r lg:border-slate-200 lg:pr-2.5">
              {isAdmin ? (
                <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0" />
              ) : (
                <UserCheck className="w-4 h-4 text-sky-600 shrink-0" />
              )}
              <span className="font-semibold text-slate-800 max-w-[120px] truncate">
                {currentUser.fullName}
              </span>
              <span className="text-slate-400 hidden xl:inline">
                ({roleLabel})
              </span>
            </div>
          )}

          <button
            type="button"
            onClick={onLogout}
            title="Đăng xuất khỏi hệ thống"
            className="hidden lg:inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 rounded-md transition-colors whitespace-nowrap shrink-0"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>Đăng xuất</span>
          </button>

          {/* Collapsible Menu Toggle for Mobile & Tablet (< lg) */}
          <button
            type="button"
            onClick={() => setMobileMenuOpen((prev) => !prev)}
            aria-expanded={mobileMenuOpen}
            aria-label={mobileMenuOpen ? 'Đóng menu điều hướng' : 'Mở menu điều hướng'}
            className="lg:hidden inline-flex items-center gap-1.5 min-h-[40px] px-2.5 py-1.5 text-xs font-bold text-slate-800 bg-slate-100 hover:bg-slate-200 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0"
          >
            {mobileMenuOpen ? <X className="w-4 h-4 text-slate-700" /> : <Menu className="w-4 h-4 text-sky-700" />}
            <span className="max-w-[96px] truncate">{activeNavItem?.mobileLabel || 'Menu'}</span>
          </button>
        </div>
      </div>

      {/* Collapsible Mobile / Tablet Drawer Menu */}
      {mobileMenuOpen && (
        <div className="lg:hidden border-t border-slate-200 bg-white shadow-lg px-3 py-3 space-y-3 max-h-[calc(100dvh-3.5rem)] overflow-y-auto">
          {accessibleBooks.length > 0 && activeBook && onSelectBook && (
            <div className="p-2.5 bg-sky-50 border border-sky-200 rounded-xl space-y-1.5">
              <label
                htmlFor="mobile-menu-book-select"
                className="text-[11px] font-bold uppercase tracking-wider text-sky-900 flex items-center gap-1.5"
              >
                <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                <span>Sổ đang chọn</span>
              </label>
              {accessibleBooks.length > 1 ? (
                <select
                  id="mobile-menu-book-select"
                  value={activeBook.id}
                  onChange={(e) => {
                    onSelectBook(e.target.value);
                    setMobileMenuOpen(false);
                  }}
                  className="w-full min-h-[42px] px-3 py-2 bg-white border border-sky-300 rounded-lg text-sm font-bold text-sky-950 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                >
                  {accessibleBooks.map((b) => (
                    <option key={b.id} value={b.id}>
                      {b.name} ({b.unitCode})
                    </option>
                  ))}
                </select>
              ) : (
                <div className="px-3 py-2 bg-white border border-sky-200 rounded-lg text-sm font-bold text-sky-950">
                  {activeBook.name} ({activeBook.unitCode})
                </div>
              )}
            </div>
          )}

          <nav className="grid grid-cols-1 sm:grid-cols-2 gap-1.5" aria-label="Điều hướng chính">
            {navItems.map((item) => (
              <button
                key={item.key}
                type="button"
                onClick={() => handleSelectTab(item.key)}
                className={`flex items-center gap-2.5 min-h-[44px] px-3.5 py-2.5 text-sm font-semibold rounded-xl border transition-colors text-left ${
                  activeTab === item.key
                    ? 'bg-sky-50 text-sky-900 border-sky-300 shadow-2xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                {item.icon}
                <span className="truncate">{item.label}</span>
              </button>
            ))}
          </nav>

          <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
            {currentUser && (
              <div className="flex items-center gap-2 text-xs text-slate-700 min-w-0">
                {isAdmin ? (
                  <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0" />
                ) : (
                  <UserCheck className="w-4 h-4 text-sky-600 shrink-0" />
                )}
                <div className="min-w-0">
                  <span className="font-bold text-slate-900 block truncate">{currentUser.fullName}</span>
                  <span className="text-[11px] text-slate-500 block truncate">
                    {roleLabel}
                    {formattedLastSync ? ` · Đồng bộ: ${formattedLastSync}` : ''}
                  </span>
                </div>
              </div>
            )}

            <button
              type="button"
              onClick={() => {
                setMobileMenuOpen(false);
                onLogout();
              }}
              className="inline-flex items-center justify-center gap-1.5 min-h-[42px] px-4 py-2 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors shrink-0"
            >
              <LogOut className="w-4 h-4" />
              <span>Đăng xuất</span>
            </button>
          </div>
        </div>
      )}
    </header>
  );
};
