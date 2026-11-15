import React from 'react';
import { FileText, FileUp, BookOpen, BarChart3, Settings, LogOut, ShieldCheck, UserCheck } from 'lucide-react';
import type { RegistryBook, User } from '../types';
import { hasAnyBookPermission, isSystemAdmin } from '../utils/permissions';

export type AppTab = 'issue' | 'import_pdf' | 'registry' | 'statistics' | 'settings';

interface NavbarProps {
  activeTab: AppTab;
  setActiveTab: (tab: AppTab) => void;
  currentUser: User | null;
  books: RegistryBook[];
  onLogout: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  setActiveTab,
  currentUser,
  books,
  onLogout,
}) => {
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

  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200 shadow-xs">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
        {/* Zone 1: Brand Wordmark */}
        <div className="flex items-center gap-2.5 sm:gap-3 min-w-0">
          <div className="w-9 h-9 sm:w-10 sm:h-10 rounded-lg bg-sky-700 flex items-center justify-center text-white font-bold shadow-xs shrink-0">
            <svg className="w-5 h-5 sm:w-6 sm:h-6 text-white" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z" />
            </svg>
          </div>
          <div className="flex flex-col min-w-0">
            <span className="text-xs sm:text-sm font-bold tracking-tight text-sky-900 truncate">
              ĐOÀN TRƯỜNG THPT NGUYỄN CHÍ THANH
            </span>
            <span className="text-[11px] text-slate-500 font-medium truncate">
              Sổ Cấp Số Văn Bản Tự Động
            </span>
          </div>
        </div>

        {/* Zone 2: Navigation tabs on desktop */}
        <nav className="hidden md:flex items-center gap-1">
          {canIssueOrRequest && (
            <button
              onClick={() => setActiveTab('issue')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'issue'
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <FileText className="w-4 h-4 text-sky-700" />
              <span>{issueDesktopLabel}</span>
            </button>
          )}

          {canImportPdf && (
            <button
              onClick={() => setActiveTab('import_pdf')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'import_pdf'
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <FileUp className="w-4 h-4 text-sky-700" />
              <span>Nhập văn bản PDF</span>
            </button>
          )}

          {canViewRegistry && (
            <button
              onClick={() => setActiveTab('registry')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'registry'
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <BookOpen className="w-4 h-4 text-sky-700" />
              <span>Sổ theo dõi</span>
            </button>
          )}

          {canViewStats && (
            <button
              onClick={() => setActiveTab('statistics')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'statistics'
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <BarChart3 className="w-4 h-4 text-sky-700" />
              <span>Thống kê</span>
            </button>
          )}

          {canManageSettings && (
            <button
              onClick={() => setActiveTab('settings')}
              className={`flex items-center gap-2 px-3.5 py-2 text-sm font-semibold rounded-md transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'settings'
                  ? 'bg-sky-50 text-sky-800'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
              }`}
            >
              <Settings className="w-4 h-4 text-sky-700" />
              <span>Cài đặt</span>
            </button>
          )}
        </nav>

        {/* Zone 3: Primary Action / User Profile */}
        <div className="flex items-center gap-2 sm:gap-3 shrink-0">
          {currentUser && (
            <div className="flex items-center gap-1.5 text-xs text-slate-700 sm:border-r sm:border-slate-200 sm:pr-3">
              {isAdmin ? (
                <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0" />
              ) : (
                <UserCheck className="w-4 h-4 text-sky-600 shrink-0" />
              )}
              <span className="font-semibold text-slate-800 max-w-[110px] sm:max-w-none truncate">
                {currentUser.fullName}
              </span>
              <span className="text-slate-400 hidden lg:inline">
                ({roleLabel})
              </span>
            </div>
          )}

          <button
            onClick={onLogout}
            title="Đăng xuất khỏi hệ thống"
            className="flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 rounded-md transition-colors whitespace-nowrap shrink-0"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Đăng xuất</span>
          </button>
        </div>
      </div>

      {/* Mobile bottom-friendly nav bar (đồng nhất quyền với desktop) */}
      <div className="md:hidden flex items-center justify-around border-t border-slate-200 bg-white py-1.5 px-1">
        {canIssueOrRequest && (
          <button
            onClick={() => setActiveTab('issue')}
            className={`flex flex-col items-center py-1 px-2 text-[11px] font-semibold rounded-lg transition-colors ${
              activeTab === 'issue' ? 'bg-sky-50 text-sky-800' : 'text-slate-500 hover:text-slate-900'
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>{issueMobileLabel}</span>
          </button>
        )}
        {canImportPdf && (
          <button
            onClick={() => setActiveTab('import_pdf')}
            className={`flex flex-col items-center py-1 px-2 text-[11px] font-semibold rounded-lg transition-colors ${
              activeTab === 'import_pdf' ? 'bg-sky-50 text-sky-800' : 'text-slate-500 hover:text-slate-900'
            }`}
          >
            <FileUp className="w-4 h-4" />
            <span>Nhập PDF</span>
          </button>
        )}
        {canViewRegistry && (
          <button
            onClick={() => setActiveTab('registry')}
            className={`flex flex-col items-center py-1 px-2 text-[11px] font-semibold rounded-lg transition-colors ${
              activeTab === 'registry' ? 'bg-sky-50 text-sky-800' : 'text-slate-500 hover:text-slate-900'
            }`}
          >
            <BookOpen className="w-4 h-4" />
            <span>Sổ theo dõi</span>
          </button>
        )}
        {canViewStats && (
          <button
            onClick={() => setActiveTab('statistics')}
            className={`flex flex-col items-center py-1 px-2 text-[11px] font-semibold rounded-lg transition-colors ${
              activeTab === 'statistics' ? 'bg-sky-50 text-sky-800' : 'text-slate-500 hover:text-slate-900'
            }`}
          >
            <BarChart3 className="w-4 h-4" />
            <span>Thống kê</span>
          </button>
        )}
        {canManageSettings && (
          <button
            onClick={() => setActiveTab('settings')}
            className={`flex flex-col items-center py-1 px-2 text-[11px] font-semibold rounded-lg transition-colors ${
              activeTab === 'settings' ? 'bg-sky-50 text-sky-800' : 'text-slate-500 hover:text-slate-900'
            }`}
          >
            <Settings className="w-4 h-4" />
            <span>Cài đặt</span>
          </button>
        )}
      </div>
    </header>
  );
};
