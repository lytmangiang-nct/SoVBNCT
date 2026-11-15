import React, { useEffect, useRef, useState } from 'react';
import {
  AlertTriangle,
  BookOpen,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Lock,
  MoreHorizontal,
  Plus,
  Settings,
  Unlock,
} from 'lucide-react';
import type { DocumentTypeConfig, RegistryBook, User } from '../../types';
import {
  DEFAULT_BOOK_ID,
  getBookEnabledTypes,
  hasBookPermission,
} from '../../utils/permissions';
import { getBookDriveStatusBadge, getBookIdentityStyle } from './settingsConstants';

interface BooksListSectionProps {
  books: RegistryBook[];
  selectedBookId: string;
  currentUser: User;
  isAdmin: boolean;
  isConfigBlocked: boolean;
  documentTypes: DocumentTypeConfig[];
  bookFeedback: { success: boolean; text: string } | null;
  onDismissFeedback: () => void;
  onConfigureBook: (bookId: string) => void;
  onOpenCreateDrawer: () => void;
  onRequestToggleLockBook: (book: RegistryBook) => void;
}

export const BooksListSection: React.FC<BooksListSectionProps> = ({
  books,
  selectedBookId,
  currentUser,
  isAdmin,
  isConfigBlocked,
  documentTypes,
  bookFeedback,
  onDismissFeedback,
  onConfigureBook,
  onOpenCreateDrawer,
  onRequestToggleLockBook,
}) => {
  const [expandedBookIds, setExpandedBookIds] = useState<Record<string, boolean>>({});
  const [openMenuBookId, setOpenMenuBookId] = useState<string | null>(null);
  const menuContainerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      if (
        openMenuBookId &&
        menuContainerRef.current &&
        !menuContainerRef.current.contains(event.target as Node)
      ) {
        setOpenMenuBookId(null);
      }
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOpenMenuBookId(null);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handleOutsideClick);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [openMenuBookId]);

  const toggleDetails = (bookId: string) => {
    setExpandedBookIds((prev) => ({ ...prev, [bookId]: !prev[bookId] }));
  };

  return (
    <div className="space-y-4">
      {bookFeedback && (
        <div
          className={`p-3.5 rounded-xl border text-xs flex items-start justify-between gap-2 ${
            bookFeedback.success
              ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
              : 'bg-rose-50 text-rose-800 border-rose-200'
          }`}
        >
          <div className="flex items-start gap-2 min-w-0">
            {bookFeedback.success ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
            ) : (
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
            )}
            <span className="break-words">{bookFeedback.text}</span>
          </div>
          <button
            type="button"
            onClick={onDismissFeedback}
            className="text-[11px] font-semibold opacity-75 hover:opacity-100 shrink-0"
          >
            Đóng
          </button>
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-xl shadow-xs">
        <div className="p-4 sm:p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="min-w-0">
            <h2 className="text-base font-bold text-slate-900">
              Danh sách sổ văn bản ({books.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Chọn <strong>Cấu hình</strong> để mở thiết lập mẫu số và dãy số cho sổ tương ứng.
            </p>
          </div>

          {isAdmin && (
            <button
              type="button"
              disabled={isConfigBlocked}
              onClick={onOpenCreateDrawer}
              className="w-full sm:w-auto min-h-[42px] px-4 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg shadow-xs transition-colors inline-flex items-center justify-center gap-1.5 shrink-0"
            >
              <Plus className="w-4 h-4 shrink-0" />
              <span>Tạo sổ mới</span>
            </button>
          )}
        </div>

        <div className="divide-y divide-slate-100 text-xs">
          {books.map((book) => {
            const isSelected = book.id === selectedBookId;
            const identity = getBookIdentityStyle(book);
            const statusBadge = getBookDriveStatusBadge(book);
            const canConfigure = isAdmin || hasBookPermission(currentUser, book.id, 'manage_book');
            const canToggleActive = isAdmin && book.id !== DEFAULT_BOOK_ID;
            const isExpanded = Boolean(expandedBookIds[book.id]);
            const isMenuOpen = openMenuBookId === book.id;
            const enabledTypes = getBookEnabledTypes(book, documentTypes);

            return (
              <div
                key={book.id}
                className={`p-4 transition-colors ${
                  isSelected ? 'bg-sky-50/40' : 'bg-white'
                }`}
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  {/* Left: Book icon + Tên sổ + Ký hiệu + Trạng thái Drive */}
                  <div className="flex items-start gap-3 min-w-0 flex-1">
                    <div
                      className={`w-9 h-9 rounded-lg flex items-center justify-center shrink-0 border ${identity.iconBg} ${identity.iconText} ${identity.badgeBorder}`}
                    >
                      <BookOpen className="w-4 h-4" />
                    </div>

                    <div className="min-w-0 flex-1 space-y-1.5">
                      <div className="flex flex-wrap items-center gap-2 min-w-0">
                        <strong className="text-sm font-bold text-slate-900 break-words">
                          {book.name}
                        </strong>
                        <span
                          className={`font-mono font-bold px-2 py-0.5 rounded border text-[11px] shrink-0 ${identity.badgeBg} ${identity.badgeText} ${identity.badgeBorder}`}
                        >
                          {book.unitCode}
                        </span>
                      </div>

                      <div className="flex flex-wrap items-center gap-2">
                        <span
                          className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded border text-[11px] font-semibold ${statusBadge.className}`}
                        >
                          {statusBadge.kind === 'error' ? (
                            <AlertTriangle className="w-3 h-3 shrink-0" />
                          ) : statusBadge.kind === 'locked' ? (
                            <Lock className="w-3 h-3 shrink-0" />
                          ) : (
                            <span
                              className={`w-1.5 h-1.5 rounded-full shrink-0 ${statusBadge.dotColor}`}
                            />
                          )}
                          <span>{statusBadge.label}</span>
                        </span>

                        {book.isActive && book.googleDriveConfig?.lastCheckStatus === 'error' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded border text-[11px] font-medium bg-amber-50 text-amber-800 border-amber-200">
                            Cần kiểm tra Drive
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Right: Cấu hình + Chi tiết + ⋯ overflow */}
                  <div className="flex flex-wrap items-center gap-2 shrink-0">
                    {canConfigure && (
                      <button
                        type="button"
                        onClick={() => onConfigureBook(book.id)}
                        className="px-3 py-2 min-h-[40px] bg-sky-700 hover:bg-sky-800 text-white rounded-lg font-semibold transition-colors inline-flex items-center gap-1.5"
                      >
                        <Settings className="w-3.5 h-3.5 shrink-0" />
                        <span>Cấu hình</span>
                      </button>
                    )}

                    <button
                      type="button"
                      onClick={() => toggleDetails(book.id)}
                      aria-expanded={isExpanded}
                      className="px-3 py-2 min-h-[40px] bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-semibold transition-colors inline-flex items-center gap-1"
                    >
                      <span>Chi tiết</span>
                      {isExpanded ? (
                        <ChevronUp className="w-3.5 h-3.5 shrink-0" />
                      ) : (
                        <ChevronDown className="w-3.5 h-3.5 shrink-0" />
                      )}
                    </button>

                    {canToggleActive && (
                      <div
                        ref={isMenuOpen ? menuContainerRef : undefined}
                        className="relative"
                      >
                        <button
                          type="button"
                          aria-label={`Thao tác khác cho ${book.name}`}
                          aria-expanded={isMenuOpen}
                          disabled={isConfigBlocked}
                          onClick={() =>
                            setOpenMenuBookId((prev) => (prev === book.id ? null : book.id))
                          }
                          className="p-2 min-w-[40px] min-h-[40px] inline-flex items-center justify-center bg-white hover:bg-slate-100 border border-slate-200 text-slate-600 rounded-lg transition-colors disabled:opacity-50"
                        >
                          <MoreHorizontal className="w-4 h-4" />
                        </button>

                        {isMenuOpen && (
                          <div className="absolute right-0 mt-1.5 w-44 bg-white border border-slate-200 rounded-xl shadow-lg py-1 z-30">
                            <button
                              type="button"
                              onClick={() => {
                                setOpenMenuBookId(null);
                                onRequestToggleLockBook(book);
                              }}
                              className={`w-full px-3 py-2.5 min-h-[40px] text-left text-xs font-semibold flex items-center gap-2 transition-colors ${
                                book.isActive
                                  ? 'text-amber-800 hover:bg-amber-50'
                                  : 'text-emerald-800 hover:bg-emerald-50'
                              }`}
                            >
                              {book.isActive ? (
                                <>
                                  <Lock className="w-3.5 h-3.5 shrink-0" />
                                  <span>Tạm khóa sổ</span>
                                </>
                              ) : (
                                <>
                                  <Unlock className="w-3.5 h-3.5 shrink-0" />
                                  <span>Mở khóa sổ</span>
                                </>
                              )}
                            </button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                {/* Expandable Chi tiết panel */}
                {isExpanded && (
                  <div className="mt-3 pt-3 border-t border-slate-100 sm:pl-12 grid grid-cols-1 sm:grid-cols-2 gap-2.5 text-[11px] text-slate-600">
                    <div className="min-w-0">
                      <span className="text-slate-400 block">Đơn vị</span>
                      <span className="font-medium text-slate-800 break-words">
                        {book.unitName} ({book.isShared ? 'Sổ dùng chung' : 'Sổ riêng đơn vị'})
                      </span>
                    </div>
                    <div className="min-w-0">
                      <span className="text-slate-400 block">Mã nội bộ (Tab Sheets)</span>
                      <code className="font-mono font-semibold text-slate-800 break-all">
                        {book.code}
                      </code>
                    </div>
                    <div className="min-w-0">
                      <span className="text-slate-400 block">Mẫu số</span>
                      <code className="font-mono font-semibold text-slate-800 break-all">
                        {book.numberFormat}
                      </code>
                    </div>
                    <div className="min-w-0">
                      <span className="text-slate-400 block">Chế độ dãy số</span>
                      <span className="font-medium text-slate-800">
                        {book.numberMode === 'shared'
                          ? 'Dãy số chung liên tục'
                          : 'Dãy số riêng theo từng loại'}
                      </span>
                    </div>
                    <div className="sm:col-span-2 min-w-0">
                      <span className="text-slate-400 block mb-1">
                        Danh sách loại văn bản ({enabledTypes.length})
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {enabledTypes.map((t) => (
                          <span
                            key={t.code}
                            className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-100 border border-slate-200 text-slate-700 font-medium"
                          >
                            <span className="font-mono font-bold text-slate-900">{t.code}</span>
                            <span>· {t.name}</span>
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
