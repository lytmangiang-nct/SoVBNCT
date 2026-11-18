import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  BookOpen,
  ExternalLink,
  Eye,
  FileSearch,
  FileText,
  Loader2,
  MessageSquareText,
  Send,
  X,
} from 'lucide-react';
import { api } from '../api';
import type { DocumentQaCitation, DocumentQaResponse, RegistryBook } from '../types';

interface QaTurn {
  id: string;
  question: string;
  loading?: boolean;
  error?: string;
  response?: DocumentQaResponse;
}

interface DocumentChatDrawerProps {
  open: boolean;
  onClose: () => void;
  viewableBooks: RegistryBook[];
  defaultBookId: string;
  onOpenDocumentDetail: (docId: string) => void;
  onOpenCitationPdfPage: (citation: DocumentQaCitation) => void;
}

const EXAMPLE_QUESTIONS = [
  'Tìm kế hoạch hoặc quyết định về hội thi 20/11 năm nay?',
  'Văn bản nào của Sở GDĐT hướng dẫn công tác chuyên môn?',
  'Quyết định nào công nhận kết quả cuộc thi video 20/11?',
];

export const DocumentChatDrawer: React.FC<DocumentChatDrawerProps> = ({
  open,
  onClose,
  viewableBooks,
  defaultBookId,
  onOpenDocumentDetail,
  onOpenCitationPdfPage,
}) => {
  const [selectedBookIds, setSelectedBookIds] = useState<string[]>([]);
  const [questionInput, setQuestionInput] = useState('');
  const [turns, setTurns] = useState<QaTurn[]>([]);
  const [isAsking, setIsAsking] = useState(false);
  const bottomRef = useRef<HTMLDivElement | null>(null);

  const permittedBookIds = useMemo(() => viewableBooks.map((b) => b.id), [viewableBooks]);

  // Reset default book selection whenever drawer opens
  useEffect(() => {
    if (!open) return;
    if (defaultBookId && defaultBookId !== 'all' && permittedBookIds.includes(defaultBookId)) {
      setSelectedBookIds([defaultBookId]);
    } else {
      setSelectedBookIds(permittedBookIds);
    }
  }, [open, defaultBookId, permittedBookIds]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [turns]);

  if (!open) return null;

  const toggleBookSelection = (bookId: string) => {
    setSelectedBookIds((prev) => {
      if (prev.includes(bookId)) {
        if (prev.length === 1) return prev; // Keep at least 1 selected book
        return prev.filter((id) => id !== bookId);
      }
      return [...prev, bookId];
    });
  };

  const handleAsk = async (rawQuestion?: string) => {
    const q = (rawQuestion ?? questionInput).trim();
    if (!q || isAsking || selectedBookIds.length === 0) return;

    const turnId = `qa-${Date.now()}`;
    setTurns((prev) => [...prev, { id: turnId, question: q, loading: true }]);
    if (!rawQuestion) setQuestionInput('');
    setIsAsking(true);

    try {
      const response = await api.askDocuments({
        question: q,
        bookIds: selectedBookIds,
      });
      setTurns((prev) =>
        prev.map((t) => (t.id === turnId ? { ...t, loading: false, response } : t))
      );
    } catch (err: any) {
      setTurns((prev) =>
        prev.map((t) =>
          t.id === turnId
            ? { ...t, loading: false, error: err?.message || 'Không thể tra cứu văn bản lúc này.' }
            : t
        )
      );
    } finally {
      setIsAsking(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex justify-end"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <aside
        aria-label="Hỏi đáp tra cứu văn bản có dẫn nguồn"
        className="bg-white w-full max-w-2xl h-[100dvh] flex flex-col shadow-2xl border-l border-slate-200 animate-slideUp sm:animate-none"
      >
        {/* Header */}
        <header className="px-4 sm:px-6 py-4 border-b border-slate-200 flex items-start justify-between gap-3 shrink-0 bg-slate-50/70">
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <MessageSquareText className="w-5 h-5 text-sky-700 shrink-0" />
              <h2 className="text-base font-bold text-slate-900 truncate">
                Hỏi về văn bản (có dẫn nguồn)
              </h2>
            </div>
            <p className="text-xs text-slate-600 mt-1">
              Hỗ trợ tiếng Việt có dấu hoặc không dấu. Hệ thống chỉ trả lời theo hồ sơ và nội dung PDF trong phạm vi bạn được phép xem.
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng khung hỏi đáp"
            className="p-2 min-w-[40px] min-h-[40px] flex items-center justify-center rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 shrink-0"
          >
            <X className="w-5 h-5" />
          </button>
        </header>

        {/* Book Scope Selector */}
        <div className="px-4 sm:px-6 py-3 border-b border-slate-200 bg-white shrink-0 space-y-2">
          <div className="flex items-center justify-between gap-2 text-xs">
            <span className="font-semibold text-slate-700 flex items-center gap-1.5">
              <BookOpen className="w-3.5 h-3.5 text-sky-700" />
              <span>Phạm vi sổ tra cứu ({selectedBookIds.length}/{viewableBooks.length}):</span>
            </span>
            {viewableBooks.length > 1 && (
              <button
                type="button"
                onClick={() =>
                  setSelectedBookIds(
                    selectedBookIds.length === viewableBooks.length
                      ? [
                          defaultBookId !== 'all' && permittedBookIds.includes(defaultBookId)
                            ? defaultBookId
                            : permittedBookIds[0],
                        ]
                      : permittedBookIds
                  )
                }
                className="text-[11px] font-semibold text-sky-700 hover:underline"
              >
                {selectedBookIds.length === viewableBooks.length
                  ? 'Chỉ sổ đang mở'
                  : 'Chọn tất cả sổ có quyền'}
              </button>
            )}
          </div>
          <div className="flex flex-wrap gap-2">
            {viewableBooks.map((book) => {
              const checked = selectedBookIds.includes(book.id);
              return (
                <label
                  key={book.id}
                  className={`inline-flex items-center gap-1.5 min-h-[34px] px-2.5 py-1 rounded-lg border text-xs font-medium cursor-pointer transition-colors ${
                    checked
                      ? 'bg-sky-50 border-sky-300 text-sky-900 font-semibold'
                      : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={() => toggleBookSelection(book.id)}
                    className="rounded border-slate-300 text-sky-700"
                  />
                  <span>{book.name}</span>
                </label>
              );
            })}
          </div>
        </div>

        {/* Conversation Area */}
        <div className="flex-1 overflow-y-auto px-4 sm:px-6 py-4 space-y-5 bg-slate-50/40">
          {turns.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-3">
              <div className="flex items-center gap-2 text-sm font-bold text-slate-900">
                <FileSearch className="w-4 h-4 text-sky-700 shrink-0" />
                <span>Đặt câu hỏi về số hiệu, nội dung hoặc văn bản đã lưu</span>
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                Mỗi câu trả lời sẽ đính kèm số/ký hiệu, tên văn bản, đoạn làm căn cứ và nút mở hồ sơ hoặc đúng trang PDF. Nếu hồ sơ chưa có nội dung PDF trích xuất, hệ thống ghi rõ căn cứ là thông tin bảng kê.
              </p>
              <div className="space-y-1.5 pt-1">
                <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wide block">
                  Gợi ý nhanh
                </span>
                <div className="flex flex-col gap-1.5">
                  {EXAMPLE_QUESTIONS.map((ex) => (
                    <button
                      key={ex}
                      type="button"
                      onClick={() => handleAsk(ex)}
                      className="text-left px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-sky-50 hover:border-sky-200 text-xs text-slate-700 hover:text-sky-900 transition-colors"
                    >
                      {ex}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          ) : (
            turns.map((turn) => (
              <div key={turn.id} className="space-y-3">
                {/* User Question */}
                <div className="flex justify-end">
                  <div className="max-w-[88%] bg-sky-700 text-white rounded-xl rounded-tr-xs px-4 py-2.5 text-xs sm:text-sm shadow-2xs">
                    {turn.question}
                  </div>
                </div>

                {/* Assistant Grounded Answer */}
                <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-2xs space-y-3">
                  {turn.loading ? (
                    <div className="flex items-center gap-2 text-xs text-slate-600 py-1">
                      <Loader2 className="w-4 h-4 animate-spin text-sky-700 shrink-0" />
                      <span>Đang tìm kiếm trong sổ và nội dung PDF được phép xem...</span>
                    </div>
                  ) : turn.error ? (
                    <div className="flex items-start gap-2 text-xs text-rose-800 bg-rose-50 border border-rose-200 rounded-lg p-3">
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                      <span>{turn.error}</span>
                    </div>
                  ) : turn.response ? (
                    <>
                      {turn.response.insufficientEvidence && (
                        <div className="p-2.5 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-950 flex items-start gap-2">
                          <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                          <div className="space-y-0.5">
                            <strong className="block">Chưa tìm thấy văn bản đủ căn cứ</strong>
                            {turn.response.mismatchNote && (
                              <span className="block text-[11px] text-amber-900">
                                {turn.response.mismatchNote}
                              </span>
                            )}
                          </div>
                        </div>
                      )}

                      <div className="text-xs sm:text-sm text-slate-800 whitespace-pre-line leading-relaxed">
                        {turn.response.answer}
                      </div>

                      <div className="text-[11px] text-slate-500 flex flex-wrap items-center gap-x-2 gap-y-1 tabular-nums">
                        <span>
                          Phạm vi: {turn.response.searchedBooks.map((b) => b.name).join(', ')}
                        </span>
                        <span aria-hidden="true">·</span>
                        <span>Đã duyệt {turn.response.totalSearchedDocs} hồ sơ</span>
                        <span aria-hidden="true">·</span>
                        <span>{turn.response.indexedPdfCount} PDF có nội dung trang</span>
                      </div>

                      {turn.response.citations.length > 0 && (
                        <div className="pt-2 border-t border-slate-100 space-y-2.5">
                          <div className="text-[11px] font-bold text-slate-600 uppercase tracking-wide">
                            Nguồn căn cứ ({turn.response.citations.length})
                          </div>
                          {turn.response.citations.map((cit, cIdx) => {
                            const isPdfSource = cit.sourceType === 'pdf_page';
                            return (
                              <div
                                key={`${cit.docId}-${cit.page || 0}-${cIdx}`}
                                className="border border-slate-200 rounded-lg p-3 bg-slate-50/70 space-y-2"
                              >
                                <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-2">
                                  <div className="min-w-0 space-y-0.5">
                                    <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs">
                                      <span className="font-mono font-bold text-sky-900">
                                        {cit.fullDocNumber}
                                      </span>
                                      <span className="text-slate-300" aria-hidden="true">·</span>
                                      <span className="font-bold text-slate-900 break-words">
                                        {cit.docName}
                                      </span>
                                    </div>
                                    <p className="text-[11px] text-slate-600 tabular-nums">
                                      {cit.bookName} · {cit.typeName}
                                      {cit.issueDate ? ` · Ngày ${cit.issueDate}` : ''}
                                      {cit.issuingAgency ? ` · ${cit.issuingAgency}` : ''}
                                      {' · '}
                                      <strong
                                        className={
                                          isPdfSource ? 'text-purple-800' : 'text-slate-800'
                                        }
                                      >
                                        {isPdfSource
                                          ? `Căn cứ: Nội dung PDF (Trang ${cit.page})`
                                          : 'Căn cứ: Thông tin hồ sơ (bảng kê)'}
                                      </strong>
                                    </p>
                                  </div>

                                  <div className="flex items-center gap-1.5 shrink-0">
                                    <button
                                      type="button"
                                      onClick={() => onOpenDocumentDetail(cit.docId)}
                                      className="inline-flex items-center gap-1 min-h-[34px] px-2.5 py-1 text-[11px] font-semibold bg-white hover:bg-slate-100 border border-slate-300 text-slate-800 rounded-md"
                                    >
                                      <Eye className="w-3.5 h-3.5 text-sky-700" />
                                      <span>Xem hồ sơ</span>
                                    </button>

                                    {isPdfSource && cit.hasPdfFile && cit.page ? (
                                      <button
                                        type="button"
                                        onClick={() => onOpenCitationPdfPage(cit)}
                                        className="inline-flex items-center gap-1 min-h-[34px] px-2.5 py-1 text-[11px] font-semibold bg-purple-700 hover:bg-purple-800 text-white rounded-md"
                                      >
                                        <FileText className="w-3.5 h-3.5" />
                                        <span>Mở trang {cit.page} PDF</span>
                                      </button>
                                    ) : cit.hasDriveFile ? (
                                      <button
                                        type="button"
                                        onClick={() => onOpenCitationPdfPage(cit)}
                                        className="inline-flex items-center gap-1 min-h-[34px] px-2.5 py-1 text-[11px] font-semibold bg-emerald-700 hover:bg-emerald-800 text-white rounded-md"
                                      >
                                        <ExternalLink className="w-3.5 h-3.5" />
                                        <span>Mở tệp</span>
                                      </button>
                                    ) : null}
                                  </div>
                                </div>

                                <p className="text-xs text-slate-700 bg-white border border-slate-200 rounded px-2.5 py-2 leading-relaxed break-words">
                                  {cit.excerpt}
                                </p>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </>
                  ) : null}
                </div>
              </div>
            ))
          )}
          <div ref={bottomRef} />
        </div>

        {/* Input Form */}
        <form
          onSubmit={(e) => {
            e.preventDefault();
            handleAsk();
          }}
          className="p-3 sm:p-4 border-t border-slate-200 bg-white shrink-0"
        >
          <div className="flex items-center gap-2">
            <input
              type="text"
              value={questionInput}
              onChange={(e) => setQuestionInput(e.target.value)}
              placeholder="Ví dụ: Tìm văn bản của Sở về hội thi 20/11 năm 2026..."
              disabled={isAsking}
              className="flex-1 min-h-[44px] px-3.5 bg-slate-50 border border-slate-300 rounded-lg text-sm text-slate-900 focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            />
            <button
              type="submit"
              disabled={isAsking || !questionInput.trim() || selectedBookIds.length === 0}
              className="inline-flex items-center justify-center gap-1.5 min-h-[44px] px-4 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs sm:text-sm font-semibold rounded-lg shrink-0"
            >
              {isAsking ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Send className="w-4 h-4" />
              )}
              <span>Hỏi</span>
            </button>
          </div>
        </form>
      </aside>
    </div>
  );
};
