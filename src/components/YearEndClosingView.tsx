import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  Archive,
  BookOpen,
  CheckCircle2,
  ClipboardCheck,
  Download,
  ExternalLink,
  Eye,
  FileEdit,
  FileSpreadsheet,
  History,
  Loader2,
  Lock,
  Plus,
  RefreshCw,
  ShieldCheck,
  Trash2,
  X,
} from 'lucide-react';
import { api } from '../api';
import type {
  ClosingAmendmentChange,
  ClosingMetricKey,
  ClosingPreviewResponse,
  ClosingSnapshotDocument,
  RegistryBook,
  User,
  YearEndClosingSnapshot,
} from '../types';
import { hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface YearEndClosingViewProps {
  currentUser: User;
  books: RegistryBook[];
  activeBookId: string;
  serverYear: number;
  onOpenDocument?: (docId: string) => void;
}

const METRIC_CARDS: Array<{
  key: ClosingMetricKey;
  label: string;
  hint: string;
  tone: 'sky' | 'purple' | 'rose' | 'amber' | 'slate';
}> = [
  { key: 'outgoing', label: 'Văn bản đi', hint: 'Bấm để xem danh sách', tone: 'sky' },
  { key: 'incoming', label: 'Văn bản đến', hint: 'Bấm để xem danh sách', tone: 'purple' },
  { key: 'cancelled', label: 'Số đã hủy', hint: 'Bấm để xem số đã hủy', tone: 'rose' },
  { key: 'missingFile', label: 'Hồ sơ thiếu tệp', hint: 'Chưa lưu tệp Drive', tone: 'amber' },
  {
    key: 'unsyncedSheets',
    label: 'Chưa đồng bộ Trang tính',
    hint: 'Chờ hoặc lỗi đồng bộ',
    tone: 'slate',
  },
];

function filterDocsByMetric(
  docs: ClosingSnapshotDocument[],
  metric: ClosingMetricKey | 'all'
): ClosingSnapshotDocument[] {
  if (metric === 'all') return docs;
  if (metric === 'outgoing') return docs.filter((d) => d.docCategory !== 'incoming');
  if (metric === 'incoming') return docs.filter((d) => d.docCategory === 'incoming');
  if (metric === 'cancelled') return docs.filter((d) => d.status === 'cancelled');
  if (metric === 'missingFile') {
    return docs.filter(
      (d) => d.status !== 'cancelled' && (d.driveStatus !== 'saved' || (!d.driveFileId && !d.driveFileUrl))
    );
  }
  if (metric === 'unsyncedSheets') {
    return docs.filter((d) => d.syncStatus !== 'synced');
  }
  return docs;
}

const TONE_CLASS: Record<string, { box: string; active: string; num: string }> = {
  sky: {
    box: 'border-slate-200 bg-white hover:border-sky-300',
    active: 'border-sky-600 bg-sky-50 ring-2 ring-sky-600/20',
    num: 'text-sky-800',
  },
  purple: {
    box: 'border-slate-200 bg-white hover:border-purple-300',
    active: 'border-purple-600 bg-purple-50 ring-2 ring-purple-600/20',
    num: 'text-purple-800',
  },
  rose: {
    box: 'border-slate-200 bg-white hover:border-rose-300',
    active: 'border-rose-600 bg-rose-50 ring-2 ring-rose-600/20',
    num: 'text-rose-800',
  },
  amber: {
    box: 'border-slate-200 bg-white hover:border-amber-300',
    active: 'border-amber-600 bg-amber-50 ring-2 ring-amber-600/20',
    num: 'text-amber-800',
  },
  slate: {
    box: 'border-slate-200 bg-white hover:border-slate-400',
    active: 'border-slate-700 bg-slate-100 ring-2 ring-slate-600/20',
    num: 'text-slate-900',
  },
};

export const YearEndClosingView: React.FC<YearEndClosingViewProps> = ({
  currentUser,
  books,
  activeBookId,
  serverYear,
  onOpenDocument,
}) => {
  const closableBooks = useMemo(
    () =>
      books.filter(
        (b) => isSystemAdmin(currentUser) || hasBookPermission(currentUser, b.id, 'close_book')
      ),
    [books, currentUser]
  );

  const viewableBooks = useMemo(
    () =>
      books.filter(
        (b) =>
          isSystemAdmin(currentUser) ||
          hasBookPermission(currentUser, b.id, 'close_book') ||
          hasBookPermission(currentUser, b.id, 'view')
      ),
    [books, currentUser]
  );

  const [selectedBookId, setSelectedBookId] = useState<string>(() => {
    if (closableBooks.some((b) => b.id === activeBookId)) return activeBookId;
    return closableBooks[0]?.id || viewableBooks[0]?.id || books[0]?.id || '';
  });
  const [selectedYear, setSelectedYear] = useState<number>(serverYear);
  const [section, setSection] = useState<'prepare' | 'snapshots'>('prepare');

  // Preview state
  const [preview, setPreview] = useState<ClosingPreviewResponse | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [activeMetric, setActiveMetric] = useState<ClosingMetricKey | 'all'>('all');
  const [metricSearch, setMetricSearch] = useState('');

  // Confirm close modal
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [closeNote, setCloseNote] = useState('');
  const [closingSubmitting, setClosingSubmitting] = useState(false);
  const [closeError, setCloseError] = useState<string | null>(null);

  // Snapshots state
  const [snapshots, setSnapshots] = useState<YearEndClosingSnapshot[]>([]);
  const [snapshotsLoading, setSnapshotsLoading] = useState(false);
  const [snapshotsError, setSnapshotsError] = useState<string | null>(null);
  const [selectedSnapshotId, setSelectedSnapshotId] = useState<string | null>(null);
  const [snapshotInnerTab, setSnapshotInnerTab] = useState<'frozen' | 'amendments'>('frozen');
  const [snapshotMetric, setSnapshotMetric] = useState<ClosingMetricKey | 'all'>('all');
  const [exportingSnapshotId, setExportingSnapshotId] = useState<string | null>(null);

  // Amendment form state
  const [amendModalOpen, setAmendModalOpen] = useState(false);
  const [amendReason, setAmendReason] = useState('');
  const [amendChanges, setAmendChanges] = useState<ClosingAmendmentChange[]>([
    { docId: '', fullDocNumber: '', fieldLabel: 'Tên văn bản', beforeValue: '', afterValue: '' },
  ]);
  const [amendSubmitting, setAmendSubmitting] = useState(false);
  const [amendError, setAmendError] = useState<string | null>(null);

  const canCloseSelectedBook = useMemo(() => {
    if (!selectedBookId) return false;
    return (
      isSystemAdmin(currentUser) || hasBookPermission(currentUser, selectedBookId, 'close_book')
    );
  }, [currentUser, selectedBookId]);

  useEffect(() => {
    if (closableBooks.length === 0 && viewableBooks.length > 0) {
      setSection('snapshots');
    }
  }, [closableBooks.length, viewableBooks.length]);

  const yearOptions = useMemo(() => {
    const arr = new Set<number>([serverYear + 1, serverYear, serverYear - 1, serverYear - 2, serverYear - 3, selectedYear]);
    return Array.from(arr).sort((a, b) => b - a);
  }, [serverYear, selectedYear]);

  const loadPreview = useCallback(async () => {
    if (!selectedBookId || !canCloseSelectedBook) {
      setPreview(null);
      return;
    }
    setPreviewLoading(true);
    setPreviewError(null);
    try {
      const res = await api.getClosingPreview({ bookId: selectedBookId, year: selectedYear });
      setPreview(res);
    } catch (err: any) {
      setPreviewError(err?.message || 'Không tải được số liệu chốt sổ.');
    } finally {
      setPreviewLoading(false);
    }
  }, [selectedBookId, selectedYear, canCloseSelectedBook]);

  const loadSnapshots = useCallback(async () => {
    if (!selectedBookId) return;
    setSnapshotsLoading(true);
    setSnapshotsError(null);
    try {
      const res = await api.getClosingSnapshots({ bookId: selectedBookId });
      setSnapshots(res.snapshots || []);
      setSelectedSnapshotId((prev) => {
        if (prev && (res.snapshots || []).some((s) => s.id === prev)) return prev;
        return res.snapshots?.[0]?.id || null;
      });
    } catch (err: any) {
      setSnapshotsError(err?.message || 'Không tải được danh sách bản chốt sổ.');
    } finally {
      setSnapshotsLoading(false);
    }
  }, [selectedBookId]);

  useEffect(() => {
    loadPreview();
    loadSnapshots();
  }, [loadPreview, loadSnapshots]);

  const selectedSnapshot = useMemo(
    () => snapshots.find((s) => s.id === selectedSnapshotId) || null,
    [snapshots, selectedSnapshotId]
  );

  const previewMetricDocs = useMemo(() => {
    if (!preview) return [];
    const base = filterDocsByMetric(preview.documents, activeMetric);
    const q = metricSearch.trim().toLowerCase();
    if (!q) return base;
    return base.filter(
      (d) =>
        d.fullDocNumber.toLowerCase().includes(q) ||
        d.docName.toLowerCase().includes(q) ||
        (d.issuingAgency || '').toLowerCase().includes(q) ||
        (d.signer || '').toLowerCase().includes(q)
    );
  }, [preview, activeMetric, metricSearch]);

  const snapshotFrozenDocs = useMemo(() => {
    if (!selectedSnapshot) return [];
    return filterDocsByMetric(selectedSnapshot.documents, snapshotMetric);
  }, [selectedSnapshot, snapshotMetric]);

  const handleConfirmClosing = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedBookId || closingSubmitting) return;
    setClosingSubmitting(true);
    setCloseError(null);
    try {
      const res = await api.confirmClosing({
        bookId: selectedBookId,
        year: selectedYear,
        note: closeNote.trim() || undefined,
      });
      setConfirmOpen(false);
      setCloseNote('');
      setSnapshots((prev) => [res.snapshot, ...prev]);
      setSelectedSnapshotId(res.snapshot.id);
      setSnapshotInnerTab('frozen');
      setSection('snapshots');
      await loadPreview();
    } catch (err: any) {
      setCloseError(err?.message || 'Không thể hoàn tất chốt sổ.');
    } finally {
      setClosingSubmitting(false);
    }
  };

  const handleDownloadHandoverExcel = async (snap: YearEndClosingSnapshot) => {
    if (exportingSnapshotId) return;
    setExportingSnapshotId(snap.id);
    setSnapshotsError(null);
    try {
      await api.downloadClosingHandoverExcel(snap.id, snap.closingCode);
    } catch (err: any) {
      setSnapshotsError(err?.message || 'Không thể tải Excel bàn giao.');
    } finally {
      setExportingSnapshotId(null);
    }
  };

  const openAmendmentModal = () => {
    setAmendReason('');
    setAmendChanges([
      { docId: '', fullDocNumber: '', fieldLabel: 'Tên văn bản', beforeValue: '', afterValue: '' },
    ]);
    setAmendError(null);
    setAmendModalOpen(true);
  };

  const handleSelectDocForAmendmentRow = (idx: number, docId: string) => {
    if (!selectedSnapshot) return;
    const found = selectedSnapshot.documents.find((d) => d.docId === docId);
    setAmendChanges((prev) =>
      prev.map((item, i) => {
        if (i !== idx) return item;
        if (!found) return { ...item, docId: '', fullDocNumber: '' };
        let autoBefore = item.beforeValue;
        if (item.fieldLabel === 'Tên văn bản') autoBefore = found.docName;
        else if (item.fieldLabel === 'Số/Ký hiệu') autoBefore = found.fullDocNumber;
        else if (item.fieldLabel === 'Ngày ban hành') autoBefore = found.issueDate;
        else if (item.fieldLabel === 'Người ký') autoBefore = found.signer;
        else if (item.fieldLabel === 'Cơ quan ban hành') autoBefore = found.issuingAgency || '';
        return {
          ...item,
          docId: found.docId,
          fullDocNumber: found.fullDocNumber,
          beforeValue: autoBefore,
        };
      })
    );
  };

  const handleSubmitAmendment = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedSnapshot || amendSubmitting) return;
    if (!amendReason.trim()) {
      setAmendError('Vui lòng nhập lý do đính chính.');
      return;
    }
    const validChanges = amendChanges.filter(
      (c) => c.fieldLabel.trim() && (c.beforeValue.trim() || c.afterValue.trim())
    );
    if (validChanges.length === 0) {
      setAmendError('Vui lòng nhập ít nhất một trường thay đổi và giá trị trước/sau.');
      return;
    }

    setAmendSubmitting(true);
    setAmendError(null);
    try {
      const res = await api.addClosingAmendment(selectedSnapshot.id, {
        reason: amendReason.trim(),
        changes: validChanges,
      });
      setSnapshots((prev) => prev.map((s) => (s.id === res.snapshot.id ? res.snapshot : s)));
      setAmendModalOpen(false);
      setSnapshotInnerTab('amendments');
    } catch (err: any) {
      setAmendError(err?.message || 'Không thể lưu đính chính.');
    } finally {
      setAmendSubmitting(false);
    }
  };

  const bookOptionsForView = section === 'prepare' && closableBooks.length > 0 ? closableBooks : viewableBooks;

  if (viewableBooks.length === 0) {
    return (
      <div className="max-w-3xl mx-auto px-4 py-12">
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 text-center space-y-2">
          <AlertTriangle className="w-8 h-8 text-amber-600 mx-auto" />
          <h2 className="text-base font-bold text-amber-950">Chưa có quyền truy cập chức năng chốt sổ</h2>
          <p className="text-xs text-amber-900">
            Vui lòng liên hệ Quản trị viên để được cấp quyền Chốt sổ cuối năm hoặc Xem sổ văn bản.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-5 sm:py-6 space-y-5">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 flex items-center gap-2">
            <ClipboardCheck className="w-6 h-6 text-sky-700 shrink-0" />
            <span>Chốt sổ &amp; Bàn giao cuối năm</span>
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            Tổng hợp số liệu trước chốt, lưu bản chụp bất biến kèm mã lần chốt, xuất Excel bàn giao và ghi đính chính nếu phát sinh.
          </p>
        </div>

        {/* Sub-tab switch */}
        <div className="inline-flex bg-slate-100 p-1 rounded-lg border border-slate-200 self-start sm:self-auto">
          {canCloseSelectedBook && (
            <button
              type="button"
              onClick={() => setSection('prepare')}
              className={`px-3.5 py-2 min-h-[38px] rounded-md text-xs font-bold transition-colors whitespace-nowrap ${
                section === 'prepare'
                  ? 'bg-sky-900 text-white shadow-2xs'
                  : 'text-slate-700 hover:text-slate-900'
              }`}
            >
              Chuẩn bị &amp; Chốt sổ
            </button>
          )}
          <button
            type="button"
            onClick={() => setSection('snapshots')}
            className={`px-3.5 py-2 min-h-[38px] rounded-md text-xs font-bold transition-colors whitespace-nowrap inline-flex items-center gap-1.5 ${
              section === 'snapshots'
                ? 'bg-sky-900 text-white shadow-2xs'
                : 'text-slate-700 hover:text-slate-900'
            }`}
          >
            <span>Bản đã chốt</span>
            <span
              className={`font-mono text-[11px] px-1.5 py-0.5 rounded ${
                section === 'snapshots' ? 'bg-sky-800 text-sky-100' : 'bg-white text-slate-700'
              }`}
            >
              {snapshots.length}
            </span>
          </button>
        </div>
      </div>

      {/* Book & Year Toolbar */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs flex flex-wrap items-end justify-between gap-3">
        <div className="flex flex-wrap items-end gap-3">
          <div className="space-y-1 min-w-[220px]">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1">
              <BookOpen className="w-3.5 h-3.5 text-sky-700" />
              <span>Sổ văn bản</span>
            </label>
            <select
              value={selectedBookId}
              onChange={(e) => {
                setSelectedBookId(e.target.value);
                setActiveMetric('all');
              }}
              className="w-full min-h-[40px] px-3 text-xs font-semibold bg-slate-50 border border-slate-300 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
            >
              {bookOptionsForView.map((b) => (
                <option key={b.id} value={b.id}>
                  {b.name} ({b.unitCode})
                </option>
              ))}
            </select>
          </div>

          {section === 'prepare' && (
            <div className="space-y-1 min-w-[140px]">
              <label className="text-xs font-bold text-slate-700">Năm chốt sổ</label>
              <select
                value={selectedYear}
                onChange={(e) => {
                  setSelectedYear(parseInt(e.target.value, 10));
                  setActiveMetric('all');
                }}
                className="w-full min-h-[40px] px-3 text-xs font-semibold bg-slate-50 border border-slate-300 rounded-lg tabular-nums focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
              >
                {yearOptions.map((y) => (
                  <option key={y} value={y}>
                    Năm {y}
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => {
              loadPreview();
              loadSnapshots();
            }}
            disabled={previewLoading || snapshotsLoading}
            className="inline-flex items-center gap-1.5 min-h-[40px] px-3.5 py-2 text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg border border-slate-200"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${previewLoading || snapshotsLoading ? 'animate-spin' : ''}`} />
            <span>Làm mới số liệu</span>
          </button>

          {section === 'prepare' && canCloseSelectedBook && preview && (
            <button
              type="button"
              onClick={() => {
                setCloseError(null);
                setConfirmOpen(true);
              }}
              className="inline-flex items-center gap-2 min-h-[40px] px-4 py-2 text-xs font-bold bg-sky-800 hover:bg-sky-900 text-white rounded-lg shadow-xs"
            >
              <Lock className="w-4 h-4" />
              <span>Chốt sổ năm {selectedYear}</span>
            </button>
          )}
        </div>
      </div>

      {/* SECTION 1: PREPARE & CLOSE */}
      {section === 'prepare' && canCloseSelectedBook && (
        <div className="space-y-4">
          {previewError && (
            <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 shrink-0" />
              <span>{previewError}</span>
            </div>
          )}

          {previewLoading && !preview ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-2">
              <Loader2 className="w-6 h-6 animate-spin text-sky-700 mx-auto" />
              <p className="text-xs text-slate-500">Đang tổng hợp số liệu kiểm tra trước khi chốt sổ...</p>
            </div>
          ) : preview ? (
            <>
              {preview.existingSnapshotsCount > 0 && (
                <div className="bg-sky-50 border border-sky-200 rounded-xl p-3.5 text-xs text-sky-950 flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <ShieldCheck className="w-4 h-4 text-sky-700 shrink-0" />
                    <span>
                      Sổ <strong>{preview.bookName}</strong> năm <strong>{preview.year}</strong> đã có{' '}
                      <strong className="font-mono">{preview.existingSnapshotsCount}</strong> bản chốt trước đó. Nếu phát sinh cập nhật mới, lần chốt mới sẽ tạo thêm bản chụp mới mà không ghi đè bản cũ.
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setSection('snapshots')}
                    className="text-xs font-bold text-sky-800 hover:underline whitespace-nowrap"
                  >
                    Xem các bản đã chốt →
                  </button>
                </div>
              )}

              {/* Clickable Metric Cards */}
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                <button
                  type="button"
                  onClick={() => setActiveMetric('all')}
                  className={`text-left p-3.5 rounded-xl border transition-all ${
                    activeMetric === 'all'
                      ? 'border-sky-700 bg-sky-50 ring-2 ring-sky-600/20'
                      : 'border-slate-200 bg-white hover:border-slate-300'
                  }`}
                >
                  <p className="text-[11px] font-semibold text-slate-500">Tổng hồ sơ năm {preview.year}</p>
                  <p className="text-2xl font-extrabold font-mono tabular-nums text-slate-900 mt-1">
                    {preview.summary.totalActiveAndCancelled}
                  </p>
                  <p className="text-[11px] text-slate-500 mt-1">Tất cả văn bản</p>
                </button>

                {METRIC_CARDS.map((card) => {
                  const count = preview.summary[card.key];
                  const tone = TONE_CLASS[card.tone];
                  const isActive = activeMetric === card.key;
                  return (
                    <button
                      key={card.key}
                      type="button"
                      onClick={() => setActiveMetric(card.key)}
                      className={`text-left p-3.5 rounded-xl border transition-all ${
                        isActive ? tone.active : tone.box
                      }`}
                    >
                      <p className="text-[11px] font-semibold text-slate-600 truncate">{card.label}</p>
                      <p className={`text-2xl font-extrabold font-mono tabular-nums mt-1 ${tone.num}`}>
                        {count}
                      </p>
                      <p className="text-[11px] text-slate-500 mt-1 truncate">{card.hint}</p>
                    </button>
                  );
                })}
              </div>

              {/* Detail Table for Selected Metric */}
              <div className="bg-white border border-slate-200 rounded-xl shadow-xs overflow-hidden">
                <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
                  <div className="text-xs font-bold text-slate-800">
                    {activeMetric === 'all'
                      ? `Tất cả hồ sơ năm ${preview.year}`
                      : METRIC_CARDS.find((c) => c.key === activeMetric)?.label}{' '}
                    <span className="font-mono text-sky-800">({previewMetricDocs.length} hồ sơ)</span>
                  </div>
                  <input
                    type="text"
                    value={metricSearch}
                    onChange={(e) => setMetricSearch(e.target.value)}
                    placeholder="Lọc nhanh số VB, tên VB, người ký..."
                    className="w-full sm:w-64 min-h-[36px] px-3 text-xs bg-white border border-slate-300 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                  />
                </div>

                {previewMetricDocs.length === 0 ? (
                  <div className="p-8 text-center text-xs text-slate-500">
                    Không có hồ sơ nào trong nhóm mục này.
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                          <th className="px-3 py-2.5">Nhóm</th>
                          <th className="px-3 py-2.5">Số/Ký hiệu</th>
                          <th className="px-3 py-2.5">Ngày ban hành</th>
                          <th className="px-3 py-2.5 min-w-[220px]">Tên văn bản</th>
                          <th className="px-3 py-2.5">Tệp Drive</th>
                          <th className="px-3 py-2.5">Trang tính</th>
                          <th className="px-3 py-2.5">Trạng thái</th>
                          {onOpenDocument && <th className="px-3 py-2.5 text-right">Chi tiết</th>}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {previewMetricDocs.map((d) => (
                          <tr key={d.docId} className="hover:bg-slate-50/70">
                            <td className="px-3 py-2.5 whitespace-nowrap">
                              <span
                                className={`inline-block px-2 py-0.5 rounded font-semibold text-[11px] ${
                                  d.docCategory === 'incoming'
                                    ? 'bg-purple-50 text-purple-800'
                                    : 'bg-sky-50 text-sky-800'
                                }`}
                              >
                                {d.docCategory === 'incoming' ? 'VB đến' : 'VB đi'}
                              </span>
                            </td>
                            <td className="px-3 py-2.5 font-mono font-bold text-slate-900 whitespace-nowrap tabular-nums">
                              {d.fullDocNumber}
                            </td>
                            <td className="px-3 py-2.5 whitespace-nowrap tabular-nums text-slate-700">
                              {d.issueDate || (d.pendingDateConfirmation ? 'Chờ xác nhận' : '—')}
                            </td>
                            <td className="px-3 py-2.5 text-slate-800">
                              <div className="font-medium line-clamp-2">{d.docName}</div>
                              {d.issuingAgency && (
                                <div className="text-[11px] text-slate-500 mt-0.5">{d.issuingAgency}</div>
                              )}
                            </td>
                            <td className="px-3 py-2.5 whitespace-nowrap">
                              {d.driveStatus === 'saved' && (d.driveFileUrl || d.driveFileId) ? (
                                <span className="text-emerald-700 font-semibold">Đã có tệp</span>
                              ) : d.status === 'cancelled' ? (
                                <span className="text-slate-400">Đã hủy</span>
                              ) : (
                                <span className="text-amber-700 font-semibold">Thiếu tệp</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 whitespace-nowrap">
                              {d.syncStatus === 'synced' ? (
                                <span className="text-emerald-700 font-semibold">Đã đồng bộ</span>
                              ) : (
                                <span className="text-amber-700 font-semibold">
                                  {d.syncStatus === 'error' ? 'Lỗi đồng bộ' : 'Chưa đồng bộ'}
                                </span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 whitespace-nowrap">
                              {d.status === 'cancelled' ? (
                                <span className="text-rose-700 font-semibold">Đã hủy</span>
                              ) : (
                                <span className="text-slate-700">Hiệu lực</span>
                              )}
                            </td>
                            {onOpenDocument && (
                              <td className="px-3 py-2.5 text-right whitespace-nowrap">
                                <button
                                  type="button"
                                  onClick={() => onOpenDocument(d.docId)}
                                  className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded border border-slate-200 hover:bg-slate-100 text-[11px] font-semibold text-slate-700"
                                >
                                  <Eye className="w-3.5 h-3.5" />
                                  <span>Mở hồ sơ</span>
                                </button>
                              </td>
                            )}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </>
          ) : null}
        </div>
      )}

      {/* SECTION 2: IMMUTABLE SNAPSHOTS & ERRATA */}
      {section === 'snapshots' && (
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-start">
          {/* Left: Snapshot list */}
          <div className="lg:col-span-4 bg-white border border-slate-200 rounded-xl shadow-xs overflow-hidden">
            <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800">Danh sách lần chốt đã lưu</span>
              <span className="text-[11px] font-mono text-slate-500">{snapshots.length} bản</span>
            </div>
            {snapshotsError && (
              <div className="p-3 bg-rose-50 border-b border-rose-200 text-xs text-rose-800">
                {snapshotsError}
              </div>
            )}
            {snapshotsLoading ? (
              <div className="p-6 text-center text-xs text-slate-500">
                <Loader2 className="w-5 h-5 animate-spin text-sky-700 mx-auto mb-2" />
                Đang tải bản chốt...
              </div>
            ) : snapshots.length === 0 ? (
              <div className="p-6 text-center text-xs text-slate-500 space-y-2">
                <Archive className="w-6 h-6 text-slate-300 mx-auto" />
                <p>Chưa có bản chốt sổ nào cho sổ đang chọn.</p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100 max-h-[600px] overflow-y-auto">
                {snapshots.map((snap) => {
                  const active = snap.id === selectedSnapshot?.id;
                  return (
                    <button
                      key={snap.id}
                      type="button"
                      onClick={() => {
                        setSelectedSnapshotId(snap.id);
                        setSnapshotMetric('all');
                      }}
                      className={`w-full text-left p-3.5 transition-colors space-y-1 ${
                        active ? 'bg-sky-50 border-l-4 border-l-sky-700' : 'hover:bg-slate-50'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-mono text-xs font-extrabold text-sky-900">
                          {snap.closingCode}
                        </span>
                        <span className="text-[11px] font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-700 tabular-nums">
                          Năm {snap.year}
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-600">
                        Chốt lúc <strong className="tabular-nums">{snap.closedAtVietnam}</strong>
                      </p>
                      <p className="text-[11px] text-slate-500">
                        Người chốt: <strong>{snap.closedBy}</strong> ·{' '}
                        <span className="font-mono">{snap.summary.totalActiveAndCancelled}</span> VB
                        {snap.amendments.length > 0 && (
                          <span className="text-amber-700 font-semibold">
                            {' '}
                            · {snap.amendments.length} đính chính
                          </span>
                        )}
                      </p>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* Right: Selected Immutable Snapshot Details */}
          <div className="lg:col-span-8 space-y-4">
            {!selectedSnapshot ? (
              <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-xs text-slate-500">
                Chọn một bản chốt ở danh sách bên trái để xem số liệu bất biến, tải Excel bàn giao hoặc xem lịch sử đính chính.
              </div>
            ) : (
              <div className="bg-white border border-slate-200 rounded-xl shadow-xs overflow-hidden">
                {/* Immutable Banner */}
                <div className="p-4 bg-sky-950 text-white flex flex-wrap items-center justify-between gap-3">
                  <div className="space-y-1 min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <Lock className="w-4 h-4 text-sky-300 shrink-0" />
                      <span className="font-mono text-sm font-extrabold">{selectedSnapshot.closingCode}</span>
                      <span className="px-2 py-0.5 rounded bg-sky-800 text-sky-100 text-[11px] font-bold">
                        Bản chụp bất biến · Chỉ đọc
                      </span>
                    </div>
                    <p className="text-xs text-sky-200">
                      {selectedSnapshot.bookName} · Năm {selectedSnapshot.year} · Chốt bởi{' '}
                      <strong className="text-white">{selectedSnapshot.closedBy}</strong> lúc{' '}
                      <span className="tabular-nums">{selectedSnapshot.closedAtVietnam}</span>
                    </p>
                    {selectedSnapshot.note && (
                      <p className="text-xs text-sky-100 italic">Ghi chú: {selectedSnapshot.note}</p>
                    )}
                  </div>

                  <button
                    type="button"
                    onClick={() => handleDownloadHandoverExcel(selectedSnapshot)}
                    disabled={exportingSnapshotId === selectedSnapshot.id}
                    className="inline-flex items-center gap-2 min-h-[40px] px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-xs font-bold rounded-lg shadow-xs shrink-0"
                  >
                    {exportingSnapshotId === selectedSnapshot.id ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <FileSpreadsheet className="w-4 h-4" />
                    )}
                    <span>Tải Excel bàn giao (.xlsx)</span>
                  </button>
                </div>

                {/* Inner Sub-tabs: Bản đã chốt vs Lịch sử đính chính */}
                <div className="px-4 pt-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setSnapshotInnerTab('frozen')}
                      className={`px-3.5 py-2 text-xs font-bold border-b-2 transition-colors ${
                        snapshotInnerTab === 'frozen'
                          ? 'border-sky-700 text-sky-900'
                          : 'border-transparent text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Bản đã chốt ({selectedSnapshot.documents.length} hồ sơ)
                    </button>
                    <button
                      type="button"
                      onClick={() => setSnapshotInnerTab('amendments')}
                      className={`px-3.5 py-2 text-xs font-bold border-b-2 transition-colors inline-flex items-center gap-1.5 ${
                        snapshotInnerTab === 'amendments'
                          ? 'border-amber-600 text-amber-900'
                          : 'border-transparent text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      <History className="w-3.5 h-3.5" />
                      <span>Lịch sử đính chính ({selectedSnapshot.amendments.length})</span>
                    </button>
                  </div>

                  {canCloseSelectedBook && (
                    <button
                      type="button"
                      onClick={openAmendmentModal}
                      className="mb-2 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold"
                    >
                      <FileEdit className="w-3.5 h-3.5" />
                      <span>Ghi đính chính</span>
                    </button>
                  )}
                </div>

                {snapshotInnerTab === 'frozen' ? (
                  <div className="p-4 space-y-4">
                    {/* Frozen summary pills/cards */}
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5">
                      <button
                        type="button"
                        onClick={() => setSnapshotMetric('all')}
                        className={`text-left p-2.5 rounded-lg border text-xs ${
                          snapshotMetric === 'all'
                            ? 'border-sky-600 bg-sky-50'
                            : 'border-slate-200 bg-white'
                        }`}
                      >
                        <div className="text-[11px] text-slate-500">Tổng đã chốt</div>
                        <div className="text-lg font-extrabold font-mono tabular-nums text-slate-900">
                          {selectedSnapshot.summary.totalActiveAndCancelled}
                        </div>
                      </button>
                      {METRIC_CARDS.map((card) => (
                        <button
                          key={card.key}
                          type="button"
                          onClick={() => setSnapshotMetric(card.key)}
                          className={`text-left p-2.5 rounded-lg border text-xs ${
                            snapshotMetric === card.key
                              ? 'border-sky-600 bg-sky-50'
                              : 'border-slate-200 bg-white'
                          }`}
                        >
                          <div className="text-[11px] text-slate-600 truncate">{card.label}</div>
                          <div className="text-lg font-extrabold font-mono tabular-nums text-slate-900">
                            {selectedSnapshot.summary[card.key]}
                          </div>
                        </button>
                      ))}
                    </div>

                    {/* Frozen documents table */}
                    <div className="overflow-x-auto border border-slate-200 rounded-lg">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                            <th className="px-3 py-2">STT</th>
                            <th className="px-3 py-2">Nhóm</th>
                            <th className="px-3 py-2">Số/Ký hiệu</th>
                            <th className="px-3 py-2">Ngày</th>
                            <th className="px-3 py-2 min-w-[200px]">Tên văn bản (đã khóa)</th>
                            <th className="px-3 py-2">Tệp Drive</th>
                            <th className="px-3 py-2">Trạng thái</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {snapshotFrozenDocs.length === 0 ? (
                            <tr>
                              <td colSpan={7} className="px-4 py-6 text-center text-slate-500">
                                Không có hồ sơ thuộc mục đã chọn trong bản chụp.
                              </td>
                            </tr>
                          ) : (
                            snapshotFrozenDocs.map((doc, idx) => (
                              <tr key={`${doc.docId}-${idx}`} className="hover:bg-slate-50/70">
                                <td className="px-3 py-2 font-mono text-slate-500 tabular-nums">{idx + 1}</td>
                                <td className="px-3 py-2 whitespace-nowrap">
                                  {doc.docCategory === 'incoming' ? 'VB đến' : 'VB đi'}
                                </td>
                                <td className="px-3 py-2 font-mono font-bold text-slate-900 whitespace-nowrap tabular-nums">
                                  {doc.fullDocNumber}
                                </td>
                                <td className="px-3 py-2 whitespace-nowrap tabular-nums text-slate-700">
                                  {doc.issueDate || '—'}
                                </td>
                                <td className="px-3 py-2 text-slate-800">{doc.docName}</td>
                                <td className="px-3 py-2 whitespace-nowrap">
                                  {doc.driveFileUrl ? (
                                    <a
                                      href={doc.driveFileUrl}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="inline-flex items-center gap-1 text-sky-700 hover:underline font-semibold"
                                    >
                                      <ExternalLink className="w-3 h-3" />
                                      <span>Mở tệp</span>
                                    </a>
                                  ) : (
                                    <span className="text-amber-700">Thiếu tệp</span>
                                  )}
                                </td>
                                <td className="px-3 py-2 whitespace-nowrap">
                                  {doc.status === 'cancelled' ? (
                                    <span className="text-rose-700 font-semibold">Đã hủy</span>
                                  ) : (
                                    <span className="text-emerald-700 font-semibold">Hiệu lực</span>
                                  )}
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : (
                  <div className="p-4 space-y-3">
                    {selectedSnapshot.amendments.length === 0 ? (
                      <div className="p-8 text-center text-xs text-slate-500 space-y-1">
                        <CheckCircle2 className="w-6 h-6 text-emerald-500 mx-auto" />
                        <p>Chưa có bản ghi đính chính nào sau khi chốt sổ {selectedSnapshot.closingCode}.</p>
                      </div>
                    ) : (
                      selectedSnapshot.amendments.map((am) => (
                        <div
                          key={am.id}
                          className="border border-amber-200 bg-amber-50/40 rounded-xl p-3.5 space-y-2.5"
                        >
                          <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
                            <div className="font-bold text-amber-950">Lý do: {am.reason}</div>
                            <div className="text-[11px] text-slate-600 tabular-nums">
                              {am.amendedBy} (@{am.amendedByUsername}) · {am.amendedAtVietnam}
                            </div>
                          </div>
                          <div className="overflow-x-auto bg-white border border-amber-200/80 rounded-lg">
                            <table className="w-full text-left text-xs border-collapse">
                              <thead>
                                <tr className="bg-amber-50/70 border-b border-amber-200/70 text-amber-950 font-bold">
                                  <th className="px-3 py-1.5">Số/Ký hiệu</th>
                                  <th className="px-3 py-1.5">Trường thông tin</th>
                                  <th className="px-3 py-1.5">Giá trị trước</th>
                                  <th className="px-3 py-1.5">Giá trị sau</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {am.changes.map((c, cIdx) => (
                                  <tr key={cIdx}>
                                    <td className="px-3 py-1.5 font-mono font-bold text-slate-800 whitespace-nowrap">
                                      {c.fullDocNumber || '—'}
                                    </td>
                                    <td className="px-3 py-1.5 font-semibold text-slate-700">{c.fieldLabel}</td>
                                    <td className="px-3 py-1.5 text-rose-800 line-through">
                                      {c.beforeValue || '—'}
                                    </td>
                                    <td className="px-3 py-1.5 text-emerald-800 font-semibold">
                                      {c.afterValue || '—'}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Modal: Xác nhận chốt sổ */}
      {confirmOpen && preview && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl border border-slate-200 max-w-lg w-full overflow-hidden">
            <div className="px-5 py-4 bg-sky-900 text-white flex items-center justify-between">
              <h3 className="text-sm font-bold flex items-center gap-2">
                <Lock className="w-4 h-4" />
                <span>Xác nhận chốt sổ cuối năm {selectedYear}</span>
              </h3>
              <button
                type="button"
                onClick={() => !closingSubmitting && setConfirmOpen(false)}
                className="p-1 text-sky-200 hover:text-white"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleConfirmClosing} className="p-5 space-y-4 text-xs">
              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                <div className="flex justify-between">
                  <span className="text-slate-500">Sổ văn bản:</span>
                  <strong className="text-slate-900">{preview.bookName}</strong>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Năm chốt sổ:</span>
                  <strong className="font-mono text-slate-900">{preview.year}</strong>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Văn bản đi / Văn bản đến:</span>
                  <strong className="font-mono text-slate-900">
                    {preview.summary.outgoing} / {preview.summary.incoming}
                  </strong>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Số đã hủy:</span>
                  <strong className="font-mono text-rose-700">{preview.summary.cancelled}</strong>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Hồ sơ thiếu tệp / Chưa đồng bộ Trang tính:</span>
                  <strong className="font-mono text-amber-700">
                    {preview.summary.missingFile} / {preview.summary.unsyncedSheets}
                  </strong>
                </div>
              </div>

              <div className="space-y-1.5">
                <label className="font-bold text-slate-700 block">Ghi chú bàn giao (tùy chọn)</label>
                <textarea
                  rows={2}
                  value={closeNote}
                  onChange={(e) => setCloseNote(e.target.value)}
                  placeholder="Ví dụ: Chốt sổ bàn giao học kỳ II / cuối năm..."
                  className="w-full px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600"
                />
              </div>

              <p className="text-[11px] text-slate-500 leading-relaxed">
                Sau khi xác nhận, hệ thống lưu một bản chụp bất biến kèm mã lần chốt. Mọi chỉnh sửa hồ sơ trực tiếp sau thời điểm này sẽ không làm thay đổi số liệu hoặc Excel bàn giao của bản đã chốt.
              </p>

              {closeError && (
                <div className="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800">
                  {closeError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  disabled={closingSubmitting}
                  onClick={() => setConfirmOpen(false)}
                  className="min-h-[38px] px-3.5 py-1.5 text-slate-600 hover:bg-slate-100 rounded-lg font-semibold"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={closingSubmitting}
                  className="min-h-[38px] px-4 py-1.5 bg-sky-800 hover:bg-sky-900 disabled:opacity-60 text-white font-bold rounded-lg inline-flex items-center gap-1.5"
                >
                  {closingSubmitting ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>Đang lưu bản chốt...</span>
                    </>
                  ) : (
                    <>
                      <Lock className="w-3.5 h-3.5" />
                      <span>Khóa bản chốt bất biến</span>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: Ghi đính chính sau chốt */}
      {amendModalOpen && selectedSnapshot && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl border border-slate-200 max-w-2xl w-full max-h-[90vh] flex flex-col overflow-hidden">
            <div className="px-5 py-4 bg-amber-700 text-white flex items-center justify-between shrink-0">
              <h3 className="text-sm font-bold flex items-center gap-2">
                <FileEdit className="w-4 h-4" />
                <span>Ghi đính chính bản chốt {selectedSnapshot.closingCode}</span>
              </h3>
              <button
                type="button"
                onClick={() => !amendSubmitting && setAmendModalOpen(false)}
                className="p-1 text-amber-100 hover:text-white"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <form onSubmit={handleSubmitAmendment} className="p-5 space-y-4 text-xs overflow-y-auto">
              <div className="space-y-1.5">
                <label className="font-bold text-slate-700 block">
                  Lý do đính chính <span className="text-rose-600">*</span>
                </label>
                <input
                  type="text"
                  required
                  value={amendReason}
                  onChange={(e) => setAmendReason(e.target.value)}
                  placeholder="Ví dụ: Điều chỉnh tên trích yếu theo biên bản bàn giao văn thư..."
                  className="w-full min-h-[38px] px-3 bg-slate-50 border border-slate-300 rounded-lg focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-amber-600"
                />
              </div>

              <div className="space-y-2.5">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-700">Chi tiết thay đổi</span>
                  <button
                    type="button"
                    onClick={() =>
                      setAmendChanges((prev) => [
                        ...prev,
                        {
                          docId: '',
                          fullDocNumber: '',
                          fieldLabel: 'Tên văn bản',
                          beforeValue: '',
                          afterValue: '',
                        },
                      ])
                    }
                    className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Thêm dòng</span>
                  </button>
                </div>

                {amendChanges.map((ch, idx) => (
                  <div key={idx} className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2.5">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">
                          Hồ sơ trong bản chốt (tùy chọn)
                        </label>
                        <select
                          value={ch.docId || ''}
                          onChange={(e) => handleSelectDocForAmendmentRow(idx, e.target.value)}
                          className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg"
                        >
                          <option value="">-- Chọn hoặc tự nhập số/ký hiệu --</option>
                          {selectedSnapshot.documents.map((d) => (
                            <option key={d.docId} value={d.docId}>
                              {d.fullDocNumber} — {d.docName.slice(0, 50)}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Trường thông tin</label>
                        <input
                          type="text"
                          value={ch.fieldLabel}
                          onChange={(e) =>
                            setAmendChanges((prev) =>
                              prev.map((item, i) =>
                                i === idx ? { ...item, fieldLabel: e.target.value } : item
                              )
                            )
                          }
                          placeholder="Ví dụ: Tên văn bản, Ngày ban hành, Số/Ký hiệu..."
                          className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Số/Ký hiệu VB</label>
                        <input
                          type="text"
                          value={ch.fullDocNumber || ''}
                          onChange={(e) =>
                            setAmendChanges((prev) =>
                              prev.map((item, i) =>
                                i === idx ? { ...item, fullDocNumber: e.target.value } : item
                              )
                            )
                          }
                          className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg font-mono"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Giá trị trước</label>
                        <input
                          type="text"
                          value={ch.beforeValue}
                          onChange={(e) =>
                            setAmendChanges((prev) =>
                              prev.map((item, i) =>
                                i === idx ? { ...item, beforeValue: e.target.value } : item
                              )
                            )
                          }
                          className="w-full min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[11px] font-bold text-slate-600">Giá trị sau</label>
                        <div className="flex items-center gap-1.5">
                          <input
                            type="text"
                            value={ch.afterValue}
                            onChange={(e) =>
                              setAmendChanges((prev) =>
                                prev.map((item, i) =>
                                  i === idx ? { ...item, afterValue: e.target.value } : item
                                )
                              )
                            }
                            className="flex-1 min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-lg"
                          />
                          {amendChanges.length > 1 && (
                            <button
                              type="button"
                              onClick={() =>
                                setAmendChanges((prev) => prev.filter((_, i) => i !== idx))
                              }
                              className="p-2 text-rose-600 hover:bg-rose-50 rounded-lg"
                              title="Xóa dòng"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {amendError && (
                <div className="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800">
                  {amendError}
                </div>
              )}

              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  disabled={amendSubmitting}
                  onClick={() => setAmendModalOpen(false)}
                  className="min-h-[38px] px-3.5 py-1.5 text-slate-600 hover:bg-slate-100 rounded-lg font-semibold"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={amendSubmitting}
                  className="min-h-[38px] px-4 py-1.5 bg-amber-700 hover:bg-amber-800 disabled:opacity-60 text-white font-bold rounded-lg inline-flex items-center gap-1.5"
                >
                  {amendSubmitting ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>Đang lưu...</span>
                    </>
                  ) : (
                    <>
                      <Download className="w-3.5 h-3.5 hidden" />
                      <CheckCircle2 className="w-3.5 h-3.5" />
                      <span>Lưu đính chính</span>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
