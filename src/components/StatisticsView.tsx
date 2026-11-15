import React, { useState, useEffect, useMemo } from 'react';
import {
  FileCheck,
  FileX,
  Layers,
  Calendar,
  RefreshCw,
  TrendingUp,
  BookOpen,
} from 'lucide-react';
import { api } from '../api';
import type { RegistryBook, StatisticsData, User } from '../types';
import { hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface StatisticsViewProps {
  currentYear: number;
  books: RegistryBook[];
  activeBookId: string;
  onSelectBookId: (bookId: string) => void;
  currentUser: User;
}

export const StatisticsView: React.FC<StatisticsViewProps> = ({
  currentYear,
  books,
  activeBookId,
  onSelectBookId,
  currentUser,
}) => {
  const viewableBooks = useMemo(
    () =>
      books.filter(
        (b) => b.isActive && (isSystemAdmin(currentUser) || hasBookPermission(currentUser, b.id, 'view'))
      ),
    [books, currentUser]
  );

  const [selectedYear, setSelectedYear] = useState<number>(currentYear);
  const prevCurrentYearRef = React.useRef<number>(currentYear);
  useEffect(() => {
    if (prevCurrentYearRef.current !== currentYear) {
      if (selectedYear === prevCurrentYearRef.current) {
        setSelectedYear(currentYear);
      }
      prevCurrentYearRef.current = currentYear;
    }
  }, [currentYear, selectedYear]);

  const [statsBookId, setStatsBookId] = useState<string>(() => {
    if (viewableBooks.some((b) => b.id === activeBookId)) return activeBookId;
    return viewableBooks.length > 1 ? 'all' : viewableBooks[0]?.id || 'all';
  });
  const [stats, setStats] = useState<StatisticsData | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const prevActiveBookRef = React.useRef<string>(activeBookId);
  useEffect(() => {
    if (prevActiveBookRef.current !== activeBookId) {
      prevActiveBookRef.current = activeBookId;
      if (viewableBooks.some((b) => b.id === activeBookId)) {
        setStatsBookId(activeBookId);
      }
    }
  }, [activeBookId, viewableBooks]);

  // Sync if viewableBooks shrinks
  useEffect(() => {
    if (statsBookId === 'all' && viewableBooks.length === 1) {
      setStatsBookId(viewableBooks[0].id);
      return;
    }
    if (statsBookId !== 'all' && !viewableBooks.some((b) => b.id === statsBookId)) {
      if (viewableBooks.some((b) => b.id === activeBookId)) {
        setStatsBookId(activeBookId);
      } else {
        setStatsBookId(viewableBooks.length > 1 ? 'all' : viewableBooks[0]?.id || 'all');
      }
    }
  }, [viewableBooks, activeBookId, statsBookId]);

  const fetchStats = async (year: number, bookFilter: string) => {
    setLoading(true);
    setError(null);
    try {
      const data = await api.getStatistics(year, bookFilter);
      setStats(data);
    } catch (err: any) {
      setError(err.message || 'Không thể tải số liệu thống kê');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchStats(selectedYear, statsBookId);
  }, [selectedYear, statsBookId, currentYear]);

  const selectedBookObj = viewableBooks.find((b) => b.id === statsBookId);

  // Compute maximum monthly count for scaling bars
  const maxMonthCount = stats
    ? Math.max(...stats.byMonth.map((m) => m.count), 1)
    : 1;

  return (
    <div className="max-w-7xl mx-auto py-5 sm:py-8 px-3 sm:px-6 lg:px-8 space-y-5 sm:space-y-8">
      {/* Header & Year / Book Selector */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 sm:gap-4">
        <div className="min-w-0">
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
            Thống Kê Sổ Văn Bản
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            {selectedBookObj
              ? `Báo cáo tổng hợp cho sổ: ${selectedBookObj.name} (${selectedBookObj.unitCode}).`
              : 'Báo cáo tổng hợp số lượng văn bản đi, văn bản đến theo năm, theo loại và theo từng tháng trên các sổ bạn có quyền xem.'}
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
          {viewableBooks.length > 0 && (
            <div className="flex items-center gap-2 bg-white border border-slate-200 px-3 py-1.5 min-h-[42px] rounded-lg shadow-2xs w-full sm:w-auto min-w-0">
              <BookOpen className="w-4 h-4 text-sky-700 shrink-0" />
              <span className="text-xs font-semibold text-slate-700 shrink-0">Sổ:</span>
              <select
                value={statsBookId}
                onChange={(e) => {
                  const val = e.target.value;
                  setStatsBookId(val);
                  if (val !== 'all') onSelectBookId(val);
                }}
                className="text-base sm:text-xs font-bold text-sky-800 bg-transparent border-0 cursor-pointer focus:outline-hidden flex-1 sm:flex-initial min-w-0 sm:max-w-[200px]"
              >
                {viewableBooks.length > 1 && <option value="all">Tất cả sổ có quyền xem</option>}
                {viewableBooks.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.name} ({b.unitCode})
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="flex items-center gap-2 flex-1 sm:flex-initial bg-white border border-slate-200 px-3 py-1.5 min-h-[42px] rounded-lg shadow-2xs min-w-0">
            <Calendar className="w-4 h-4 text-slate-500 shrink-0" />
            <span className="text-xs font-semibold text-slate-700 shrink-0">Năm:</span>
            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(parseInt(e.target.value, 10))}
              className="text-base sm:text-xs font-bold text-sky-800 bg-transparent border-0 cursor-pointer focus:outline-hidden flex-1 sm:flex-initial min-w-0"
            >
              {[currentYear - 2, currentYear - 1, currentYear, currentYear + 1].map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>

          <button
            type="button"
            onClick={() => fetchStats(selectedYear, statsBookId)}
            disabled={loading}
            className="w-10 h-10 text-slate-600 hover:text-sky-700 bg-white border border-slate-200 rounded-lg shadow-2xs transition-colors inline-flex items-center justify-center shrink-0 disabled:opacity-60"
            title="Làm mới thống kê"
            aria-label="Làm mới thống kê"
          >
            <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
          </button>
        </div>
      </div>

      {loading && !stats ? (
        <div className="p-12 text-center text-slate-400">
          <div className="w-8 h-8 border-3 border-sky-600 border-t-transparent rounded-full animate-spin mx-auto mb-3" />
          <p className="text-sm">Đang tính toán thống kê...</p>
        </div>
      ) : error ? (
        <div className="p-4 bg-rose-50 text-rose-700 rounded-xl text-sm border border-rose-200">
          {error}
        </div>
      ) : stats ? (
        <>
          {/* Key Metric Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {/* 1. Tổng văn bản */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs">
              <div className="flex items-center justify-between text-slate-500 mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                  Tổng văn bản đã ghi
                </span>
                <span className="p-2 bg-sky-50 text-sky-700 rounded-lg">
                  <Layers className="w-4 h-4" />
                </span>
              </div>
              <div className="text-3xl font-extrabold text-slate-900 font-mono tracking-tight">
                {stats.totalIssued}
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Toàn bộ văn bản đã phát sinh trong năm {stats.year}
              </div>
            </div>

            {/* 2. Đang hiệu lực */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs">
              <div className="flex items-center justify-between text-slate-500 mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                  Đang hiệu lực
                </span>
                <span className="p-2 bg-emerald-50 text-emerald-700 rounded-lg">
                  <FileCheck className="w-4 h-4" />
                </span>
              </div>
              <div className="text-3xl font-extrabold text-emerald-700 font-mono tracking-tight">
                {stats.activeCount}
              </div>
              <div className="text-xs text-slate-500 mt-1">
                {stats.totalIssued > 0
                  ? `${Math.round((stats.activeCount / stats.totalIssued) * 100)}% tổng số văn bản`
                  : 'Chưa có văn bản'}
              </div>
            </div>

            {/* 3. Đã hủy */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs">
              <div className="flex items-center justify-between text-slate-500 mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                  Số văn bản đã hủy
                </span>
                <span className="p-2 bg-rose-50 text-rose-700 rounded-lg">
                  <FileX className="w-4 h-4" />
                </span>
              </div>
              <div className="text-3xl font-extrabold text-rose-700 font-mono tracking-tight">
                {stats.cancelledCount}
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Giữ trong sổ để kiểm tra, số không tái sử dụng
              </div>
            </div>

            {/* 4. Chờ đồng bộ Excel */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs">
              <div className="flex items-center justify-between text-slate-500 mb-2">
                <span className="text-xs font-semibold uppercase tracking-wider text-slate-600">
                  Chờ đồng bộ Excel
                </span>
                <span className="p-2 bg-amber-50 text-amber-700 rounded-lg">
                  <TrendingUp className="w-4 h-4" />
                </span>
              </div>
              <div className="text-3xl font-extrabold text-amber-700 font-mono tracking-tight">
                {stats.pendingSyncCount}
              </div>
              <div className="text-xs text-slate-500 mt-1">
                Văn bản đang chờ đẩy lên bảng tính trực tuyến
              </div>
            </div>
          </div>

          {/* Section: Breakdown by Document Type */}
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs">
            <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
              <h2 className="text-base font-bold text-slate-900">
                Số Lượng Chi Tiết Theo Loại Văn Bản
              </h2>
              <span className="text-xs text-slate-500">
                Bao gồm văn bản đi theo từng loại và văn bản đến
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-4">
              {stats.byType.map((t) => {
                const isIncoming = t.typeCode === 'VBD';
                return (
                  <div
                    key={t.typeCode}
                    className={`p-3.5 sm:p-4 border rounded-lg flex items-center justify-between gap-2 min-w-0 ${
                      isIncoming
                        ? 'bg-indigo-50/60 border-indigo-200'
                        : 'bg-slate-50 border-slate-200'
                    }`}
                  >
                    <div className="min-w-0">
                      <span
                        className={`text-xs font-bold px-2 py-0.5 rounded font-mono ${
                          isIncoming
                            ? 'text-indigo-800 bg-indigo-100'
                            : 'text-sky-800 bg-sky-100'
                        }`}
                      >
                        {t.typeCode}
                      </span>
                      <h3 className="text-sm font-semibold text-slate-900 mt-1.5 break-words">
                        {t.typeName}
                      </h3>
                      <p className="text-xs text-slate-500 mt-0.5">
                        Đã hủy: <span className="font-mono text-rose-600 tabular-nums">{t.cancelledCount}</span>
                      </p>
                    </div>

                    <div className="text-right shrink-0">
                      <span className="text-2xl font-bold font-mono text-slate-900 tabular-nums">
                        {t.count}
                      </span>
                      <span className="block text-[11px] text-slate-400">văn bản</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Section: Monthly Breakdown */}
          <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs">
            <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
              <h2 className="text-base font-bold text-slate-900">
                Phân Bổ Theo 12 Tháng Trong Năm {stats.year}
              </h2>
              <span className="text-xs text-slate-500">Múi giờ Asia/Ho_Chi_Minh</span>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-3">
              {stats.byMonth.map((m) => {
                const percent = Math.round((m.count / maxMonthCount) * 100);
                return (
                  <div
                    key={m.month}
                    className="p-3 bg-slate-50 border border-slate-200 rounded-lg flex flex-col justify-between"
                  >
                    <div className="flex items-center justify-between text-xs text-slate-600 mb-2">
                      <span className="font-semibold">Tháng {m.month}</span>
                      <span className="font-mono font-bold text-sky-900">{m.count} VB</span>
                    </div>

                    {/* Progress visual bar */}
                    <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden mt-1">
                      <div
                        className="h-full bg-sky-700 rounded-full transition-all duration-300"
                        style={{ width: `${percent}%` }}
                      />
                    </div>

                    <div className="text-[11px] text-slate-400 mt-2 flex justify-between">
                      <span>Hiệu lực: {m.count - m.cancelledCount}</span>
                      {m.cancelledCount > 0 && (
                        <span className="text-rose-600 font-mono">Hủy: {m.cancelledCount}</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      ) : null}
    </div>
  );
};
