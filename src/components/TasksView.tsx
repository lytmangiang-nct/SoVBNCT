import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  AlertCircle,
  BookOpen,
  CalendarClock,
  CheckCircle2,
  Clock,
  ExternalLink,
  FileText,
  Filter,
  ListChecks,
  RefreshCw,
  Search,
  UserCheck,
} from 'lucide-react';
import { api } from '../api';
import type {
  DocumentTask,
  DocumentTaskStatus,
  RegistryBook,
  User,
} from '../types';
import { hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface TasksViewProps {
  currentUser: User;
  accessibleBooks: RegistryBook[];
  activeBookId: string;
  onOpenDocument: (docId: string) => void;
}

const STATUS_META: Record<
  DocumentTaskStatus,
  { label: string; badge: string }
> = {
  pending: {
    label: 'Chờ xử lý',
    badge: 'bg-amber-50 text-amber-900 border-amber-200',
  },
  in_progress: {
    label: 'Đang làm',
    badge: 'bg-sky-50 text-sky-900 border-sky-200',
  },
  done: {
    label: 'Hoàn thành',
    badge: 'bg-emerald-50 text-emerald-900 border-emerald-200',
  },
};

export const TasksView: React.FC<TasksViewProps> = ({
  currentUser,
  accessibleBooks,
  activeBookId,
  onOpenDocument,
}) => {
  const [scope, setScope] = useState<'mine' | 'all'>('mine');
  const [statusFilter, setStatusFilter] = useState<'all' | DocumentTaskStatus>('all');
  const [bookFilter, setBookFilter] = useState<string>('all');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [tasks, setTasks] = useState<DocumentTask[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [updatingTaskId, setUpdatingTaskId] = useState<string | null>(null);

  const loadTasks = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await api.getTasks({
        scope,
        bookId: bookFilter !== 'all' ? bookFilter : undefined,
        status: statusFilter,
      });
      setTasks(res.tasks || []);
    } catch (err: any) {
      setError(err?.message || 'Không thể tải danh sách việc cần làm.');
    } finally {
      setLoading(false);
    }
  }, [scope, bookFilter, statusFilter]);

  useEffect(() => {
    loadTasks();
  }, [loadTasks]);

  const filteredTasks = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    if (!q) return tasks;
    return tasks.filter((t) =>
      `${t.title} ${t.fullDocNumber} ${t.docName} ${t.recipient || ''} ${t.assigneeName} ${t.dueDate || ''}`
        .toLowerCase()
        .includes(q)
    );
  }, [tasks, searchTerm]);

  const counts = useMemo(() => {
    let pending = 0;
    let inProgress = 0;
    let done = 0;
    for (const t of tasks) {
      if (t.status === 'pending') pending++;
      else if (t.status === 'in_progress') inProgress++;
      else if (t.status === 'done') done++;
    }
    return { total: tasks.length, pending, inProgress, done };
  }, [tasks]);

  const canViewDoc = (bookId: string): boolean => {
    if (isSystemAdmin(currentUser)) return true;
    return hasBookPermission(currentUser, bookId, 'view');
  };

  const canEditStatus = (task: DocumentTask): boolean => {
    if (isSystemAdmin(currentUser)) return true;
    if (task.assigneeUserId === currentUser.id) return true;
    if (task.createdByUserId === currentUser.id) return true;
    return (
      hasBookPermission(currentUser, task.bookId, 'manage_book') ||
      hasBookPermission(currentUser, task.bookId, 'edit')
    );
  };

  const handleChangeStatus = async (task: DocumentTask, nextStatus: DocumentTaskStatus) => {
    if (task.status === nextStatus || updatingTaskId) return;
    setUpdatingTaskId(task.id);
    setError(null);
    try {
      const res = await api.updateTask(task.id, { status: nextStatus });
      setTasks((prev) => prev.map((t) => (t.id === task.id ? res.task : t)));
    } catch (err: any) {
      setError(err?.message || 'Không thể cập nhật trạng thái công việc.');
    } finally {
      setUpdatingTaskId(null);
    }
  };

  return (
    <div className="max-w-7xl mx-auto py-5 sm:py-8 px-3 sm:px-6 lg:px-8 space-y-5">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <ListChecks className="w-5 h-5 text-sky-700 shrink-0" />
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
              Việc của tôi &amp; Theo dõi thời hạn
            </h1>
          </div>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            Tổng hợp đầu việc và hạn nộp đã xác nhận từ các văn bản đến theo từng sổ.
          </p>
        </div>

        <button
          type="button"
          onClick={loadTasks}
          disabled={loading}
          className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-3.5 py-2 text-xs font-semibold bg-white hover:bg-slate-50 border border-slate-300 text-slate-700 rounded-lg shrink-0 self-start sm:self-center disabled:opacity-50"
        >
          <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
          <span>Làm mới</span>
        </button>
      </div>

      {/* KPI Strip */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs">
          <span className="text-xs font-medium text-slate-500 block">Tổng đầu việc</span>
          <span className="text-2xl font-bold font-mono text-slate-900 mt-1 block tabular-nums">
            {counts.total}
          </span>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs">
          <span className="text-xs font-medium text-amber-700 block">Chờ xử lý</span>
          <span className="text-2xl font-bold font-mono text-amber-800 mt-1 block tabular-nums">
            {counts.pending}
          </span>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs">
          <span className="text-xs font-medium text-sky-700 block">Đang thực hiện</span>
          <span className="text-2xl font-bold font-mono text-sky-800 mt-1 block tabular-nums">
            {counts.inProgress}
          </span>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs">
          <span className="text-xs font-medium text-emerald-700 block">Hoàn thành</span>
          <span className="text-2xl font-bold font-mono text-emerald-800 mt-1 block tabular-nums">
            {counts.done}
          </span>
        </div>
      </div>

      {/* Filter Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 shadow-xs space-y-3">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          {/* Scope Switcher */}
          <div className="inline-grid grid-cols-2 bg-slate-100 p-1 rounded-lg border border-slate-200 w-full sm:w-auto">
            <button
              type="button"
              onClick={() => setScope('mine')}
              className={`min-h-[38px] px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                scope === 'mine'
                  ? 'bg-white text-sky-900 shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Việc của tôi
            </button>
            <button
              type="button"
              onClick={() => setScope('all')}
              className={`min-h-[38px] px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors ${
                scope === 'all'
                  ? 'bg-white text-sky-900 shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Tất cả trong sổ có quyền
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2 flex-1 lg:justify-end">
            <div className="relative flex-1 min-w-[200px] sm:max-w-xs">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder="Tìm nội dung việc, số văn bản..."
                className="w-full min-h-[40px] pl-9 pr-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600"
              />
            </div>

            <div className="inline-flex items-center gap-1.5">
              <BookOpen className="w-3.5 h-3.5 text-slate-500 shrink-0" />
              <select
                aria-label="Lọc theo sổ"
                value={bookFilter}
                onChange={(e) => setBookFilter(e.target.value)}
                className="min-h-[40px] px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-800"
              >
                <option value="all">Tất cả sổ ({accessibleBooks.length})</option>
                {accessibleBooks.map((b) => (
                  <option key={b.id} value={b.id}>
                    {b.name} {b.id === activeBookId ? '(Đang chọn)' : ''}
                  </option>
                ))}
              </select>
            </div>

            <div className="inline-flex items-center gap-1.5">
              <Filter className="w-3.5 h-3.5 text-slate-500 shrink-0" />
              <select
                aria-label="Lọc trạng thái"
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value as 'all' | DocumentTaskStatus)}
                className="min-h-[40px] px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-800"
              >
                <option value="all">Tất cả trạng thái</option>
                <option value="pending">Chờ xử lý</option>
                <option value="in_progress">Đang làm</option>
                <option value="done">Hoàn thành</option>
              </select>
            </div>
          </div>
        </div>
      </div>

      {error && (
        <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-center gap-2">
          <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
          <span>{error}</span>
        </div>
      )}

      {/* Task List */}
      {loading ? (
        <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-xs text-slate-500">
          Đang tải danh sách công việc...
        </div>
      ) : filteredTasks.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-2">
          <CheckCircle2 className="w-8 h-8 text-slate-300 mx-auto" />
          <p className="text-sm font-semibold text-slate-700">
            Chưa có công việc nào khớp bộ lọc
          </p>
          <p className="text-xs text-slate-500">
            Khi nhập PDF có thời hạn thực hiện và bấm xác nhận, đầu việc sẽ hiển thị tại đây.
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {filteredTasks.map((task) => {
            const meta = STATUS_META[task.status];
            const editable = canEditStatus(task);
            const openable = canViewDoc(task.bookId);

            return (
              <div
                key={task.id}
                className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3"
              >
                <div className="flex flex-col md:flex-row md:items-start justify-between gap-3">
                  <div className="space-y-1.5 min-w-0 flex-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className={`text-[11px] font-bold px-2 py-0.5 rounded border ${meta.badge}`}>
                        {meta.label}
                      </span>
                      {task.dueDate ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-800 border border-slate-200 tabular-nums">
                          <CalendarClock className="w-3 h-3 text-sky-700" />
                          <span>Hạn: {task.dueDate}</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-[11px] font-bold px-2 py-0.5 rounded bg-amber-100 text-amber-900 border border-amber-300">
                          <Clock className="w-3 h-3" />
                          <span>
                            Cần xác nhận{task.rawDeadlineText ? `: ${task.rawDeadlineText}` : ''}
                          </span>
                        </span>
                      )}
                      {task.bookName && (
                        <span className="text-[11px] font-medium text-slate-600 bg-slate-50 px-2 py-0.5 rounded border border-slate-200">
                          {task.bookName}
                        </span>
                      )}
                    </div>

                    <h3 className="text-sm font-bold text-slate-900 break-words">
                      {task.title}
                    </h3>

                    <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-600">
                      <span className="inline-flex items-center gap-1">
                        <UserCheck className="w-3.5 h-3.5 text-sky-700 shrink-0" />
                        <span>
                          Phụ trách: <strong className="text-slate-900">{task.assigneeName}</strong> (@{task.assigneeUsername})
                        </span>
                      </span>
                      {task.recipient && (
                        <span>
                          Nơi nhận: <strong className="text-slate-800">{task.recipient}</strong>
                        </span>
                      )}
                    </div>

                    <div className="flex flex-wrap items-center gap-2 text-xs pt-0.5">
                      <span className="inline-flex items-center gap-1 text-slate-600">
                        <FileText className="w-3.5 h-3.5 text-purple-700 shrink-0" />
                        <span>Văn bản nguồn:</span>
                        <strong className="font-mono text-slate-900">{task.fullDocNumber}</strong>
                        <span className="truncate max-w-md">— {task.docName}</span>
                      </span>
                      {openable && (
                        <button
                          type="button"
                          onClick={() => onOpenDocument(task.docId)}
                          className="inline-flex items-center gap-1 text-xs font-semibold text-sky-700 hover:underline shrink-0"
                        >
                          <span>Mở hồ sơ</span>
                          <ExternalLink className="w-3 h-3" />
                        </button>
                      )}
                    </div>

                    {task.sourceExcerpt && (
                      <p className="text-[11px] text-slate-600 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5 break-words">
                        <strong className="text-slate-700">
                          Trích đoạn{task.sourcePage ? ` (Trang ${task.sourcePage})` : ''}:
                        </strong>{' '}
                        {task.sourceExcerpt}
                      </p>
                    )}
                  </div>

                  {editable && (
                    <div className="flex flex-wrap sm:flex-nowrap items-center gap-1.5 shrink-0">
                      {(['pending', 'in_progress', 'done'] as DocumentTaskStatus[]).map((st) => (
                        <button
                          key={st}
                          type="button"
                          disabled={updatingTaskId === task.id}
                          onClick={() => handleChangeStatus(task, st)}
                          className={`min-h-[36px] px-3 py-1.5 text-xs font-semibold rounded-lg border transition-colors whitespace-nowrap ${
                            task.status === st
                              ? st === 'done'
                                ? 'bg-emerald-700 text-white border-emerald-700'
                                : st === 'in_progress'
                                ? 'bg-sky-700 text-white border-sky-700'
                                : 'bg-amber-600 text-white border-amber-600'
                              : 'bg-white hover:bg-slate-50 text-slate-700 border-slate-300'
                          }`}
                        >
                          {STATUS_META[st].label}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
