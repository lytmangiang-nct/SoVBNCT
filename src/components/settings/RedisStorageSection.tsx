import React, { useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Database,
  Download,
  FileUp,
  Info,
  RefreshCw,
  Server,
  ShieldAlert,
  ShieldCheck,
  Upload,
} from 'lucide-react';
import { api } from '../../api';
import type {
  LegacySnapshotSummary,
  PersistenceStatus,
  StorageConnectionState,
  StorageTestResult,
} from '../../types';

interface RedisStorageSectionProps {
  persistenceStatus: PersistenceStatus | null;
  onStatusUpdated: (nextStatus: PersistenceStatus) => void;
  onFullRefresh: () => void;
}

const STATE_BADGE_MAP: Record<
  StorageConnectionState,
  { label: string; badgeClass: string; boxClass: string }
> = {
  connected: {
    label: 'Đã kết nối',
    badgeClass: 'bg-emerald-50 text-emerald-800 border-emerald-300',
    boxClass: 'bg-emerald-50/60 border-emerald-200 text-emerald-950',
  },
  missing_config: {
    label: 'Thiếu cấu hình',
    badgeClass: 'bg-rose-50 text-rose-800 border-rose-300',
    boxClass: 'bg-rose-50 border-rose-200 text-rose-950',
  },
  conflict: {
    label: 'Cấu hình mâu thuẫn',
    badgeClass: 'bg-rose-50 text-rose-800 border-rose-300',
    boxClass: 'bg-rose-50 border-rose-200 text-rose-950',
  },
  connection_error: {
    label: 'Lỗi kết nối',
    badgeClass: 'bg-rose-50 text-rose-800 border-rose-300',
    boxClass: 'bg-rose-50 border-rose-200 text-rose-950',
  },
  awaiting_migration: {
    label: 'Chờ chuyển dữ liệu',
    badgeClass: 'bg-amber-50 text-amber-900 border-amber-300',
    boxClass: 'bg-amber-50 border-amber-200 text-amber-950',
  },
  corrupt_data: {
    label: 'Dữ liệu cần kiểm tra',
    badgeClass: 'bg-rose-50 text-rose-800 border-rose-300',
    boxClass: 'bg-rose-50 border-rose-200 text-rose-950',
  },
};

function summarizeUploadedSnapshot(raw: any): LegacySnapshotSummary | null {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null;
  const books = Array.isArray(raw.books) ? raw.books : [];
  const docs = Array.isArray(raw.documents) ? raw.documents : [];
  const users = Array.isArray(raw.users) ? raw.users : [];
  if (books.length === 0 && docs.length === 0 && users.length === 0 && !raw.settings) return null;

  let startMilestonesCount = 0;
  for (const b of books) {
    if (b?.customStartSequences && typeof b.customStartSequences === 'object') {
      startMilestonesCount += Object.keys(b.customStartSequences).length;
    }
  }
  if (startMilestonesCount === 0 && raw.settings?.customStartSequences) {
    startMilestonesCount = Object.keys(raw.settings.customStartSequences).length;
  }
  return {
    usersCount: users.length,
    booksCount: books.length,
    documentsCount: docs.filter((d: any) => d?.status !== 'deleted').length,
    deletedDocumentsCount: docs.filter((d: any) => d?.status === 'deleted').length,
    startMilestonesCount,
    driveLinkedCount: docs.filter((d: any) => Boolean(d?.driveFileId || d?.driveFileUrl)).length,
    bookNames: books.map((b: any) => String(b?.name || b?.code || '')).filter(Boolean),
  };
}

export const RedisStorageSection: React.FC<RedisStorageSectionProps> = ({
  persistenceStatus,
  onStatusUpdated,
  onFullRefresh,
}) => {
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<StorageTestResult | null>(null);
  const [actionBusy, setActionBusy] = useState(false);
  const [actionFeedback, setActionFeedback] = useState<{
    success: boolean;
    message: string;
  } | null>(null);

  const [initConfirmText, setInitConfirmText] = useState('');
  const [initChecked, setInitChecked] = useState(false);

  const [uploadedFileName, setUploadedFileName] = useState<string>('');
  const [uploadedSnapshot, setUploadedSnapshot] = useState<Record<string, any> | null>(null);
  const [uploadedSummary, setUploadedSummary] = useState<LegacySnapshotSummary | null>(null);
  const [allowOverwriteCorrupt, setAllowOverwriteCorrupt] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const status = testResult?.persistenceStatus || persistenceStatus;
  const connState: StorageConnectionState = status?.connectionState || 'missing_config';
  const badgeMeta = STATE_BADGE_MAP[connState] || STATE_BADGE_MAP.missing_config;

  const handleTestConnection = async () => {
    setTesting(true);
    setActionFeedback(null);
    try {
      const res = await api.testStorageConnection();
      setTestResult(res);
      if (res.persistenceStatus) {
        onStatusUpdated(res.persistenceStatus);
      }
      if (res.success && res.keyExists) {
        onFullRefresh();
      }
    } catch (err: any) {
      setTestResult({
        success: false,
        readOk: false,
        writeOk: false,
        deleteOk: false,
        keyExists: false,
        latencyMs: 0,
        message: err?.message || 'Không thể kiểm tra kết nối lưu trữ từ máy chủ.',
        persistenceStatus:
          status ||
          ({
            durable: false,
            storageMode: 'ephemeral_tmp',
            storageLabel: 'Chưa xác định',
            connectionState: 'connection_error',
            connectionLabel: 'Lỗi kết nối',
          } as PersistenceStatus),
      });
    } finally {
      setTesting(false);
    }
  };

  const handleExportLegacySnapshot = async () => {
    setActionFeedback(null);
    try {
      const res = await api.exportLegacySnapshot();
      const blob = new Blob([JSON.stringify(res.snapshot, null, 2)], {
        type: 'application/json;charset=utf-8',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `sao-luu-so-van-ban-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setActionFeedback({
        success: true,
        message: `Đã tải xuống bản sao đối chiếu (${res.summary.booksCount} sổ, ${res.summary.documentsCount} văn bản, ${res.summary.startMilestonesCount} mốc số).`,
      });
    } catch (err: any) {
      setActionFeedback({
        success: false,
        message: err?.message || 'Không thể xuất bản sao dữ liệu cũ.',
      });
    }
  };

  const handleImportBundledLegacy = async () => {
    setActionBusy(true);
    setActionFeedback(null);
    try {
      const res = await api.importBundledLegacyToStorage(connState === 'corrupt_data' && allowOverwriteCorrupt);
      onStatusUpdated(res.persistenceStatus);
      setActionFeedback({ success: true, message: res.message });
      onFullRefresh();
    } catch (err: any) {
      if (err?.persistenceStatus) onStatusUpdated(err.persistenceStatus);
      setActionFeedback({
        success: false,
        message: err?.message || 'Không thể nhập dữ liệu nguồn cũ vào Upstash Redis.',
      });
    } finally {
      setActionBusy(false);
    }
  };

  const handleFileSelected = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    setActionFeedback(null);
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const summary = summarizeUploadedSnapshot(parsed);
      if (!summary) {
        setUploadedFileName('');
        setUploadedSnapshot(null);
        setUploadedSummary(null);
        setActionFeedback({
          success: false,
          message: 'Tệp JSON không đúng cấu trúc sao lưu của Sổ cấp số văn bản.',
        });
        return;
      }
      setUploadedFileName(file.name);
      setUploadedSnapshot(parsed);
      setUploadedSummary(summary);
    } catch {
      setUploadedFileName('');
      setUploadedSnapshot(null);
      setUploadedSummary(null);
      setActionFeedback({
        success: false,
        message: 'Không đọc được tệp JSON. Vui lòng chọn đúng tệp sao lưu hợp lệ.',
      });
    }
  };

  const handleImportUploadedFile = async () => {
    if (!uploadedSnapshot) return;
    setActionBusy(true);
    setActionFeedback(null);
    try {
      const res = await api.importSnapshotToStorage(
        uploadedSnapshot,
        connState === 'corrupt_data' && allowOverwriteCorrupt
      );
      onStatusUpdated(res.persistenceStatus);
      setUploadedSnapshot(null);
      setUploadedSummary(null);
      setUploadedFileName('');
      if (fileInputRef.current) fileInputRef.current.value = '';
      setActionFeedback({ success: true, message: res.message });
      onFullRefresh();
    } catch (err: any) {
      if (err?.persistenceStatus) onStatusUpdated(err.persistenceStatus);
      setActionFeedback({
        success: false,
        message: err?.message || 'Không thể nhập tệp sao lưu vào Upstash Redis.',
      });
    } finally {
      setActionBusy(false);
    }
  };

  const handleInitializeNew = async () => {
    if (!initChecked || initConfirmText.trim() !== 'XÁC NHẬN HỆ THỐNG MỚI') {
      setActionFeedback({
        success: false,
        message: 'Vui lòng đánh dấu xác nhận và nhập chính xác "XÁC NHẬN HỆ THỐNG MỚI".',
      });
      return;
    }
    setActionBusy(true);
    setActionFeedback(null);
    try {
      const res = await api.initializeNewStorage(initConfirmText.trim());
      onStatusUpdated(res.persistenceStatus);
      setInitConfirmText('');
      setInitChecked(false);
      setActionFeedback({ success: true, message: res.message });
      onFullRefresh();
    } catch (err: any) {
      if (err?.persistenceStatus) onStatusUpdated(err.persistenceStatus);
      setActionFeedback({
        success: false,
        message: err?.message || 'Không thể khởi tạo dữ liệu mới trên Upstash Redis.',
      });
    } finally {
      setActionBusy(false);
    }
  };

  const showSetupGuide =
    connState === 'missing_config' ||
    connState === 'conflict' ||
    connState === 'connection_error' ||
    status?.storageMode === 'ephemeral_tmp';

  return (
    <div className="space-y-5 max-w-3xl min-w-0">
      {/* 1. Connection Overview & Live Test */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="min-w-0 flex-1">
            <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Database className="w-5 h-5 text-sky-700 shrink-0" />
              <span>Lưu Trữ Chính — Upstash Redis (Vercel Production)</span>
            </h2>
            <p className="text-xs text-slate-600 mt-1 break-words">
              Toàn bộ sổ văn bản, tài khoản, mốc số bắt đầu, lịch sử thao tác và liên kết Drive được lưu tập trung trên máy chủ.
            </p>
          </div>
          <span
            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-bold border shrink-0 ${badgeMeta.badgeClass}`}
          >
            {connState === 'connected' ? (
              <ShieldCheck className="w-3.5 h-3.5 shrink-0" />
            ) : (
              <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
            )}
            <span>{status?.connectionLabel || badgeMeta.label}</span>
          </span>
        </div>

        {/* Metadata Grid (Secrets never rendered) */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs bg-slate-50 border border-slate-200 rounded-lg p-3.5">
          <div className="space-y-0.5 min-w-0">
            <span className="text-[11px] font-semibold text-slate-500 uppercase">Chế độ lưu trữ</span>
            <p className="font-semibold text-slate-900 break-words">
              {status?.storageLabel || 'Đang kiểm tra...'}
            </p>
          </div>
          <div className="space-y-0.5 min-w-0">
            <span className="text-[11px] font-semibold text-slate-500 uppercase">Cặp biến môi trường</span>
            <p className="font-mono text-slate-800 break-all">
              {status?.envPairUsed || 'NONE'}
              {status?.maskedHost ? ` · ${status.maskedHost}` : ''}
            </p>
          </div>
          <div className="space-y-0.5 min-w-0">
            <span className="text-[11px] font-semibold text-slate-500 uppercase">Khóa dữ liệu (PERSISTENT_DB_KEY)</span>
            <p className="font-mono text-slate-800 break-all">
              {status?.dbKey || 'nct_vanban_registry_db_v1 (cục bộ nếu chưa gắn Redis)'}
            </p>
          </div>
          <div className="space-y-0.5 min-w-0">
            <span className="text-[11px] font-semibold text-slate-500 uppercase">Phiên bản cấu trúc &amp; Bảo mật phiên</span>
            <p className="text-slate-800">
              Schema v{status?.schemaVersion || 2} ·{' '}
              {status?.sessionSecretConfigured ? (
                <span className="text-emerald-700 font-semibold">SESSION_SECRET đã đặt</span>
              ) : (
                <span className="text-amber-700 font-medium">SESSION_SECRET mặc định ổn định</span>
              )}
            </p>
          </div>
          {(status?.lastSavedAt || status?.lastVerifiedAt) && (
            <div className="sm:col-span-2 pt-2 border-t border-slate-200 flex flex-wrap gap-x-6 gap-y-1 text-[11px] text-slate-600">
              {status.lastSavedAt && (
                <span>
                  Lần ghi gần nhất:{' '}
                  <strong className="text-slate-800">
                    {new Date(status.lastSavedAt).toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' })}
                  </strong>
                </span>
              )}
              {status.lastVerifiedAt && (
                <span>
                  Xác thực đọc lại:{' '}
                  <strong className="text-slate-800">
                    {new Date(status.lastVerifiedAt).toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' })}
                  </strong>
                </span>
              )}
            </div>
          )}
        </div>

        {/* Status detail box */}
        {status?.errorDetail && connState !== 'connected' && (
          <div className={`p-3 rounded-lg border text-xs flex items-start gap-2.5 ${badgeMeta.boxClass}`}>
            <ShieldAlert className="w-4 h-4 shrink-0 mt-0.5" />
            <div className="space-y-1 min-w-0">
              <strong className="block font-bold">{badgeMeta.label}</strong>
              <p className="break-words">{status.errorDetail}</p>
            </div>
          </div>
        )}

        {/* Test Connection Button */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
          <p className="text-[11px] text-slate-500">
            Thử ghi và đọc lại một khóa kiểm tra riêng có thời hạn 30 giây (tách biệt hoàn toàn với dữ liệu sổ).
          </p>
          <button
            type="button"
            onClick={handleTestConnection}
            disabled={testing}
            className="inline-flex items-center justify-center gap-2 min-h-[42px] px-4 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg shadow-xs transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${testing ? 'animate-spin' : ''}`} />
            <span>{testing ? 'Đang kiểm tra kết nối...' : 'Kiểm tra kết nối'}</span>
          </button>
        </div>

        {/* Live Test Output */}
        {testResult && (
          <div
            className={`p-3 rounded-lg border text-xs space-y-1.5 ${
              testResult.success
                ? 'bg-emerald-50 border-emerald-200 text-emerald-950'
                : 'bg-rose-50 border-rose-200 text-rose-950'
            }`}
          >
            <div className="flex items-start gap-2">
              {testResult.success ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
              ) : (
                <AlertTriangle className="w-4 h-4 text-rose-700 shrink-0 mt-0.5" />
              )}
              <div className="min-w-0 space-y-1">
                <p className="font-semibold break-words">{testResult.message}</p>
                <div className="flex flex-wrap gap-x-4 gap-y-1 text-[11px]">
                  <span>Ghi khóa thử: {testResult.writeOk ? 'Đạt' : 'Thất bại'}</span>
                  <span>Đọc khóa thử: {testResult.readOk ? 'Đạt' : 'Thất bại'}</span>
                  <span>Dọn khóa thử: {testResult.deleteOk ? 'Đạt' : 'Tự hết hạn 30s'}</span>
                  <span>Độ trễ: {testResult.latencyMs} ms</span>
                </div>
              </div>
            </div>
          </div>
        )}

        {actionFeedback && (
          <div
            className={`p-3 rounded-lg border text-xs flex items-start gap-2 ${
              actionFeedback.success
                ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                : 'bg-rose-50 border-rose-200 text-rose-900'
            }`}
          >
            {actionFeedback.success ? (
              <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0 mt-0.5" />
            ) : (
              <AlertTriangle className="w-4 h-4 text-rose-700 shrink-0 mt-0.5" />
            )}
            <span className="break-words">{actionFeedback.message}</span>
          </div>
        )}
      </div>

      {/* 2. Awaiting Migration / Corrupt Data Recovery Panel */}
      {(connState === 'awaiting_migration' || connState === 'corrupt_data') && (
        <div className="bg-white border-2 border-amber-300 rounded-xl p-4 sm:p-6 shadow-xs space-y-5">
          <div className="space-y-1">
            <h3 className="text-sm font-bold text-amber-950 flex items-center gap-2">
              <Server className="w-4 h-4 text-amber-700 shrink-0" />
              <span>
                {connState === 'awaiting_migration'
                  ? 'Chờ chuyển dữ liệu — Khóa Redis đích đang trống'
                  : 'Dữ liệu cần kiểm tra — Khóa Redis không đọc được'}
              </span>
            </h3>
            <p className="text-xs text-slate-700 break-words">
              {connState === 'awaiting_migration'
                ? 'Để tránh ghi đè mất dữ liệu cũ, hệ thống không tự tạo dữ liệu trắng khi khóa Redis chưa tồn tại. Nếu đang chuyển từ hệ thống cũ sang Redis, hãy đối chiếu và nhập bản sao dữ liệu trước khi mở lại cấp số.'
                : 'Khóa dữ liệu hiện tại trên Redis không đúng định dạng JSON hợp lệ. Hệ thống giữ nguyên dữ liệu để tránh mất mát. Bạn có thể tải lại từ bản sao hợp lệ nếu xác nhận ghi đè.'}
            </p>
          </div>

          {connState === 'corrupt_data' && (
            <label className="flex items-start gap-2 p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-900 cursor-pointer">
              <input
                type="checkbox"
                checked={allowOverwriteCorrupt}
                onChange={(e) => setAllowOverwriteCorrupt(e.target.checked)}
                className="mt-0.5 rounded text-rose-700"
              />
              <span>
                Tôi đã kiểm tra trên Upstash Console và xác nhận cho phép ghi đè khóa bị hỏng bằng bản sao hợp lệ bên dưới.
              </span>
            </label>
          )}

          {/* Option A1: Legacy bundled data on server */}
          {status?.hasLegacyBundledSnapshot && status.legacySnapshotSummary && (
            <div className="p-4 bg-sky-50/70 border border-sky-200 rounded-lg space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <h4 className="text-xs font-bold text-sky-950">
                  1A. Nguồn dữ liệu cũ phát hiện trên máy chủ (đối chiếu trước khi nhập)
                </h4>
                <button
                  type="button"
                  onClick={handleExportLegacySnapshot}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-white hover:bg-sky-50 text-sky-800 border border-sky-300 rounded-md text-[11px] font-semibold"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Tải bản sao JSON</span>
                </button>
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 text-xs">
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">Số sổ văn bản</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.booksCount}</strong>
                </div>
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">Tài khoản</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.usersCount}</strong>
                </div>
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">Văn bản hiện hành/hủy</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.documentsCount}</strong>
                </div>
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">VB đã xóa</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.deletedDocumentsCount}</strong>
                </div>
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">Mốc số bắt đầu</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.startMilestonesCount}</strong>
                </div>
                <div className="bg-white border border-sky-100 rounded px-2.5 py-2">
                  <span className="text-[10px] text-slate-500 block">Tệp Drive liên kết</span>
                  <strong className="text-slate-900">{status.legacySnapshotSummary.driveLinkedCount}</strong>
                </div>
              </div>
              <button
                type="button"
                disabled={actionBusy || (connState === 'corrupt_data' && !allowOverwriteCorrupt)}
                onClick={handleImportBundledLegacy}
                className="inline-flex items-center gap-2 min-h-[42px] px-4 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg"
              >
                <Upload className="w-3.5 h-3.5" />
                <span>{actionBusy ? 'Đang nhập vào Redis...' : 'Chuyển dữ liệu nguồn cũ vào Upstash Redis'}</span>
              </button>
            </div>
          )}

          {/* Option A2: Upload JSON backup file */}
          <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-3">
            <h4 className="text-xs font-bold text-slate-900">
              1B. Nhập từ tệp sao lưu JSON đã xuất trước đó
            </h4>
            <div className="flex flex-wrap items-center gap-2.5">
              <input
                ref={fileInputRef}
                type="file"
                accept="application/json,.json"
                onChange={handleFileSelected}
                className="block text-xs text-slate-700 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border file:border-slate-300 file:text-xs file:font-semibold file:bg-white file:text-slate-800 hover:file:bg-slate-100"
              />
            </div>
            {uploadedSummary && (
              <div className="p-3 bg-white border border-slate-200 rounded-lg space-y-2 text-xs">
                <p className="font-semibold text-slate-900">
                  Đối chiếu tệp: <span className="font-mono">{uploadedFileName}</span>
                </p>
                <p className="text-slate-600">
                  {uploadedSummary.booksCount} sổ ({uploadedSummary.bookNames.join(', ')}) ·{' '}
                  {uploadedSummary.usersCount} tài khoản · {uploadedSummary.documentsCount} văn bản ·{' '}
                  {uploadedSummary.startMilestonesCount} mốc số · {uploadedSummary.driveLinkedCount} tệp Drive
                </p>
                <button
                  type="button"
                  disabled={actionBusy || (connState === 'corrupt_data' && !allowOverwriteCorrupt)}
                  onClick={handleImportUploadedFile}
                  className="inline-flex items-center gap-2 min-h-[40px] px-4 py-2 bg-emerald-700 hover:bg-emerald-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg"
                >
                  <FileUp className="w-3.5 h-3.5" />
                  <span>{actionBusy ? 'Đang nhập...' : 'Xác nhận nhập tệp này vào Redis'}</span>
                </button>
              </div>
            )}
          </div>

          {/* Option B: Fresh empty system confirmation (only when key does not exist) */}
          {connState === 'awaiting_migration' && (
            <div className="p-4 bg-white border border-slate-300 rounded-lg space-y-3">
              <h4 className="text-xs font-bold text-slate-900">
                2. Chỉ dùng khi đây là hệ thống mới hoàn toàn (chưa có dữ liệu cũ)
              </h4>
              <label className="flex items-start gap-2 text-xs text-slate-700 cursor-pointer">
                <input
                  type="checkbox"
                  checked={initChecked}
                  onChange={(e) => setInitChecked(e.target.checked)}
                  className="mt-0.5 rounded text-sky-700"
                />
                <span>
                  Tôi xác nhận đây là hệ thống mới, không cần chuyển dữ liệu cũ và đồng ý tạo sổ mặc định (Đoàn trường, Chi bộ, Nhà trường) trên Redis.
                </span>
              </label>
              <div className="flex flex-col sm:flex-row gap-2">
                <input
                  type="text"
                  value={initConfirmText}
                  onChange={(e) => setInitConfirmText(e.target.value)}
                  placeholder='Gõ chính xác: XÁC NHẬN HỆ THỐNG MỚI'
                  className="flex-1 min-h-[42px] px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs font-mono"
                />
                <button
                  type="button"
                  disabled={
                    actionBusy ||
                    !initChecked ||
                    initConfirmText.trim() !== 'XÁC NHẬN HỆ THỐNG MỚI'
                  }
                  onClick={handleInitializeNew}
                  className="min-h-[42px] px-4 py-2 bg-slate-900 hover:bg-slate-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg shrink-0"
                >
                  {actionBusy ? 'Đang khởi tạo...' : 'Khởi tạo hệ thống mới'}
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* 3. Setup Guide on Vercel Production */}
      {showSetupGuide && (
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 sm:p-5 space-y-3 text-xs text-slate-800">
          <div className="flex items-start gap-2">
            <Info className="w-4 h-4 text-sky-700 shrink-0 mt-0.5" />
            <div className="space-y-2 min-w-0">
              <strong className="block font-bold text-slate-900">
                Hướng dẫn gắn Upstash Redis cho project trên Vercel
              </strong>
              <p className="font-semibold text-sky-900 bg-sky-50 border border-sky-200 rounded-lg px-3 py-2">
                Vercel → Marketplace → Upstash for Redis → Kết nối đúng project → Redeploy
              </p>
              <p className="text-slate-600">
                Biến môi trường Production cần có (chỉ dùng 1 cặp URL + Token thuộc cùng một database):
              </p>
              <ul className="list-disc pl-5 space-y-1 font-mono text-[11px] text-slate-800">
                <li>UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN (ưu tiên)</li>
                <li>Hoặc KV_REST_API_URL + KV_REST_API_TOKEN</li>
                <li>PERSISTENT_DB_KEY (tùy chọn, mặc định: nct_vanban_registry_db_v1)</li>
                <li>SESSION_SECRET (chuỗi bí mật cố định để giữ phiên đăng nhập giữa các bản triển khai)</li>
              </ul>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
