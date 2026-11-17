import React, { useCallback, useEffect, useState } from 'react';
import {
  FileText,
  Plus,
  Trash2,
  Upload,
  CheckCircle2,
  AlertTriangle,
  Copy,
  Check,
  RefreshCw,
} from 'lucide-react';
import { api } from '../../api';
import type { DocumentTypeConfig, RegistryBook, WordTemplate } from '../../types';

interface WordTemplatesSectionProps {
  selectedBook: RegistryBook;
  enabledTypes: DocumentTypeConfig[];
  canManage: boolean;
  isBlocked: boolean;
  onRequestConfirmDelete: (req: {
    title: string;
    description: string;
    confirmLabel: string;
    onConfirm: () => Promise<void>;
  }) => void;
}

const DEFAULT_PLACEHOLDERS: Array<{ key: string; label: string }> = [
  { key: 'so_ky_hieu', label: 'Số và ký hiệu văn bản' },
  { key: 'don_vi', label: 'Tên đơn vị / bộ phận' },
  { key: 'ten_van_ban', label: 'Tên văn bản' },
  { key: 'trich_yeu', label: 'Trích yếu / Nội dung' },
  { key: 'ngay_ban_hanh', label: 'Ngày ban hành (dd/MM/yyyy)' },
  { key: 'nguoi_ky', label: 'Họ và tên người ký' },
  { key: 'chuc_vu', label: 'Chức vụ người ký' },
];

export const WordTemplatesSection: React.FC<WordTemplatesSectionProps> = ({
  selectedBook,
  enabledTypes,
  canManage,
  isBlocked,
  onRequestConfirmDelete,
}) => {
  const [templates, setTemplates] = useState<WordTemplate[]>([]);
  const [supportedPlaceholders, setSupportedPlaceholders] = useState(DEFAULT_PLACEHOLDERS);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [typeCode, setTypeCode] = useState<string>(enabledTypes[0]?.code || 'KH');
  const [name, setName] = useState('');
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);
  const [formSuccess, setFormSuccess] = useState<string | null>(null);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  useEffect(() => {
    if (enabledTypes.length > 0 && !enabledTypes.some((t) => t.code === typeCode)) {
      setTypeCode(enabledTypes[0].code);
    }
  }, [enabledTypes, typeCode]);

  const loadTemplates = useCallback(async () => {
    setLoading(true);
    setLoadError(null);
    try {
      const res = await api.getWordTemplates(selectedBook.id);
      setTemplates(res.templates || []);
      if (Array.isArray(res.supportedPlaceholders) && res.supportedPlaceholders.length > 0) {
        setSupportedPlaceholders(res.supportedPlaceholders);
      }
    } catch (err: any) {
      setLoadError(err?.message || 'Không tải được danh sách mẫu Word.');
    } finally {
      setLoading(false);
    }
  }, [selectedBook.id]);

  useEffect(() => {
    loadTemplates();
  }, [loadTemplates]);

  const handleUpload = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManage || isBlocked) return;
    if (!selectedFile) {
      setFormError('Vui lòng chọn tệp mẫu Word (.docx).');
      return;
    }
    setSaving(true);
    setFormError(null);
    setFormSuccess(null);
    try {
      const res = await api.createWordTemplate({
        bookId: selectedBook.id,
        typeCode,
        name: name.trim() || selectedFile.name.replace(/\.docx$/i, ''),
        file: selectedFile,
      });
      setTemplates((prev) => [res.template, ...prev]);
      setName('');
      setSelectedFile(null);
      setFormSuccess(
        `Đã lưu mẫu "${res.template.name}" (${res.placeholders.length} trường giữ chỗ được nhận diện).`
      );
      setTimeout(() => setFormSuccess(null), 5000);
    } catch (err: any) {
      setFormError(err?.message || 'Không thể tải lên mẫu Word.');
    } finally {
      setSaving(false);
    }
  };

  const handleCopyPlaceholder = async (key: string) => {
    const token = `{{${key}}}`;
    try {
      await navigator.clipboard.writeText(token);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(null), 2000);
    } catch {
      // ignore clipboard errors
    }
  };

  const handleDeleteTemplate = (tpl: WordTemplate) => {
    onRequestConfirmDelete({
      title: 'Xóa mẫu Word (.docx)',
      description: `Bạn có chắc chắn muốn xóa mẫu "${tpl.name}" (${tpl.typeCode}) khỏi ${selectedBook.name}? Các văn bản đã cấp số trước đây vẫn được giữ nguyên.`,
      confirmLabel: 'Xóa mẫu',
      onConfirm: async () => {
        await api.deleteWordTemplate(tpl.id);
        setTemplates((prev) => prev.filter((t) => t.id !== tpl.id));
      },
    });
  };

  return (
    <div className="space-y-6 max-w-4xl">
      {/* Header & Placeholder Reference */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Mẫu Word (.docx) — {selectedBook.name}
            </h2>
            <p className="text-xs text-slate-600 mt-1">
              Tải lên mẫu .docx có sẵn các trường giữ chỗ. Khi cấp số chính thức, hệ thống tự điền thông tin và lưu bản hoàn chỉnh lên Google Drive.
            </p>
          </div>
          <button
            type="button"
            onClick={loadTemplates}
            disabled={loading}
            className="inline-flex items-center gap-1.5 h-9 px-3 text-xs font-semibold text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />
            <span>Làm mới</span>
          </button>
        </div>

        <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 space-y-2">
          <p className="text-[11px] font-bold uppercase tracking-wider text-slate-600">
            Trường giữ chỗ hỗ trợ trong tệp .docx (bấm để sao chép)
          </p>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
            {supportedPlaceholders.map((p) => (
              <button
                key={p.key}
                type="button"
                onClick={() => handleCopyPlaceholder(p.key)}
                className="flex items-center justify-between gap-2 px-2.5 py-2 bg-white border border-slate-200 hover:border-sky-400 rounded-lg text-left transition-colors"
              >
                <div className="min-w-0">
                  <code className="text-xs font-mono font-bold text-sky-800 block truncate">
                    {`{{${p.key}}}`}
                  </code>
                  <span className="text-[11px] text-slate-500 block truncate">{p.label}</span>
                </div>
                {copiedKey === p.key ? (
                  <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                ) : (
                  <Copy className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                )}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Upload Form */}
      {canManage && (
        <form
          onSubmit={handleUpload}
          className="bg-white border border-slate-200 rounded-xl p-4 sm:p-6 shadow-xs space-y-4 text-xs"
        >
          <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
            <Plus className="w-4 h-4 text-sky-700 shrink-0" />
            <span>Thêm mẫu Word mới cho {selectedBook.name}</span>
          </h3>

          <div className="grid grid-cols-1 sm:grid-cols-12 gap-3">
            <div className="sm:col-span-4 space-y-1">
              <label className="font-bold text-slate-700">Loại văn bản *</label>
              <select
                value={typeCode}
                onChange={(e) => setTypeCode(e.target.value)}
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs font-medium"
              >
                {enabledTypes.map((t) => (
                  <option key={t.code} value={t.code}>
                    {t.code} — {t.name}
                  </option>
                ))}
              </select>
            </div>

            <div className="sm:col-span-8 space-y-1">
              <label className="font-bold text-slate-700">Tên mẫu hiển thị</label>
              <input
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Ví dụ: Mẫu Kế hoạch hoạt động Đoàn trường"
                className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
              />
            </div>

            <div className="sm:col-span-9 space-y-1">
              <label className="font-bold text-slate-700">Tệp mẫu Word (.docx, tối đa 5 MB) *</label>
              <label className="flex items-center gap-2 min-h-[44px] px-3 py-2 bg-slate-50 hover:bg-slate-100 border border-dashed border-slate-300 rounded-lg cursor-pointer transition-colors">
                <Upload className="w-4 h-4 text-sky-700 shrink-0" />
                <span className="truncate text-slate-700 font-medium">
                  {selectedFile
                    ? `${selectedFile.name} (${(selectedFile.size / 1024).toFixed(0)} KB)`
                    : 'Chọn tệp .docx từ máy...'}
                </span>
                <input
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  onChange={(e) => {
                    const f = e.target.files?.[0] || null;
                    setSelectedFile(f);
                    if (f && !name.trim()) {
                      setName(f.name.replace(/\.docx$/i, ''));
                    }
                  }}
                  className="hidden"
                />
              </label>
            </div>

            <div className="sm:col-span-3 flex items-end">
              <button
                type="submit"
                disabled={saving || isBlocked || !selectedFile}
                className="w-full min-h-[44px] h-11 sm:h-10 px-4 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg transition-colors"
              >
                {saving ? 'Đang tải...' : 'Lưu mẫu'}
              </button>
            </div>
          </div>

          {formError && (
            <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-800 rounded-lg flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
              <span>{formError}</span>
            </div>
          )}

          {formSuccess && (
            <div className="p-2.5 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-lg flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{formSuccess}</span>
            </div>
          )}
        </form>
      )}

      {/* Existing Templates List */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
        <div className="p-4 border-b border-slate-100 flex items-center justify-between">
          <h3 className="text-sm font-bold text-slate-900">
            Danh sách mẫu đang lưu ({templates.length})
          </h3>
        </div>

        {loadError && (
          <div className="p-4 bg-rose-50 text-rose-800 text-xs">{loadError}</div>
        )}

        {templates.length === 0 ? (
          <div className="p-8 text-center text-xs text-slate-500">
            Chưa có mẫu Word nào cho {selectedBook.name}. Hãy tải lên tệp .docx để tự động hóa soạn thảo khi cấp số.
          </div>
        ) : (
          <div className="divide-y divide-slate-100">
            {templates.map((tpl) => (
              <div
                key={tpl.id}
                className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs"
              >
                <div className="flex items-start gap-3 min-w-0 flex-1">
                  <div className="w-9 h-9 rounded-lg bg-sky-50 border border-sky-200 flex items-center justify-center shrink-0">
                    <FileText className="w-4 h-4 text-sky-700" />
                  </div>
                  <div className="min-w-0 flex-1 space-y-1">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="font-mono font-bold px-2 py-0.5 rounded bg-sky-50 text-sky-800 border border-sky-200">
                        {tpl.typeCode}
                      </span>
                      <span className="font-bold text-slate-900 text-sm break-words">
                        {tpl.name}
                      </span>
                    </div>
                    <div className="text-slate-500 flex flex-wrap items-center gap-x-3 gap-y-1">
                      <span>Tệp gốc: {tpl.originalFileName}</span>
                      <span className="tabular-nums">
                        {(tpl.fileSize / 1024).toFixed(0)} KB
                      </span>
                      <span>Người tải: {tpl.createdBy}</span>
                    </div>
                    <div className="flex flex-wrap items-center gap-1.5 pt-0.5">
                      <span className="text-[11px] text-slate-500 font-medium">
                        Trường giữ chỗ:
                      </span>
                      {tpl.placeholders.length > 0 ? (
                        tpl.placeholders.map((ph) => (
                          <code
                            key={ph}
                            className="text-[11px] font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200"
                          >
                            {`{{${ph}}}`}
                          </code>
                        ))
                      ) : (
                        <span className="text-[11px] text-amber-700">
                          Không phát hiện trường giữ chỗ {`{{...}}`} nào
                        </span>
                      )}
                    </div>
                  </div>
                </div>

                {canManage && (
                  <button
                    type="button"
                    onClick={() => handleDeleteTemplate(tpl)}
                    className="inline-flex items-center gap-1.5 h-9 px-3 text-xs font-semibold text-rose-700 hover:bg-rose-50 border border-rose-200 rounded-lg transition-colors shrink-0 self-start sm:self-center"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Xóa mẫu</span>
                  </button>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
