import React, { useEffect, useState } from 'react';
import { AlertTriangle, ChevronDown, ChevronRight, Plus, X } from 'lucide-react';
import type { DocumentTypeConfig, RegistryBook } from '../../types';
import { normalizeBookCode } from '../../utils/permissions';

export type NewBookFieldErrorKey =
  | 'name'
  | 'unitName'
  | 'unitCode'
  | 'code'
  | 'numberFormat'
  | 'enabledTypeCodes';

export interface CreateBookFormValues {
  code?: string;
  name: string;
  unitName: string;
  unitCode: string;
  isShared: boolean;
  numberMode: 'separate' | 'shared';
  numberFormat: string;
  enabledTypeCodes: string[];
}

interface CreateBookDrawerProps {
  open: boolean;
  onClose: () => void;
  existingBooks: RegistryBook[];
  documentTypes: DocumentTypeConfig[];
  saving: boolean;
  isConfigBlocked: boolean;
  serverError: string | null;
  onClearServerError: () => void;
  onSubmitBook: (values: CreateBookFormValues) => Promise<boolean>;
}

export const CreateBookDrawer: React.FC<CreateBookDrawerProps> = ({
  open,
  onClose,
  existingBooks,
  documentTypes,
  saving,
  isConfigBlocked,
  serverError,
  onClearServerError,
  onSubmitBook,
}) => {
  const activeDocTypes = documentTypes.filter((t) => t.isActive);

  const [name, setName] = useState('');
  const [unitName, setUnitName] = useState('');
  const [unitCode, setUnitCode] = useState('');
  const [code, setCode] = useState('');
  const [codeEdited, setCodeEdited] = useState(false);

  const [advancedOpen, setAdvancedOpen] = useState(false);
  const [numberFormat, setNumberFormat] = useState('{so}-{loai}/{don_vi}');
  const [isShared, setIsShared] = useState(false);
  const [numberMode, setNumberMode] = useState<'separate' | 'shared'>('separate');
  const [enabledTypeCodes, setEnabledTypeCodes] = useState<string[]>(() =>
    activeDocTypes.map((t) => t.code)
  );
  const [fieldErrors, setFieldErrors] = useState<Partial<Record<NewBookFieldErrorKey, string>>>({});

  // Reset form each time drawer opens
  useEffect(() => {
    if (open) {
      setFieldErrors({});
      setAdvancedOpen(false);
      if (enabledTypeCodes.length === 0 && activeDocTypes.length > 0) {
        setEnabledTypeCodes(activeDocTypes.map((t) => t.code));
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  // Close on Escape
  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !saving) onClose();
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [open, saving, onClose]);

  if (!open) return null;

  const clearFieldError = (key: NewBookFieldErrorKey) => {
    setFieldErrors((prev) => {
      if (!prev[key]) return prev;
      const next = { ...prev };
      delete next[key];
      return next;
    });
    if (serverError) onClearServerError();
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isConfigBlocked || saving) return;

    const trimmedName = name.trim();
    const trimmedUnitName = unitName.trim();
    const trimmedUnitCode = unitCode.trim().toUpperCase();
    const resolvedCode = normalizeBookCode(code || trimmedName);
    const trimmedFormat = numberFormat.trim();

    const errors: Partial<Record<NewBookFieldErrorKey, string>> = {};
    if (!trimmedName) {
      errors.name = 'Vui lòng nhập tên sổ văn bản.';
    }
    if (!trimmedUnitName) {
      errors.unitName = 'Vui lòng nhập tên đơn vị hoặc bộ phận.';
    }
    if (!trimmedUnitCode) {
      errors.unitCode = 'Vui lòng nhập ký hiệu đơn vị chính thức (ví dụ: TNV, CB).';
    }
    if (!resolvedCode || resolvedCode.length < 2) {
      errors.code = 'Mã sổ nội bộ phải có ít nhất 2 ký tự chữ hoặc số.';
    } else if (existingBooks.some((b) => b.code.toUpperCase() === resolvedCode)) {
      errors.code = `Mã sổ "${resolvedCode}" đã tồn tại. Vui lòng chọn mã khác.`;
    }

    if (!trimmedFormat || !trimmedFormat.includes('{so}')) {
      errors.numberFormat = 'Mẫu định dạng số phải chứa biến {so} (ví dụ: {so}-{loai}/{don_vi}).';
    }
    if (enabledTypeCodes.length === 0) {
      errors.enabledTypeCodes = 'Vui lòng chọn ít nhất 1 loại văn bản được sử dụng trong sổ mới.';
    }

    if (Object.keys(errors).length > 0) {
      setFieldErrors(errors);
      if (errors.numberFormat || errors.enabledTypeCodes) {
        setAdvancedOpen(true);
      }
      return;
    }

    const ok = await onSubmitBook({
      code: resolvedCode,
      name: trimmedName,
      unitName: trimmedUnitName,
      unitCode: trimmedUnitCode,
      isShared,
      numberMode,
      numberFormat: trimmedFormat,
      enabledTypeCodes,
    });

    if (ok) {
      setName('');
      setUnitName('');
      setUnitCode('');
      setCode('');
      setCodeEdited(false);
      setIsShared(false);
      setNumberMode('separate');
      setNumberFormat('{so}-{loai}/{don_vi}');
      setEnabledTypeCodes(activeDocTypes.map((t) => t.code));
      setFieldErrors({});
      setAdvancedOpen(false);
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="create-book-drawer-title"
    >
      {/* Backdrop click to close */}
      <button
        type="button"
        aria-label="Đóng bảng tạo sổ mới"
        onClick={() => !saving && onClose()}
        className="hidden sm:block flex-1 cursor-default"
      />

      {/* Full screen on mobile, right slide-over panel on sm+ */}
      <div className="w-full sm:max-w-lg bg-white h-full flex flex-col shadow-2xl border-l border-slate-200 overflow-hidden">
        {/* Drawer Header */}
        <div className="px-4 sm:px-6 py-4 border-b border-slate-200 flex items-center justify-between gap-3 shrink-0 bg-white">
          <div className="min-w-0">
            <h2
              id="create-book-drawer-title"
              className="text-base font-bold text-slate-900 flex items-center gap-2"
            >
              <Plus className="w-4 h-4 text-sky-700 shrink-0" />
              <span className="truncate">Tạo sổ văn bản mới</span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Tạo sổ riêng theo bộ phận / tổ chuyên môn hoặc sổ dùng chung toàn trường.
            </p>
          </div>
          <button
            type="button"
            onClick={() => !saving && onClose()}
            className="p-2 min-w-[40px] min-h-[40px] inline-flex items-center justify-center rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 transition-colors shrink-0"
            title="Đóng"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Drawer Body */}
        <form
          id="create-book-drawer-form"
          onSubmit={handleSubmit}
          noValidate
          className="flex-1 overflow-y-auto px-4 sm:px-6 py-5 space-y-4 text-xs"
        >
          {/* Common Fields First */}
          <div className="space-y-1 min-w-0">
            <label htmlFor="new-book-name" className="font-bold text-slate-700 block">
              Tên sổ văn bản *
            </label>
            <input
              id="new-book-name"
              type="text"
              value={name}
              onChange={(e) => {
                const val = e.target.value;
                setName(val);
                clearFieldError('name');
                if (!codeEdited) {
                  setCode(normalizeBookCode(val));
                  clearFieldError('code');
                }
              }}
              placeholder="Ví dụ: Tổ Ngữ văn / Công đoàn / Văn thư"
              aria-invalid={Boolean(fieldErrors.name)}
              className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg text-base sm:text-sm focus:outline-hidden focus:ring-2 ${
                fieldErrors.name
                  ? 'border-rose-400 focus:ring-rose-500'
                  : 'border-slate-300 focus:ring-sky-600'
              }`}
            />
            {fieldErrors.name && (
              <p className="text-[11px] font-medium text-rose-600">{fieldErrors.name}</p>
            )}
          </div>

          <div className="space-y-1 min-w-0">
            <label htmlFor="new-book-unit-name" className="font-bold text-slate-700 block">
              Tên đơn vị / bộ phận *
            </label>
            <input
              id="new-book-unit-name"
              type="text"
              value={unitName}
              onChange={(e) => {
                setUnitName(e.target.value);
                clearFieldError('unitName');
              }}
              placeholder="Ví dụ: Tổ Ngữ văn - Trường THPT Nguyễn Chí Thanh"
              aria-invalid={Boolean(fieldErrors.unitName)}
              className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg text-base sm:text-sm focus:outline-hidden focus:ring-2 ${
                fieldErrors.unitName
                  ? 'border-rose-400 focus:ring-rose-500'
                  : 'border-slate-300 focus:ring-sky-600'
              }`}
            />
            {fieldErrors.unitName && (
              <p className="text-[11px] font-medium text-rose-600">{fieldErrors.unitName}</p>
            )}
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div className="space-y-1 min-w-0">
              <label htmlFor="new-book-unit-code" className="font-bold text-slate-700 block">
                Ký hiệu đơn vị *
              </label>
              <input
                id="new-book-unit-code"
                type="text"
                value={unitCode}
                onChange={(e) => {
                  setUnitCode(e.target.value);
                  clearFieldError('unitCode');
                }}
                placeholder="Ví dụ: TNV, CB..."
                aria-invalid={Boolean(fieldErrors.unitCode)}
                className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg font-mono uppercase text-base sm:text-sm focus:outline-hidden focus:ring-2 ${
                  fieldErrors.unitCode
                    ? 'border-rose-400 focus:ring-rose-500'
                    : 'border-slate-300 focus:ring-sky-600'
                }`}
              />
              {fieldErrors.unitCode && (
                <p className="text-[11px] font-medium text-rose-600">{fieldErrors.unitCode}</p>
              )}
            </div>

            <div className="space-y-1 min-w-0">
              <label htmlFor="new-book-code" className="font-bold text-slate-700 block">
                Mã sổ nội bộ cố định *
              </label>
              <input
                id="new-book-code"
                type="text"
                value={code}
                onChange={(e) => {
                  setCodeEdited(true);
                  setCode(normalizeBookCode(e.target.value));
                  clearFieldError('code');
                }}
                placeholder="Ví dụ: TO_NGU_VAN"
                aria-invalid={Boolean(fieldErrors.code)}
                className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg font-mono uppercase text-base sm:text-sm focus:outline-hidden focus:ring-2 ${
                  fieldErrors.code
                    ? 'border-rose-400 focus:ring-rose-500'
                    : 'border-slate-300 focus:ring-sky-600'
                }`}
              />
              {fieldErrors.code ? (
                <p className="text-[11px] font-medium text-rose-600">{fieldErrors.code}</p>
              ) : (
                <p className="text-[11px] text-slate-500">
                  Dùng làm tên tab Google Sheets, không thể đổi sau khi tạo.
                </p>
              )}
            </div>
          </div>

          {/* Expandable "Tùy chọn cấp số" section */}
          <div className="border border-slate-200 rounded-xl overflow-hidden bg-slate-50/70">
            <button
              type="button"
              onClick={() => setAdvancedOpen((prev) => !prev)}
              aria-expanded={advancedOpen}
              className="w-full px-3.5 py-3 min-h-[44px] flex items-center justify-between gap-2 text-left bg-slate-50 hover:bg-slate-100/80 transition-colors"
            >
              <div className="min-w-0">
                <span className="font-bold text-slate-800 block">Tùy chọn cấp số</span>
                <span className="text-[11px] text-slate-500 block truncate">
                  Mẫu: {numberFormat || '{so}-{loai}/{don_vi}'} ·{' '}
                  {numberMode === 'shared' ? 'Dãy chung' : 'Dãy riêng'} · {enabledTypeCodes.length} loại VB
                </span>
              </div>
              {advancedOpen ? (
                <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
              ) : (
                <ChevronRight className="w-4 h-4 text-slate-500 shrink-0" />
              )}
            </button>

            {advancedOpen && (
              <div className="p-3.5 sm:p-4 bg-white border-t border-slate-200 space-y-4">
                <div className="space-y-1 min-w-0">
                  <label htmlFor="new-book-format" className="font-bold text-slate-700 block">
                    Mẫu định dạng số *
                  </label>
                  <input
                    id="new-book-format"
                    type="text"
                    value={numberFormat}
                    onChange={(e) => {
                      setNumberFormat(e.target.value);
                      clearFieldError('numberFormat');
                    }}
                    aria-invalid={Boolean(fieldErrors.numberFormat)}
                    className={`w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border rounded-lg font-mono text-base sm:text-xs focus:outline-hidden focus:ring-2 ${
                      fieldErrors.numberFormat
                        ? 'border-rose-400 focus:ring-rose-500'
                        : 'border-slate-300 focus:ring-sky-600'
                    }`}
                  />
                  {fieldErrors.numberFormat ? (
                    <p className="text-[11px] font-medium text-rose-600">{fieldErrors.numberFormat}</p>
                  ) : (
                    <p className="text-[11px] text-slate-500">
                      Biến hỗ trợ: <code>{'{so}'}</code>, <code>{'{loai}'}</code>, <code>{'{don_vi}'}</code>.
                    </p>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="space-y-1 min-w-0">
                    <label htmlFor="new-book-scope" className="font-bold text-slate-700 block">
                      Phạm vi sổ
                    </label>
                    <select
                      id="new-book-scope"
                      value={isShared ? 'shared' : 'unit'}
                      onChange={(e) => setIsShared(e.target.value === 'shared')}
                      className="w-full min-h-[44px] h-11 sm:h-10 px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                    >
                      <option value="unit">Sổ riêng đơn vị</option>
                      <option value="shared">Dùng chung trường</option>
                    </select>
                  </div>

                  <div className="space-y-1 min-w-0">
                    <label htmlFor="new-book-mode" className="font-bold text-slate-700 block">
                      Chế độ dãy số
                    </label>
                    <select
                      id="new-book-mode"
                      value={numberMode}
                      onChange={(e) =>
                        setNumberMode(e.target.value === 'shared' ? 'shared' : 'separate')
                      }
                      className="w-full min-h-[44px] h-11 sm:h-10 px-2.5 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
                    >
                      <option value="separate">Dãy riêng theo loại</option>
                      <option value="shared">Dãy chung liên tục</option>
                    </select>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="flex items-center justify-between gap-2">
                    <label className="font-bold text-slate-700 block">
                      Loại văn bản được sử dụng *
                    </label>
                    <div className="flex items-center gap-2 text-[11px]">
                      <button
                        type="button"
                        onClick={() => {
                          setEnabledTypeCodes(activeDocTypes.map((t) => t.code));
                          clearFieldError('enabledTypeCodes');
                        }}
                        className="font-semibold text-sky-700 hover:underline"
                      >
                        Chọn tất cả
                      </button>
                      <span className="text-slate-300">|</span>
                      <button
                        type="button"
                        onClick={() => setEnabledTypeCodes([])}
                        className="font-semibold text-slate-500 hover:underline"
                      >
                        Bỏ chọn
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    {activeDocTypes.map((t) => {
                      const checked = enabledTypeCodes.includes(t.code);
                      return (
                        <label
                          key={t.code}
                          className={`flex items-center gap-2.5 p-2.5 min-h-[40px] rounded-lg border cursor-pointer transition-colors ${
                            checked
                              ? 'bg-sky-50 border-sky-300 text-sky-950 font-semibold'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => {
                              clearFieldError('enabledTypeCodes');
                              setEnabledTypeCodes((prev) =>
                                prev.includes(t.code)
                                  ? prev.filter((c) => c !== t.code)
                                  : [...prev, t.code]
                              );
                            }}
                            className="rounded text-sky-700"
                          />
                          <span className="truncate">
                            {t.name} ({t.code})
                          </span>
                        </label>
                      );
                    })}
                  </div>
                  {fieldErrors.enabledTypeCodes && (
                    <p className="text-[11px] font-medium text-rose-600">
                      {fieldErrors.enabledTypeCodes}
                    </p>
                  )}
                </div>
              </div>
            )}
          </div>

          {serverError && (
            <div className="p-3 bg-rose-50 border border-rose-200 text-rose-700 rounded-lg flex items-start gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
              <span className="break-words">{serverError}</span>
            </div>
          )}
        </form>

        {/* Drawer Footer */}
        <div className="px-4 sm:px-6 py-3.5 border-t border-slate-200 bg-slate-50 flex flex-col-reverse sm:flex-row sm:justify-end gap-2 shrink-0">
          <button
            type="button"
            disabled={saving}
            onClick={onClose}
            className="w-full sm:w-auto min-h-[42px] px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs transition-colors"
          >
            Hủy
          </button>
          <button
            type="submit"
            form="create-book-drawer-form"
            disabled={saving || isConfigBlocked}
            className="w-full sm:w-auto min-h-[42px] px-5 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg text-xs shadow-xs transition-colors inline-flex items-center justify-center gap-1.5"
          >
            <Plus className="w-4 h-4" />
            <span>{saving ? 'Đang tạo sổ...' : 'Tạo sổ văn bản'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};
