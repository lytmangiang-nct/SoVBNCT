import React, { useEffect, useState } from 'react';
import { AlertTriangle, Lock, Pencil, Save, Trash2, Unlock, X } from 'lucide-react';
import type { RegistryBook, User } from '../../types';
import { getUserDisplayName } from '../../utils/permissions';

interface PendingBookSwitchModalProps {
  pendingSwitch: {
    targetBookId: string;
    targetBookName: string;
    currentBookName: string;
  } | null;
  saving: boolean;
  errorMessage: string | null;
  onSaveAndSwitch: () => void;
  onDiscardAndSwitch: () => void;
  onCancel: () => void;
}

export const PendingBookSwitchModal: React.FC<PendingBookSwitchModalProps> = ({
  pendingSwitch,
  saving,
  errorMessage,
  onSaveAndSwitch,
  onDiscardAndSwitch,
  onCancel,
}) => {
  if (!pendingSwitch) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="pending-book-switch-title"
    >
      <div className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden">
        <div className="px-4 sm:px-5 py-4 border-b border-slate-100 flex items-start justify-between gap-3">
          <div className="flex items-start gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-lg bg-amber-100 text-amber-700 flex items-center justify-center shrink-0">
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h3 id="pending-book-switch-title" className="text-sm font-bold text-slate-900">
                Có thay đổi chưa lưu
              </h3>
              <p className="text-xs text-slate-600 mt-0.5 break-words">
                Sổ <strong>{pendingSwitch.currentBookName}</strong> đang có thay đổi chưa lưu trước khi chuyển sang{' '}
                <strong>{pendingSwitch.targetBookName}</strong>.
              </p>
            </div>
          </div>
          <button
            type="button"
            disabled={saving}
            onClick={onCancel}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="px-4 sm:px-5 py-4 space-y-3 text-xs text-slate-700">
          <p>Bạn muốn lưu lại cấu hình vừa sửa trước khi chuyển sổ hay bỏ qua các thay đổi chưa lưu?</p>
          {errorMessage && (
            <div className="p-2.5 bg-rose-50 border border-rose-200 text-rose-700 rounded-lg break-words">
              {errorMessage}
            </div>
          )}
        </div>

        <div className="px-4 sm:px-5 py-3.5 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row sm:justify-end gap-2">
          <button
            type="button"
            disabled={saving}
            onClick={onCancel}
            className="w-full sm:w-auto min-h-[40px] px-3.5 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs"
          >
            Ở lại tiếp tục sửa
          </button>
          <button
            type="button"
            disabled={saving}
            onClick={onDiscardAndSwitch}
            className="w-full sm:w-auto min-h-[40px] px-3.5 py-2 bg-rose-50 hover:bg-rose-100 text-rose-800 border border-rose-200 font-semibold rounded-lg text-xs"
          >
            Bỏ thay đổi
          </button>
          <button
            type="button"
            disabled={saving}
            onClick={onSaveAndSwitch}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white font-semibold rounded-lg text-xs inline-flex items-center justify-center gap-1.5"
          >
            <Save className="w-3.5 h-3.5" />
            <span>{saving ? 'Đang lưu...' : 'Lưu thay đổi'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};

interface ToggleBookLockModalProps {
  book: RegistryBook | null;
  saving: boolean;
  onConfirm: () => void;
  onCancel: () => void;
}

export const ToggleBookLockModal: React.FC<ToggleBookLockModalProps> = ({
  book,
  saving,
  onConfirm,
  onCancel,
}) => {
  if (!book) return null;
  const isLocking = book.isActive;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="toggle-book-lock-title"
    >
      <div className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden">
        <div className="px-4 sm:px-5 py-4 border-b border-slate-100 flex items-start gap-3">
          <div
            className={`w-9 h-9 rounded-lg flex items-center justify-center shrink-0 ${
              isLocking ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'
            }`}
          >
            {isLocking ? <Lock className="w-5 h-5" /> : <Unlock className="w-5 h-5" />}
          </div>
          <div className="min-w-0 flex-1">
            <h3 id="toggle-book-lock-title" className="text-sm font-bold text-slate-900">
              {isLocking ? `Tạm khóa sổ "${book.name}"?` : `Mở khóa sổ "${book.name}"?`}
            </h3>
            <p className="text-xs text-slate-600 mt-1 break-words">
              {isLocking
                ? 'Khi tạm khóa, người dùng sẽ tạm ngừng cấp số hoặc nhập văn bản mới trên sổ này; toàn bộ hồ sơ và số đã cấp trước đó vẫn được giữ nguyên.'
                : 'Sổ sẽ hoạt động trở lại bình thường cho các tài khoản đã được phân quyền.'}
            </p>
          </div>
        </div>

        <div className="px-4 sm:px-5 py-3.5 bg-slate-50 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
          <button
            type="button"
            disabled={saving}
            onClick={onCancel}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={saving}
            onClick={onConfirm}
            className={`w-full sm:w-auto min-h-[40px] px-4 py-2 text-white font-semibold rounded-lg text-xs disabled:opacity-50 ${
              isLocking ? 'bg-amber-700 hover:bg-amber-800' : 'bg-emerald-700 hover:bg-emerald-800'
            }`}
          >
            {saving ? 'Đang cập nhật...' : isLocking ? 'Xác nhận tạm khóa' : 'Xác nhận mở khóa'}
          </button>
        </div>
      </div>
    </div>
  );
};

interface ConfirmDeleteModalProps {
  request: {
    title: string;
    description: string;
    confirmLabel: string;
    onConfirm: () => Promise<void>;
  } | null;
  busy: boolean;
  onClose: () => void;
}

export const ConfirmDeleteModal: React.FC<ConfirmDeleteModalProps> = ({
  request,
  busy,
  onClose,
}) => {
  if (!request) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
    >
      <div className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden">
        <div className="px-4 sm:px-5 py-4 border-b border-slate-100 flex items-start gap-3">
          <div className="w-9 h-9 rounded-lg bg-rose-100 text-rose-700 flex items-center justify-center shrink-0">
            <Trash2 className="w-5 h-5" />
          </div>
          <div className="min-w-0 flex-1">
            <h3 className="text-sm font-bold text-slate-900">{request.title}</h3>
            <p className="text-xs text-slate-600 mt-1 break-words">{request.description}</p>
          </div>
        </div>
        <div className="px-4 sm:px-5 py-3.5 bg-slate-50 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
          <button
            type="button"
            disabled={busy}
            onClick={onClose}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={busy}
            onClick={request.onConfirm}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-lg text-xs disabled:opacity-50"
          >
            {busy ? 'Đang xử lý...' : request.confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
};

interface ResetPasswordModalProps {
  user: User | null;
  saving: boolean;
  onSubmit: (newPassword: string) => Promise<void>;
  onClose: () => void;
}

export const ResetPasswordModal: React.FC<ResetPasswordModalProps> = ({
  user,
  saving,
  onSubmit,
  onClose,
}) => {
  const [password, setPassword] = useState('');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (user) {
      setPassword('');
      setError(null);
    }
  }, [user]);

  if (!user) return null;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (password.trim().length < 6) {
      setError('Mật khẩu mới phải có tối thiểu 6 ký tự.');
      return;
    }
    setError(null);
    await onSubmit(password.trim());
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
    >
      <form
        onSubmit={handleSubmit}
        noValidate
        className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden"
      >
        <div className="px-4 sm:px-5 py-4 border-b border-slate-100 flex items-center justify-between gap-2">
          <div className="min-w-0">
            <h3 className="text-sm font-bold text-slate-900">
              Đổi mật khẩu tài khoản @{user.username}
            </h3>
            <p className="text-xs text-slate-500 mt-0.5 truncate">{getUserDisplayName(user)}</p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="px-4 sm:px-5 py-4 space-y-2 text-xs">
          <label htmlFor="reset-user-password-input" className="font-bold text-slate-700 block">
            Mật khẩu mới (tối thiểu 6 ký tự)
          </label>
          <input
            id="reset-user-password-input"
            type="password"
            value={password}
            onChange={(e) => {
              setPassword(e.target.value);
              if (error) setError(null);
            }}
            placeholder="Nhập mật khẩu mới..."
            className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-base sm:text-xs"
          />
          {error && <p className="text-[11px] font-medium text-rose-600">{error}</p>}
        </div>

        <div className="px-4 sm:px-5 py-3.5 bg-slate-50 border-t border-slate-100 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
          <button
            type="button"
            disabled={saving}
            onClick={onClose}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs"
          >
            Hủy
          </button>
          <button
            type="submit"
            disabled={saving}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg text-xs disabled:opacity-50"
          >
            {saving ? 'Đang cập nhật...' : 'Cập nhật mật khẩu'}
          </button>
        </div>
      </form>
    </div>
  );
};

interface EditDisplayNameModalProps {
  user: User | null;
  saving: boolean;
  onSubmit: (displayName: string) => Promise<void>;
  onClose: () => void;
}

export const EditDisplayNameModal: React.FC<EditDisplayNameModalProps> = ({
  user,
  saving,
  onSubmit,
  onClose,
}) => {
  const [displayName, setDisplayName] = useState('');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (user) {
      setDisplayName(user.displayName || '');
      setError(null);
    }
  }, [user]);

  if (!user) return null;

  const fallbackLabel = user.fullName?.trim() || user.username;
  const fallbackDesc = user.fullName?.trim()
    ? `Họ và tên (${user.fullName.trim()})`
    : `tên đăng nhập (@${user.username})`;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = displayName.trim();
    if (trimmed.length > 80) {
      setError('Tên hiển thị tối đa 80 ký tự.');
      return;
    }
    setError(null);
    await onSubmit(trimmed);
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/50 backdrop-blur-[1px]"
      role="dialog"
      aria-modal="true"
      aria-labelledby="edit-display-name-title"
    >
      <form
        onSubmit={handleSubmit}
        noValidate
        className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden"
      >
        <div className="px-4 sm:px-5 py-4 border-b border-slate-100 flex items-start justify-between gap-2">
          <div className="flex items-start gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-lg bg-sky-50 text-sky-700 border border-sky-200 flex items-center justify-center shrink-0">
              <Pencil className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h3 id="edit-display-name-title" className="text-sm font-bold text-slate-900">
                Sửa tên hiển thị tài khoản @{user.username}
              </h3>
              <p className="text-xs text-slate-500 mt-0.5 break-words">
                {user.fullName?.trim()
                  ? `Họ và tên gốc: ${user.fullName.trim()}`
                  : `Tài khoản chưa có họ tên (mặc định dùng @${user.username})`}
              </p>
            </div>
          </div>
          <button
            type="button"
            disabled={saving}
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="px-4 sm:px-5 py-4 space-y-2.5 text-xs">
          <div className="flex items-center justify-between gap-2">
            <label htmlFor="edit-user-display-name-input" className="font-bold text-slate-700">
              Tên hiển thị
            </label>
            {displayName.trim() && (
              <button
                type="button"
                disabled={saving}
                onClick={() => {
                  setDisplayName('');
                  if (error) setError(null);
                }}
                className="text-[11px] font-semibold text-sky-700 hover:underline"
              >
                Xóa để dùng {user.fullName?.trim() ? 'Họ và tên' : 'tên đăng nhập'}
              </button>
            )}
          </div>
          <input
            id="edit-user-display-name-input"
            type="text"
            maxLength={80}
            value={displayName}
            onChange={(e) => {
              setDisplayName(e.target.value);
              if (error) setError(null);
            }}
            placeholder="Ví dụ: Hiệu trưởng, Phó Bí thư ĐT hoặc Nguyễn Văn A"
            className="w-full min-h-[44px] h-11 sm:h-10 px-3 bg-white border border-slate-300 rounded-lg text-base sm:text-xs"
          />
          <p className="text-[11px] text-slate-500 leading-relaxed">
            Gợi ý dùng chức vụ hoặc họ tên tùy mục đích. Để trống sẽ hiển thị theo{' '}
            <strong className="text-slate-700">{fallbackDesc}</strong> (hiển thị:{' '}
            <span className="font-semibold text-slate-700">{fallbackLabel}</span>). Tên hiển thị chỉ
            là nhãn nhận diện, không làm thay đổi tên đăng nhập hoặc quyền tài khoản.
          </p>
          {error && <p className="text-[11px] font-medium text-rose-600">{error}</p>}
        </div>

        <div className="px-4 sm:px-5 py-3.5 bg-slate-50 border-t border-slate-100 flex flex-col-reverse sm:flex-row sm:justify-end gap-2">
          <button
            type="button"
            disabled={saving}
            onClick={onClose}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-semibold rounded-lg text-xs"
          >
            Hủy
          </button>
          <button
            type="submit"
            disabled={saving}
            className="w-full sm:w-auto min-h-[40px] px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg text-xs disabled:opacity-50 inline-flex items-center justify-center gap-1.5"
          >
            <Save className="w-3.5 h-3.5" />
            <span>{saving ? 'Đang lưu...' : 'Lưu tên hiển thị'}</span>
          </button>
        </div>
      </form>
    </div>
  );
};
