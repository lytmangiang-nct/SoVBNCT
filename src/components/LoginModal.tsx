import React, { useState } from 'react';
import { Lock, User as UserIcon, ArrowRight, AlertCircle, Eye, EyeOff } from 'lucide-react';
import { api, saveAuth } from '../api';
import type { User } from '../types';

interface LoginModalProps {
  onLoginSuccess: (user: User) => void;
}

export const LoginModal: React.FC<LoginModalProps> = ({ onLoginSuccess }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const cleanUser = username.trim();
    if (!cleanUser || !password) {
      setError('Vui lòng nhập tên đăng nhập và mật khẩu');
      return;
    }

    setLoading(true);
    setError(null);
    try {
      const res = await api.login(cleanUser, password, rememberMe);
      saveAuth(res);

      // Kiểm tra lại thông tin tài khoản và phân quyền sổ ngay sau khi đăng nhập
      let verifiedUser = res.user;
      try {
        const meRes = await api.getCurrentUser();
        if (meRes?.user) {
          verifiedUser = meRes.user;
        }
      } catch {
        // Fallback giữ res.user nếu phiên vừa cấp hợp lệ
      }

      onLoginSuccess(verifiedUser);
    } catch (err: any) {
      if (err?.status === 401) {
        setError('Sai tên đăng nhập hoặc mật khẩu');
      } else if (err?.status === 404) {
        setError('Không tìm thấy dịch vụ đăng nhập (404). Vui lòng kiểm tra cấu hình backend hoặc quy tắc API.');
      } else if (err?.status === 0 || err?.code === 'NETWORK_ERROR') {
        setError('Không kết nối được máy chủ. Vui lòng kiểm tra kết nối mạng hoặc trạng thái dịch vụ.');
      } else {
        setError(err?.message || 'Đăng nhập không thành công');
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-[100dvh] bg-slate-100 flex flex-col justify-center items-center px-3 py-4 sm:p-6 overflow-y-auto">
      <div className="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-lg p-5 sm:p-8 space-y-5 sm:space-y-6 my-auto">
        {/* Emblem & Title */}
        <div className="text-center space-y-2">
          <div className="w-12 h-12 sm:w-14 sm:h-14 rounded-2xl bg-sky-700 flex items-center justify-center text-white mx-auto shadow-md">
            <svg className="w-7 h-7 sm:w-8 sm:h-8 text-white" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2L15.09 8.26L22 9.27L17 14.14L18.18 21.02L12 17.77L5.82 21.02L7 14.14L2 9.27L8.91 8.26L12 2Z" />
            </svg>
          </div>
          <div>
            <span className="text-[11px] sm:text-xs font-bold text-sky-800 uppercase tracking-wider sm:tracking-widest block">
              TRƯỜNG THPT NGUYỄN CHÍ THANH
            </span>
            <h1 className="text-lg sm:text-xl font-extrabold text-slate-900 tracking-tight mt-1">
              SỔ CẤP SỐ VĂN BẢN
            </h1>
            <p className="text-xs text-slate-500 mt-1 leading-relaxed">
              Hệ thống cấp số văn bản tự động và lưu trữ trực tuyến
            </p>
          </div>
        </div>

        {/* Standard Browser-Compatible Login Form */}
        <form name="login" method="post" onSubmit={handleSubmit} className="space-y-4 text-xs sm:text-sm">
          <div className="space-y-1.5">
            <label htmlFor="login-username" className="font-bold text-slate-700 block text-xs sm:text-sm">
              Tên đăng nhập
            </label>
            <div className="relative">
              <UserIcon className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                id="login-username"
                name="username"
                type="text"
                autoComplete="username"
                autoCapitalize="none"
                autoCorrect="off"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Nhập tên đăng nhập..."
                required
                className="w-full min-h-[44px] h-11 pl-10 pr-3 bg-slate-50 border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600 transition-colors"
                disabled={loading}
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label htmlFor="login-password" className="font-bold text-slate-700 block text-xs sm:text-sm">
              Mật khẩu
            </label>
            <div className="relative">
              <Lock className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                id="login-password"
                name="password"
                type={showPassword ? 'text' : 'password'}
                autoComplete="current-password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
                required
                className="w-full min-h-[44px] h-11 pl-10 pr-12 bg-slate-50 border border-slate-300 rounded-lg text-base sm:text-sm text-slate-900 focus:bg-white focus:outline-hidden focus:ring-2 focus:ring-sky-600 transition-colors"
                disabled={loading}
              />
              <button
                type="button"
                onClick={() => setShowPassword((prev) => !prev)}
                tabIndex={-1}
                aria-label={showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                className="absolute right-0 inset-y-0 min-w-[44px] flex items-center justify-center text-slate-500 hover:text-slate-700 transition-colors rounded-r-lg"
              >
                {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <div className="pt-0.5">
            <label
              htmlFor="remember-me"
              className="flex items-start gap-3 min-h-[44px] p-2 -mx-2 rounded-lg hover:bg-slate-50 active:bg-slate-100 cursor-pointer select-none group transition-colors"
            >
              <input
                id="remember-me"
                name="rememberMe"
                type="checkbox"
                checked={rememberMe}
                onChange={(e) => setRememberMe(e.target.checked)}
                disabled={loading}
                className="mt-0.5 h-5 w-5 shrink-0 rounded border-slate-300 text-sky-700 focus:ring-sky-600 cursor-pointer"
              />
              <span className="leading-snug min-w-0">
                <span className="text-xs sm:text-sm font-semibold text-slate-800 group-hover:text-sky-800 block">
                  Duy trì đăng nhập trên thiết bị này
                </span>
                <span className="text-[11px] text-slate-500 block mt-0.5">
                  Giữ phiên đăng nhập tối đa 30 ngày trên thiết bị tin cậy
                </span>
              </span>
            </label>
          </div>

          {error && (
            <div
              role="alert"
              className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs sm:text-sm flex items-start gap-2 border border-rose-200"
            >
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
              <span className="break-words min-w-0">{error}</span>
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full min-h-[48px] h-12 bg-sky-700 hover:bg-sky-800 active:bg-sky-900 text-white font-semibold text-sm sm:text-base rounded-lg shadow-sm transition-all flex items-center justify-center gap-2 disabled:opacity-50"
          >
            {loading ? (
              <span>Đang xác thực...</span>
            ) : (
              <>
                <span>Đăng nhập</span>
                <ArrowRight className="w-4 h-4" />
              </>
            )}
          </button>
        </form>
      </div>
    </div>
  );
};
