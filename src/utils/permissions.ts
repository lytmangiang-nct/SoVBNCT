import type { BookPermission, DocumentRecord, DocumentTypeConfig, RegistryBook, User } from '../types.js';

export const DEFAULT_BOOK_ID = 'book-doan-truong';
export const CHI_BO_BOOK_ID = 'book-chi-bo';
export const NHA_TRUONG_BOOK_ID = 'book-nha-truong';

export const ALL_BOOK_PERMISSIONS: BookPermission[] = [
  'view',
  'issue_number',
  'request_number',
  'import_incoming',
  'edit',
  'replace_file',
  'cancel',
  'delete',
  'export',
  'manage_book',
];

export const BOOK_PERMISSION_META: Record<
  BookPermission,
  { label: string; shortLabel: string; description: string }
> = {
  view: {
    label: 'Chỉ xem sổ & mở tệp',
    shortLabel: 'Chỉ xem',
    description: 'Xem danh sách, tra cứu, chi tiết văn bản và mở liên kết tệp đính kèm trong sổ.',
  },
  issue_number: {
    label: 'Lấy số / Cấp số văn bản',
    shortLabel: 'Lấy số',
    description: 'Cấp số văn bản chính thức và đồng thời xem, tra cứu các số đã cấp trong sổ này.',
  },
  request_number: {
    label: 'Xin số văn bản (chờ duyệt)',
    shortLabel: 'Xin số',
    description: 'Gửi yêu cầu xin cấp số văn bản để người có quyền Lấy số duyệt.',
  },
  import_incoming: {
    label: 'Nhập văn bản đến',
    shortLabel: 'Nhập văn bản đến',
    description: 'Phân tích và nhập văn bản PDF đến từ Sở hoặc cơ quan khác vào sổ.',
  },
  edit: {
    label: 'Sửa thông tin',
    shortLabel: 'Sửa',
    description: 'Chỉnh sửa tên văn bản, trích yếu, ngày ban hành, ngày ký, người ký và ghi chú.',
  },
  replace_file: {
    label: 'Thay tệp / Thử lưu lại',
    shortLabel: 'Thay tệp',
    description: 'Thay thế tệp đính kèm trên Google Drive hoặc thử lưu lại khi tải tệp thất bại.',
  },
  cancel: {
    label: 'Hủy số văn bản',
    shortLabel: 'Hủy số',
    description: 'Hủy hiệu lực văn bản trong sổ (giữ lại số và lịch sử, không xóa vĩnh viễn).',
  },
  delete: {
    label: 'Xóa văn bản (VB đã xóa)',
    shortLabel: 'Xóa VB',
    description: 'Xóa văn bản khỏi sổ hiệu lực (bắt buộc nhập lý do, chuyển sang VB đã xóa, giữ số và tệp Drive).',
  },
  export: {
    label: 'Xuất sổ Excel',
    shortLabel: 'Xuất sổ',
    description: 'Xuất dữ liệu văn bản trong sổ ra tệp Excel (.xlsx).',
  },
  manage_book: {
    label: 'Quản trị sổ',
    shortLabel: 'Quản trị sổ',
    description: 'Cấu hình dãy số, số bắt đầu, loại văn bản dùng trong sổ, người ký và thư mục Google Drive của sổ.',
  },
};

export function normalizeBookCode(raw: string): string {
  return (raw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 40);
}

export function normalizeBookPermissionList(perms: BookPermission[]): BookPermission[] {
  const validSet = new Set<string>(ALL_BOOK_PERMISSIONS);
  const set = new Set<BookPermission>(
    perms.filter((p): p is BookPermission => typeof p === 'string' && validSet.has(p))
  );
  // Khi gán quyền Lấy số (issue_number) ở một sổ, đồng thời tự động có quyền Xem & tra cứu trong sổ đó
  if (set.has('issue_number')) {
    set.add('view');
  }
  return ALL_BOOK_PERMISSIONS.filter((p) => set.has(p));
}

export function sanitizeBookPermissionsMap(raw: any): Record<string, BookPermission[]> {
  if (!raw || typeof raw !== 'object') return {};
  const cleaned: Record<string, BookPermission[]> = {};
  for (const [bookId, perms] of Object.entries(raw)) {
    if (!bookId || !Array.isArray(perms)) continue;
    const normalized = normalizeBookPermissionList(perms as BookPermission[]);
    if (normalized.length > 0) {
      cleaned[bookId] = normalized;
    }
  }
  return cleaned;
}

export function isSystemAdmin(user: Pick<User, 'role' | 'username'> | null | undefined): boolean {
  if (!user) return false;
  return user.role === 'admin' || user.username.toLowerCase() === 'admin';
}

export function getUserBookPermissions(
  user: User | null | undefined,
  bookId: string = DEFAULT_BOOK_ID
): BookPermission[] {
  if (!user) return [];
  if (isSystemAdmin(user)) return [...ALL_BOOK_PERMISSIONS];
  const perms = user.bookPermissions?.[bookId] || [];
  return normalizeBookPermissionList(perms);
}

export function hasBookPermission(
  user: User | null | undefined,
  bookId: string | undefined,
  permission: BookPermission
): boolean {
  if (!user) return false;
  if (isSystemAdmin(user)) return true;
  const targetBookId = bookId || DEFAULT_BOOK_ID;
  const perms = user.bookPermissions?.[targetBookId] || [];
  if (permission === 'view') {
    return perms.includes('view') || perms.includes('issue_number');
  }
  return perms.includes(permission);
}

export function hasAnyBookPermission(
  user: User | null | undefined,
  permission: BookPermission,
  books: RegistryBook[] = []
): boolean {
  if (!user) return false;
  if (isSystemAdmin(user)) return true;
  const bookIds = books.length > 0 ? books.filter((b) => b.isActive).map((b) => b.id) : Object.keys(user.bookPermissions || {});
  return bookIds.some((bid) => hasBookPermission(user, bid, permission));
}

export function getAccessibleBooks(user: User | null | undefined, books: RegistryBook[]): RegistryBook[] {
  const activeBooks = books.filter((b) => b.isActive);
  if (!user) return [];
  if (isSystemAdmin(user)) return activeBooks;
  return activeBooks.filter((b) => (user.bookPermissions?.[b.id]?.length || 0) > 0);
}

export function getDocBookId(doc: Pick<DocumentRecord, 'bookId'>): string {
  return doc.bookId || DEFAULT_BOOK_ID;
}

export function isTypeEnabledInBook(book: Pick<RegistryBook, 'enabledTypeCodes'> | undefined, typeCode: string): boolean {
  if (!book || !Array.isArray(book.enabledTypeCodes) || book.enabledTypeCodes.length === 0) {
    return true;
  }
  const upper = (typeCode || '').trim().toUpperCase();
  return book.enabledTypeCodes.some((c) => c.toUpperCase() === upper);
}

export function getBookEnabledTypes(
  book: Pick<RegistryBook, 'enabledTypeCodes'> | undefined,
  allTypes: DocumentTypeConfig[]
): DocumentTypeConfig[] {
  const active = allTypes.filter((t) => t.isActive);
  if (!book || !Array.isArray(book.enabledTypeCodes) || book.enabledTypeCodes.length === 0) {
    return active;
  }
  const enabledSet = new Set(book.enabledTypeCodes.map((c) => c.toUpperCase()));
  const filtered = active.filter((t) => enabledSet.has(t.code.toUpperCase()));
  return filtered.length > 0 ? filtered : active;
}
