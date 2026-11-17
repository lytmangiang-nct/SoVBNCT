export interface DocumentColorStyle {
  key: string;
  label: string;
  hex: string;
  // Badge for table document number (Số/Ký hiệu)
  badgeBg: string;
  badgeText: string;
  badgeBorder: string;
  // Tag pill for type name (Loại văn bản)
  tagBg: string;
  tagText: string;
  tagBorder: string;
  // Next number preview banner in IssueNumberView
  previewBg: string;
  previewBorder: string;
  previewText: string;
  previewTag: string;
  // Indicator dot
  dotColor: string;
}

export const DOCUMENT_COLOR_PALETTE: Record<string, DocumentColorStyle> = {
  blue: {
    key: 'blue',
    label: 'Xanh dương (Kế hoạch)',
    hex: '#2563eb',
    badgeBg: 'bg-blue-50/90',
    badgeText: 'text-blue-800',
    badgeBorder: 'border-blue-300',
    tagBg: 'bg-blue-100 text-blue-800',
    tagText: 'text-blue-800',
    tagBorder: 'border-blue-300',
    previewBg: 'bg-gradient-to-br from-blue-50 via-sky-50 to-indigo-50/60',
    previewBorder: 'border-blue-300',
    previewText: 'text-blue-900',
    previewTag: 'bg-blue-100 text-blue-800 border-blue-300',
    dotColor: 'bg-blue-600',
  },
  emerald: {
    key: 'emerald',
    label: 'Xanh ngọc / lục (Báo cáo)',
    hex: '#059669',
    badgeBg: 'bg-emerald-50/90',
    badgeText: 'text-emerald-800',
    badgeBorder: 'border-emerald-300',
    tagBg: 'bg-emerald-100 text-emerald-800',
    tagText: 'text-emerald-800',
    tagBorder: 'border-emerald-300',
    previewBg: 'bg-gradient-to-br from-emerald-50 via-teal-50 to-green-50/60',
    previewBorder: 'border-emerald-300',
    previewText: 'text-emerald-900',
    previewTag: 'bg-emerald-100 text-emerald-800 border-emerald-300',
    dotColor: 'bg-emerald-600',
  },
  amber: {
    key: 'amber',
    label: 'Hổ phách / Vàng (Thông báo)',
    hex: '#d97706',
    badgeBg: 'bg-amber-50/90',
    badgeText: 'text-amber-900',
    badgeBorder: 'border-amber-300',
    tagBg: 'bg-amber-100 text-amber-900',
    tagText: 'text-amber-900',
    tagBorder: 'border-amber-300',
    previewBg: 'bg-gradient-to-br from-amber-50 via-yellow-50 to-orange-50/50',
    previewBorder: 'border-amber-300',
    previewText: 'text-amber-900',
    previewTag: 'bg-amber-100 text-amber-900 border-amber-300',
    dotColor: 'bg-amber-600',
  },
  red: {
    key: 'red',
    label: 'Đỏ hành chính (Quyết định)',
    hex: '#dc2626',
    badgeBg: 'bg-rose-50/90',
    badgeText: 'text-rose-800',
    badgeBorder: 'border-rose-300',
    tagBg: 'bg-rose-100 text-rose-800',
    tagText: 'text-rose-800',
    tagBorder: 'border-rose-300',
    previewBg: 'bg-gradient-to-br from-rose-50 via-red-50 to-orange-50/50',
    previewBorder: 'border-rose-300',
    previewText: 'text-rose-900',
    previewTag: 'bg-rose-100 text-rose-800 border-rose-300',
    dotColor: 'bg-rose-600',
  },
  purple: {
    key: 'purple',
    label: 'Tím hoa cà (Đề nghị)',
    hex: '#9333ea',
    badgeBg: 'bg-purple-50/90',
    badgeText: 'text-purple-800',
    badgeBorder: 'border-purple-300',
    tagBg: 'bg-purple-100 text-purple-800',
    tagText: 'text-purple-800',
    tagBorder: 'border-purple-300',
    previewBg: 'bg-gradient-to-br from-purple-50 via-fuchsia-50 to-indigo-50/50',
    previewBorder: 'border-purple-300',
    previewText: 'text-purple-900',
    previewTag: 'bg-purple-100 text-purple-800 border-purple-300',
    dotColor: 'bg-purple-600',
  },
  teal: {
    key: 'teal',
    label: 'Xanh lục lam (Nội dung)',
    hex: '#0d9488',
    badgeBg: 'bg-teal-50/90',
    badgeText: 'text-teal-800',
    badgeBorder: 'border-teal-300',
    tagBg: 'bg-teal-100 text-teal-800',
    tagText: 'text-teal-800',
    tagBorder: 'border-teal-300',
    previewBg: 'bg-gradient-to-br from-teal-50 via-cyan-50 to-emerald-50/50',
    previewBorder: 'border-teal-300',
    previewText: 'text-teal-900',
    previewTag: 'bg-teal-100 text-teal-800 border-teal-300',
    dotColor: 'bg-teal-600',
  },
  indigo: {
    key: 'indigo',
    label: 'Xanh chàm (Tờ trình)',
    hex: '#4f46e5',
    badgeBg: 'bg-indigo-50/90',
    badgeText: 'text-indigo-800',
    badgeBorder: 'border-indigo-300',
    tagBg: 'bg-indigo-100 text-indigo-800',
    tagText: 'text-indigo-800',
    tagBorder: 'border-indigo-300',
    previewBg: 'bg-gradient-to-br from-indigo-50 via-blue-50 to-purple-50/50',
    previewBorder: 'border-indigo-300',
    previewText: 'text-indigo-900',
    previewTag: 'bg-indigo-100 text-indigo-800 border-indigo-300',
    dotColor: 'bg-indigo-600',
  },
  orange: {
    key: 'orange',
    label: 'Cam nhiệt đới (Quy chế)',
    hex: '#ea580c',
    badgeBg: 'bg-orange-50/90',
    badgeText: 'text-orange-900',
    badgeBorder: 'border-orange-300',
    tagBg: 'bg-orange-100 text-orange-900',
    tagText: 'text-orange-900',
    tagBorder: 'border-orange-300',
    previewBg: 'bg-gradient-to-br from-orange-50 via-amber-50 to-red-50/50',
    previewBorder: 'border-orange-300',
    previewText: 'text-orange-900',
    previewTag: 'bg-orange-100 text-orange-900 border-orange-300',
    dotColor: 'bg-orange-600',
  },
  cyan: {
    key: 'cyan',
    label: 'Xanh lơ / da trời (Công văn)',
    hex: '#0891b2',
    badgeBg: 'bg-cyan-50/90',
    badgeText: 'text-cyan-800',
    badgeBorder: 'border-cyan-300',
    tagBg: 'bg-cyan-100 text-cyan-800',
    tagText: 'text-cyan-800',
    tagBorder: 'border-cyan-300',
    previewBg: 'bg-gradient-to-br from-cyan-50 via-sky-50 to-blue-50/50',
    previewBorder: 'border-cyan-300',
    previewText: 'text-cyan-900',
    previewTag: 'bg-cyan-100 text-cyan-800 border-cyan-300',
    dotColor: 'bg-cyan-600',
  },
  pink: {
    key: 'pink',
    label: 'Hồng sen (Biên bản)',
    hex: '#db2777',
    badgeBg: 'bg-pink-50/90',
    badgeText: 'text-pink-800',
    badgeBorder: 'border-pink-300',
    tagBg: 'bg-pink-100 text-pink-800',
    tagText: 'text-pink-800',
    tagBorder: 'border-pink-300',
    previewBg: 'bg-gradient-to-br from-pink-50 via-rose-50 to-purple-50/50',
    previewBorder: 'border-pink-300',
    previewText: 'text-pink-900',
    previewTag: 'bg-pink-100 text-pink-800 border-pink-300',
    dotColor: 'bg-pink-600',
  },
};

// Default mapping by type code
export const DEFAULT_TYPE_COLORS: Record<string, string> = {
  KH: 'blue',
  BC: 'emerald',
  ĐN: 'purple',
  DN: 'purple',
  ND: 'teal',
  TB: 'amber',
  QĐ: 'red',
  QD: 'red',
  TTR: 'indigo',
  CV: 'cyan',
  BB: 'pink',
  HD: 'cyan',
  QC: 'orange',
  CTR: 'emerald',
  VBD: 'purple',
};

export const SEMANTIC_STATUS_STYLES = {
  outgoing: {
    label: 'Văn bản đi',
    badge: 'bg-blue-50 text-blue-800 border-blue-200',
    strongBadge: 'bg-blue-100 text-blue-900 border-blue-300',
    text: 'text-blue-800',
    dot: 'bg-blue-600',
  },
  incoming: {
    label: 'Văn bản đến',
    badge: 'bg-purple-50 text-purple-800 border-purple-200',
    strongBadge: 'bg-purple-100 text-purple-900 border-purple-300',
    text: 'text-purple-800',
    dot: 'bg-purple-600',
  },
  saved: {
    label: 'Đã lưu lên Drive',
    badge: 'bg-emerald-50 text-emerald-800 border-emerald-200',
    strongBadge: 'bg-emerald-100 text-emerald-900 border-emerald-300',
    text: 'text-emerald-800',
    dot: 'bg-emerald-600',
  },
  pending: {
    label: 'Cần kiểm tra / Đang chờ',
    badge: 'bg-amber-50 text-amber-900 border-amber-300',
    strongBadge: 'bg-amber-100 text-amber-900 border-amber-300',
    text: 'text-amber-900',
    dot: 'bg-amber-600',
  },
  danger: {
    label: 'Lỗi / Đã hủy / Đã xóa',
    badge: 'bg-rose-50 text-rose-800 border-rose-200',
    strongBadge: 'bg-rose-100 text-rose-900 border-rose-300',
    text: 'text-rose-800',
    dot: 'bg-rose-600',
  },
} as const;

const COLOR_KEYS = Object.keys(DOCUMENT_COLOR_PALETTE);

/**
 * Get the color configuration for a given document type code and optional explicit color key
 */
export function getTypeColorStyle(typeCode: string, colorKey?: string): DocumentColorStyle {
  if (colorKey && DOCUMENT_COLOR_PALETTE[colorKey]) {
    return DOCUMENT_COLOR_PALETTE[colorKey];
  }

  const cleanCode = (typeCode || '').toUpperCase().trim();
  const defaultKey = DEFAULT_TYPE_COLORS[cleanCode];
  if (defaultKey && DOCUMENT_COLOR_PALETTE[defaultKey]) {
    return DOCUMENT_COLOR_PALETTE[defaultKey];
  }

  // Consistent fallback hash
  let hash = 0;
  for (let i = 0; i < cleanCode.length; i++) {
    hash = (hash << 5) - hash + cleanCode.charCodeAt(i);
    hash |= 0;
  }
  const index = Math.abs(hash) % COLOR_KEYS.length;
  return DOCUMENT_COLOR_PALETTE[COLOR_KEYS[index]] || DOCUMENT_COLOR_PALETTE.blue;
}
