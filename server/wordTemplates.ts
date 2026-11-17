import PizZip from 'pizzip';

export interface WordTemplatePlaceholderMeta {
  key: string;
  label: string;
  required: boolean;
}

export const SUPPORTED_WORD_PLACEHOLDERS: WordTemplatePlaceholderMeta[] = [
  { key: 'so_ky_hieu', label: 'Số/Ký hiệu văn bản', required: true },
  { key: 'don_vi', label: 'Đơn vị ban hành', required: false },
  { key: 'ten_van_ban', label: 'Tên văn bản', required: true },
  { key: 'trich_yeu', label: 'Trích yếu', required: false },
  { key: 'ngay_ban_hanh', label: 'Ngày ban hành', required: false },
  { key: 'nguoi_ky', label: 'Người ký', required: true },
  { key: 'chuc_vu', label: 'Chức vụ người ký', required: false },
];

const CANONICAL_PLACEHOLDER_SET = new Set(SUPPORTED_WORD_PLACEHOLDERS.map((p) => p.key));

function escapeXml(value: string): string {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function normalizePlaceholderKey(rawKey: string): string {
  return String(rawKey || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_');
}

/**
 * Khi soạn thảo trong Word, thẻ {{so_ky_hieu}} thường bị chia nhỏ qua nhiều thẻ XML
 * (ví dụ <w:t>{{</w:t></w:r><w:r><w:t>so_ky_hieu</w:t></w:r><w:r><w:t>}}</w:t>).
 * Hàm này xử lý từng đoạn <w:p>...</w:p> hoặc toàn bộ XML text node để gộp các thẻ {{...}} bị cắt đôi
 * mà vẫn bảo toàn định dạng văn bản gốc.
 */
function collapseSplitPlaceholdersInXml(xml: string): string {
  // Tìm các cụm bắt đầu bằng {{ cho tới }} dù bên trong có kẹp thẻ XML <...>
  return xml.replace(/\{\{(?:(?!(?:\{\{|\}\})).)*?\}\}/gs, (match) => {
    const innerText = match.replace(/<[^>]+>/g, '');
    const normalized = normalizePlaceholderKey(innerText.slice(2, -2));
    return `{{${normalized}}}`;
  });
}

function getTargetXmlPaths(zip: PizZip): string[] {
  return Object.keys(zip.files).filter((filePath) => {
    const file = zip.files[filePath];
    if (!file || file.dir) return false;
    return (
      filePath === 'word/document.xml' ||
      /^word\/header\d*\.xml$/i.test(filePath) ||
      /^word\/footer\d*\.xml$/i.test(filePath)
    );
  });
}

/**
 * Kiểm tra tệp .docx mẫu hợp lệ và trích danh sách thẻ {{...}} xuất hiện trong tài liệu.
 */
export function inspectDocxTemplate(buffer: Buffer): {
  placeholders: string[];
  unknownPlaceholders: string[];
} {
  let zip: PizZip;
  try {
    zip = new PizZip(buffer);
  } catch {
    throw new Error('Tệp Word (.docx) không hợp lệ hoặc bị hỏng cấu trúc.');
  }

  if (!zip.files['word/document.xml']) {
    throw new Error('Tệp .docx thiếu nội dung word/document.xml hợp lệ.');
  }

  const foundSet = new Set<string>();
  const xmlPaths = getTargetXmlPaths(zip);
  for (const xmlPath of xmlPaths) {
    const rawXml = zip.files[xmlPath].asText();
    const cleanedXml = collapseSplitPlaceholdersInXml(rawXml);
    const regex = /\{\{([a-z0-9_]+)\}\}/gi;
    let m: RegExpExecArray | null;
    while ((m = regex.exec(cleanedXml)) !== null) {
      foundSet.add(normalizePlaceholderKey(m[1]));
    }
  }

  const placeholders = Array.from(foundSet);
  const unknownPlaceholders = placeholders.filter((k) => !CANONICAL_PLACEHOLDER_SET.has(k));
  return { placeholders, unknownPlaceholders };
}

export interface FillDocxValues {
  so_ky_hieu?: string;
  don_vi?: string;
  ten_van_ban?: string;
  trich_yeu?: string;
  ngay_ban_hanh?: string;
  nguoi_ky?: string;
  chuc_vu?: string;
  [key: string]: string | undefined;
}

/**
 * Điền dữ liệu vào mẫu .docx. Khi isDraft=true sẽ đóng dấu "DỰ THẢO" rõ ràng và không dùng số chính thức.
 */
export function fillDocxTemplate(
  buffer: Buffer,
  values: FillDocxValues,
  options?: { isDraft?: boolean }
): {
  outputBuffer: Buffer;
  placeholders: string[];
  missingFields: Array<{ key: string; label: string }>;
} {
  let zip: PizZip;
  try {
    zip = new PizZip(buffer);
  } catch {
    throw new Error('Không thể đọc tệp mẫu Word (.docx).');
  }

  if (!zip.files['word/document.xml']) {
    throw new Error('Tệp mẫu Word (.docx) không hợp lệ.');
  }

  const isDraft = Boolean(options?.isDraft);
  const effectiveValues: Record<string, string> = {};
  for (const [k, v] of Object.entries(values)) {
    effectiveValues[normalizePlaceholderKey(k)] = typeof v === 'string' ? v.trim() : '';
  }

  if (isDraft) {
    effectiveValues.so_ky_hieu = 'DỰ THẢO';
  }

  const detectedSet = new Set<string>();
  const missingSet = new Set<string>();

  const xmlPaths = getTargetXmlPaths(zip);
  for (const xmlPath of xmlPaths) {
    const rawXml = zip.files[xmlPath].asText();
    let processedXml = collapseSplitPlaceholdersInXml(rawXml);

    processedXml = processedXml.replace(/\{\{([a-z0-9_]+)\}\}/gi, (_full, rawKey) => {
      const key = normalizePlaceholderKey(rawKey);
      detectedSet.add(key);
      const val = effectiveValues[key] ?? '';
      if (!val) {
        missingSet.add(key);
        return '';
      }
      return escapeXml(val);
    });

    if (isDraft && xmlPath === 'word/document.xml') {
      const draftBannerXml =
        '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="160"/></w:pPr>' +
        '<w:r><w:rPr><w:b/><w:color w:val="B91C1C"/><w:sz w:val="24"/></w:rPr>' +
        '<w:t>BẢN DỰ THẢO — CHƯA CẤP SỐ CHÍNH THỨC</w:t></w:r></w:p>';
      processedXml = processedXml.replace(/(<w:body[^>]*>)/i, `$1${draftBannerXml}`);
    }

    zip.file(xmlPath, processedXml);
  }

  const outputBuffer = zip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  }) as Buffer;

  const missingFields = Array.from(missingSet).map((key) => {
    const meta = SUPPORTED_WORD_PLACEHOLDERS.find((p) => p.key === key);
    return {
      key,
      label: meta ? meta.label : key,
    };
  });

  return {
    outputBuffer,
    placeholders: Array.from(detectedSet),
    missingFields,
  };
}
