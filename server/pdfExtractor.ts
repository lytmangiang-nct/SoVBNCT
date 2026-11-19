import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import zlib from 'zlib';
import { createRequire } from 'module';
import { execFile } from 'child_process';
import { pathToFileURL } from 'url';
import { promisify } from 'util';
import type {
  DateComparisonStatus,
  ExtractableFieldKey,
  ExtractedDateSource,
  FieldLocation,
  NormalizedBBox,
  PdfExtractionResult,
  PdfPageContent,
} from '../src/types.js';
import {
  getMaxWordFileSizeBytes,
  getMaxWordFileSizeMB,
  validatePdfFileBuffer,
} from './googleDrive.js';
import { normalizeDateOnlyDdMmYyyy, removeVietnameseDiacritics } from '../src/utils/fileNaming.js';

const execFileAsync = promisify(execFile);
const requireFromHere = createRequire(import.meta.url);

export class PdfUrlFetchError extends Error {
  public readonly statusCode: number;
  constructor(message: string, statusCode = 400) {
    super(message);
    this.name = 'PdfUrlFetchError';
    this.statusCode = statusCode;
  }
}

interface NormalizedBox {
  text: string;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  height: number;
  fontName?: string;
}

interface NormalizedLine {
  text: string;
  x0: number;
  x1: number;
  xCenter: number;
  y0: number;
  y1: number;
  yCenter: number;
  avgHeight: number;
  fontNames: string[];
  words: NormalizedBox[];
  isDivider: boolean;
}

interface PdfPageTextInfo {
  pageNumber: number;
  fullText: string;
  leftTopText: string;
  rightTopText: string;
  rightBottomText: string;
  rightHalfText: string;
  bottomText: string;
  annotationText: string;
  alphaNumCount: number;
  boxes: NormalizedBox[];
  dividerYs: number[];
}

function countMeaningfulChars(str: string): number {
  const matches = str.match(/[\p{L}\p{N}]/gu);
  return matches ? matches.length : 0;
}

function isGarbledText(str: string): boolean {
  if (!str) return true;
  const meaningful = countMeaningfulChars(str);
  if (meaningful < 35) return true;
  const replacementCount = (str.match(/\uFFFD/g) || []).length;
  return replacementCount > 5 || replacementCount / Math.max(str.length, 1) > 0.08;
}

function normalizeUrlForPdfFetch(rawUrl: string): string {
  const trimmed = rawUrl.trim();
  // Chuyển liên kết xem Google Drive sang liên kết tải trực tiếp
  const driveFileMatch = trimmed.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveFileMatch) {
    return `https://drive.google.com/uc?export=download&id=${driveFileMatch[1]}`;
  }
  const driveOpenMatch = trimmed.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (driveOpenMatch) {
    return `https://drive.google.com/uc?export=download&id=${driveOpenMatch[1]}`;
  }
  return trimmed;
}

export async function fetchPdfFromUrl(rawUrl: string): Promise<{
  buffer: Buffer;
  fileName: string;
  finalUrl: string;
}> {
  let parsed: URL;
  try {
    parsed = new URL((rawUrl || '').trim());
  } catch {
    throw new PdfUrlFetchError(
      'Đường dẫn PDF không đúng định dạng URL (cần bắt đầu bằng http:// hoặc https://).'
    );
  }

  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    throw new PdfUrlFetchError('Chỉ hỗ trợ đường dẫn giao thức http:// hoặc https://.');
  }

  // Chặn truy cập metadata nội bộ đám mây
  const blockedHosts = ['169.254.169.254', 'metadata.google.internal', 'metadata.goog'];
  if (blockedHosts.includes(parsed.hostname.toLowerCase())) {
    throw new PdfUrlFetchError('Đường dẫn không hợp lệ hoặc bị từ chối vì lý do bảo mật.');
  }

  const targetUrl = normalizeUrlForPdfFetch(parsed.toString());
  const maxBytes = getMaxWordFileSizeBytes();
  const maxMB = getMaxWordFileSizeMB();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20000);

  try {
    const res = await fetch(targetUrl, {
      method: 'GET',
      redirect: 'follow',
      signal: controller.signal,
      headers: {
        'User-Agent': 'NCT-DocumentRegistry/1.0 (+PDF-Import)',
        Accept: 'application/pdf,application/octet-stream,*/*;q=0.8',
      },
    });

    if (res.status === 401 || res.status === 403) {
      throw new PdfUrlFetchError(
        `Đường dẫn yêu cầu đăng nhập hoặc không có quyền truy cập (HTTP ${res.status}). Vui lòng tải tệp PDF về máy và dùng cách "Tải PDF từ máy".`
      );
    }
    if (!res.ok) {
      throw new PdfUrlFetchError(
        `Không thể tải tệp từ đường dẫn (HTTP ${res.status}). Vui lòng kiểm tra lại liên kết hoặc tải PDF về máy để nhập thủ công.`
      );
    }

    const contentLengthHeader = res.headers.get('content-length');
    if (contentLengthHeader) {
      const expectedLen = parseInt(contentLengthHeader, 10);
      if (Number.isFinite(expectedLen) && expectedLen > maxBytes) {
        throw new PdfUrlFetchError(`Dung lượng tệp PDF từ đường dẫn vượt quá giới hạn tối đa ${maxMB} MB.`);
      }
    }

    const contentType = (res.headers.get('content-type') || '').toLowerCase();
    if (contentType.includes('text/html')) {
      throw new PdfUrlFetchError(
        'Đường dẫn trả về trang web hoặc trang đăng nhập thay vì tệp PDF trực tiếp. Vui lòng tải tệp PDF về máy rồi dùng cách "Tải PDF từ máy".'
      );
    }

    if (!res.body) {
      throw new PdfUrlFetchError('Máy chủ nguồn không trả về dữ liệu tệp.');
    }

    const reader = res.body.getReader();
    const chunks: Uint8Array[] = [];
    let totalBytes = 0;

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value) {
        totalBytes += value.byteLength;
        if (totalBytes > maxBytes) {
          controller.abort();
          throw new PdfUrlFetchError(`Dung lượng tệp PDF từ đường dẫn vượt quá giới hạn tối đa ${maxMB} MB.`);
        }
        chunks.push(value);
      }
    }

    const buffer = Buffer.concat(chunks.map((c) => Buffer.from(c)));
    const cd = res.headers.get('content-disposition') || '';
    let fileName = '';
    const cdUtf8 = cd.match(/filename\*=UTF-8''([^;\n]+)/i);
    const cdPlain = cd.match(/filename="?([^";\n]+)"?/i);
    if (cdUtf8) {
      fileName = decodeURIComponent(cdUtf8[1].trim());
    } else if (cdPlain) {
      fileName = cdPlain[1].trim();
    } else {
      const pathPart = new URL(res.url || targetUrl).pathname.split('/').filter(Boolean).pop() || '';
      fileName = decodeURIComponent(pathPart);
    }
    if (!fileName || !fileName.toLowerCase().endsWith('.pdf')) {
      fileName = fileName ? `${fileName}.pdf` : 'van-ban-nhap.pdf';
    }

    const validation = validatePdfFileBuffer(fileName, contentType, buffer);
    if (!validation.valid) {
      throw new PdfUrlFetchError(
        validation.error + ' Vui lòng kiểm tra lại đường dẫn hoặc tải tệp PDF về máy để nhập thủ công.'
      );
    }

    return {
      buffer,
      fileName,
      finalUrl: res.url || targetUrl,
    };
  } catch (err: any) {
    if (err instanceof PdfUrlFetchError) throw err;
    if (err?.name === 'AbortError') {
      throw new PdfUrlFetchError(
        'Quá thời gian tải PDF từ đường dẫn (20 giây). Vui lòng tải tệp PDF về máy và dùng cách "Tải PDF từ máy".'
      );
    }
    throw new PdfUrlFetchError(
      err?.message || 'Không thể tải tệp PDF từ đường dẫn. Vui lòng kiểm tra lại liên kết hoặc tải PDF từ máy.'
    );
  } finally {
    clearTimeout(timeout);
  }
}

function clamp01(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Math.max(0, Math.min(1, val));
}

function isDividerLineText(str: string): boolean {
  const t = (str || '').trim();
  if (!t) return true;
  if (/^[\s_\-—–=~*•·.,|\\\/]{2,}$/.test(t) && countMeaningfulChars(t) === 0) {
    return true;
  }
  return countMeaningfulChars(t) < 2 && t.length <= 4;
}

function groupBoxesIntoLines(boxes: NormalizedBox[]): NormalizedLine[] {
  if (boxes.length === 0) return [];
  const sorted = [...boxes].sort((a, b) => {
    const ay = (a.y0 + a.y1) / 2;
    const by = (b.y0 + b.y1) / 2;
    if (Math.abs(ay - by) > 0.004) return ay - by;
    return a.x0 - b.x0;
  });

  const rows: NormalizedBox[][] = [];
  for (const box of sorted) {
    const boxYCenter = (box.y0 + box.y1) / 2;
    let matchedRow: NormalizedBox[] | null = null;
    let bestScore = Infinity;

    for (let i = Math.max(0, rows.length - 4); i < rows.length; i++) {
      const row = rows[i];
      const rowYCenter =
        row.reduce((acc, w) => acc + (w.y0 + w.y1) / 2, 0) / row.length;
      const rowY0 = Math.min(...row.map((w) => w.y0));
      const rowY1 = Math.max(...row.map((w) => w.y1));
      const rowAvgH = row.reduce((acc, w) => acc + w.height, 0) / row.length;
      const minH = Math.max(0.005, Math.min(rowAvgH, box.height));
      const tol = Math.max(0.006, 0.58 * minH);

      const overlap = Math.min(rowY1, box.y1) - Math.max(rowY0, box.y0);
      const overlapRatio = overlap / minH;
      const dy = Math.abs(boxYCenter - rowYCenter);

      if (dy <= tol || overlapRatio >= 0.42) {
        if (dy < bestScore) {
          bestScore = dy;
          matchedRow = row;
        }
      }
    }

    if (matchedRow) {
      matchedRow.push(box);
    } else {
      rows.push([box]);
    }
  }

  const lines: NormalizedLine[] = [];
  for (const row of rows) {
    row.sort((a, b) => a.x0 - b.x0);
    const avgHeight = row.reduce((acc, w) => acc + w.height, 0) / row.length;

    let joined = '';
    for (let i = 0; i < row.length; i++) {
      const cur = row[i];
      const t = cur.text.trim();
      if (!t) continue;
      if (!joined) {
        joined = t;
        continue;
      }
      const prev = row[i - 1];
      const gap = cur.x0 - prev.x1;
      const hasPrevWidth = prev.x1 > prev.x0 + 0.0002;
      if (hasPrevWidth && gap <= 0.0012 && gap >= -0.005) {
        joined += t;
      } else {
        joined += ` ${t}`;
      }
    }

    const text = joined.replace(/\s+/g, ' ').trim();
    if (!text) continue;

    const x0 = Math.min(...row.map((w) => w.x0));
    const x1 = Math.max(...row.map((w) => w.x1));
    const y0 = Math.min(...row.map((w) => w.y0));
    const y1 = Math.max(...row.map((w) => w.y1));
    const fontNames = Array.from(
      new Set(row.map((w) => w.fontName).filter((f): f is string => Boolean(f)))
    );

    lines.push({
      text,
      x0,
      x1,
      xCenter: (x0 + x1) / 2,
      y0,
      y1,
      yCenter: (y0 + y1) / 2,
      avgHeight,
      fontNames,
      words: row,
      isDivider: isDividerLineText(text),
    });
  }

  lines.sort((a, b) => a.yCenter - b.yCenter);
  return lines;
}

function boxesToText(boxes: NormalizedBox[]): string {
  return groupBoxesIntoLines(boxes)
    .map((l) => l.text)
    .filter(Boolean)
    .join('\n');
}

async function detectHorizontalDividerYsFromPdfOps(
  page: any,
  pdfjsOps: any,
  viewportWidth: number,
  viewportHeight: number
): Promise<number[]> {
  const dividerYs: number[] = [];
  if (!page || typeof page.getOperatorList !== 'function' || !viewportWidth || !viewportHeight) {
    return dividerYs;
  }

  try {
    const opList = await page.getOperatorList();
    const fnArray: number[] = opList?.fnArray || [];
    const argsArray: any[] = opList?.argsArray || [];

    const checkHorizontalSegment = (xA: number, yA: number, xB: number, yB: number) => {
      const x0Frac = Math.min(xA, xB) / viewportWidth;
      const x1Frac = Math.max(xA, xB) / viewportWidth;
      const yTopFrac = 1 - Math.max(yA, yB) / viewportHeight;
      const yBotFrac = 1 - Math.min(yA, yB) / viewportHeight;
      const wFrac = x1Frac - x0Frac;
      const hFrac = Math.abs(yBotFrac - yTopFrac);
      const xCenterFrac = (x0Frac + x1Frac) / 2;
      const yCenterFrac = (yTopFrac + yBotFrac) / 2;

      if (
        yCenterFrac >= 0.04 &&
        yCenterFrac <= 0.42 &&
        xCenterFrac <= 0.52 &&
        wFrac >= 0.05 &&
        wFrac <= 0.46 &&
        hFrac <= 0.008
      ) {
        dividerYs.push(yCenterFrac);
      }
    };

    let curX = 0;
    let curY = 0;

    for (let i = 0; i < fnArray.length; i++) {
      const fn = fnArray[i];
      const args = argsArray[i];
      if (!args) continue;

      if (fn === pdfjsOps?.constructPath && Array.isArray(args[0]) && Array.isArray(args[1])) {
        const subOps: number[] = args[0];
        const coords: number[] = args[1];
        let ci = 0;
        let lx = 0;
        let ly = 0;
        for (const sop of subOps) {
          if (sop === pdfjsOps.moveTo && ci + 1 < coords.length) {
            lx = coords[ci++];
            ly = coords[ci++];
          } else if (sop === pdfjsOps.lineTo && ci + 1 < coords.length) {
            const nx = coords[ci++];
            const ny = coords[ci++];
            checkHorizontalSegment(lx, ly, nx, ny);
            lx = nx;
            ly = ny;
          } else if (sop === pdfjsOps.rectangle && ci + 3 < coords.length) {
            const rx = coords[ci++];
            const ry = coords[ci++];
            const rw = coords[ci++];
            const rh = coords[ci++];
            checkHorizontalSegment(rx, ry, rx + rw, ry + rh);
          } else if (sop === 15 || sop === 16 || sop === 17) {
            // curveTo / curveTo2 / curveTo3
            ci += sop === 15 ? 6 : 4;
          }
        }
      } else if (fn === pdfjsOps?.moveTo && args.length >= 2) {
        curX = Number(args[0]) || 0;
        curY = Number(args[1]) || 0;
      } else if (fn === pdfjsOps?.lineTo && args.length >= 2) {
        const nx = Number(args[0]) || 0;
        const ny = Number(args[1]) || 0;
        checkHorizontalSegment(curX, curY, nx, ny);
        curX = nx;
        curY = ny;
      } else if (fn === pdfjsOps?.rectangle && args.length >= 4) {
        const rx = Number(args[0]) || 0;
        const ry = Number(args[1]) || 0;
        const rw = Number(args[2]) || 0;
        const rh = Number(args[3]) || 0;
        checkHorizontalSegment(rx, ry, rx + rw, ry + rh);
      }
    }
  } catch {
    // ignore operator list errors
  }

  return Array.from(new Set(dividerYs.map((y) => Math.round(y * 1000) / 1000))).sort((a, b) => a - b);
}

let serverPdfJsLoadPromise: Promise<{
  pdfjs: any;
  standardFontDataUrl?: string;
  cMapUrl?: string;
  workerResolvedPath: string;
}> | null = null;

export function getServerPdfWorkerDiagnostics(): {
  workerPath: string;
  workerExists: boolean;
  standardFontsExists: boolean;
  cmapsExists: boolean;
  pdfjsVersion: string;
} {
  let workerPath = '';
  let workerExists = false;
  let standardFontsExists = false;
  let cmapsExists = false;
  let pdfjsVersion = '';
  try {
    workerPath = requireFromHere.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
    workerExists = fs.existsSync(workerPath);
  } catch {
    workerPath = path.resolve(process.cwd(), 'node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs');
    workerExists = fs.existsSync(workerPath);
  }
  try {
    const pkgPath = requireFromHere.resolve('pdfjs-dist/package.json');
    const pkgDir = path.dirname(pkgPath);
    pdfjsVersion = JSON.parse(fs.readFileSync(pkgPath, 'utf8'))?.version || '';
    standardFontsExists = fs.existsSync(path.join(pkgDir, 'standard_fonts'));
    cmapsExists = fs.existsSync(path.join(pkgDir, 'cmaps'));
  } catch {
    // ignore
  }
  return {
    workerPath,
    workerExists,
    standardFontsExists,
    cmapsExists,
    pdfjsVersion,
  };
}

async function loadServerPdfJs() {
  if (!serverPdfJsLoadPromise) {
    serverPdfJsLoadPromise = (async () => {
      // Import tường minh cả pdf.mjs và pdf.worker.mjs cùng phiên bản để trình đóng gói Vercel (@vercel/nft)
      // luôn theo dõi và đưa tệp worker vào gói Serverless Function (/var/task/node_modules/pdfjs-dist/legacy/build/)
      const [pdfjs, pdfjsWorker] = await Promise.all([
        import('pdfjs-dist/legacy/build/pdf.mjs'),
        // @ts-ignore pdfjs-dist legacy worker module has no .d.mts declarations
        import('pdfjs-dist/legacy/build/pdf.worker.mjs'),
      ]);

      // Cấu hình worker riêng cho runtime máy chủ (tách biệt hoàn toàn khỏi PdfPreviewPane phía trình duyệt)
      if (pdfjsWorker?.WorkerMessageHandler) {
        (globalThis as any).pdfjsWorker = pdfjsWorker;
      }

      let workerResolvedPath = '';
      try {
        workerResolvedPath = requireFromHere.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
        if (workerResolvedPath && fs.existsSync(workerResolvedPath) && pdfjs?.GlobalWorkerOptions) {
          pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(workerResolvedPath).href;
        }
      } catch {
        // Fallback sử dụng globalThis.pdfjsWorker đã nạp trực tiếp ở trên
      }

      let standardFontDataUrl: string | undefined;
      let cMapUrl: string | undefined;
      try {
        const pdfPkgDir = path.dirname(requireFromHere.resolve('pdfjs-dist/package.json'));
        const fontsDir = path.join(pdfPkgDir, 'standard_fonts');
        const cmapsDir = path.join(pdfPkgDir, 'cmaps');
        if (fs.existsSync(fontsDir)) {
          standardFontDataUrl = pathToFileURL(fontsDir + path.sep).href;
        }
        if (fs.existsSync(cmapsDir)) {
          cMapUrl = pathToFileURL(cmapsDir + path.sep).href;
        }
      } catch {
        // ignore optional font/cmap assets
      }

      return { pdfjs, standardFontDataUrl, cMapUrl, workerResolvedPath };
    })().catch((err) => {
      serverPdfJsLoadPromise = null;
      throw err;
    });
  }
  return serverPdfJsLoadPromise;
}

async function extractPdfTextWithPdfJs(buffer: Buffer): Promise<{
  pageCount: number;
  pages: PdfPageTextInfo[];
  rawSigDates: string[];
  rawSigSigners: string[];
}> {
  const { pdfjs, standardFontDataUrl, cMapUrl } = await loadServerPdfJs();
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(buffer),
    verbosity: 0,
    useSystemFonts: true,
    isEvalSupported: false,
    ...(standardFontDataUrl ? { standardFontDataUrl } : {}),
    ...(cMapUrl ? { cMapUrl, cMapPacked: true } : {}),
  });

  const pdfDoc = await loadingTask.promise;
  const pageCount = pdfDoc.numPages || 1;
  const pages: PdfPageTextInfo[] = [];

  for (let p = 1; p <= pageCount; p++) {
    const page = await pdfDoc.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();

    const allBoxes: NormalizedBox[] = [];
    const leftTopBoxes: NormalizedBox[] = [];
    const rightTopBoxes: NormalizedBox[] = [];
    const rightBottomBoxes: NormalizedBox[] = [];
    const rightHalfBoxes: NormalizedBox[] = [];
    const bottomBoxes: NormalizedBox[] = [];

    for (const item of textContent.items as any[]) {
      if (!item || typeof item.str !== 'string' || !item.str.trim()) continue;
      const txX = Number(item.transform?.[4] ?? 0);
      const txY = Number(item.transform?.[5] ?? 0);
      const itemW = Math.max(0, Number(item.width ?? 0));
      const itemH = Math.max(4, Math.abs(Number(item.height || item.transform?.[3] || 10)));

      const x0 = clamp01(txX / viewport.width);
      const x1 = clamp01((txX + itemW) / viewport.width);
      const y1 = clamp01(1 - txY / viewport.height);
      const y0 = clamp01(1 - (txY + itemH) / viewport.height);
      const height = Math.max(0.004, y1 - y0);
      const yCenter = (y0 + y1) / 2;
      const xCenter = (x0 + x1) / 2;

      const box: NormalizedBox = {
        text: item.str,
        x0,
        x1: Math.max(x0 + 0.001, x1),
        y0,
        y1: Math.max(y0 + 0.004, y1),
        height,
        fontName: typeof item.fontName === 'string' ? item.fontName : undefined,
      };
      allBoxes.push(box);

      if (yCenter <= 0.45) {
        if (xCenter < 0.52) leftTopBoxes.push(box);
        if (x0 >= 0.45) rightTopBoxes.push(box);
      }
      if (yCenter >= 0.35) {
        bottomBoxes.push(box);
        if (x0 >= 0.40) rightBottomBoxes.push(box);
      }
      if (x0 >= 0.40) {
        rightHalfBoxes.push(box);
      }
    }

    const dividerYs =
      p === 1
        ? await detectHorizontalDividerYsFromPdfOps(page, pdfjs.OPS, viewport.width, viewport.height)
        : [];

    let annotationText = '';
    try {
      const annots = await page.getAnnotations();
      const annotParts: string[] = [];
      for (const a of annots || []) {
        if (a.contentsObj?.str) annotParts.push(a.contentsObj.str);
        else if (typeof a.contents === 'string' && a.contents) annotParts.push(a.contents);
        if (typeof a.fieldValue === 'string' && a.fieldValue) annotParts.push(a.fieldValue);
      }
      annotationText = annotParts.join('\n').trim();
    } catch {
      // ignore annotation read errors
    }

    const fullText = boxesToText(allBoxes);
    pages.push({
      pageNumber: p,
      fullText,
      leftTopText: boxesToText(leftTopBoxes),
      rightTopText: boxesToText(rightTopBoxes),
      rightBottomText: boxesToText(rightBottomBoxes),
      rightHalfText: boxesToText(rightHalfBoxes),
      bottomText: boxesToText(bottomBoxes),
      annotationText,
      alphaNumCount: countMeaningfulChars(fullText),
      boxes: allBoxes,
      dividerYs,
    });
  }

  // Quét metadata chữ ký số nhúng trong raw PDF (/M (D:YYYYMMDD...) và /Name (...)) — chỉ lấy ngày/tháng/năm
  const rawLatin = buffer.toString('latin1');
  const rawSigDates: string[] = [];
  const rawSigSigners: string[] = [];

  const sigBlockMatches = rawLatin.match(/\/Type\s*\/Sig[\s\S]{0,800}/g) || [];
  for (const blk of sigBlockMatches) {
    const mMatch = blk.match(/\/M\s*\(D:(\d{4})(\d{2})(\d{2})/);
    if (mMatch) {
      const [, yyyy, mm, dd] = mMatch;
      const dateOnly = formatDateDdMmYyyy(dd, mm, yyyy);
      if (dateOnly) rawSigDates.push(dateOnly);
    }
    const nameMatch = blk.match(/\/Name\s*\(([^)]{2,100})\)/);
    if (nameMatch && !nameMatch[1].startsWith('þÿ')) {
      rawSigSigners.push(nameMatch[1].trim());
    }
  }

  if (typeof pdfDoc.destroy === 'function') {
    await pdfDoc.destroy().catch(() => {});
  } else if (typeof loadingTask.destroy === 'function') {
    await loadingTask.destroy().catch(() => {});
  }
  return { pageCount, pages, rawSigDates, rawSigSigners };
}

function getPngDimensions(pngPath: string): { width: number; height: number } {
  try {
    const fd = fs.openSync(pngPath, 'r');
    const hdr = Buffer.alloc(24);
    fs.readSync(fd, hdr, 0, 24, 0);
    fs.closeSync(fd);
    if (hdr.toString('ascii', 1, 4) === 'PNG') {
      const width = hdr.readUInt32BE(16);
      const height = hdr.readUInt32BE(20);
      if (width > 0 && height > 0) return { width, height };
    }
  } catch {
    // ignore
  }
  return { width: 1654, height: 2339 };
}

function extractBoxesFromTesseractData(
  data: any,
  imgWidth: number,
  imgHeight: number
): NormalizedBox[] {
  const boxes: NormalizedBox[] = [];
  if (!data?.blocks || !Array.isArray(data.blocks)) return boxes;

  for (const block of data.blocks) {
    for (const para of block?.paragraphs || []) {
      for (const line of para?.lines || []) {
        for (const word of line?.words || []) {
          const text = (word?.text || '').trim();
          const bbox = word?.bbox;
          if (!text || !bbox) continue;
          const x0 = clamp01(Number(bbox.x0) / imgWidth);
          const x1 = clamp01(Number(bbox.x1) / imgWidth);
          const y0 = clamp01(Number(bbox.y0) / imgHeight);
          const y1 = clamp01(Number(bbox.y1) / imgHeight);
          if (x1 <= x0 || y1 <= y0) continue;
          boxes.push({
            text,
            x0,
            x1,
            y0,
            y1,
            height: Math.max(0.004, y1 - y0),
          });
        }
      }
    }
  }
  return boxes;
}

async function renderPdfPageToPng(
  pdfPath: string,
  pageNumber: number,
  outPngPath: string,
  dpi = 200
): Promise<void> {
  await execFileAsync('gs', [
    '-dNOPAUSE',
    '-dBATCH',
    '-dQUIET',
    '-sDEVICE=png16m',
    `-r${dpi}`,
    `-dFirstPage=${pageNumber}`,
    `-dLastPage=${pageNumber}`,
    `-sOutputFile=${outPngPath}`,
    pdfPath,
  ]);
}

const PNG_CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function pngCrc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = PNG_CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function makePngChunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, 'ascii');
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcInput = Buffer.concat([typeBuf, data]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(pngCrc32(crcInput), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function cropRgbPngRegion(
  pngPath: string,
  outCropPath: string,
  xFrac0: number,
  yFrac0: number,
  xFrac1: number,
  yFrac1: number
): { width: number; height: number } | null {
  try {
    const fileBuf = fs.readFileSync(pngPath);
    if (fileBuf.length < 32 || fileBuf.toString('ascii', 1, 4) !== 'PNG') return null;

    let offset = 8;
    let width = 0;
    let height = 0;
    let bitDepth = 8;
    let colorType = 2;
    const idatChunks: Buffer[] = [];

    while (offset + 8 <= fileBuf.length) {
      const len = fileBuf.readUInt32BE(offset);
      const type = fileBuf.toString('ascii', offset + 4, offset + 8);
      const dataStart = offset + 8;
      const dataEnd = dataStart + len;
      if (dataEnd + 4 > fileBuf.length) break;
      const data = fileBuf.subarray(dataStart, dataEnd);
      if (type === 'IHDR' && data.length >= 13) {
        width = data.readUInt32BE(0);
        height = data.readUInt32BE(4);
        bitDepth = data[8];
        colorType = data[9];
      } else if (type === 'IDAT') {
        idatChunks.push(data);
      } else if (type === 'IEND') {
        break;
      }
      offset = dataEnd + 4;
    }

    if (!width || !height || bitDepth !== 8 || (colorType !== 2 && colorType !== 6) || idatChunks.length === 0) {
      return null;
    }

    const bpp = colorType === 6 ? 4 : 3;
    const stride = width * bpp;
    const raw = zlib.inflateSync(Buffer.concat(idatChunks));
    if (raw.length < height * (stride + 1)) return null;

    const recon = Buffer.alloc(height * stride);
    const paethPredictor = (a: number, b: number, c: number) => {
      const p = a + b - c;
      const pa = Math.abs(p - a);
      const pb = Math.abs(p - b);
      const pc = Math.abs(p - c);
      if (pa <= pb && pa <= pc) return a;
      if (pb <= pc) return b;
      return c;
    };

    for (let y = 0; y < height; y++) {
      const filterType = raw[y * (stride + 1)];
      const rowInStart = y * (stride + 1) + 1;
      const rowOutStart = y * stride;
      const prevOutStart = (y - 1) * stride;
      for (let x = 0; x < stride; x++) {
        const val = raw[rowInStart + x];
        const a = x >= bpp ? recon[rowOutStart + x - bpp] : 0;
        const b = y > 0 ? recon[prevOutStart + x] : 0;
        const c = y > 0 && x >= bpp ? recon[prevOutStart + x - bpp] : 0;
        let resByte = val;
        if (filterType === 1) resByte = (val + a) & 0xff;
        else if (filterType === 2) resByte = (val + b) & 0xff;
        else if (filterType === 3) resByte = (val + ((a + b) >> 1)) & 0xff;
        else if (filterType === 4) resByte = (val + paethPredictor(a, b, c)) & 0xff;
        recon[rowOutStart + x] = resByte;
      }
    }

    const cx0 = Math.max(0, Math.min(width - 1, Math.floor(xFrac0 * width)));
    const cy0 = Math.max(0, Math.min(height - 1, Math.floor(yFrac0 * height)));
    const cx1 = Math.max(cx0 + 1, Math.min(width, Math.ceil(xFrac1 * width)));
    const cy1 = Math.max(cy0 + 1, Math.min(height, Math.ceil(yFrac1 * height)));
    const cropW = cx1 - cx0;
    const cropH = cy1 - cy0;
    const cropStride = cropW * bpp;
    const cropRaw = Buffer.alloc(cropH * (cropStride + 1));

    for (let y = 0; y < cropH; y++) {
      const srcOffset = (cy0 + y) * stride + cx0 * bpp;
      const dstOffset = y * (cropStride + 1);
      cropRaw[dstOffset] = 0; // None filter
      recon.copy(cropRaw, dstOffset + 1, srcOffset, srcOffset + cropStride);
    }

    const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(cropW, 0);
    ihdr.writeUInt32BE(cropH, 4);
    ihdr[8] = 8;
    ihdr[9] = colorType;
    ihdr[10] = 0;
    ihdr[11] = 0;
    ihdr[12] = 0;
    const compressed = zlib.deflateSync(cropRaw, { level: 4 });
    const outPng = Buffer.concat([
      sig,
      makePngChunk('IHDR', ihdr),
      makePngChunk('IDAT', compressed),
      makePngChunk('IEND', Buffer.alloc(0)),
    ]);
    fs.writeFileSync(outCropPath, outPng);
    return { width: cropW, height: cropH };
  } catch {
    return null;
  }
}

interface OcrPageResult {
  text: string;
  confidence: number;
  boxes: NormalizedBox[];
}

async function extractEmbeddedJpegFromPdfPage(
  buffer: Buffer,
  pageNumber: number
): Promise<Buffer | null> {
  try {
    const { PDFDocument, PDFDict, PDFName, PDFRawStream } = await import('pdf-lib');
    const doc = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const pages = doc.getPages();
    const page = pages[pageNumber - 1];
    if (!page) return null;
    const resources = page.node.Resources();
    if (!resources) return null;
    const xObject = resources.lookup(PDFName.of('XObject'));
    if (!(xObject instanceof PDFDict)) return null;

    let bestJpeg: Buffer | null = null;
    for (const [, val] of xObject.entries()) {
      const resolved = doc.context.lookup(val);
      if (resolved instanceof PDFRawStream) {
        const subtype = resolved.dict.lookup(PDFName.of('Subtype'));
        const filter = resolved.dict.lookup(PDFName.of('Filter'));
        const isImage = subtype === PDFName.of('Image');
        const filterText = filter?.toString?.() || '';
        if (isImage && filterText.includes('DCTDecode') && resolved.contents.length > 0) {
          const candidate = Buffer.from(resolved.contents);
          if (!bestJpeg || candidate.length > bestJpeg.length) {
            bestJpeg = candidate;
          }
        }
      }
    }
    return bestJpeg;
  } catch {
    return null;
  }
}

async function runVietnameseOcrOnPages(
  buffer: Buffer,
  pageNumbers: number[]
): Promise<{
  results: Map<number, OcrPageResult>;
  ocrUnavailableReason?: string;
}> {
  const results = new Map<number, OcrPageResult>();
  if (pageNumbers.length === 0) return { results };

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'nct-pdf-ocr-'));
  const pdfPath = path.join(tmpDir, 'input.pdf');
  fs.writeFileSync(pdfPath, buffer);

  let worker: any = null;
  let ocrUnavailableReason: string | undefined;
  try {
    const { createWorker } = await import('tesseract.js');
    const langPath = path.resolve(process.cwd(), 'node_modules/@tesseract.js-data/vie/4.0.0');
    const workerOpts: any = fs.existsSync(path.join(langPath, 'vie.traineddata.gz'))
      ? { langPath, gzip: true }
      : {};

    worker = await createWorker('vie', 1, workerOpts);

    for (const p of pageNumbers) {
      let imagePathToRecognize = '';
      const outPng = path.join(tmpDir, `page-${p}.png`);
      try {
        await renderPdfPageToPng(pdfPath, p, outPng);
        if (fs.existsSync(outPng)) {
          imagePathToRecognize = outPng;
        }
      } catch (renderErr: any) {
        // Trên môi trường serverless không cài Ghostscript (gs), fallback trích ảnh JPEG nhúng trong trang scan
        const jpegBuf = await extractEmbeddedJpegFromPdfPage(buffer, p);
        if (jpegBuf) {
          const outJpg = path.join(tmpDir, `page-${p}.jpg`);
          fs.writeFileSync(outJpg, jpegBuf);
          imagePathToRecognize = outJpg;
        } else {
          console.warn(
            `[PDF OCR] Không thể chuyển trang ${p} sang ảnh để OCR (${renderErr?.code || renderErr?.message || 'render_error'}).`
          );
          ocrUnavailableReason =
            'Bản PDF dạng scan trên máy chủ hiện cần kiểm tra trực tiếp từ bản gốc bên cạnh; vui lòng đối chiếu và bổ sung thông tin còn thiếu trước khi lưu.';
        }
      }

      if (imagePathToRecognize && fs.existsSync(imagePathToRecognize)) {
        const dims = imagePathToRecognize.endsWith('.png')
          ? getPngDimensions(imagePathToRecognize)
          : { width: 1654, height: 2339 };
        const ret = await worker.recognize(imagePathToRecognize, {}, { blocks: true });
        const boxes = extractBoxesFromTesseractData(ret?.data, dims.width, dims.height);
        const rawText = (ret?.data?.text || '').trim();
        const reconstructed = boxes.length > 0 ? boxesToText(boxes) : rawText;
        results.set(p, {
          text: reconstructed || rawText,
          confidence: Number(ret?.data?.confidence ?? 0),
          boxes,
        });
      }
    }
  } catch (ocrErr: any) {
    console.warn('[PDF OCR] Lỗi chạy OCR tiếng Việt, giữ lại kết quả đọc PDF & xác nhận thủ công:', ocrErr?.message);
    ocrUnavailableReason =
      'Không thể chạy OCR tự động cho bản scan này; vui lòng đối chiếu bản PDF gốc bên cạnh và kiểm tra các trường thông tin.';
  } finally {
    if (worker) {
      await worker.terminate().catch(() => {});
    }
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }

  return { results, ocrUnavailableReason };
}

function formatDateDdMmYyyy(day: string, month: string, year: string): string {
  const d = day.padStart(2, '0');
  const m = month.padStart(2, '0');
  const y = year.length === 2 ? `20${year}` : year;
  return normalizeDateOnlyDdMmYyyy(`${d}/${m}/${y}`);
}

const BODY_DATE_CONTEXT_REGEX =
  /\b(?:can\s+cu|dieu\s+\d+|khoan\s+\d+|diem\s+[a-z]|ngay\s+hieu\s+luc|co\s+hieu\s+luc|ke\s+tu\s+ngay|han\s+(?:nop|cuoi)|cham\s+nhat|truoc\s+ngay|tu\s+ngay|den\s+(?:het\s+)?ngay|lich\s+trinh|thoi\s+gian\s+(?:to\s+chuc|nhan|nop|thi)|quyet\s+dinh\s+so|cong\s+van\s+so|thong\s+tu\s+so|nghi\s+dinh\s+so|ke\s+hoach\s+so|ban\s+hanh\s+kem\s+theo)\b/i;

function extractDateFromHeaderOrSigLine(lineText: string): string {
  if (!lineText) return '';
  const norm = removeVietnameseDiacritics(lineText);
  if (BODY_DATE_CONTEXT_REGEX.test(norm)) return '';

  // 1. "ngày10 tháng 11 năm 2026" hoặc "An Giang, ngày 10 tháng 11 năm 2026"
  const mText = norm.match(/\bngay\s*(\d{1,2})\s*thang\s*(\d{1,2})\s*nam\s*(\d{4})\b/i);
  if (mText) {
    return formatDateDdMmYyyy(mText[1], mText[2], mText[3]);
  }
  // 2. "ngày 10/11/2026" hoặc "ngày 10-11-2026"
  const mSlash = norm.match(/\bngay\s*[:\-]?\s*(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/i);
  if (mSlash) {
    return formatDateDdMmYyyy(mSlash[1], mSlash[2], mSlash[3]);
  }
  return '';
}

function extractTopRightHeaderDate(
  page1Boxes: NormalizedBox[],
  rightTopText: string,
  splitX: number
): ExtractedDateSource | null {
  const topRightBoxes = page1Boxes.filter(
    (b) => (b.y0 + b.y1) / 2 <= 0.44 && b.x0 >= Math.max(0.32, splitX - 0.04)
  );
  const topRightLines = groupBoxesIntoLines(topRightBoxes).filter((l) => !l.isDivider);

  let foundMottoIdx = -1;
  for (let i = 0; i < topRightLines.length; i++) {
    const norm = removeVietnameseDiacritics(topRightLines[i].text);
    if (/(CONG HOA XA HOI|DOC LAP\s*[-–—]?\s*TU DO|HANH PHUC)/i.test(norm)) {
      foundMottoIdx = i;
    }
  }

  const candidateLines =
    foundMottoIdx !== -1 ? topRightLines.slice(foundMottoIdx + 1) : topRightLines;

  for (const line of candidateLines) {
    const norm = removeVietnameseDiacritics(line.text);
    if (/^(KE HOACH|QUYET DINH|THONG BAO|CONG VAN|CAN CU|KINH GUI|DIEU \d+)/i.test(norm)) {
      break;
    }
    const parsed = extractDateFromHeaderOrSigLine(line.text);
    if (parsed) {
      return {
        date: parsed,
        positionLabel: 'Góc trên bên phải (Trang 1)',
        page: 1,
        bbox: unionLineBBoxes([line]),
        rawText: line.text,
      };
    }
  }

  // Fallback từ text cột phải trên
  if (rightTopText) {
    for (const rawLine of rightTopText.split('\n').map((s) => s.trim()).filter(Boolean)) {
      const norm = removeVietnameseDiacritics(rawLine);
      if (/(CONG HOA|DOC LAP)/i.test(norm)) continue;
      const parsed = extractDateFromHeaderOrSigLine(rawLine);
      if (parsed) {
        return {
          date: parsed,
          positionLabel: 'Góc trên bên phải (Trang 1)',
          page: 1,
          estimated: true,
          rawText: rawLine,
        };
      }
    }
  }

  return null;
}

function extractExplicitSignDateFromText(text: string): { date: string; rawText: string } | null {
  if (!text) return null;
  const lines = text.split('\n').map((s) => s.trim()).filter(Boolean);
  for (const line of lines) {
    const m1 = line.match(
      /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y|Signed\s+on|Signing\s+time)\s*[:\-]?\s*(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:[T\s,]+\d{1,2}:\d{2}(?::\d{2})?)?/i
    );
    if (m1) {
      const d = formatDateDdMmYyyy(m1[1], m1[2], m1[3]);
      if (d) return { date: d, rawText: line };
    }
    const m2 = line.match(
      /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y|Signed\s+on|Signing\s+time)\s*[:\-]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s,]+\d{1,2}:\d{2}(?::\d{2})?)?/i
    );
    if (m2) {
      const d = formatDateDdMmYyyy(m2[3], m2[2], m2[1]);
      if (d) return { date: d, rawText: line };
    }
    const m3 = line.match(
      /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y)\s*[:\-]?\s*(?:ng[àa]y\s*)?(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})/i
    );
    if (m3) {
      const d = formatDateDdMmYyyy(m3[1], m3[2], m3[3]);
      if (d) return { date: d, rawText: line };
    }
  }
  return null;
}

function extractSignatureZoneDate(params: {
  signaturePage: number;
  pageBoxes: NormalizedBox[];
  annotationText: string;
  rawSigDates: string[];
  signerPageText: string;
}): ExtractedDateSource | null {
  const { signaturePage, pageBoxes, annotationText, rawSigDates, signerPageText } = params;

  // 1. Tìm nhãn tường minh "Ngày ký:" / "Thời gian ký:" trong vùng chữ ký hoặc annotation PDF
  const explicitAnnot = extractExplicitSignDateFromText(annotationText);
  if (explicitAnnot) {
    const bbox = findLinesBBoxByQuery(pageBoxes, explicitAnnot.date, 0.2, 1);
    return {
      date: explicitAnnot.date,
      positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
      page: signaturePage,
      ...(bbox ? { bbox } : { estimated: true }),
      rawText: explicitAnnot.rawText,
    };
  }

  // Lọc các hộp thuộc phần ký (bỏ phần tiêu đề góc trên trang 1 để không đếm trùng Vị trí 1)
  const sigZoneBoxes = pageBoxes.filter((b) => {
    const yc = (b.y0 + b.y1) / 2;
    if (signaturePage === 1 && yc <= 0.40) return false;
    return yc >= 0.25;
  });
  const sigZoneLines = groupBoxesIntoLines(sigZoneBoxes).filter((l) => !l.isDivider);

  // Tìm dòng chức vụ ký đầu tiên ở vùng chữ ký
  let roleLineIdx = -1;
  for (let i = 0; i < sigZoneLines.length; i++) {
    const norm = removeVietnameseDiacritics(sigZoneLines[i].text);
    if (SIGNER_ROLE_REGEX.test(norm) && sigZoneLines[i].x0 >= 0.30) {
      roleLineIdx = i;
      break;
    }
  }

  const sigScopeLines =
    roleLineIdx !== -1
      ? sigZoneLines.slice(Math.max(0, roleLineIdx - 4), Math.min(sigZoneLines.length, roleLineIdx + 14))
      : sigZoneLines.filter((l) => l.x0 >= 0.35 && l.yCenter >= 0.45);

  const explicitZone = extractExplicitSignDateFromText(sigScopeLines.map((l) => l.text).join('\n'));
  if (explicitZone) {
    const matchedLine = sigScopeLines.find((l) => l.text.includes(explicitZone.rawText) || l.text.includes(explicitZone.date));
    return {
      date: explicitZone.date,
      positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
      page: signaturePage,
      bbox: matchedLine ? unionLineBBoxes([matchedLine]) : findLinesBBoxByQuery(sigZoneBoxes, explicitZone.date, 0.25, 1),
      rawText: explicitZone.rawText,
    };
  }

  // 2. Dòng ngày ngay phía trên hoặc trong khối ký ở nửa dưới/phải trang ký (vd: "An Giang, ngày 12 tháng 11 năm 2026" hoặc con dấu ngày "12/11/2026")
  for (const line of sigScopeLines) {
    if (signaturePage === 1 && line.yCenter <= 0.40) continue;
    const norm = removeVietnameseDiacritics(line.text);
    if (BODY_DATE_CONTEXT_REGEX.test(norm)) continue;

    const fromVn = extractDateFromHeaderOrSigLine(line.text);
    if (fromVn) {
      return {
        date: fromVn,
        positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
        page: signaturePage,
        bbox: unionLineBBoxes([line]),
        rawText: line.text,
      };
    }

    // Ngày dạng dd/MM/yyyy đứng độc lập trong khối ký (cạnh chữ ký / dấu mộc)
    if (line.x0 >= 0.35 && line.yCenter >= 0.42) {
      const mBare = line.text.match(/\b(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/);
      if (mBare && line.text.length <= 48) {
        const d = formatDateDdMmYyyy(mBare[1], mBare[2], mBare[3]);
        if (d) {
          return {
            date: d,
            positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
            page: signaturePage,
            bbox: unionLineBBoxes([line]),
            rawText: line.text,
          };
        }
      }
    }
  }

  // Fallback từ text trang ký nếu chưa có boxes (tránh phần đầu trang 1)
  if (sigZoneBoxes.length === 0 && signerPageText) {
    const rawLines = signerPageText.split('\n').map((s) => s.trim()).filter(Boolean);
    const searchLines = signaturePage === 1 ? rawLines.slice(Math.floor(rawLines.length * 0.45)) : rawLines;
    const explicitText = extractExplicitSignDateFromText(searchLines.join('\n'));
    if (explicitText) {
      return {
        date: explicitText.date,
        positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
        page: signaturePage,
        estimated: true,
        rawText: explicitText.rawText,
      };
    }
    for (const l of searchLines) {
      const d = extractDateFromHeaderOrSigLine(l);
      if (d) {
        return {
          date: d,
          positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
          page: signaturePage,
          estimated: true,
          rawText: l,
        };
      }
    }
  }

  // 3. Ngày trong chữ ký số PDF (/Type /Sig /M D:...) — bỏ giờ phút giây
  if (rawSigDates.length > 0) {
    const validRaw = normalizeDateOnlyDdMmYyyy(rawSigDates[0]);
    if (validRaw) {
      return {
        date: validRaw,
        positionLabel: `Chữ ký số PDF (Trang ${signaturePage})`,
        page: signaturePage,
        estimated: true,
        rawText: rawSigDates[0],
      };
    }
  }

  return null;
}

function extractDocNumberFromSoValue(rawAfterLabel: string): string {
  if (!rawAfterLabel) return '';
  const cleaned = rawAfterLabel
    .replace(/[ \t]*\/[ \t]*/g, '/')
    .replace(/[ \t]*[-–—][ \t]*/g, '-')
    .trim();
  if (!cleaned) return '';

  // Lấy giá trị số/ký hiệu đầy đủ (giữ số 0 ở đầu, dấu /, dấu -, chữ cái)
  const primaryMatch = cleaned.match(
    /^([0-9]{1,5}[\/\-][\p{L}\p{N}_\-\.\/]+)/u
  );
  let candidate = primaryMatch ? primaryMatch[1] : '';
  if (!candidate) {
    // Ghi nhận cả trường hợp thiếu phần số đứng trước dấu "/" (ví dụ "/QĐ-SGDĐT") để cảnh báo cần kiểm tra
    const suffixOnlyMatch = cleaned.match(/^(\/[\p{L}][\p{L}\p{N}_\-\.\/]*)/u);
    if (suffixOnlyMatch) {
      candidate = suffixOnlyMatch[1];
    } else {
      candidate = cleaned.split(/\s+/)[0] || '';
    }
  }
  candidate = candidate.replace(/[.,;:\)]+$/g, '').trim();
  if (!candidate) return '';
  if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(candidate)) return '';
  return candidate;
}

function isSoLabelLine(lineText: string): boolean {
  const cleaned = lineText.replace(/^[\s_\-—–=~*•·.]+/, '').trim();
  if (!cleaned) return false;
  if (/^(?:S[ốoôơ6](?:\s*\/\s*K[ýy]\s*hi[ệe]u)?)\s*[:\.]/i.test(cleaned)) {
    return true;
  }
  // Hỗ trợ "Số2937 /QĐ-SGDĐT", "Số 2937/QĐ-SGDĐT" lẫn "Số /QĐ-SGDĐT"
  if (/^S[ốoôơ6](?:\s*\d{1,5})?\s*[\/\-]/i.test(cleaned) || /^S[ốoôơ6]\s+\d{1,5}(?:\s|$)/i.test(cleaned)) {
    return true;
  }
  return false;
}

function extractValueFromSoLine(lineText: string): string {
  const cleaned = lineText.replace(/^[\s_\-—–=~*•·.]+/, '').trim();
  const afterLabel = cleaned.replace(
    /^(?:S[ốoôơ6](?:\s*\/\s*K[ýy]\s*hi[ệe]u)?)\s*[:\.]?\s*/i,
    ''
  );
  return extractDocNumberFromSoValue(afterLabel);
}

function extractDocNumberCandidates(text: string): string[] {
  if (!text) return [];
  const regex =
    /(?:S[ốoôơ6]|Số\s*\/\s*Ký\s*hiệu)\s*[:\.]?[ \t]*([0-9]{0,5}[ \t]*[\/\-][ \t]*[\p{L}\p{N}_\-\.\/]+(?:[ \t]*[\-\/][ \t]*[\p{L}\p{N}_\-\.]+)*)/giu;
  const results: string[] = [];
  const seen = new Set<string>();

  for (const m of text.matchAll(regex)) {
    let num = (m[1] || '')
      .replace(/[ \t]*\/[ \t]*/g, '/')
      .replace(/[ \t]*[-–—][ \t]*/g, '-')
      .replace(/[.,;:\)]+$/g, '')
      .trim();

    num = num.split(/\s+/)[0] || '';
    if (!num) continue;
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(num)) continue;
    const key = num.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      results.push(num);
    }
  }
  return results;
}

function isIncompleteOrSuspiciousDocNumber(val: string): { suspicious: boolean; reason: string } {
  const trimmed = (val || '').trim();
  if (!trimmed) {
    return {
      suspicious: true,
      reason: 'Chưa nhận diện được Số/ký hiệu văn bản ở góc trên bên trái trang 1.',
    };
  }
  if (/^[\/\-]/.test(trimmed)) {
    return {
      suspicious: true,
      reason: `Số/ký hiệu đọc được ("${trimmed}") đang thiếu chữ số gốc đứng trước dấu "/". Vui lòng đối chiếu bản PDF gốc và bổ sung đủ số văn bản.`,
    };
  }
  if (!/\d/.test(trimmed)) {
    return {
      suspicious: true,
      reason: `Số/ký hiệu đọc được ("${trimmed}") chưa có phần chữ số rõ ràng. Vui lòng kiểm tra lại với bản PDF gốc.`,
    };
  }
  if (!/[\/\-]/.test(trimmed)) {
    return {
      suspicious: true,
      reason: `Số/ký hiệu đọc được ("${trimmed}") có thể bị thiếu phần ký hiệu sau dấu "/". Vui lòng kiểm tra lại.`,
    };
  }
  if (/[\uFFFD?]/.test(trimmed)) {
    return {
      suspicious: true,
      reason: `Số/ký hiệu đọc được ("${trimmed}") có ký tự lỗi phông chữ. Vui lòng kiểm tra lại.`,
    };
  }
  return { suspicious: false, reason: '' };
}

export function extractFilenameDocNumberHint(fileName: string): string {
  const base = (fileName || '').replace(/\.[^.]+$/, '').trim();
  if (!base) return '';
  // Tìm mẫu số văn bản phổ biến trong tên tệp (vd "2937-QD-SGDDT", "2937.QD-SGDDT", "03-TB.DTN", "2937_QD-SGDDT")
  const match = base.match(/(?:^|[^\d])(\d{1,5}[.\-\/_][A-Za-zÀ-ỹĐđ]{1,8}(?:[.\-\/_][A-Za-zÀ-ỹĐđ0-9]+)*)/i);
  if (!match) return '';
  const raw = match[1].trim();
  if (/^\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{2,4}$/.test(raw)) return '';
  const normalized = raw.replace(/^(\d{1,5})[._\-/]/, '$1/').replace(/[._]+/g, '-');
  return normalized;
}

function isMostlyUppercaseLine(str: string): boolean {
  const letters = (str || '').match(/\p{L}/gu) || [];
  if (letters.length === 0) return false;
  const upperCount = letters.filter((ch) => ch === ch.toUpperCase() && ch !== ch.toLowerCase()).length;
  return upperCount / letters.length >= 0.82;
}

function startsWithAgencyNoun(str: string): boolean {
  const norm = removeVietnameseDiacritics(str).trim();
  return /^(?:SO|PHONG|TRUONG|TRUNG TAM|UY BAN|UBND|BO|CUC|VU|CHI CUC|CHI BO|DOAN|BAN|VIEN|HOC VIEN|CONG AN|TINH DOAN|HUYEN DOAN|THANH DOAN|HDND|DAI HOC)\b/i.test(
    norm
  );
}

function isWrappedAgencyContinuation(prevText: string, curText: string): boolean {
  const normPrev = removeVietnameseDiacritics(prevText).trim();
  const normCur = removeVietnameseDiacritics(curText).trim();
  if (!normPrev || !normCur) return false;
  if (startsWithAgencyNoun(curText)) return false;
  if (/^(?:VA|CUA|TINH|THANH PHO|THPT|THCS|TH|QUAN|HUYEN|THI XA|TP\.|VIET NAM|NGUYEN|AN GIANG|DAO TAO)\b/i.test(normCur)) {
    return true;
  }
  if (/\b(?:VA|CUA|GIAO DUC|DAO TAO|NHAN DAN|THPT|TRUONG|TRUNG TAM|PHONG)\s*$/i.test(normPrev)) {
    return true;
  }
  return false;
}

function detectHeaderColumnSplitX(topBoxes: NormalizedBox[]): number {
  const headerTopBoxes = topBoxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.34);
  if (headerTopBoxes.length === 0) return 0.48;

  // 1. Neo theo khối quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" / "Độc lập..." bên phải
  const rightMottoRegex = /^(?:CONG|HOA|XA|HOI|CHU|NGHIA|VIET|NAM|DOC|LAP|TU|DO|HANH|PHUC)$/i;
  const rightAnchors = headerTopBoxes.filter((b) => {
    const norm = removeVietnameseDiacritics(b.text).replace(/[^A-Za-z]/g, '');
    return b.x0 >= 0.34 && rightMottoRegex.test(norm);
  });
  if (rightAnchors.length >= 2) {
    const minRightX = Math.min(...rightAnchors.map((b) => b.x0));
    return Math.min(0.55, Math.max(0.36, minRightX - 0.01));
  }

  // 2. Fallback: tìm khoảng trắng dọc (gutter) rộng nhất trong khoảng x = 0.35..0.56
  const leftCandidates = headerTopBoxes.filter((b) => b.x1 <= 0.52);
  const rightCandidates = headerTopBoxes.filter((b) => b.x0 >= 0.38);
  if (leftCandidates.length > 0 && rightCandidates.length > 0) {
    const intervals = headerTopBoxes
      .map((b) => ({ x0: b.x0, x1: b.x1 }))
      .sort((a, b) => a.x0 - b.x0);
    let bestGap = 0;
    let bestSplit = 0.48;
    let curMaxX = intervals[0].x1;
    for (let i = 1; i < intervals.length; i++) {
      if (intervals[i].x0 > curMaxX) {
        const gapMid = (curMaxX + intervals[i].x0) / 2;
        const gapSize = intervals[i].x0 - curMaxX;
        if (gapMid >= 0.35 && gapMid <= 0.56 && gapSize > bestGap) {
          bestGap = gapSize;
          bestSplit = gapMid;
        }
      }
      curMaxX = Math.max(curMaxX, intervals[i].x1);
    }
    if (bestGap >= 0.022) {
      return bestSplit;
    }
  }

  return 0.48;
}

function unionLineBBoxes(lines: NormalizedLine[]): NormalizedBBox | undefined {
  const valid = lines.filter((l) => l && Number.isFinite(l.x0) && Number.isFinite(l.y0));
  if (valid.length === 0) return undefined;
  return {
    x0: Math.max(0, Math.min(...valid.map((l) => l.x0)) - 0.006),
    y0: Math.max(0, Math.min(...valid.map((l) => l.y0)) - 0.004),
    x1: Math.min(1, Math.max(...valid.map((l) => l.x1)) + 0.006),
    y1: Math.min(1, Math.max(...valid.map((l) => l.y1)) + 0.004),
  };
}

function findLinesBBoxByQuery(boxes: NormalizedBox[], query: string, yMin = 0, yMax = 1): NormalizedBBox | undefined {
  if (!boxes || boxes.length === 0 || !query || !query.trim()) return undefined;
  const scoped = boxes.filter((b) => {
    const yc = (b.y0 + b.y1) / 2;
    return yc >= yMin && yc <= yMax;
  });
  const lines = groupBoxesIntoLines(scoped).filter((l) => !l.isDivider);
  if (lines.length === 0) return undefined;

  const normQuery = removeVietnameseDiacritics(query).toLowerCase().replace(/\s+/g, ' ').trim();
  const queryWords = normQuery.split(/\s+/).filter((w) => w.length >= 2);
  if (normQuery.length < 2) return undefined;

  // 1. Direct single-line match first (tightest bbox), then multi-line window substring match
  for (let i = 0; i < lines.length; i++) {
    const norm1 = removeVietnameseDiacritics(lines[i].text).toLowerCase().replace(/\s+/g, ' ').trim();
    if (norm1.includes(normQuery) || (norm1.length >= 6 && normQuery.includes(norm1) && norm1.length >= normQuery.length * 0.45)) {
      return unionLineBBoxes([lines[i]]);
    }
  }
  for (let len = 2; len <= 5; len++) {
    for (let i = 0; i + len <= lines.length; i++) {
      const slice = lines.slice(i, i + len);
      const combined = slice
        .map((l) => removeVietnameseDiacritics(l.text).toLowerCase().replace(/\s+/g, ' ').trim())
        .join(' ');
      if (combined.includes(normQuery)) {
        return unionLineBBoxes(slice);
      }
    }
  }

  // 2. Best token overlap match across 1..5 consecutive lines
  if (queryWords.length >= 2) {
    let bestScore = 0;
    let bestSlice: NormalizedLine[] = [];
    for (let i = 0; i < lines.length; i++) {
      for (let len = 1; len <= 5 && i + len <= lines.length; len++) {
        const slice = lines.slice(i, i + len);
        const sliceNorm = removeVietnameseDiacritics(slice.map((l) => l.text).join(' '))
          .toLowerCase()
          .replace(/\s+/g, ' ');
        const matchedCount = queryWords.filter((w) => sliceNorm.includes(w)).length;
        const ratio = matchedCount / queryWords.length;
        if (matchedCount >= Math.min(2, queryWords.length) && ratio > bestScore) {
          bestScore = ratio;
          bestSlice = slice;
        }
      }
    }
    if (bestScore >= 0.55 && bestSlice.length > 0) {
      return unionLineBBoxes(bestSlice);
    }
  }

  return undefined;
}

interface TopLeftHeaderAnalysis {
  parentAgency: string;
  parentAgencyBBox?: NormalizedBBox;
  issuingAgency: string;
  agencyBBox?: NormalizedBBox;
  agencyNeedsReview: boolean;
  agencyReviewReason: string;
  docNumber: string;
  docNumberBBox?: NormalizedBBox;
  docNumberNeedsReview: boolean;
  docNumberReviewReason: string;
  topLeftHeaderDocNumbers: string[];
  rightTopText: string;
  leftTopText: string;
  splitX: number;
}

function analyzeTopLeftHeader(
  page1Boxes: NormalizedBox[],
  graphicDividerYs: number[] = []
): TopLeftHeaderAnalysis {
  if (page1Boxes.length === 0) {
    return {
      parentAgency: '',
      issuingAgency: '',
      agencyNeedsReview: false,
      agencyReviewReason: '',
      docNumber: '',
      docNumberNeedsReview: false,
      docNumberReviewReason: '',
      topLeftHeaderDocNumbers: [],
      rightTopText: '',
      leftTopText: '',
      splitX: 0.48,
    };
  }

  const topBoxes = page1Boxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.48);
  const splitX = detectHeaderColumnSplitX(topBoxes);

  const leftTopBoxes = topBoxes.filter((b) => {
    const xCenter = (b.x0 + b.x1) / 2;
    return xCenter < splitX + 0.01 && b.x0 < splitX - 0.002;
  });
  const rightTopBoxes = topBoxes.filter((b) => b.x0 >= splitX - 0.02);

  const leftTopLines = groupBoxesIntoLines(leftTopBoxes);
  const rightTopText = boxesToText(rightTopBoxes);
  const leftTopText = leftTopLines
    .filter((l) => !l.isDivider)
    .map((l) => l.text)
    .join('\n');

  const headerStopRegex =
    /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|TO TRINH|CHUONG TRINH|HUONG DAN|CONG VAN|QUY CHE|GIAY MOI|THU MOI|CAN CU|KINH GUI|V\/V\b|VE VIEC\b|DIEU 1|I\.\s)/i;

  // Tìm dòng "Số:" trong khối phía trên bên trái (ưu tiên ngay dưới khối cơ quan)
  let firstSoIndex = -1;
  const headerSoNumbers: string[] = [];
  const headerSoLines: NormalizedLine[] = [];

  for (let i = 0; i < leftTopLines.length; i++) {
    const line = leftTopLines[i];
    if (line.isDivider) continue;
    const norm = removeVietnameseDiacritics(line.text);
    if (headerStopRegex.test(norm)) {
      break;
    }
    if (isSoLabelLine(line.text)) {
      if (firstSoIndex === -1) firstSoIndex = i;
      let val = extractValueFromSoLine(line.text);
      const matchedSoLines: NormalizedLine[] = [line];
      // Nếu "Số:" đứng một mình trên dòng và giá trị nằm ngay dòng sát bên dưới
      if (!val && i + 1 < leftTopLines.length) {
        const nextLine = leftTopLines[i + 1];
        if (!nextLine.isDivider && nextLine.yCenter - line.yCenter <= 0.035) {
          val = extractDocNumberFromSoValue(nextLine.text);
          if (val) matchedSoLines.push(nextLine);
        }
      }
      if (val && !headerSoNumbers.some((n) => n.toLowerCase() === val.toLowerCase())) {
        headerSoNumbers.push(val);
        if (headerSoLines.length === 0) {
          headerSoLines.push(...matchedSoLines);
        }
      } else if (headerSoLines.length === 0) {
        headerSoLines.push(line);
      }
    }
  }

  let docNumber = headerSoNumbers[0] || '';
  let docNumberNeedsReview = false;
  let docNumberReviewReason = '';

  if (firstSoIndex !== -1 && !docNumber) {
    docNumberNeedsReview = true;
    docNumberReviewReason =
      'Đã thấy nhãn "Số" ở góc trên bên trái nhưng chưa đọc chắc chắn phần số/ký hiệu phía sau.';
  } else if (headerSoNumbers.length > 1) {
    docNumberNeedsReview = true;
    docNumberReviewReason = `Phát hiện nhiều dòng "Số" trong khối tiêu đề góc trên bên trái (${headerSoNumbers.join(', ')}), vui lòng xác nhận đúng số gốc.`;
  }

  // Thu thập các dòng tên cơ quan STRICTLY phía trên dòng "Số:" (bỏ đường kẻ, dòng trống, quốc hiệu)
  const searchLimitIndex = firstSoIndex !== -1 ? firstSoIndex : leftTopLines.length;
  const rawAgencyLines: NormalizedLine[] = [];
  const textDividerYs: number[] = [];

  for (let i = 0; i < searchLimitIndex; i++) {
    const line = leftTopLines[i];
    if (line.isDivider) {
      textDividerYs.push(line.yCenter);
      continue;
    }
    const cleanText = line.text
      .replace(/CỘNG\s+HÒA\s+XÃ\s+HỘI[\s\S]*/i, '')
      .replace(/Độc\s+lập\s*-\s*Tự\s+do[\s\S]*/i, '')
      .trim();
    if (!cleanText || countMeaningfulChars(cleanText) < 3) continue;
    const normClean = removeVietnameseDiacritics(cleanText);
    if (firstSoIndex === -1 && headerStopRegex.test(normClean)) break;
    if (/^(CONG HOA|DOC LAP|NGAY\b)/i.test(normClean)) continue;

    rawAgencyLines.push({
      ...line,
      text: cleanText,
    });
  }

  const allDividerYs = [...graphicDividerYs, ...textDividerYs].sort((a, b) => a - b);
  const hasDividerBetween = (yTop: number, yBottom: number) =>
    allDividerYs.some((dy) => dy > yTop + 0.001 && dy < yBottom - 0.001);

  let parentAgency = '';
  let parentUsedLines: NormalizedLine[] = [];
  let issuingAgency = '';
  let agencyUsedLines: NormalizedLine[] = [];
  let agencyNeedsReview = false;
  let agencyReviewReason = '';

  if (rawAgencyLines.length === 0) {
    issuingAgency = '';
    agencyNeedsReview = true;
    agencyReviewReason = 'Chưa nhận diện được khối tên cơ quan ban hành ở góc trên bên trái trang đầu.';
  } else if (rawAgencyLines.length === 1) {
    issuingAgency = rawAgencyLines[0].text;
    agencyUsedLines = [rawAgencyLines[0]];
  } else if (rawAgencyLines.length === 2) {
    const [line1, line2] = rawAgencyLines;
    const divider12 = hasDividerBetween(line1.yCenter, line2.yCenter);
    const wrapped12 = !divider12 && isWrappedAgencyContinuation(line1.text, line2.text);
    if (wrapped12) {
      issuingAgency = `${line1.text} ${line2.text}`.replace(/\s+/g, ' ').trim();
      agencyUsedLines = [line1, line2];
    } else {
      parentAgency = line1.text;
      parentUsedLines = [line1];
      issuingAgency = line2.text;
      agencyUsedLines = [line2];
    }
  } else if (rawAgencyLines.length === 3) {
    const [line1, line2, line3] = rawAgencyLines;
    const dividerBetween23 = hasDividerBetween(line2.yCenter, line3.yCenter);
    const dividerBetween12 = hasDividerBetween(line1.yCenter, line2.yCenter);

    const aligned23 = Math.abs(line2.xCenter - line3.xCenter) <= 0.055 || Math.abs(line2.x0 - line3.x0) <= 0.035;
    const aligned12 = Math.abs(line1.xCenter - line2.xCenter) <= 0.055 || Math.abs(line1.x0 - line2.x0) <= 0.035;

    const dy12 = Math.max(0.004, line2.yCenter - line1.yCenter);
    const dy23 = Math.max(0.004, line3.yCenter - line2.yCenter);

    const heightSimilar23 =
      Math.abs(line2.avgHeight - line3.avgHeight) / Math.max(line2.avgHeight, line3.avgHeight, 0.001) <= 0.30;
    const heightSimilar12 =
      Math.abs(line1.avgHeight - line2.avgHeight) / Math.max(line1.avgHeight, line2.avgHeight, 0.001) <= 0.30;
    const bothUpper23 = isMostlyUppercaseLine(line2.text) && isMostlyUppercaseLine(line3.text);
    const bothUpper12 = isMostlyUppercaseLine(line1.text) && isMostlyUppercaseLine(line2.text);

    const wrapped23 =
      !dividerBetween23 &&
      aligned23 &&
      heightSimilar23 &&
      bothUpper23 &&
      (isWrappedAgencyContinuation(line2.text, line3.text) || (!dividerBetween12 && dy23 <= dy12 * 1.25));
    const wrapped12 =
      !dividerBetween12 &&
      aligned12 &&
      heightSimilar12 &&
      bothUpper12 &&
      (isWrappedAgencyContinuation(line1.text, line2.text) || dy12 * 1.25 < dy23);

    if (dividerBetween23) {
      parentAgency = line1.text;
      parentUsedLines = [line1];
      issuingAgency = line2.text;
      agencyUsedLines = [line2];
      agencyNeedsReview = true;
      agencyReviewReason =
        'Khối cơ quan ban hành có 3 dòng và có đường kẻ phân tách trước dòng thứ 3, vui lòng kiểm tra lại.';
    } else if (wrapped23 && !wrapped12) {
      parentAgency = line1.text;
      parentUsedLines = [line1];
      issuingAgency = `${line2.text} ${line3.text}`.replace(/\s+/g, ' ').trim();
      agencyUsedLines = [line2, line3];
    } else if (wrapped12 && !wrapped23) {
      parentAgency = `${line1.text} ${line2.text}`.replace(/\s+/g, ' ').trim();
      parentUsedLines = [line1, line2];
      issuingAgency = line3.text;
      agencyUsedLines = [line3];
    } else if (isWrappedAgencyContinuation(line1.text, line2.text) && isWrappedAgencyContinuation(line2.text, line3.text)) {
      issuingAgency = `${line1.text} ${line2.text} ${line3.text}`.replace(/\s+/g, ' ').trim();
      agencyUsedLines = [line1, line2, line3];
    } else {
      parentAgency = line1.text;
      parentUsedLines = [line1];
      issuingAgency = `${line2.text} ${line3.text}`.replace(/\s+/g, ' ').trim();
      agencyUsedLines = [line2, line3];
      agencyNeedsReview = true;
      agencyReviewReason =
        'Khối tên cơ quan ban hành có 3 dòng, vui lòng kiểm tra lại phần cơ quan chủ quản và cơ quan ban hành.';
    }
  } else {
    // >= 4 dòng: gom thành cụm theo khoảng cách hoặc từ khóa cơ quan
    const clusters: NormalizedLine[][] = [[rawAgencyLines[0]]];
    const gaps: number[] = [];
    for (let i = 1; i < rawAgencyLines.length; i++) {
      gaps.push(Math.max(0.002, rawAgencyLines[i].yCenter - rawAgencyLines[i - 1].yCenter));
    }
    const minGap = Math.min(...gaps);
    for (let i = 1; i < rawAgencyLines.length; i++) {
      const prev = rawAgencyLines[i - 1];
      const cur = rawAgencyLines[i];
      const gap = cur.yCenter - prev.yCenter;
      const newCluster =
        hasDividerBetween(prev.yCenter, cur.yCenter) ||
        startsWithAgencyNoun(cur.text) ||
        (gap > minGap * 1.45 && !isWrappedAgencyContinuation(prev.text, cur.text));
      if (newCluster) {
        clusters.push([cur]);
      } else {
        clusters[clusters.length - 1].push(cur);
      }
    }
    if (clusters.length === 2) {
      parentUsedLines = clusters[0];
      agencyUsedLines = clusters[1];
      parentAgency = parentUsedLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
      issuingAgency = agencyUsedLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
    } else if (clusters.length === 1) {
      agencyUsedLines = clusters[0];
      issuingAgency = agencyUsedLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
    } else {
      parentUsedLines = clusters[0];
      agencyUsedLines = clusters.slice(1).flat();
      parentAgency = parentUsedLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
      issuingAgency = agencyUsedLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
      agencyNeedsReview = true;
      agencyReviewReason = `Khối tên cơ quan có ${rawAgencyLines.length} dòng phía trên dòng Số, vui lòng kiểm tra lại.`;
    }
  }

  return {
    parentAgency,
    parentAgencyBBox: unionLineBBoxes(parentUsedLines),
    issuingAgency,
    agencyBBox: unionLineBBoxes(agencyUsedLines),
    agencyNeedsReview,
    agencyReviewReason,
    docNumber,
    docNumberBBox: unionLineBBoxes(headerSoLines),
    docNumberNeedsReview,
    docNumberReviewReason,
    topLeftHeaderDocNumbers: headerSoNumbers,
    rightTopText,
    leftTopText,
    splitX,
  };
}

function extractIssuingAgencyFallback(page1Text: string): string {
  const lines = page1Text
    .split('\n')
    .slice(0, 12)
    .map((l) => l.trim())
    .filter(Boolean);

  const ignoreRegex = /^(CỘNG HÒA|ĐỘC LẬP|SỐ\s*:|S6\s*:|SO\s*:|NGÀY|V\/V|VỀ VIỆC|KÍNH GỬI)/i;
  const agencyLines: string[] = [];

  for (const line of lines) {
    if (isSoLabelLine(line)) break;
    if (ignoreRegex.test(line)) continue;
    const cleanPart = line
      .replace(/CỘNG\s+HÒA\s+XÃ\s+HỘI[\s\S]*/i, '')
      .replace(/Độc\s+lập\s*-\s*Tự\s+do[\s\S]*/i, '')
      .trim();
    if (cleanPart.length >= 4 && !isDividerLineText(cleanPart)) {
      agencyLines.push(cleanPart);
    }
    if (agencyLines.length >= 2) break;
  }

  if (agencyLines.length === 0) return '';
  if (agencyLines.length >= 2) return agencyLines[1];
  return agencyLines[0];
}

function capitalizeSentenceStart(str: string): string {
  const s = (str || '').trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function extractSubjectAndTitle(
  page1Boxes: NormalizedBox[],
  page1Text: string
): {
  docName: string;
  title: string;
  hasVv: boolean;
  docNameBBox?: NormalizedBBox;
  titleBBox?: NormalizedBBox;
} {
  const boxLines =
    page1Boxes.length > 0
      ? groupBoxesIntoLines(page1Boxes).filter((l) => !l.isDivider)
      : page1Text
          .split('\n')
          .map((l) => l.trim())
          .filter(Boolean)
          .map((text, idx) => ({
            text,
            x0: 0.1,
            x1: 0.9,
            xCenter: 0.5,
            y0: idx * 0.03,
            y1: idx * 0.03 + 0.02,
            yCenter: idx * 0.03 + 0.01,
            avgHeight: 0.015,
            fontNames: [],
            words: [],
            isDivider: false,
          }));

  // 1. Ưu tiên tìm dòng "V/v" hoặc "Về việc" (công văn)
  for (let i = 0; i < boxLines.length; i++) {
    const vvMatch = boxLines[i].text.match(/^(?:[-•*]\s*)?(?:V\/v|V[ềe]\s+vi[ệe]c)\s*[:\-]?\s*(.*)$/i);
    if (vvMatch) {
      const parts: string[] = [];
      const usedTitleLines: NormalizedLine[] = [boxLines[i]];
      if (vvMatch[1]?.trim()) parts.push(vvMatch[1].trim());
      for (let j = i + 1; j < Math.min(i + 6, boxLines.length); j++) {
        const nextLine = boxLines[j];
        const prevLine = boxLines[j - 1];
        if (nextLine.yCenter - prevLine.yCenter > 0.045) break;
        const normNext = removeVietnameseDiacritics(nextLine.text);
        if (
          /^(Kinh gui|Can cu|Thuc hien|Theo |I\.|1\.|Noi nhan|Doc lap|Dieu \d+)/i.test(normNext) ||
          (/^[A-ZĐÂĂÊÔƠƯ\s]{8,}$/.test(nextLine.text) && nextLine.text.length >= 12)
        ) {
          break;
        }
        parts.push(nextLine.text);
        usedTitleLines.push(nextLine);
      }
      const afterVv = capitalizeSentenceStart(parts.join(' ').replace(/\s+/g, ' ').trim());
      if (afterVv) {
        const prevHeading = i > 0 ? boxLines[i - 1] : null;
        const isHeading = prevHeading
          ? /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|TO TRINH|CHUONG TRINH|HUONG DAN|GIAY MOI)$/i.test(
              removeVietnameseDiacritics(prevHeading.text)
            )
          : false;
        const docNameLines = isHeading && prevHeading ? [prevHeading, ...usedTitleLines] : usedTitleLines;
        return {
          docName: isHeading && prevHeading ? `${prevHeading.text} ${afterVv}`.replace(/\s+/g, ' ').trim() : afterVv,
          title: afterVv,
          hasVv: true,
          docNameBBox: unionLineBBoxes(docNameLines),
          titleBBox: unionLineBBoxes(usedTitleLines),
        };
      }
    }
  }

  // 2. Nếu không có "V/v", tìm tiêu đề loại văn bản in hoa (QUYẾT ĐỊNH / KẾ HOẠCH / THÔNG BÁO...) + dòng tên/trích yếu ngay sau
  const headingKeywords =
    /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|CHUONG TRINH|HUONG DAN|TO TRINH|QUY CHE|GIAY MOI|THU MOI|CONG VAN|NGHI QUYET|CHI THI)\b(?!\s*:)/i;
  const subStopRegex =
    /^(Can cu|Kinh gui|Dieu \d+|I\.|1\.|So\s*[:.]|Noi nhan|Thuc hien|Theo |Xet de nghi|Ban chap hanh|Uy ban nhan dan|Giam doc\b|Pho giam doc\b|Hieu truong\b|Chu tich\b|Truong phong\b)/i;

  for (let i = 0; i < boxLines.length; i++) {
    const curLine = boxLines[i];
    const normCur = removeVietnameseDiacritics(curLine.text);
    // Bỏ qua khối tiêu đề cơ quan/quốc hiệu quá cao nếu chưa qua dòng quốc hiệu/Số
    if (curLine.yCenter < 0.14) continue;
    if (headingKeywords.test(normCur)) {
      const subParts: string[] = [];
      const subLines: NormalizedLine[] = [];
      for (let j = i + 1; j < Math.min(i + 7, boxLines.length); j++) {
        const cand = boxLines[j];
        const prev = boxLines[j - 1];
        if (cand.yCenter - prev.yCenter > 0.045) break;
        const normSub = removeVietnameseDiacritics(cand.text);
        if (
          subStopRegex.test(normSub) ||
          SIGNER_ROLE_REGEX.test(normSub) ||
          isDividerLineText(cand.text)
        ) {
          break;
        }
        subParts.push(cand.text);
        subLines.push(cand);
      }
      const subtitle = subParts.join(' ').replace(/\s+/g, ' ').trim();
      const combined = [curLine.text, subtitle].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
      return {
        docName: combined,
        title: subtitle || combined,
        hasVv: false,
        docNameBBox: unionLineBBoxes([curLine, ...subLines]),
        titleBBox: unionLineBBoxes(subLines.length > 0 ? subLines : [curLine]),
      };
    }
  }

  return { docName: '', title: '', hasVv: false };
}

function isAppendixPage(text: string): boolean {
  if (!text) return false;
  const topLines = text
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean)
    .slice(0, 8);

  for (let i = 0; i < topLines.length; i++) {
    const line = removeVietnameseDiacritics(topLines[i]);
    if (
      /^(?:PHU\s*LUC(?:\s*[IVXLCDM0-9]+|\b)|THE\s*LE(?:\b|\s+CUOC\s+THI)|QUY\s*CHE\s*(?:KEM\s*THEO|DINH\s*KEM)|DANH\s*SACH\s*KEM\s*THEO|BANG\s*(?:KEM\s*THEO|PHAN\s*CONG|TONG\s*HOP|DU\s*TOAN|THEO\s*DOI)|BIEU\s*MAU|MAU\s*SO)\b/i.test(
        line
      )
    ) {
      return true;
    }
    if (
      i < 5 &&
      /^\(?\s*(?:Ban\s+hanh\s+)?kem\s+theo\s+(?:Cong\s+van|Quyet\s+dinh|Ke\s+hoach|Thong\s+bao|To\s+trinh|Huong\s+dan|Chuong\s+trinh)\b/i.test(
        line
      )
    ) {
      return true;
    }
  }
  return false;
}

const DELEGATION_PREFIX_REGEX = /^(?:KT\.|TM\.|Q\.|TL\.|TUQ\.)(?:\s|$)/i;

const CONCRETE_ROLE_REGEX =
  /^(?:GIAM DOC|PHO GIAM DOC|CHU TICH|PHO CHU TICH|HIEU TRUONG|PHO HIEU TRUONG|BI THU|PHO BI THU|TRUONG PHONG|PHO TRUONG PHONG|CHANH VAN PHONG|PHO CHANH VAN PHONG|TRUONG BAN|PHO TRUONG BAN|THU TRUONG|CUC TRUONG|PHO CUC TRUONG|VU TRUONG|CHU NHIEM|PHO CHU NHIEM|BO TRUONG|THU TUONG|PHO THU TUONG|QUYEN HIEU TRUONG|QUYEN GIAM DOC|TONG BIEN TAP|PHO TONG BIEN TAP|UY VIEN)/i;

const SIGNER_ROLE_REGEX =
  /^(?:KT\.|TM\.|Q\.|TL\.|TUQ\.|GIAM DOC|PHO GIAM DOC|CHU TICH|PHO CHU TICH|HIEU TRUONG|PHO HIEU TRUONG|BI THU|PHO BI THU|TRUONG PHONG|PHO TRUONG PHONG|CHANH VAN PHONG|PHO CHANH VAN PHONG|TRUONG BAN|PHO TRUONG BAN|THU TRUONG|CUC TRUONG|PHO CUC TRUONG|VU TRUONG|CHU NHIEM|PHO CHU NHIEM|BO TRUONG|THU TUONG|PHO THU TUONG|QUYEN HIEU TRUONG|QUYEN GIAM DOC|TONG BIEN TAP|PHO TONG BIEN TAP|UY VIEN)/i;

const NON_NAME_LINE_REGEX =
  /^(CONG HOA|DOC LAP|HANH PHUC|CHU NGHIA|VIET NAM|NOI NHAN|DA KY|SO GIAO DUC|UY BAN NHAN DAN|THANH PHO|TINH |TRUONG THPT|DOAN TRUONG|PHONG GIAO DUC|BO GIAO DUC|PHU LUC|THE LE|KT\.|TM\.|TL\.|TUQ\.|Q\.|BAN GIAM HIEU|BAN TO CHUC|KE HOACH|QUYET DINH|THONG BAO|BAO CAO|CHUONG TRINH|CONG VAN|DINH KEM|BAN HANH|VAN BAN|LUU|KINH GUI)/i;

const SKIP_SIG_LINE_REGEX =
  /^(Noi nhan|Luu|-\s*Nhu tren|Kinh gui|Ngay ky|Thoi gian ky|Ky boi|\(?Da ky\)?|\(?Ky,?\s*(?:ghi ro ho ten|dong dau|ten).*?\)?)/i;

const DIGITAL_SIG_NOISE_REGEX =
  /(?:smartca|vnpt|viettel[\-\s]?ca|bkav|newca|trustca|\bca2\b|fpt[\-\s]?ca|safe[\-\s]?ca|efy[\-\s]?ca|itext|itextsharp|adobe|acrobat|foxit|signature\s+valid|digitally\s+signed|signed\s+by\s+system|certificate|serial(?:\s*number)?|issuer|timestamp|revocation|valid\s+from|chung\s+thu\s+so|to\s+chuc\s+cap|phan\s+mem\s+ky|chu\s+ky\s+so\s+hop\s+le|\bCN\s*=|\bO\s*=|\bOU\s*=|\bC\s*=\s*VN\b)/i;

function isDigitalSignatureNoise(rawLine: string): boolean {
  const s = (rawLine || '').trim();
  if (!s) return true;
  const norm = removeVietnameseDiacritics(s);
  if (DIGITAL_SIG_NOISE_REGEX.test(norm)) return true;
  // Chuỗi phiên bản phần mềm (vd 5.5.13, v2.1.0) hoặc OID kỹ thuật
  if (/\bv?\d+\.\d+\.\d+(?:\.\d+)?\b/i.test(s)) return true;
  if (/\b1\.\d+\.\d+\.\d+\b/.test(s)) return true;
  return false;
}

function cleanOcrStampNoiseFromName(rawLine: string): string {
  return (rawLine || '')
    .replace(/^[\s\-–—•*·|~`'":;.,()\[\]{}\/\\0-9]+/, '')
    .replace(/[\s\-–—•*·|~`'":;.,()\[\]{}\/\\0-9]+$/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isLikelyVietnamesePersonName(cand: string): boolean {
  const s = (cand || '').trim();
  if (!s || s.length < 4 || s.length > 60) return false;
  if (/\d/.test(s)) return false;
  if (isDigitalSignatureNoise(s)) return false;

  const norm = removeVietnameseDiacritics(s);
  if (
    DELEGATION_PREFIX_REGEX.test(norm) ||
    SIGNER_ROLE_REGEX.test(norm) ||
    NON_NAME_LINE_REGEX.test(norm) ||
    SKIP_SIG_LINE_REGEX.test(norm)
  ) {
    return false;
  }

  const words = s.split(/\s+/).filter(Boolean);
  if (words.length < 2 || words.length > 5) return false;

  // Cho phép cả Title Case ("Huỳnh Văn Hóa") lẫn IN HOA toàn bộ ("HUỲNH VĂN HÓA")
  const allWordsCapitalized = words.every((w) => /^[\p{Lu}][\p{L}]{0,18}$/u.test(w));
  if (!allWordsCapitalized) return false;

  // Nếu dòng IN HOA toàn bộ, loại trừ các cụm hành chính thường đứng gần khối ký
  if (isMostlyUppercaseLine(s)) {
    if (
      /\b(?:CONG HOA|XA HOI|DOC LAP|HANH PHUC|GIAO DUC|DAO TAO|TRUONG|SO |PHONG |BAN |UBND|BO |CHI BO|DOAN|TINH |HUYEN|THANH PHO|QUYET DINH|THONG BAO|KE HOACH|PHU LUC)\b/i.test(
        norm
      )
    ) {
      return false;
    }
  }
  return true;
}

function chooseBestSignerRole(roleLines: string[]): string {
  const cleaned = roleLines.map((r) => r.trim()).filter(Boolean);
  if (cleaned.length === 0) return '';
  // Ưu tiên chức vụ cụ thể (vd "PHÓ GIÁM ĐỐC") thay vì dòng ủy quyền ("KT. GIÁM ĐỐC")
  const concrete = cleaned.filter((r) => {
    const norm = removeVietnameseDiacritics(r);
    return !DELEGATION_PREFIX_REGEX.test(norm) && CONCRETE_ROLE_REGEX.test(norm);
  });
  if (concrete.length > 0) {
    return concrete[concrete.length - 1];
  }
  return cleaned.join(' - ');
}

function extractSignersFromZoneText(
  zoneText: string,
  pageNumber: number
): Array<{ name: string; role: string; page: number }> {
  const found: Array<{ name: string; role: string; page: number }> = [];
  if (!zoneText) return found;

  // 1. Mẫu chức vụ + tên người ký ở vùng chữ ký hiển thị trên trang PDF (ưu tiên hơn text chữ ký số)
  const lines = zoneText
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);

  for (let i = 0; i < lines.length; i++) {
    if (isDigitalSignatureNoise(lines[i])) continue;
    const normLine = removeVietnameseDiacritics(lines[i]);
    if (SIGNER_ROLE_REGEX.test(normLine)) {
      const roleParts = [lines[i]];
      let nameLine = '';
      for (let j = i + 1; j < Math.min(i + 16, lines.length); j++) {
        const rawL = lines[j];
        if (isDigitalSignatureNoise(rawL)) continue;
        const l = cleanOcrStampNoiseFromName(rawL);
        if (!l || isDigitalSignatureNoise(l)) continue;
        const normL = removeVietnameseDiacritics(l);
        if (SKIP_SIG_LINE_REGEX.test(normL)) {
          continue;
        }
        if (SIGNER_ROLE_REGEX.test(normL) && j <= i + 3) {
          roleParts.push(l);
          continue;
        }
        if (isLikelyVietnamesePersonName(l)) {
          nameLine = l;
          break;
        }
      }
      if (nameLine) {
        found.push({
          name: nameLine,
          role: chooseBestSignerRole(roleParts),
          page: pageNumber,
        });
      }
    }
  }

  // 2. Mẫu nhãn tường minh ("Ký bởi: Nguyễn Văn A", "Người ký: ...", "Signed by: ...") - lọc bỏ CA/phần mềm
  const digitalMatches = zoneText.matchAll(
    /(?:K[ýy]\s+b[ởơ]i|Ng[ườờ]i\s+k[ýy]|Signed\s+by)\s*[:\-]\s*([^\n,;]{3,70})/gi
  );
  for (const dm of digitalMatches) {
    const rawSeg = dm[1]
      .replace(/\([^)]*\)/g, '')
      .replace(/(?:Ng[àa]y\s+k[ýy]|Time|Date|Serial).*$/i, '')
      .trim();
    const candName = cleanOcrStampNoiseFromName(rawSeg);
    if (
      candName.length >= 4 &&
      isLikelyVietnamesePersonName(candName) &&
      !found.some((f) => f.name.toLowerCase() === candName.toLowerCase())
    ) {
      found.push({ name: candName, role: '', page: pageNumber });
    }
  }

  return found;
}

interface SignatureZonePage {
  pageNumber: number;
  rightBottomText: string;
  rightHalfText: string;
  bottomText: string;
  fullText: string;
  annotationText: string;
}

function extractSignerInfo(
  signaturePages: SignatureZonePage[],
  rawSigSigners: string[]
): {
  signer: string;
  signerRole: string;
  signerCandidates: string[];
  signaturePage: number;
} {
  const candidates: Array<{ name: string; role: string; page: number }> = [];

  for (const p of signaturePages) {
    // Ưu tiên các vùng chữ ký hiển thị trên trang TRƯỚC; KHÔNG trộn annotation chữ ký số vào trước khi đọc chữ ký hiển thị
    const zones = [p.rightBottomText, p.rightHalfText, p.bottomText, p.fullText].filter(
      (z) => z && z.trim().length > 0
    );

    let pageCandidates: Array<{ name: string; role: string; page: number }> = [];
    for (const zone of zones) {
      const matches = extractSignersFromZoneText(zone, p.pageNumber);
      if (matches.length > 0) {
        pageCandidates = matches;
        break;
      }
    }

    // Chỉ dùng annotationText làm nguồn phụ khi vùng hiển thị trên trang không có tên người ký hợp lệ
    if (pageCandidates.length === 0 && p.annotationText) {
      const cleanAnnotLines = p.annotationText
        .split('\n')
        .filter((l) => !isDigitalSignatureNoise(l))
        .join('\n');
      if (cleanAnnotLines) {
        pageCandidates = extractSignersFromZoneText(cleanAnnotLines, p.pageNumber);
      }
    }

    if (pageCandidates.length > 0) {
      candidates.push(...pageCandidates);
      // Khi đã tìm thấy khối chữ ký ở trang chính gần cuối nhất (trước phụ lục), dừng lùi về các trang trước
      break;
    }
  }

  for (const rs of rawSigSigners) {
    let cleanRs = rs.trim();
    const cnMatch = cleanRs.match(/CN=([^,\/]+)/i);
    if (cnMatch) cleanRs = cnMatch[1].trim();
    cleanRs = cleanOcrStampNoiseFromName(cleanRs);
    if (
      cleanRs &&
      isLikelyVietnamesePersonName(cleanRs) &&
      !candidates.some((c) => c.name.toLowerCase() === cleanRs.toLowerCase())
    ) {
      candidates.push({
        name: cleanRs,
        role: '',
        page: signaturePages[0]?.pageNumber || 1,
      });
    }
  }

  // Khử trùng giữ thứ tự và loại bỏ hoàn toàn các tên nhiễu CA/phần mềm nếu còn sót
  const filteredCandidates = candidates.filter((c) => !isDigitalSignatureNoise(c.name));
  const uniqueNames: string[] = [];
  for (const c of filteredCandidates) {
    if (!uniqueNames.some((n) => n.toLowerCase() === c.name.toLowerCase())) {
      uniqueNames.push(c.name);
    }
  }

  const primary = filteredCandidates[0];
  return {
    signer: primary?.name || '',
    signerRole: primary?.role || '',
    signerCandidates: uniqueNames,
    signaturePage: primary?.page || signaturePages[0]?.pageNumber || 1,
  };
}

export async function extractPdfPageContentsForIndex(buffer: Buffer): Promise<{
  pageTexts: PdfPageContent[];
}> {
  const { pageCount, pages } = await extractPdfTextWithPdfJs(buffer);
  const pageTextByNum = new Map<number, string>();
  const ocrPages: number[] = [];

  for (const p of pages) {
    pageTextByNum.set(p.pageNumber, p.fullText);
    if (isGarbledText(p.fullText) && ocrPages.length < 10) {
      ocrPages.push(p.pageNumber);
    }
  }

  if (ocrPages.length > 0) {
    try {
      const ocrRun = await runVietnameseOcrOnPages(buffer, ocrPages);
      for (const [pNum, ocrRes] of ocrRun.results.entries()) {
        const existing = pageTextByNum.get(pNum) || '';
        if (countMeaningfulChars(ocrRes.text) > countMeaningfulChars(existing) || isGarbledText(existing)) {
          pageTextByNum.set(pNum, ocrRes.text);
        }
      }
    } catch {
      // ignore OCR failure during background indexing
    }
  }

  const pageTexts: PdfPageContent[] = [];
  for (let p = 1; p <= pageCount; p++) {
    pageTexts.push({
      page: p,
      text: (pageTextByNum.get(p) || '').trim(),
    });
  }

  return {
    pageTexts,
  };
}

interface GeminiExtractedMetadata {
  docNumber?: string;
  parentAgency?: string;
  issuingAgency?: string;
  docName?: string;
  title?: string;
  topRightDate?: string;
  signatureDate?: string;
  signer?: string;
  signerRole?: string;
}

async function extractMetadataWithGemini(params: {
  pdfBuffer: Buffer;
  page1Text: string;
  signaturePageText: string;
  signaturePageNumber: number;
  pageCount: number;
}): Promise<GeminiExtractedMetadata | null> {
  const apiKey = (process.env.GEMINI_API_KEY || '').trim();
  if (!apiKey) return null;

  try {
    const { GoogleGenAI, Type } = await import('@google/genai');
    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        },
      },
    });

    const parts: any[] = [];
    if (params.pdfBuffer.length <= 8 * 1024 * 1024) {
      parts.push({
        inlineData: {
          mimeType: 'application/pdf',
          data: params.pdfBuffer.toString('base64'),
        },
      });
    }

    const promptText = [
      'Bạn là hệ thống trích xuất thông tin hành chính văn bản tiếng Việt của trường học.',
      'CHỈ trích xuất 5 nhóm thông tin hành chính sau đây (KHÔNG trích xuất thời hạn xử lý, hạn nộp, nơi nhận, người phụ trách hay danh sách công việc):',
      '1. docNumber: Số/ký hiệu gốc ở dòng "Số" góc trên bên trái trang 1. Ghép các cụm bị tách khoảng trắng, chuẩn hóa dạng "2937/QĐ-SGDĐT" (kể cả khi bản gốc viết "Số2937 /QĐ-SGDĐT"), giữ nguyên số 0 ở đầu nếu có. Không lấy số văn bản được viện dẫn trong phần Căn cứ/nội dung.',
      '2. parentAgency & issuingAgency: Khối tên cơ quan góc trên bên trái trang 1 phía trên dòng Số. Nếu có 2 cấp cơ quan (vd cấp 1 UBND TỈNH AN GIANG, cấp 2 SỞ GIÁO DỤC VÀ ĐÀO TẠO), đưa cấp 1 vào parentAgency và cấp 2 vào issuingAgency; ghép các dòng xuống hàng thuộc cùng một tên cơ quan. Nếu chỉ có 1 cấp thì để parentAgency rỗng.',
      '3. docName & title: Tên văn bản đầy đủ bằng cách nối tiêu đề loại văn bản (vd QUYẾT ĐỊNH, KẾ HOẠCH...) với dòng trích yếu bên dưới, bỏ xuống dòng và khoảng trắng thừa. Nếu là công văn V/v thì title là nội dung sau V/v.',
      '4. topRightDate & signatureDate: Tìm ngày tháng năm ban hành/ký ở HAI vị trí riêng biệt và chuẩn hóa về dd/MM/yyyy (không gồm giờ:phút:giây, trả về chuỗi rỗng nếu không có hoặc ngày không hợp lệ):',
      '   - topRightDate: Ngày ở góc trên bên phải trang 1 phía dưới dòng Độc lập - Tự do - Hạnh phúc (kể cả khi thiếu khoảng trắng như "ngày10 tháng 11 năm 2026").',
      '   - signatureDate: Ngày ở vùng chữ ký / cuối văn bản chính (hoặc Ngày ký chữ ký số), bỏ qua các trang Phụ lục / Thể lệ đính kèm phía sau. KHÔNG lấy ngày của văn bản viện dẫn hay lịch trình trong nội dung.',
      '5. signer & signerRole: Họ tên người ký và chức vụ ở khối chữ ký cuối văn bản chính (bỏ qua các trang Phụ lục / Thể lệ kèm theo). Không nhầm chữ ký tắt/ký thay (KT., TL., TUQ., TM.) thành họ tên người ký; ưu tiên chức vụ thực tế của người ký (vd "PHÓ GIÁM ĐỐC").',
      '',
      `Tổng số trang PDF: ${params.pageCount}. Trang chữ ký chính ước tính: ${params.signaturePageNumber}.`,
      params.page1Text ? `--- Văn bản đọc được từ Trang 1 ---\n${params.page1Text.slice(0, 3500)}` : '',
      params.signaturePageText && params.signaturePageNumber > 1
        ? `--- Văn bản đọc được từ Trang ${params.signaturePageNumber} ---\n${params.signaturePageText.slice(0, 3500)}`
        : '',
    ]
      .filter(Boolean)
      .join('\n');

    parts.push({ text: promptText });

    const response = await ai.models.generateContent({
      model: 'gemini-4-hydrogen',
      contents: [{ role: 'user', parts }],
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            docNumber: { type: Type.STRING },
            parentAgency: { type: Type.STRING },
            issuingAgency: { type: Type.STRING },
            docName: { type: Type.STRING },
            title: { type: Type.STRING },
            topRightDate: { type: Type.STRING },
            signatureDate: { type: Type.STRING },
            signer: { type: Type.STRING },
            signerRole: { type: Type.STRING },
          },
        },
      },
    });

    const rawText = (response.text || '').trim();
    if (!rawText) return null;
    const parsed = JSON.parse(rawText);
    return parsed && typeof parsed === 'object' ? (parsed as GeminiExtractedMetadata) : null;
  } catch (err: any) {
    console.warn('[PDF Gemini] Không dùng được Gemini fallback/assist, tiếp tục với kết quả phân tích bố cục + OCR:', err?.message);
    return null;
  }
}

async function runTopLeftHeader300DpiFallback(buffer: Buffer): Promise<{
  text: string;
  boxes: NormalizedBox[];
  visionDocNumber?: string;
}> {
  const cropWFrac = 0.58;
  const cropHFrac = 0.42;
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'nct-hdr-ocr-'));
  const pdfPath = path.join(tmpDir, 'input.pdf');
  const pagePng = path.join(tmpDir, 'page1-300.png');
  const cropPng = path.join(tmpDir, 'page1-top-left.png');
  fs.writeFileSync(pdfPath, buffer);

  let worker: any = null;
  let ocrText = '';
  let mappedBoxes: NormalizedBox[] = [];
  let visionDocNumber = '';

  try {
    await renderPdfPageToPng(pdfPath, 1, pagePng, 300);
    if (fs.existsSync(pagePng)) {
      const cropMeta = cropRgbPngRegion(pagePng, cropPng, 0, 0, cropWFrac, cropHFrac);
      const imagePath = cropMeta && fs.existsSync(cropPng) ? cropPng : pagePng;
      const effXScale = cropMeta ? cropWFrac : 1;
      const effYScale = cropMeta ? cropHFrac : 1;

      try {
        const { createWorker } = await import('tesseract.js');
        const langPath = path.resolve(process.cwd(), 'node_modules/@tesseract.js-data/vie/4.0.0');
        const workerOpts: any = fs.existsSync(path.join(langPath, 'vie.traineddata.gz'))
          ? { langPath, gzip: true }
          : {};
        worker = await createWorker('vie', 1, workerOpts);
        const dims = getPngDimensions(imagePath);
        const ret = await worker.recognize(imagePath, {}, { blocks: true });
        const rawBoxes = extractBoxesFromTesseractData(ret?.data, dims.width, dims.height);
        mappedBoxes = rawBoxes.map((b) => ({
          ...b,
          x0: clamp01(b.x0 * effXScale),
          x1: clamp01(b.x1 * effXScale),
          y0: clamp01(b.y0 * effYScale),
          y1: clamp01(b.y1 * effYScale),
          height: b.height * effYScale,
        }));
        ocrText = (mappedBoxes.length > 0 ? boxesToText(mappedBoxes) : (ret?.data?.text || '')).trim();
      } catch (ocrErr: any) {
        console.warn('[PDF 300DPI Header OCR] Lỗi OCR vùng tiêu đề góc trên trái:', ocrErr?.message);
      }

      const apiKey = (process.env.GEMINI_API_KEY || '').trim();
      if (apiKey && fs.existsSync(imagePath)) {
        try {
          const { GoogleGenAI, Type } = await import('@google/genai');
          const ai = new GoogleGenAI({
            apiKey,
            httpOptions: { headers: { 'User-Agent': 'aistudio-build' } },
          });
          const imgBuf = fs.readFileSync(imagePath);
          const resp = await ai.models.generateContent({
            model: 'gemini-4-hydrogen',
            contents: [
              {
                role: 'user',
                parts: [
                  { inlineData: { mimeType: 'image/png', data: imgBuf.toString('base64') } },
                  {
                    text: 'Đọc chính xác Số/ký hiệu văn bản ở dòng "Số:" trong ảnh vùng góc trên bên trái trang 1 (giữ nguyên số 0 đứng đầu nếu có, chuẩn hóa dạng 2937/QĐ-SGDĐT). Nếu không chắc chắn từng chữ số, trả về chuỗi rỗng.',
                  },
                ],
              },
            ],
            config: {
              responseMimeType: 'application/json',
              responseSchema: {
                type: Type.OBJECT,
                properties: {
                  docNumber: { type: Type.STRING },
                },
              },
            },
          });
          const parsed = JSON.parse((resp.text || '{}').trim());
          if (typeof parsed?.docNumber === 'string' && parsed.docNumber.trim()) {
            visionDocNumber = extractDocNumberFromSoValue(
              parsed.docNumber.replace(/^S[ốo]\s*[:\.]?\s*/i, '')
            );
          }
        } catch {
          // ignore optional vision assist errors
        }
      }
    }
  } catch (err: any) {
    console.warn('[PDF 300DPI Header] Không thể render trang 1 ở 300 DPI:', err?.message);
  } finally {
    if (worker) {
      await worker.terminate().catch(() => {});
    }
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }

  return { text: ocrText, boxes: mappedBoxes, visionDocNumber: visionDocNumber || undefined };
}

export async function analyzePdfBuffer(buffer: Buffer): Promise<{
  sha256: string;
  extraction: PdfExtractionResult;
  pageTexts: PdfPageContent[];
}> {
  const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
  const { pageCount, pages, rawSigDates, rawSigSigners } = await extractPdfTextWithPdfJs(buffer);

  // Ưu tiên trang 1 và các trang cuối (khối chữ ký)
  const targetPagesSet = new Set<number>([1]);
  if (pageCount >= 2) targetPagesSet.add(pageCount);
  if (pageCount >= 3) targetPagesSet.add(pageCount - 1);

  const ocrNeededPages: number[] = [];
  for (const pNum of targetPagesSet) {
    const info = pages.find((p) => p.pageNumber === pNum);
    if (!info || isGarbledText(info.fullText)) {
      ocrNeededPages.push(pNum);
    }
  }

  let usedOcr = false;
  const ocrConfidences: number[] = [];
  const pageTextByNum = new Map<number, string>();
  for (const p of pages) {
    pageTextByNum.set(p.pageNumber, p.fullText);
  }
  const ocrBoxesByNum = new Map<number, NormalizedBox[]>();

  const mergeOcrResults = (ocrResults: Map<number, OcrPageResult>) => {
    for (const [pNum, ocrRes] of ocrResults.entries()) {
      const existing = pageTextByNum.get(pNum) || '';
      if (countMeaningfulChars(ocrRes.text) > countMeaningfulChars(existing) || isGarbledText(existing)) {
        pageTextByNum.set(pNum, ocrRes.text);
        if (ocrRes.boxes.length > 0) {
          ocrBoxesByNum.set(pNum, ocrRes.boxes);
        }
      }
      if (ocrRes.confidence > 0) {
        ocrConfidences.push(ocrRes.confidence);
      }
    }
  };

  let ocrUnavailableReason: string | undefined;
  if (ocrNeededPages.length > 0) {
    usedOcr = true;
    const ocrRun = await runVietnameseOcrOnPages(buffer, ocrNeededPages);
    mergeOcrResults(ocrRun.results);
    if (ocrRun.ocrUnavailableReason) {
      ocrUnavailableReason = ocrRun.ocrUnavailableReason;
    }
  }

  // Nếu PDF scan nhiều trang (> 3) và các trang cuối là Phụ lục / Thể lệ hoặc chưa có chữ ký, lùi thêm về các trang trước phụ lục
  if (pageCount > 3 && usedOcr) {
    const trailingHasNonAppendixSigner = [pageCount, pageCount - 1].some((pNum) => {
      const txt = pageTextByNum.get(pNum) || '';
      if (!txt || isAppendixPage(txt)) return false;
      const lines = txt.split('\n').filter(Boolean);
      const bottomApprox = lines.slice(Math.floor(lines.length * 0.30)).join('\n');
      return extractSignersFromZoneText(bottomApprox || txt, pNum).length > 0;
    });

    if (!trailingHasNonAppendixSigner) {
      const extraPagesToOcr: number[] = [];
      for (let pNum = pageCount - 2; pNum >= 2 && extraPagesToOcr.length < 5; pNum--) {
        const info = pages.find((p) => p.pageNumber === pNum);
        if (!info || isGarbledText(info.fullText)) {
          extraPagesToOcr.push(pNum);
        }
      }
      if (extraPagesToOcr.length > 0) {
        const extraOcrRun = await runVietnameseOcrOnPages(buffer, extraPagesToOcr);
        mergeOcrResults(extraOcrRun.results);
      }
    }
  }

  const ocrConfidence =
    ocrConfidences.length > 0
      ? Math.round(ocrConfidences.reduce((a, b) => a + b, 0) / ocrConfidences.length)
      : undefined;

  const page1 = pages.find((p) => p.pageNumber === 1);
  const page1Text = pageTextByNum.get(1) || '';

  // 1. Phân tích khối tiêu đề phía trên bên trái trang 1 bằng tọa độ thực (PDF spans hoặc OCR word bboxes)
  const page1HasPdfBoxes = Boolean(page1 && !isGarbledText(page1.fullText) && page1.boxes.length > 0);
  const page1Boxes = page1HasPdfBoxes
    ? page1!.boxes
    : ocrBoxesByNum.get(1) || [];
  const page1DividerYs = page1HasPdfBoxes ? page1!.dividerYs : [];

  const headerAnalysis = analyzeTopLeftHeader(page1Boxes, page1DividerYs);

  let docNumber = headerAnalysis.docNumber;
  let usedFallbackDocNumber = false;
  let fallbackDocNumbers: string[] = [];
  let docNumberBBoxOverride: NormalizedBBox | undefined = headerAnalysis.docNumberBBox;
  let docNumberConflictDetected = false;
  let docNumberConflictDetail = '';

  if (!docNumber) {
    fallbackDocNumbers = extractDocNumberCandidates(page1Text);
    docNumber = fallbackDocNumbers[0] || '';
    usedFallbackDocNumber = Boolean(docNumber);
  }

  // Kiểm tra bất thường Số văn bản ở góc trên bên trái: thiếu phần số ("/QĐ-...") hoặc trống hoặc có cảnh báo
  const initialSuspicious = isIncompleteOrSuspiciousDocNumber(docNumber);
  if (
    initialSuspicious.suspicious ||
    !page1HasPdfBoxes ||
    headerAnalysis.docNumberNeedsReview
  ) {
    const croppedFallback = await runTopLeftHeader300DpiFallback(buffer);
    const cropHeader =
      croppedFallback.boxes.length > 0
        ? analyzeTopLeftHeader(croppedFallback.boxes, [])
        : null;
    const cropCandidates = [
      cropHeader?.docNumber || '',
      ...extractDocNumberCandidates(croppedFallback.text),
      croppedFallback.visionDocNumber || '',
    ]
      .map((s) => s.trim())
      .filter(Boolean);

    const completeCropCandidate = cropCandidates.find(
      (c) => !isIncompleteOrSuspiciousDocNumber(c).suspicious
    );

    if (completeCropCandidate) {
      usedOcr = true;
      if (!docNumber || initialSuspicious.suspicious) {
        docNumber = completeCropCandidate;
        if (!docNumberBBoxOverride && cropHeader?.docNumberBBox) {
          docNumberBBoxOverride = cropHeader.docNumberBBox;
        }
      } else if (docNumber.toLowerCase() !== completeCropCandidate.toLowerCase()) {
        docNumberConflictDetected = true;
        docNumberConflictDetail = `Số văn bản từ lớp chữ ("${docNumber}") khác với kết quả đọc vùng tiêu đề ("${completeCropCandidate}"). Vui lòng đối chiếu bản PDF gốc.`;
      }
    }
  }

  let parentAgency = headerAnalysis.parentAgency;
  let issuingAgency =
    headerAnalysis.issuingAgency || extractIssuingAgencyFallback(page1Text);

  const subjectRes = extractSubjectAndTitle(page1Boxes, page1Text);
  let docName = subjectRes.docName;
  let title = subjectRes.title;

  // 2. Xác định các trang ứng viên cho khối chữ ký: loại bỏ trang Phụ lục / Thể lệ phía sau, quét lùi từ trang chính cuối về trang 1
  const orderedPageNums: number[] = [];
  for (let p = pageCount; p >= 1; p--) {
    orderedPageNums.push(p);
  }
  const nonAppendixPages = orderedPageNums.filter((pNum) => {
    const txt = pageTextByNum.get(pNum) || '';
    return !isAppendixPage(txt);
  });
  const appendixPages = orderedPageNums.filter((pNum) => !nonAppendixPages.includes(pNum));

  const getBoxesForPage = (pNum: number): NormalizedBox[] => {
    const pInfo = pages.find((p) => p.pageNumber === pNum);
    if (pInfo && !isGarbledText(pInfo.fullText) && pInfo.boxes.length > 0) {
      return pInfo.boxes;
    }
    return ocrBoxesByNum.get(pNum) || [];
  };

  const buildZonePage = (pNum: number): SignatureZonePage => {
    const pInfo = pages.find((x) => x.pageNumber === pNum);
    const full = pageTextByNum.get(pNum) || '';
    const hasReliableSpans = Boolean(pInfo && !isGarbledText(pInfo.fullText));

    if (hasReliableSpans && pInfo) {
      // Cho phép lệch nhẹ vùng ký (yCenter >= 0.30, x0 >= 0.35)
      const widerRightBottom = pInfo.boxes.filter((b) => (b.y0 + b.y1) / 2 >= 0.30 && b.x0 >= 0.35);
      return {
        pageNumber: pNum,
        rightBottomText: boxesToText(widerRightBottom) || pInfo.rightBottomText,
        rightHalfText: pInfo.rightHalfText,
        bottomText: pInfo.bottomText,
        fullText: full,
        annotationText: pInfo.annotationText || '',
      };
    }

    const ocrBoxes = ocrBoxesByNum.get(pNum) || [];
    if (ocrBoxes.length > 0) {
      const rightBottomBoxes = ocrBoxes.filter((b) => (b.y0 + b.y1) / 2 >= 0.30 && b.x0 >= 0.35);
      const rightHalfBoxes = ocrBoxes.filter((b) => b.x0 >= 0.35);
      const bottomBoxes = ocrBoxes.filter((b) => (b.y0 + b.y1) / 2 >= 0.30);
      return {
        pageNumber: pNum,
        rightBottomText: boxesToText(rightBottomBoxes),
        rightHalfText: boxesToText(rightHalfBoxes),
        bottomText: boxesToText(bottomBoxes),
        fullText: full,
        annotationText: pInfo?.annotationText || '',
      };
    }

    const ocrLines = full.split('\n').map((l) => l.trim()).filter(Boolean);
    const ocrBottom = ocrLines.slice(Math.floor(ocrLines.length * 0.30)).join('\n');
    return {
      pageNumber: pNum,
      rightBottomText: '',
      rightHalfText: '',
      bottomText: ocrBottom,
      fullText: full,
      annotationText: pInfo?.annotationText || '',
    };
  };

  let sigResult = extractSignerInfo(nonAppendixPages.map(buildZonePage), rawSigSigners);
  if (!sigResult.signer && appendixPages.length > 0) {
    sigResult = extractSignerInfo(appendixPages.map(buildZonePage), rawSigSigners);
  }
  let { signer, signerRole } = sigResult;
  const { signerCandidates, signaturePage } = sigResult;

  // 3. Đối chiếu ngày ở HAI vị trí:
  //    (1) Góc trên bên phải trang 1 dưới "Độc lập - Tự do - Hạnh phúc"
  //    (2) Vùng chữ ký cuối văn bản chính (hoặc ngày ký chữ ký số)
  let topRightDate = extractTopRightHeaderDate(
    page1Boxes,
    headerAnalysis.rightTopText || page1?.rightTopText || '',
    headerAnalysis.splitX
  );
  const sigPageInfo = pages.find((p) => p.pageNumber === signaturePage);
  let signatureDate = extractSignatureZoneDate({
    signaturePage,
    pageBoxes: getBoxesForPage(signaturePage),
    annotationText: sigPageInfo?.annotationText || '',
    rawSigDates,
    signerPageText: pageTextByNum.get(signaturePage) || '',
  });

  // 4. Optional server-side Gemini assist (strictly 5 administrative metadata fields, only when GEMINI_API_KEY is configured)
  let aiUsed = false;
  const geminiResult = await extractMetadataWithGemini({
    pdfBuffer: buffer,
    page1Text,
    signaturePageText: pageTextByNum.get(signaturePage) || '',
    signaturePageNumber: signaturePage,
    pageCount,
  });

  if (geminiResult) {
    aiUsed = true;
    const cleanGeminiDocNum = geminiResult.docNumber
      ? extractDocNumberFromSoValue(geminiResult.docNumber.replace(/^S[ốo]\s*[:\.]?\s*/i, ''))
      : '';
    if ((!docNumber || isIncompleteOrSuspiciousDocNumber(docNumber).suspicious) && cleanGeminiDocNum) {
      if (!isIncompleteOrSuspiciousDocNumber(cleanGeminiDocNum).suspicious || !docNumber) {
        docNumber = cleanGeminiDocNum;
      }
    } else if (
      docNumber &&
      cleanGeminiDocNum &&
      !isIncompleteOrSuspiciousDocNumber(cleanGeminiDocNum).suspicious &&
      docNumber.toLowerCase() !== cleanGeminiDocNum.toLowerCase()
    ) {
      docNumberConflictDetected = true;
      docNumberConflictDetail =
        docNumberConflictDetail ||
        `Số văn bản đọc từ bố cục ("${docNumber}") khác với nhận diện bổ sung ("${cleanGeminiDocNum}"). Vui lòng đối chiếu bản PDF gốc.`;
    }
    if (!parentAgency && geminiResult.parentAgency?.trim()) {
      parentAgency = geminiResult.parentAgency.replace(/\s+/g, ' ').trim();
    }
    if (!issuingAgency && geminiResult.issuingAgency?.trim()) {
      issuingAgency = geminiResult.issuingAgency.replace(/\s+/g, ' ').trim();
    }
    if (!docName && geminiResult.docName?.trim()) {
      docName = geminiResult.docName.replace(/\s+/g, ' ').trim();
    }
    if (!title && geminiResult.title?.trim()) {
      title = geminiResult.title.replace(/\s+/g, ' ').trim();
    }
    if (!signer && geminiResult.signer?.trim()) {
      const candSigner = cleanOcrStampNoiseFromName(geminiResult.signer.trim());
      if (candSigner && isLikelyVietnamesePersonName(candSigner)) {
        signer = candSigner;
      }
    }
    if (!signerRole && geminiResult.signerRole?.trim()) {
      signerRole = geminiResult.signerRole.replace(/\s+/g, ' ').trim();
    }

    if (!topRightDate && geminiResult.topRightDate) {
      const validGTop = normalizeDateOnlyDdMmYyyy(geminiResult.topRightDate);
      if (validGTop) {
        const bbox = findLinesBBoxByQuery(page1Boxes, validGTop, 0, 0.44);
        topRightDate = {
          date: validGTop,
          positionLabel: 'Góc trên bên phải (Trang 1)',
          page: 1,
          ...(bbox ? { bbox } : { estimated: true }),
          rawText: geminiResult.topRightDate,
        };
      }
    }
    if (!signatureDate && geminiResult.signatureDate) {
      const validGSig = normalizeDateOnlyDdMmYyyy(geminiResult.signatureDate);
      if (validGSig) {
        const sigBoxesForLoc = getBoxesForPage(signaturePage);
        const bbox = findLinesBBoxByQuery(sigBoxesForLoc, validGSig, 0.25, 1);
        signatureDate = {
          date: validGSig,
          positionLabel: `Khối chữ ký (Trang ${signaturePage})`,
          page: signaturePage,
          ...(bbox ? { bbox } : { estimated: true }),
          rawText: geminiResult.signatureDate,
        };
      }
    }
  }

  // 5. Xử lý kết quả đối chiếu ngày ở 2 vị trí
  const d1 = topRightDate?.date || '';
  const d2 = signatureDate?.date || '';
  let dateComparisonStatus: DateComparisonStatus = 'missing';
  let documentDate = '';
  let issueDate = '';
  let signDate = '';

  if (d1 && d2) {
    if (d1 === d2) {
      dateComparisonStatus = 'matched';
      documentDate = d1;
      issueDate = d1;
      signDate = d2;
    } else {
      dateComparisonStatus = 'mismatch';
      documentDate = '';
      issueDate = '';
      signDate = d2;
    }
  } else if (d1) {
    dateComparisonStatus = 'single';
    documentDate = d1;
    issueDate = d1;
    signDate = '';
  } else if (d2) {
    dateComparisonStatus = 'single';
    documentDate = d2;
    issueDate = d2;
    signDate = d2;
  }

  const finalDocNumCheck = isIncompleteOrSuspiciousDocNumber(docNumber);
  let docNumberNeedsConfirmation = false;
  let docNumberReviewReason = '';

  if (finalDocNumCheck.suspicious) {
    docNumberNeedsConfirmation = true;
    docNumberReviewReason = finalDocNumCheck.reason;
  } else if (docNumberConflictDetected && docNumberConflictDetail) {
    docNumberNeedsConfirmation = true;
    docNumberReviewReason = docNumberConflictDetail;
  } else if (headerAnalysis.docNumberNeedsReview && headerAnalysis.docNumberReviewReason) {
    docNumberNeedsConfirmation = true;
    docNumberReviewReason = headerAnalysis.docNumberReviewReason;
  } else if (usedFallbackDocNumber) {
    docNumberNeedsConfirmation = true;
    docNumberReviewReason =
      fallbackDocNumbers.length > 1
        ? `Chưa xác định rõ dòng "Số" ở góc trên bên trái; phát hiện nhiều số trong trang đầu (${fallbackDocNumbers.join(', ')}), vui lòng kiểm tra đúng số gốc.`
        : 'Chưa xác định rõ dòng "Số" ngay dưới khối cơ quan ở góc trên bên trái, vui lòng kiểm tra lại Số/ký hiệu đã trích xuất.';
  }

  const reviewReasons: string[] = [];
  if (usedOcr) {
    if (ocrUnavailableReason && ocrConfidence === undefined) {
      reviewReasons.push(ocrUnavailableReason);
    } else if (ocrConfidence !== undefined && ocrConfidence < 75) {
      reviewReasons.push(`Chất lượng quét/OCR thấp (${ocrConfidence}%), vui lòng đối chiếu lại với bản PDF gốc.`);
    } else {
      reviewReasons.push('Văn bản dạng bản scan đã nhận diện bằng OCR tiếng Việt, vui lòng rà soát trước khi lưu.');
    }
  }
  if (docNumberReviewReason) {
    reviewReasons.push(docNumberReviewReason);
  }
  if (headerAnalysis.agencyNeedsReview && headerAnalysis.agencyReviewReason) {
    reviewReasons.push(headerAnalysis.agencyReviewReason);
  }
  if (!docName && !title) reviewReasons.push('Chưa nhận diện được Tên văn bản hoặc Trích yếu.');
  if (dateComparisonStatus === 'mismatch') {
    reviewReasons.push(
      `Ngày trên văn bản chưa thống nhất: ${topRightDate?.positionLabel} (${d1}) khác với ${signatureDate?.positionLabel} (${d2}). Cần quản trị viên xác nhận ngày.`
    );
  } else if (dateComparisonStatus === 'missing') {
    reviewReasons.push('Chưa nhận diện được Ngày văn bản.');
  }
  if (!signer) reviewReasons.push('Chưa nhận diện được Họ tên người ký.');
  if (signerCandidates.length > 1) {
    reviewReasons.push(`Phát hiện nhiều tên người ký (${signerCandidates.join(', ')}), vui lòng xác nhận người ký chính.`);
  }

  const relevantPagesSet = new Set<number>([1]);
  if (signaturePage >= 1 && signaturePage <= pageCount) {
    relevantPagesSet.add(signaturePage);
  }
  const relevantPages = Array.from(relevantPagesSet).sort((a, b) => a - b);

  // 6. Ghi nhận vị trí trang và vùng chứa (bbox chuẩn hóa 0..1) cho từng trường để hỗ trợ "Xem vị trí"
  const fieldLocations: Partial<Record<ExtractableFieldKey, FieldLocation>> = {};

  if (docNumber) {
    const bbox = docNumberBBoxOverride || findLinesBBoxByQuery(page1Boxes, docNumber, 0, 0.48);
    fieldLocations.docNumber = bbox
      ? { page: 1, bbox }
      : { page: 1, estimated: true, bbox: { x0: 0.04, y0: 0.08, x1: 0.48, y1: 0.30 } };
  } else {
    fieldLocations.docNumber = docNumberBBoxOverride
      ? { page: 1, estimated: true, bbox: docNumberBBoxOverride }
      : { page: 1, estimated: true, bbox: { x0: 0.04, y0: 0.08, x1: 0.48, y1: 0.30 } };
  }

  if (parentAgency) {
    const bbox = headerAnalysis.parentAgencyBBox || findLinesBBoxByQuery(page1Boxes, parentAgency, 0, 0.35);
    fieldLocations.parentAgency = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  }

  if (issuingAgency) {
    const bbox = headerAnalysis.agencyBBox || findLinesBBoxByQuery(page1Boxes, issuingAgency, 0, 0.45);
    fieldLocations.issuingAgency = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  } else {
    fieldLocations.issuingAgency = { page: 1, estimated: true };
  }

  if (docName) {
    const bbox = subjectRes.docNameBBox || findLinesBBoxByQuery(page1Boxes, docName, 0.08, 0.88);
    fieldLocations.docName = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  }

  if (title) {
    const bbox = subjectRes.titleBBox || findLinesBBoxByQuery(page1Boxes, title, 0.08, 0.88);
    fieldLocations.title = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  }

  if (topRightDate) {
    const loc: FieldLocation = {
      page: topRightDate.page,
      ...(topRightDate.bbox ? { bbox: topRightDate.bbox } : { estimated: true }),
    };
    fieldLocations.topRightDate = loc;
    fieldLocations.issueDate = loc;
  }

  if (signatureDate) {
    const loc: FieldLocation = {
      page: signatureDate.page,
      ...(signatureDate.bbox ? { bbox: signatureDate.bbox } : { estimated: true }),
    };
    fieldLocations.signatureDate = loc;
    fieldLocations.signDate = loc;
  }

  if (documentDate) {
    fieldLocations.documentDate = fieldLocations.topRightDate || fieldLocations.signatureDate;
  }

  const sigBoxes = getBoxesForPage(signaturePage);
  if (signer) {
    const bbox = findLinesBBoxByQuery(sigBoxes, signer, 0.1, 1);
    fieldLocations.signer = bbox ? { page: signaturePage, bbox } : { page: signaturePage, estimated: true };
  }

  if (signerRole) {
    const firstRolePart = signerRole.split(' - ')[0] || signerRole;
    const bbox =
      findLinesBBoxByQuery(sigBoxes, signerRole, 0.1, 1) ||
      findLinesBBoxByQuery(sigBoxes, firstRolePart, 0.1, 1);
    fieldLocations.signerRole = bbox ? { page: signaturePage, bbox } : { page: signaturePage, estimated: true };
  }

  const needsReview = reviewReasons.length > 0;

  const pageTexts: PdfPageContent[] = [];
  for (let p = 1; p <= pageCount; p++) {
    pageTexts.push({
      page: p,
      text: (pageTextByNum.get(p) || '').trim(),
    });
  }
  const indexedPageCount = pageTexts.filter((p) => countMeaningfulChars(p.text) >= 10).length;

  return {
    sha256,
    pageTexts,
    extraction: {
      docNumber,
      docNumberNeedsConfirmation,
      docNumberReviewReason: docNumberReviewReason || undefined,
      docName,
      title,
      documentDate,
      issueDate,
      signDate,
      dateComparisonStatus,
      topRightDate,
      signatureDate,
      signer,
      signerRole,
      parentAgency,
      issuingAgency,
      needsReview,
      reviewReasons,
      relevantPages,
      usedOcr,
      aiUsed,
      ocrConfidence,
      pageCount,
      signerCandidates,
      fieldLocations,
      indexedPageCount,
    },
  };
}
