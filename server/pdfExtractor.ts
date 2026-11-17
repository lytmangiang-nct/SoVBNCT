import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createRequire } from 'module';
import { execFile } from 'child_process';
import { pathToFileURL } from 'url';
import { promisify } from 'util';
import type {
  ExtractableFieldKey,
  FieldLocation,
  NormalizedBBox,
  PdfExtractionResult,
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

async function renderPdfPageToPng(pdfPath: string, pageNumber: number, outPngPath: string): Promise<void> {
  await execFileAsync('gs', [
    '-dNOPAUSE',
    '-dBATCH',
    '-dQUIET',
    '-sDEVICE=png16m',
    '-r200',
    `-dFirstPage=${pageNumber}`,
    `-dLastPage=${pageNumber}`,
    `-sOutputFile=${outPngPath}`,
    pdfPath,
  ]);
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
  const dn = parseInt(d, 10);
  const mn = parseInt(m, 10);
  const yn = parseInt(y, 10);
  if (dn < 1 || dn > 31 || mn < 1 || mn > 12 || yn < 1990 || yn > 2100) return '';
  return `${d}/${m}/${y}`;
}

function extractIssueDate(text: string): string {
  if (!text) return '';
  const norm = removeVietnameseDiacritics(text);
  // 1. Dạng "ngày 12 tháng 11 năm 2026"
  const mText = norm.match(/\bngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})\b/i);
  if (mText) {
    return formatDateDdMmYyyy(mText[1], mText[2], mText[3]);
  }
  // 2. Dạng "ngày 12/11/2026" hoặc "ngày 12-11-2026"
  const mSlash = norm.match(/\bngay\s+(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})\b/i);
  if (mSlash) {
    return formatDateDdMmYyyy(mSlash[1], mSlash[2], mSlash[3]);
  }
  return '';
}

function extractSeparateSignDate(allTexts: string[], rawSigDates: string[]): string {
  const combined = allTexts.join('\n');
  // 1. Khối chữ ký số: "Ngày ký: ...", "Thời gian ký: ...", "Ký ngày: ..." — chỉ lưu ngày/tháng/năm (dd/MM/yyyy)
  const explicitPatterns = [
    /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y|Signed\s+on|Signing\s+time)\s*[:\-]?\s*(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?/i,
    /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y|Signed\s+on|Signing\s+time)\s*[:\-]?\s*(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?/i,
    /(?:Ng[àa]y\s+k[ýy]|Th[ờơ]i\s+gian\s+k[ýy]|K[ýy]\s+ng[àa]y)\s*[:\-]?\s*(?:ng[àa]y\s+)?(\d{1,2})\s+th[áa]ng\s+(\d{1,2})\s+n[ăa]m\s+(\d{4})/i,
  ];

  const m1 = combined.match(explicitPatterns[0]);
  if (m1) {
    const base = formatDateDdMmYyyy(m1[1], m1[2], m1[3]);
    if (base) return normalizeDateOnlyDdMmYyyy(base);
  }

  const m2 = combined.match(explicitPatterns[1]);
  if (m2) {
    const base = formatDateDdMmYyyy(m2[3], m2[2], m2[1]);
    if (base) return normalizeDateOnlyDdMmYyyy(base);
  }

  const m3 = combined.match(explicitPatterns[2]);
  if (m3) {
    const base = formatDateDdMmYyyy(m3[1], m3[2], m3[3]);
    if (base) return normalizeDateOnlyDdMmYyyy(base);
  }

  // 2. Ngày trong chữ ký số PDF (/M D:...) — đã bỏ giờ phút giây
  if (rawSigDates.length > 0) {
    return normalizeDateOnlyDdMmYyyy(rawSigDates[0]);
  }

  return '';
}

function extractDocNumberFromSoValue(rawAfterLabel: string): string {
  if (!rawAfterLabel) return '';
  const cleaned = rawAfterLabel
    .replace(/[ \t]*\/[ \t]*/g, '/')
    .replace(/[ \t]*-[ \t]*/g, '-')
    .trim();
  if (!cleaned) return '';

  // Lấy giá trị số/ký hiệu đầy đủ (giữ số 0 ở đầu, dấu /, dấu -, chữ cái)
  const primaryMatch = cleaned.match(
    /^([0-9]{1,5}[\/\-][\p{L}\p{N}_\-\.\/]+)/u
  );
  let candidate = primaryMatch ? primaryMatch[1] : cleaned.split(/\s+/)[0] || '';
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
  if (/^S[ốoôơ6]\s+\d{1,5}\s*[\/\-]/i.test(cleaned)) {
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
    /(?:S[ốoôơ6]|Số\/Ký hiệu)\s*[:\.]?[ \t]*([0-9]{1,5}[ \t]*[\/\-][ \t]*[\p{L}\p{N}_\-\.\/]+(?:[ \t]*[\-\/][ \t]*[\p{L}\p{N}_\-\.]+)*)/giu;
  const results: string[] = [];
  const seen = new Set<string>();

  for (const m of text.matchAll(regex)) {
    let num = (m[1] || '')
      .replace(/[ \t]*\/[ \t]*/g, '/')
      .replace(/[ \t]*-[ \t]*/g, '-')
      .replace(/[.,;:\)]+$/g, '')
      .trim();

    num = num.split(/\s+/)[0] || '';
    if (!num) continue;
    // Bỏ qua nếu vô tình bắt nhầm ngày dạng 14/11/2026
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(num)) continue;
    const key = num.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      results.push(num);
    }
  }
  return results;
}

function isMostlyUppercaseLine(str: string): boolean {
  const letters = (str || '').match(/\p{L}/gu) || [];
  if (letters.length === 0) return false;
  const upperCount = letters.filter((ch) => ch === ch.toUpperCase() && ch !== ch.toLowerCase()).length;
  return upperCount / letters.length >= 0.85;
}

function detectHeaderColumnSplitX(topBoxes: NormalizedBox[]): number {
  const headerTopBoxes = topBoxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.30);
  if (headerTopBoxes.length === 0) return 0.48;

  // 1. Neo theo khối quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" / "Độc lập..." bên phải
  const rightMottoRegex = /^(?:CONG|HOA|XA|HOI|CHU|NGHIA|VIET|NAM|DOC|LAP|TU|DO|HANH|PHUC)$/i;
  const rightAnchors = headerTopBoxes.filter((b) => {
    const norm = removeVietnameseDiacritics(b.text).replace(/[^A-Za-z]/g, '');
    return b.x0 >= 0.36 && rightMottoRegex.test(norm);
  });
  if (rightAnchors.length >= 2) {
    const minRightX = Math.min(...rightAnchors.map((b) => b.x0));
    return Math.min(0.54, Math.max(0.38, minRightX - 0.01));
  }

  // 2. Fallback: tìm khoảng trắng dọc (gutter) rộng nhất trong khoảng x = 0.36..0.56
  const leftCandidates = headerTopBoxes.filter((b) => b.x1 <= 0.52);
  const rightCandidates = headerTopBoxes.filter((b) => b.x0 >= 0.40);
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
        if (gapMid >= 0.36 && gapMid <= 0.56 && gapSize > bestGap) {
          bestGap = gapSize;
          bestSplit = gapMid;
        }
      }
      curMaxX = Math.max(curMaxX, intervals[i].x1);
    }
    if (bestGap >= 0.025) {
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

  // 1. Direct single-line or 2-line window substring match
  for (let i = 0; i < lines.length; i++) {
    const norm1 = removeVietnameseDiacritics(lines[i].text).toLowerCase().replace(/\s+/g, ' ').trim();
    if (norm1.includes(normQuery) || (norm1.length >= 6 && normQuery.includes(norm1) && norm1.length >= normQuery.length * 0.5)) {
      return unionLineBBoxes([lines[i]]);
    }
    if (i + 1 < lines.length) {
      const combined = `${norm1} ${removeVietnameseDiacritics(lines[i + 1].text).toLowerCase().replace(/\s+/g, ' ').trim()}`;
      if (combined.includes(normQuery)) {
        return unionLineBBoxes([lines[i], lines[i + 1]]);
      }
    }
  }

  // 2. Best token overlap match across 1..3 consecutive lines
  if (queryWords.length >= 2) {
    let bestScore = 0;
    let bestSlice: NormalizedLine[] = [];
    for (let i = 0; i < lines.length; i++) {
      for (let len = 1; len <= 3 && i + len <= lines.length; len++) {
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
    if (bestScore >= 0.6 && bestSlice.length > 0) {
      return unionLineBBoxes(bestSlice);
    }
  }

  return undefined;
}

interface TopLeftHeaderAnalysis {
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
}

function analyzeTopLeftHeader(
  page1Boxes: NormalizedBox[],
  graphicDividerYs: number[] = []
): TopLeftHeaderAnalysis {
  if (page1Boxes.length === 0) {
    return {
      issuingAgency: '',
      agencyNeedsReview: false,
      agencyReviewReason: '',
      docNumber: '',
      docNumberNeedsReview: false,
      docNumberReviewReason: '',
      topLeftHeaderDocNumbers: [],
      rightTopText: '',
      leftTopText: '',
    };
  }

  const topBoxes = page1Boxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.45);
  const splitX = detectHeaderColumnSplitX(topBoxes);

  const leftTopBoxes = topBoxes.filter((b) => {
    const xCenter = (b.x0 + b.x1) / 2;
    return xCenter < splitX && b.x0 < splitX - 0.004;
  });
  const rightTopBoxes = topBoxes.filter((b) => b.x0 >= splitX - 0.01);

  const leftTopLines = groupBoxesIntoLines(leftTopBoxes);
  const rightTopText = boxesToText(rightTopBoxes);
  const leftTopText = leftTopLines
    .filter((l) => !l.isDivider)
    .map((l) => l.text)
    .join('\n');

  const headerStopRegex =
    /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|TO TRINH|CHUONG TRINH|HUONG DAN|CONG VAN|QUY CHE|GIAY MOI|THU MOI|CAN CU|KINH GUI|V\/V\b|VE VIEC\b|DIEU 1|I\.\s)/i;

  // Tìm dòng "Số:" trong khối phía trên bên trái (trước phần thân văn bản)
  let firstSoIndex = -1;
  const headerSoNumbers: string[] = [];
  const headerSoLines: NormalizedLine[] = [];

  for (let i = 0; i < leftTopLines.length; i++) {
    const line = leftTopLines[i];
    if (line.isDivider) continue;
    const norm = removeVietnameseDiacritics(line.text);
    if (firstSoIndex !== -1 && headerStopRegex.test(norm)) {
      break;
    }
    if (firstSoIndex === -1 && headerStopRegex.test(norm)) {
      break;
    }
    if (isSoLabelLine(line.text)) {
      if (firstSoIndex === -1) firstSoIndex = i;
      let val = extractValueFromSoLine(line.text);
      const matchedSoLines: NormalizedLine[] = [line];
      // Nếu "Số:" đứng một mình trên dòng và giá trị nằm ngay dòng sát bên dưới
      if (!val && i + 1 < leftTopLines.length) {
        const nextLine = leftTopLines[i + 1];
        if (!nextLine.isDivider && nextLine.yCenter - line.yCenter <= 0.03) {
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
      'Đã thấy nhãn "Số:" ở góc trên bên trái nhưng chưa đọc chắc chắn phần số/ký hiệu phía sau.';
  } else if (headerSoNumbers.length > 1) {
    docNumberNeedsReview = true;
    docNumberReviewReason = `Phát hiện nhiều dòng "Số:" trong khối tiêu đề góc trên bên trái (${headerSoNumbers.join(', ')}), vui lòng xác nhận đúng số gốc.`;
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
    // 2 dòng: dòng 1 là cơ quan chủ quản, dòng 2 là cơ quan ban hành
    issuingAgency = rawAgencyLines[1].text;
    agencyUsedLines = [rawAgencyLines[1]];
  } else if (rawAgencyLines.length === 3) {
    const [line1, line2, line3] = rawAgencyLines;
    const dividerBetween23 = hasDividerBetween(line2.yCenter, line3.yCenter);
    const dividerBetween12 = hasDividerBetween(line1.yCenter, line2.yCenter);
    const dividerBelow3 =
      firstSoIndex !== -1
        ? hasDividerBetween(line3.yCenter, leftTopLines[firstSoIndex].yCenter)
        : allDividerYs.some((dy) => dy > line3.yCenter && dy - line3.yCenter <= 0.045);

    const centerDiff23 = Math.abs(line2.xCenter - line3.xCenter);
    const leftDiff23 = Math.abs(line2.x0 - line3.x0);
    const aligned23 = centerDiff23 <= 0.045 || leftDiff23 <= 0.025;

    const dy12 = Math.max(0.004, line2.yCenter - line1.yCenter);
    const dy23 = Math.max(0.004, line3.yCenter - line2.yCenter);
    const spacing23Tight = dy23 <= dy12 * 1.15;

    const maxH23 = Math.max(line2.avgHeight, line3.avgHeight, 0.001);
    const heightSimilar23 = Math.abs(line2.avgHeight - line3.avgHeight) / maxH23 <= 0.28;
    const bothUpper23 = isMostlyUppercaseLine(line2.text) && isMostlyUppercaseLine(line3.text);
    const fontsAvailable = line2.fontNames.length > 0 && line3.fontNames.length > 0;
    const fontConsistent23 =
      !fontsAvailable || line2.fontNames.some((f) => line3.fontNames.includes(f));

    if (dividerBetween23) {
      // Có đường kẻ nằm giữa dòng 2 và dòng 3 => dòng 3 không phải phần xuống dòng của dòng 2
      issuingAgency = line2.text;
      agencyUsedLines = [line2];
      agencyNeedsReview = true;
      agencyReviewReason =
        'Khối cơ quan ban hành có 3 dòng và có đường kẻ phân tách trước dòng thứ 3, vui lòng kiểm tra lại.';
    } else if (
      !dividerBetween12 &&
      aligned23 &&
      spacing23Tight &&
      heightSimilar23 &&
      bothUpper23 &&
      fontConsistent23
    ) {
      // Dòng 3 chắc chắn là phần xuống dòng của dòng 2 (cùng căn giữa/lề, khoảng cách dòng gần, cùng kiểu chữ)
      issuingAgency = `${line2.text} ${line3.text}`.replace(/\s+/g, ' ').trim();
      agencyUsedLines = [line2, line3];
      agencyNeedsReview = false;
    } else {
      // Chưa chắc chắn (ví dụ dòng 1-2 xuống dòng còn dòng 3 là cơ quan, hoặc khoảng cách/kiểu chữ lệch)
      const looksLike12Wrapped =
        dy12 * 1.2 < dy23 &&
        Math.abs(line1.xCenter - line2.xCenter) <= 0.045 &&
        isMostlyUppercaseLine(line3.text);

      if (looksLike12Wrapped && !dividerBelow3) {
        issuingAgency = line3.text;
        agencyUsedLines = [line3];
      } else if (bothUpper23) {
        issuingAgency = `${line2.text} ${line3.text}`.replace(/\s+/g, ' ').trim();
        agencyUsedLines = [line2, line3];
      } else {
        issuingAgency = line2.text;
        agencyUsedLines = [line2];
      }
      agencyNeedsReview = true;
      agencyReviewReason =
        'Khối tên cơ quan ban hành có 3 dòng nhưng chưa đủ chắc chắn phần xuống dòng, vui lòng kiểm tra và chỉnh sửa nếu cần.';
    }
  } else {
    // >= 4 dòng phía trên Số:
    const tail = rawAgencyLines.slice(1);
    issuingAgency = tail.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim();
    agencyUsedLines = tail;
    agencyNeedsReview = true;
    agencyReviewReason = `Khối tên cơ quan ban hành có ${rawAgencyLines.length} dòng phía trên dòng Số:, vui lòng kiểm tra và xác nhận đúng cơ quan ban hành.`;
  }

  return {
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

function extractSubjectAndTitle(page1Text: string): {
  docName: string;
  title: string;
  hasVv: boolean;
} {
  const lines = page1Text
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);

  // 1. Ưu tiên tìm dòng "V/v" hoặc "Về việc" (công văn)
  for (let i = 0; i < lines.length; i++) {
    const vvMatch = lines[i].match(/^(?:[-•*]\s*)?(?:V\/v|V[ềe]\s+vi[ệe]c)\s*[:\-]?\s*(.*)$/i);
    if (vvMatch) {
      const parts: string[] = [];
      if (vvMatch[1]?.trim()) parts.push(vvMatch[1].trim());
      for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
        const nextLine = lines[j];
        const normNext = removeVietnameseDiacritics(nextLine);
        if (
          /^(Kinh gui|Can cu|Thuc hien|Theo|I\.|1\.|Noi nhan|Doc lap)/i.test(normNext) ||
          /^[A-ZĐÂĂÊÔƠƯ\s]{8,}$/.test(nextLine)
        ) {
          break;
        }
        parts.push(nextLine);
      }
      const afterVv = capitalizeSentenceStart(parts.join(' ').replace(/\s+/g, ' ').trim());
      if (afterVv) {
        const prevHeading = i > 0 ? lines[i - 1] : '';
        const isHeading = /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|TO TRINH|CHUONG TRINH|HUONG DAN|GIAY MOI)$/i.test(
          removeVietnameseDiacritics(prevHeading)
        );
        return {
          docName: isHeading ? `${prevHeading} ${afterVv}` : afterVv,
          title: afterVv,
          hasVv: true,
        };
      }
    }
  }

  // 2. Nếu không có "V/v", tìm tiêu đề văn bản in hoa (KẾ HOẠCH / QUYẾT ĐỊNH / THÔNG BÁO...) + dòng trích yếu ngay sau
  const headingKeywords =
    /^(KE HOACH|QUYET DINH|THONG BAO|BAO CAO|CHUONG TRINH|HUONG DAN|TO TRINH|QUY CHE|GIAY MOI|THU MOI|CONG VAN)\b/i;
  for (let i = 0; i < lines.length; i++) {
    if (headingKeywords.test(removeVietnameseDiacritics(lines[i]))) {
      const subParts: string[] = [];
      for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
        const normSub = removeVietnameseDiacritics(lines[j]);
        if (
          /^(Can cu|Kinh gui|Dieu 1|I\.|1\.|So\s*:|Noi nhan|Thuc hien|Theo )/i.test(normSub) ||
          SIGNER_ROLE_REGEX.test(normSub)
        ) {
          break;
        }
        subParts.push(lines[j]);
      }
      const subtitle = subParts.join(' ').replace(/\s+/g, ' ').trim();
      const combined = [lines[i], subtitle].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim();
      return {
        docName: combined,
        title: subtitle || combined,
        hasVv: false,
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
    .slice(0, 6);

  for (let i = 0; i < topLines.length; i++) {
    const line = removeVietnameseDiacritics(topLines[i]);
    if (
      /^(?:PHU\s*LUC(?:\s*[IVXLCDM0-9]+|\b)|DANH\s*SACH\s*KEM\s*THEO|BANG\s*(?:KEM\s*THEO|PHAN\s*CONG|TONG\s*HOP|DU\s*TOAN|THEO\s*DOI)|BIEU\s*MAU|MAU\s*SO)\b/i.test(
        line
      )
    ) {
      return true;
    }
    if (
      i < 4 &&
      /^\(?\s*(?:Ban\s+hanh\s+)?kem\s+theo\s+(?:Cong\s+van|Quyet\s+dinh|Ke\s+hoach|Thong\s+bao|To\s+trinh|Huong\s+dan|Chuong\s+trinh)\s+so\b/i.test(
        line
      )
    ) {
      return true;
    }
  }
  return false;
}

const SIGNER_ROLE_REGEX =
  /^(KT\.|TM\.|Q\.|TL\.|TUQ\.|GIAM DOC|PHO GIAM DOC|CHU TICH|PHO CHU TICH|HIEU TRUONG|PHO HIEU TRUONG|BI THU|PHO BI THU|TRUONG PHONG|PHO TRUONG PHONG|CHANH VAN PHONG|PHO CHANH VAN PHONG|TRUONG BAN|PHO TRUONG BAN|THU TRUONG|CUC TRUONG|PHO CUC TRUONG|VU TRUONG|CHU NHIEM|PHO CHU NHIEM|BO TRUONG|UY VIEN)/i;

const NON_NAME_LINE_REGEX =
  /^(CONG HOA|DOC LAP|NOI NHAN|DA KY|SO GIAO DUC|UY BAN NHAN DAN|THANH PHO|TINH |TRUONG THPT|DOAN TRUONG|PHONG GIAO DUC|BO GIAO DUC|PHU LUC)/i;

const SKIP_SIG_LINE_REGEX =
  /^(Noi nhan|Luu|-\s*Nhu tren|Kinh gui|Ngay ky|Thoi gian ky|\(?Da ky\)?|\(?Ky,?\s*(?:ghi ro ho ten|dong dau|ten).*?\)?)/i;

function extractSignersFromZoneText(
  zoneText: string,
  pageNumber: number
): Array<{ name: string; role: string; page: number }> {
  const found: Array<{ name: string; role: string; page: number }> = [];
  if (!zoneText) return found;

  // 1. Mẫu chữ ký số ("Ký bởi: Nguyễn Văn A", "Người ký: ...", "Signed by: ...")
  const digitalMatches = zoneText.matchAll(
    /(?:K[ýy]\s+b[ởơ]i|Ng[ườờ]i\s+k[ýy]|Signed\s+by)\s*[:\-]\s*([^\n,;]{3,70})/gi
  );
  for (const dm of digitalMatches) {
    const candName = dm[1].replace(/\([^)]*\)/g, '').trim();
    if (candName.length >= 4) {
      found.push({ name: candName, role: '', page: pageNumber });
    }
  }

  // 2. Khối chức vụ + tên người ký bên dưới
  const lines = zoneText
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean);

  for (let i = 0; i < lines.length; i++) {
    const normLine = removeVietnameseDiacritics(lines[i]);
    if (SIGNER_ROLE_REGEX.test(normLine)) {
      const roleParts = [lines[i]];
      let nameLine = '';
      for (let j = i + 1; j < Math.min(i + 8, lines.length); j++) {
        const l = lines[j];
        const normL = removeVietnameseDiacritics(l);
        if (SKIP_SIG_LINE_REGEX.test(normL)) {
          continue;
        }
        if (SIGNER_ROLE_REGEX.test(normL) && j <= i + 2) {
          roleParts.push(l);
          continue;
        }
        // Tên người Việt: 2-5 từ bắt đầu chữ hoa, không chứa số hoặc dấu câu dài
        if (
          /^[\p{Lu}][\p{L}]+(?:\s+[\p{Lu}][\p{L}]+){1,4}$/u.test(l) &&
          !NON_NAME_LINE_REGEX.test(normL)
        ) {
          nameLine = l;
          break;
        }
      }
      if (nameLine) {
        found.push({
          name: nameLine,
          role: roleParts.join(' - '),
          page: pageNumber,
        });
      }
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
    // Ưu tiên góc dưới bên phải -> nửa phải -> nửa dưới -> toàn trang (kèm annotation chữ ký số)
    const zones = [p.rightBottomText, p.rightHalfText, p.bottomText, p.fullText].filter(
      (z) => z && z.trim().length > 0
    );

    let pageCandidates: Array<{ name: string; role: string; page: number }> = [];
    for (const zone of zones) {
      const mergedZone = [zone, p.annotationText].filter(Boolean).join('\n');
      const matches = extractSignersFromZoneText(mergedZone, p.pageNumber);
      if (matches.length > 0) {
        pageCandidates = matches;
        break;
      }
    }

    if (pageCandidates.length === 0 && p.annotationText) {
      pageCandidates = extractSignersFromZoneText(p.annotationText, p.pageNumber);
    }

    if (pageCandidates.length > 0) {
      candidates.push(...pageCandidates);
      // Khi đã tìm thấy khối chữ ký ở trang chính gần cuối nhất (trước phụ lục), dừng lùi về các trang trước
      break;
    }
  }

  for (const rs of rawSigSigners) {
    if (rs && !candidates.some((c) => c.name.toLowerCase() === rs.toLowerCase())) {
      candidates.push({
        name: rs,
        role: '',
        page: signaturePages[0]?.pageNumber || 1,
      });
    }
  }

  // Khử trùng giữ thứ tự
  const uniqueNames: string[] = [];
  for (const c of candidates) {
    if (!uniqueNames.some((n) => n.toLowerCase() === c.name.toLowerCase())) {
      uniqueNames.push(c.name);
    }
  }

  const primary = candidates[0];
  return {
    signer: primary?.name || '',
    signerRole: primary?.role || '',
    signerCandidates: uniqueNames,
    signaturePage: primary?.page || signaturePages[0]?.pageNumber || 1,
  };
}

export async function analyzePdfBuffer(buffer: Buffer): Promise<{
  sha256: string;
  extraction: PdfExtractionResult;
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

  // Nếu PDF scan nhiều trang (> 3) và các trang cuối là Phụ lục hoặc chưa có chữ ký, lùi thêm về các trang trước phụ lục
  if (pageCount > 3 && usedOcr) {
    const trailingHasNonAppendixSigner = [pageCount, pageCount - 1].some((pNum) => {
      const txt = pageTextByNum.get(pNum) || '';
      if (!txt || isAppendixPage(txt)) return false;
      const lines = txt.split('\n').filter(Boolean);
      const bottomApprox = lines.slice(Math.floor(lines.length * 0.35)).join('\n');
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

  if (!docNumber) {
    fallbackDocNumbers = extractDocNumberCandidates(page1Text);
    docNumber = fallbackDocNumbers[0] || '';
    usedFallbackDocNumber = Boolean(docNumber);
  }

  const issuingAgency =
    headerAnalysis.issuingAgency || extractIssuingAgencyFallback(page1Text);

  const issueDateSource = headerAnalysis.rightTopText || page1?.rightTopText || '';
  const page1Top35Text =
    page1Boxes.length > 0
      ? boxesToText(page1Boxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.35))
      : page1Text.split('\n').slice(0, 10).join('\n');
  const issueDate = normalizeDateOnlyDdMmYyyy(
    extractIssueDate(issueDateSource) || extractIssueDate(page1Top35Text)
  );
  const { docName, title } = extractSubjectAndTitle(page1Text);

  // 2. Xác định các trang ứng viên cho khối chữ ký: loại bỏ trang Phụ lục phía sau, quét lùi từ trang chính cuối về trang 1
  const orderedPageNums: number[] = [];
  for (let p = pageCount; p >= 1; p--) {
    orderedPageNums.push(p);
  }
  const nonAppendixPages = orderedPageNums.filter((pNum) => {
    const txt = pageTextByNum.get(pNum) || '';
    return !isAppendixPage(txt);
  });
  const appendixPages = orderedPageNums.filter((pNum) => !nonAppendixPages.includes(pNum));

  const buildZonePage = (pNum: number): SignatureZonePage => {
    const pInfo = pages.find((x) => x.pageNumber === pNum);
    const full = pageTextByNum.get(pNum) || '';
    const hasReliableSpans = Boolean(pInfo && !isGarbledText(pInfo.fullText));

    if (hasReliableSpans && pInfo) {
      return {
        pageNumber: pNum,
        rightBottomText: pInfo.rightBottomText,
        rightHalfText: pInfo.rightHalfText,
        bottomText: pInfo.bottomText,
        fullText: full,
        annotationText: pInfo.annotationText || '',
      };
    }

    const ocrBoxes = ocrBoxesByNum.get(pNum) || [];
    if (ocrBoxes.length > 0) {
      const rightBottomBoxes = ocrBoxes.filter((b) => (b.y0 + b.y1) / 2 >= 0.35 && b.x0 >= 0.40);
      const rightHalfBoxes = ocrBoxes.filter((b) => b.x0 >= 0.40);
      const bottomBoxes = ocrBoxes.filter((b) => (b.y0 + b.y1) / 2 >= 0.35);
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
    const ocrBottom = ocrLines.slice(Math.floor(ocrLines.length * 0.35)).join('\n');
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
  const { signer, signerRole, signerCandidates, signaturePage } = sigResult;

  const allTexts = [
    ...Array.from(pageTextByNum.values()),
    ...pages.map((p) => p.annotationText).filter(Boolean),
  ];
  // Chỉ lưu ngày ký (dd/MM/yyyy) khi phát hiện riêng ngày ký; giữ trống nếu chỉ có ngày ban hành
  const signDate = normalizeDateOnlyDdMmYyyy(extractSeparateSignDate(allTexts, rawSigDates));

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
  if (!docNumber) {
    reviewReasons.push('Chưa nhận diện được Số/ký hiệu văn bản.');
  } else if (headerAnalysis.docNumberNeedsReview && headerAnalysis.docNumberReviewReason) {
    reviewReasons.push(headerAnalysis.docNumberReviewReason);
  } else if (usedFallbackDocNumber) {
    reviewReasons.push(
      fallbackDocNumbers.length > 1
        ? `Chưa xác định rõ dòng "Số:" ở góc trên bên trái; phát hiện nhiều số trong trang đầu (${fallbackDocNumbers.join(', ')}), vui lòng kiểm tra đúng số gốc.`
        : 'Chưa xác định rõ dòng "Số:" ngay dưới khối cơ quan ở góc trên bên trái, vui lòng kiểm tra lại Số/ký hiệu đã trích xuất.'
    );
  }
  if (headerAnalysis.agencyNeedsReview && headerAnalysis.agencyReviewReason) {
    reviewReasons.push(headerAnalysis.agencyReviewReason);
  }
  if (!docName && !title) reviewReasons.push('Chưa nhận diện được Tên văn bản hoặc Trích yếu.');
  if (!issueDate) reviewReasons.push('Chưa nhận diện được Ngày ban hành.');
  if (!signer) reviewReasons.push('Chưa nhận diện được Họ tên người ký.');
  if (signerCandidates.length > 1) {
    reviewReasons.push(`Phát hiện nhiều tên người ký (${signerCandidates.join(', ')}), vui lòng xác nhận người ký chính.`);
  }

  const relevantPagesSet = new Set<number>([1]);
  if (signaturePage >= 1 && signaturePage <= pageCount) {
    relevantPagesSet.add(signaturePage);
  }
  const relevantPages = Array.from(relevantPagesSet).sort((a, b) => a - b);

  // 3. Ghi nhận vị trí trang và vùng chứa (bbox chuẩn hóa 0..1) cho từng trường để hỗ trợ "Xem vị trí"
  const getBoxesForPage = (pNum: number): NormalizedBox[] => {
    const pInfo = pages.find((p) => p.pageNumber === pNum);
    if (pInfo && !isGarbledText(pInfo.fullText) && pInfo.boxes.length > 0) {
      return pInfo.boxes;
    }
    return ocrBoxesByNum.get(pNum) || [];
  };

  const fieldLocations: Partial<Record<ExtractableFieldKey, FieldLocation>> = {};

  if (docNumber) {
    const bbox = headerAnalysis.docNumberBBox || findLinesBBoxByQuery(page1Boxes, docNumber, 0, 0.45);
    fieldLocations.docNumber = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  } else {
    fieldLocations.docNumber = { page: 1, estimated: true };
  }

  if (issuingAgency) {
    const bbox = headerAnalysis.agencyBBox || findLinesBBoxByQuery(page1Boxes, issuingAgency, 0, 0.42);
    fieldLocations.issuingAgency = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  } else {
    fieldLocations.issuingAgency = { page: 1, estimated: true };
  }

  if (docName) {
    const bbox = findLinesBBoxByQuery(page1Boxes, docName, 0.08, 0.85);
    fieldLocations.docName = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  }

  if (title) {
    const bbox = findLinesBBoxByQuery(page1Boxes, title, 0.08, 0.85);
    fieldLocations.title = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
  }

  if (issueDate) {
    const [dd, mm, yyyy] = issueDate.split('/');
    const topRightBoxes = page1Boxes.filter((b) => (b.y0 + b.y1) / 2 <= 0.40);
    const dateLine = groupBoxesIntoLines(topRightBoxes).find((l) => {
      const norm = removeVietnameseDiacritics(l.text).toLowerCase();
      return (
        norm.includes(issueDate) ||
        (dd && mm && yyyy && (norm.includes(`ngay ${parseInt(dd, 10)} thang ${parseInt(mm, 10)} nam ${yyyy}`) || norm.includes(`${dd}/${mm}/${yyyy}`)))
      );
    });
    const bbox = dateLine ? unionLineBBoxes([dateLine]) : undefined;
    fieldLocations.issueDate = bbox ? { page: 1, bbox } : { page: 1, estimated: true };
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

  if (signDate) {
    let foundSignLoc: FieldLocation | undefined;
    const candidatePages = [signaturePage, ...Array.from({ length: pageCount }, (_, i) => i + 1).filter((p) => p !== signaturePage)];
    for (const pNum of candidatePages) {
      const pBoxes = getBoxesForPage(pNum);
      const bbox = findLinesBBoxByQuery(pBoxes, signDate, 0, 1);
      if (bbox) {
        foundSignLoc = { page: pNum, bbox };
        break;
      }
    }
    fieldLocations.signDate = foundSignLoc || { page: signaturePage, estimated: true };
  }

  const needsReview = reviewReasons.length > 0;

  return {
    sha256,
    extraction: {
      docNumber,
      docName,
      title,
      issueDate,
      signDate,
      signer,
      signerRole,
      issuingAgency,
      needsReview,
      reviewReasons,
      relevantPages,
      usedOcr,
      ocrConfidence,
      pageCount,
      signerCandidates,
      fieldLocations,
    },
  };
}
