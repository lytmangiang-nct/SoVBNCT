import React, { useEffect, useRef, useState } from 'react';
import {
  ChevronLeft,
  ChevronRight,
  ZoomIn,
  ZoomOut,
  ExternalLink,
  Loader2,
  AlertCircle,
  FileText,
} from 'lucide-react';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

import type { NormalizedBBox } from '../types';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export interface PdfHighlightTarget {
  page: number;
  bbox?: NormalizedBBox;
  label?: string;
  estimated?: boolean;
  nonce?: number;
}

interface PdfPreviewPaneProps {
  pdfUrl: string;
  fileName?: string;
  relevantPages?: number[];
  jumpToPage?: number | null;
  highlightBox?: PdfHighlightTarget | null;
}

export const PdfPreviewPane: React.FC<PdfPreviewPaneProps> = ({
  pdfUrl,
  fileName,
  relevantPages = [],
  jumpToPage,
  highlightBox,
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [pdfDoc, setPdfDoc] = useState<any>(null);
  const [numPages, setNumPages] = useState<number>(0);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [scale, setScale] = useState<number>(1.1);
  const [loading, setLoading] = useState<boolean>(true);
  const [rendering, setRendering] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  // Load PDF document
  useEffect(() => {
    let cancelled = false;
    let loadedDoc: any = null;

    let loadedTask: any = null;

    const loadPdf = async () => {
      if (!pdfUrl) return;
      setLoading(true);
      setError(null);
      setPdfDoc(null);
      setCurrentPage(1);

      try {
        const res = await fetch(pdfUrl);
        if (!res.ok) {
          throw new Error(`Không thể tải bản xem trước PDF (HTTP ${res.status})`);
        }
        const arrayBuffer = await res.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({
          data: new Uint8Array(arrayBuffer),
        });
        loadedTask = loadingTask;
        const doc: any = await loadingTask.promise;
        if (cancelled) {
          doc?.destroy?.().catch?.(() => {});
          loadingTask?.destroy?.().catch?.(() => {});
          return;
        }
        loadedDoc = doc;
        setPdfDoc(doc);
        setNumPages(doc.numPages || 1);
      } catch (err: any) {
        if (!cancelled) {
          setError(err.message || 'Không thể hiển thị xem trước tệp PDF.');
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    loadPdf();

    return () => {
      cancelled = true;
      if (loadedDoc?.destroy) {
        loadedDoc.destroy().catch(() => {});
      } else if (loadedTask?.destroy) {
        loadedTask.destroy().catch(() => {});
      }
    };
  }, [pdfUrl]);

  // Sync external page jump requests
  useEffect(() => {
    if (jumpToPage && numPages > 0 && jumpToPage >= 1 && jumpToPage <= numPages) {
      setCurrentPage(jumpToPage);
    }
  }, [jumpToPage, numPages]);

  // Sync highlightBox jump & smooth scroll to region
  useEffect(() => {
    if (!highlightBox || numPages <= 0) return;
    const targetPage = Math.max(1, Math.min(numPages, highlightBox.page || 1));
    setCurrentPage(targetPage);

    const timer = setTimeout(() => {
      if (!containerRef.current || !canvasRef.current) return;
      if (highlightBox.bbox) {
        const canvasHeight = canvasRef.current.clientHeight || 0;
        const containerHeight = containerRef.current.clientHeight || 0;
        if (canvasHeight > 0) {
          const targetTop = Math.max(0, highlightBox.bbox.y0 * canvasHeight - containerHeight * 0.25);
          containerRef.current.scrollTo({ top: targetTop, behavior: 'smooth' });
        }
      } else {
        containerRef.current.scrollTo({ top: 0, behavior: 'smooth' });
      }
    }, 180);

    return () => clearTimeout(timer);
  }, [highlightBox, numPages]);

  // Render active page to canvas
  useEffect(() => {
    let renderTask: any = null;
    let cancelled = false;

    const renderPage = async () => {
      if (!pdfDoc || !canvasRef.current) return;
      setRendering(true);

      try {
        const page = await pdfDoc.getPage(currentPage);
        if (cancelled) return;

        const baseViewport = page.getViewport({ scale: 1 });
        const containerWidth = containerRef.current?.clientWidth
          ? Math.max(containerRef.current.clientWidth - 24, 280)
          : 540;
        const fitScale = containerWidth / baseViewport.width;
        const effectiveScale = Math.max(0.4, Math.min(2.6, fitScale * scale));
        const dpr = Math.min(window.devicePixelRatio || 1, 2);

        const viewport = page.getViewport({ scale: effectiveScale });
        const canvas = canvasRef.current;
        const context = canvas.getContext('2d');
        if (!context) return;

        canvas.width = Math.floor(viewport.width * dpr);
        canvas.height = Math.floor(viewport.height * dpr);
        canvas.style.width = `${Math.floor(viewport.width)}px`;
        canvas.style.height = `${Math.floor(viewport.height)}px`;

        context.setTransform(dpr, 0, 0, dpr, 0, 0);

        renderTask = page.render({
          canvasContext: context,
          viewport,
        });
        await renderTask.promise;
      } catch (err: any) {
        if (err?.name !== 'RenderingCancelledException' && !cancelled) {
          console.error('Lỗi vẽ trang PDF:', err);
        }
      } finally {
        if (!cancelled) setRendering(false);
      }
    };

    renderPage();

    return () => {
      cancelled = true;
      if (renderTask) {
        try {
          renderTask.cancel();
        } catch {
          // ignore
        }
      }
    };
  }, [pdfDoc, currentPage, scale]);

  const uniqueRelevantPages = Array.from(
    new Set(relevantPages.filter((p) => p >= 1 && p <= (numPages || 999)))
  ).sort((a, b) => a - b);

  return (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs flex flex-col">
      {/* Toolbar */}
      <div className="px-3.5 py-2.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2 min-w-0">
          <FileText className="w-4 h-4 text-sky-700 shrink-0" />
          <span className="text-xs font-bold text-slate-800 truncate max-w-[200px]" title={fileName}>
            {fileName || 'Bản PDF gốc'}
          </span>
        </div>

        <div className="flex items-center gap-1.5 flex-wrap">
          {/* Page navigation */}
          <div className="flex items-center bg-white border border-slate-200 rounded-lg overflow-hidden">
            <button
              type="button"
              disabled={currentPage <= 1}
              onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              className="p-1.5 text-slate-600 hover:bg-slate-100 disabled:opacity-40"
              title="Trang trước"
            >
              <ChevronLeft className="w-3.5 h-3.5" />
            </button>
            <span className="px-2 text-[11px] font-mono font-semibold text-slate-700">
              {numPages > 0 ? `${currentPage}/${numPages}` : '...'}
            </span>
            <button
              type="button"
              disabled={currentPage >= numPages}
              onClick={() => setCurrentPage((p) => Math.min(numPages, p + 1))}
              className="p-1.5 text-slate-600 hover:bg-slate-100 disabled:opacity-40"
              title="Trang sau"
            >
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Zoom */}
          <div className="flex items-center bg-white border border-slate-200 rounded-lg overflow-hidden">
            <button
              type="button"
              onClick={() => setScale((s) => Math.max(0.6, +(s - 0.15).toFixed(2)))}
              className="p-1.5 text-slate-600 hover:bg-slate-100"
              title="Thu nhỏ"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="px-1.5 text-[11px] font-mono text-slate-600">
              {Math.round(scale * 100)}%
            </span>
            <button
              type="button"
              onClick={() => setScale((s) => Math.min(2.0, +(s + 0.15).toFixed(2)))}
              className="p-1.5 text-slate-600 hover:bg-slate-100"
              title="Phóng to"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          <a
            href={pdfUrl}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1 px-2.5 py-1.5 text-[11px] font-semibold text-sky-700 bg-white hover:bg-sky-50 border border-slate-200 rounded-lg"
            title="Mở bản PDF ở tab mới"
          >
            <ExternalLink className="w-3 h-3" />
            <span>Tab mới</span>
          </a>
        </div>
      </div>

      {/* Quick jump to relevant pages */}
      {uniqueRelevantPages.length > 0 && (
        <div className="px-3.5 py-2 bg-amber-50/70 border-b border-amber-100 flex flex-wrap items-center gap-1.5 text-[11px]">
          <span className="font-semibold text-amber-900 mr-1">Chuyển nhanh:</span>
          {uniqueRelevantPages.map((p) => (
            <button
              key={p}
              type="button"
              onClick={() => setCurrentPage(p)}
              className={`px-2 py-0.5 rounded font-semibold border transition-colors ${
                currentPage === p
                  ? 'bg-amber-600 text-white border-amber-700'
                  : 'bg-white text-amber-900 border-amber-200 hover:bg-amber-100'
              }`}
            >
              {p === 1
                ? 'Trang 1 (Đầu VB)'
                : p === numPages
                ? `Trang ${p} (Chữ ký)`
                : `Trang ${p}`}
            </button>
          ))}
        </div>
      )}

      {/* Estimated / Page-only highlight warning */}
      {highlightBox && (highlightBox.estimated || !highlightBox.bbox) && (
        <div className="px-3.5 py-2 bg-amber-50 border-b border-amber-200 flex items-center justify-between gap-2 text-[11px] text-amber-900">
          <div className="flex items-center gap-2">
            <span className="px-1.5 py-0.5 rounded bg-amber-600 text-white font-bold uppercase text-[10px]">
              Cần kiểm tra
            </span>
            <span>
              Đã chuyển tới <strong>Trang {highlightBox.page}</strong> cho trường{' '}
              <strong>{highlightBox.label || 'đang chọn'}</strong> (chỉ xác định được trang, chưa khoanh chính xác vùng chữ).
            </span>
          </div>
        </div>
      )}

      {/* Canvas area */}
      <div
        ref={containerRef}
        className="bg-slate-100 p-3 overflow-auto max-h-[680px] min-h-[380px] flex flex-col items-center justify-start relative"
      >
        {loading ? (
          <div className="my-auto flex flex-col items-center gap-2 text-slate-500 py-16">
            <Loader2 className="w-6 h-6 animate-spin text-sky-700" />
            <span className="text-xs font-medium">Đang tải xem trước bản PDF gốc...</span>
          </div>
        ) : error ? (
          <div className="my-auto p-4 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-start gap-2 max-w-md">
            <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
            <span>{error}</span>
          </div>
        ) : (
          <div className="relative shadow-md bg-white border border-slate-200 inline-block">
            {rendering && (
              <div className="absolute top-2 right-2 z-20 bg-slate-900/75 text-white text-[10px] px-2 py-0.5 rounded flex items-center gap-1">
                <Loader2 className="w-3 h-3 animate-spin" />
                <span>Đang hiển thị...</span>
              </div>
            )}
            <canvas ref={canvasRef} className="block max-w-full" />
            {highlightBox &&
              highlightBox.page === currentPage &&
              highlightBox.bbox && (
                <div
                  aria-label={`Vùng nguồn: ${highlightBox.label || ''}`}
                  className="pointer-events-none absolute z-10 border-2 border-amber-500 bg-amber-300/25 rounded-sm shadow-[0_0_0_2px_rgba(245,158,11,0.35)] transition-all duration-200"
                  style={{
                    left: `${Math.max(0, highlightBox.bbox.x0 * 100)}%`,
                    top: `${Math.max(0, highlightBox.bbox.y0 * 100)}%`,
                    width: `${Math.max((highlightBox.bbox.x1 - highlightBox.bbox.x0) * 100, 3)}%`,
                    height: `${Math.max((highlightBox.bbox.y1 - highlightBox.bbox.y0) * 100, 1.8)}%`,
                  }}
                >
                  {highlightBox.label && (
                    <span className="absolute -top-5 left-0 px-1.5 py-0.5 bg-amber-600 text-white text-[10px] font-bold rounded shadow-xs whitespace-nowrap">
                      {highlightBox.label}
                    </span>
                  )}
                </div>
              )}
          </div>
        )}
      </div>
    </div>
  );
};
