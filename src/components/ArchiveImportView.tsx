import React, { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  ExternalLink,
  Eye,
  FileSpreadsheet,
  FolderOpen,
  Loader2,
  RefreshCw,
  RotateCw,
  Upload,
} from 'lucide-react';
import { api } from '../api';
import type {
  ArchiveImportColumnMapping,
  ArchiveImportCommitRowResult,
  ArchiveImportEditableRow,
  ArchiveImportPreviewResponse,
  ArchiveImportPreviewRow,
  ArchiveImportRowStatus,
  ArchiveParseExcelResponse,
  DocumentRecord,
  RegistryBook,
  User,
} from '../types';
import { hasBookPermission, isSystemAdmin } from '../utils/permissions';

interface ArchiveImportViewProps {
  currentUser: User;
  books: RegistryBook[];
  activeBookId: string;
  onDocumentsImported: (createdDocs: DocumentRecord[], updatedBook?: RegistryBook) => void;
  onOpenExistingDocument: (docId: string) => void;
}

const MAPPING_FIELDS: Array<{
  key: keyof ArchiveImportColumnMapping;
  label: string;
  required?: boolean;
}> = [
  { key: 'docCategory', label: 'Nhóm văn bản (Đến/Đi)' },
  { key: 'fullDocNumber', label: 'Số/Ký hiệu', required: true },
  { key: 'docName', label: 'Tên văn bản', required: true },
  { key: 'issueDate', label: 'Ngày ban hành' },
  { key: 'year', label: 'Năm' },
  { key: 'issuingAgency', label: 'Cơ quan ban hành' },
  { key: 'signer', label: 'Người ký' },
  { key: 'driveLinkOrFileName', label: 'Liên kết hoặc tên tệp Drive' },
  { key: 'typeCode', label: 'Loại văn bản (tùy chọn)' },
  { key: 'signerRole', label: 'Chức vụ người ký (tùy chọn)' },
  { key: 'parentAgency', label: 'Cơ quan chủ quản (tùy chọn)' },
  { key: 'title', label: 'Trích yếu (tùy chọn)' },
  { key: 'notes', label: 'Ghi chú (tùy chọn)' },
];

function parseCategoryInput(raw: string): 'outgoing' | 'incoming' {
  const lower = (raw || '').trim().toLowerCase();
  if (lower.includes('đến') || lower.includes('den') || lower === 'incoming') {
    return 'incoming';
  }
  return 'outgoing';
}

export const ArchiveImportView: React.FC<ArchiveImportViewProps> = ({
  currentUser,
  books,
  activeBookId,
  onDocumentsImported,
  onOpenExistingDocument,
}) => {
  const isAdmin = isSystemAdmin(currentUser);
  const importableBooks = useMemo(
    () =>
      books.filter(
        (b) => b.isActive && (isAdmin || hasBookPermission(currentUser, b.id, 'import_archive'))
      ),
    [books, currentUser, isAdmin]
  );

  const [targetBookId, setTargetBookId] = useState<string>('');
  const [folderUrlOrId, setFolderUrlOrId] = useState<string>('');
  const [downloadingTpl, setDownloadingTpl] = useState(false);

  // Step 1: Excel Parsing
  const [parsingExcel, setParsingExcel] = useState(false);
  const [excelError, setExcelError] = useState<string | null>(null);
  const [parsedWorkbook, setParsedWorkbook] = useState<ArchiveParseExcelResponse | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState<string>('');
  const [mapping, setMapping] = useState<ArchiveImportColumnMapping>({});

  // Step 2: Preview & Disambiguation
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [previewResult, setPreviewResult] = useState<ArchiveImportPreviewResponse | null>(null);
  const [editableRows, setEditableRows] = useState<ArchiveImportEditableRow[]>([]);
  const [selectedRowIds, setSelectedRowIds] = useState<Record<string, boolean>>({});
  const [allowDuplicateMap, setAllowDuplicateMap] = useState<Record<string, boolean>>({});
  const [statusFilter, setStatusFilter] = useState<'all' | ArchiveImportRowStatus | 'unmatched_any'>('all');

  // Step 3: Batch Commit
  const [committing, setCommitting] = useState(false);
  const [commitProgress, setCommitProgress] = useState<{ done: number; total: number } | null>(null);
  const [commitResults, setCommitResults] = useState<ArchiveImportCommitRowResult[]>([]);
  const [commitSummary, setCommitSummary] = useState<{
    batchId: string;
    imported: number;
    skipped: number;
    errors: number;
  } | null>(null);
  const [lastBatchId, setLastBatchId] = useState<string>('');

  useEffect(() => {
    if (importableBooks.length === 0) return;
    if (!targetBookId || !importableBooks.some((b) => b.id === targetBookId)) {
      const preferred = importableBooks.find((b) => b.id === activeBookId) || importableBooks[0];
      setTargetBookId(preferred.id);
    }
  }, [importableBooks, targetBookId, activeBookId]);

  const activeTargetBook = useMemo(
    () => importableBooks.find((b) => b.id === targetBookId) || importableBooks[0],
    [importableBooks, targetBookId]
  );

  const selectedSheet = useMemo(() => {
    if (!parsedWorkbook) return null;
    return (
      parsedWorkbook.sheets.find((s) => s.name === selectedSheetName) ||
      parsedWorkbook.sheets[0] ||
      null
    );
  }, [parsedWorkbook, selectedSheetName]);

  const handleDownloadTemplate = async () => {
    setDownloadingTpl(true);
    setExcelError(null);
    try {
      await api.downloadArchiveTemplate();
    } catch (err: any) {
      setExcelError(err?.message || 'Không tải được tệp mẫu.');
    } finally {
      setDownloadingTpl(false);
    }
  };

  const handleExcelFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    setParsingExcel(true);
    setExcelError(null);
    setPreviewResult(null);
    setCommitResults([]);
    setCommitSummary(null);

    try {
      const parsed = await api.parseArchiveExcel(file);
      setParsedWorkbook(parsed);
      const firstSheet = parsed.sheets[0];
      if (firstSheet) {
        setSelectedSheetName(firstSheet.name);
        setMapping(firstSheet.suggestedMapping || {});
      }
      setLastBatchId(`arch-${Date.now().toString(36)}`);
    } catch (err: any) {
      setExcelError(err?.message || 'Không đọc được tệp Excel.');
    } finally {
      setParsingExcel(false);
    }
  };

  const handleChangeSheet = (sheetName: string) => {
    setSelectedSheetName(sheetName);
    const s = parsedWorkbook?.sheets.find((x) => x.name === sheetName);
    if (s) {
      setMapping(s.suggestedMapping || {});
    }
    setPreviewResult(null);
  };

  const buildRowsFromSheet = (): ArchiveImportEditableRow[] => {
    if (!selectedSheet) return [];
    const getCell = (row: Record<string, { value: string; hyperlink?: string }>, col?: string) => {
      if (!col || !row[col]) return { value: '', hyperlink: undefined };
      return row[col];
    };

    return selectedSheet.rows.map((rawRow, idx) => {
      const catCell = getCell(rawRow, mapping.docCategory);
      const numCell = getCell(rawRow, mapping.fullDocNumber);
      const nameCell = getCell(rawRow, mapping.docName);
      const dateCell = getCell(rawRow, mapping.issueDate);
      const yearCell = getCell(rawRow, mapping.year);
      const agencyCell = getCell(rawRow, mapping.issuingAgency);
      const signerCell = getCell(rawRow, mapping.signer);
      const roleCell = getCell(rawRow, mapping.signerRole);
      const typeCell = getCell(rawRow, mapping.typeCode);
      const parentCell = getCell(rawRow, mapping.parentAgency);
      const titleCell = getCell(rawRow, mapping.title);
      const notesCell = getCell(rawRow, mapping.notes);
      const driveCell = getCell(rawRow, mapping.driveLinkOrFileName);

      return {
        rowNumber: idx + 2,
        docCategory: parseCategoryInput(catCell.value),
        fullDocNumber: numCell.value,
        docName: nameCell.value,
        issueDate: dateCell.value,
        year: yearCell.value,
        issuingAgency: agencyCell.value,
        signer: signerCell.value,
        signerRole: roleCell.value,
        typeCode: typeCell.value,
        parentAgency: parentCell.value,
        title: titleCell.value,
        notes: notesCell.value,
        driveLinkOrFileName: driveCell.value,
        driveHyperlink: driveCell.hyperlink,
      };
    });
  };

  const runPreview = async (rowsToPreview: ArchiveImportEditableRow[], preserveSelection = false) => {
    if (!targetBookId || rowsToPreview.length === 0) return;
    setPreviewLoading(true);
    setPreviewError(null);
    setCommitResults([]);
    setCommitSummary(null);

    try {
      const res = await api.previewArchiveImport({
        bookId: targetBookId,
        folderUrlOrId: folderUrlOrId.trim() || undefined,
        rows: rowsToPreview,
      });
      setPreviewResult(res);
      setEditableRows(
        res.rows.map((r) => ({
          rowNumber: r.rowNumber,
          docCategory: r.docCategory,
          fullDocNumber: r.fullDocNumber,
          docName: r.docName,
          issueDate: r.issueDate,
          year: r.year,
          issuingAgency: r.issuingAgency,
          signer: r.signer,
          signerRole: r.signerRole,
          typeCode: r.typeCode,
          parentAgency: r.parentAgency,
          title: r.title,
          notes: r.notes,
          driveLinkOrFileName: r.driveLinkOrFileName,
          driveHyperlink: r.driveHyperlink,
          selectedDriveFileId: r.selectedDriveFileId || r.matchedDriveFile?.fileId,
        }))
      );
      if (!preserveSelection) {
        const nextSel: Record<string, boolean> = {};
        for (const r of res.rows) {
          nextSel[r.rowId] = r.status === 'ready';
        }
        setSelectedRowIds(nextSel);
      }
    } catch (err: any) {
      setPreviewError(err?.message || 'Lỗi khi kiểm tra dữ liệu và đối chiếu Google Drive.');
    } finally {
      setPreviewLoading(false);
    }
  };

  const updateEditableRowField = (
    idx: number,
    patch: Partial<ArchiveImportEditableRow>
  ) => {
    setEditableRows((prev) => prev.map((r, i) => (i === idx ? { ...r, ...patch } : r)));
  };

  const filteredPreviewRows = useMemo(() => {
    if (!previewResult) return [];
    if (statusFilter === 'all') return previewResult.rows;
    if (statusFilter === 'unmatched_any') {
      return previewResult.rows.filter(
        (r) => r.status === 'unmatched_file' || r.status === 'ambiguous_file'
      );
    }
    return previewResult.rows.filter((r) => r.status === statusFilter);
  }, [previewResult, statusFilter]);

  const selectedCount = useMemo(() => {
    if (!previewResult) return 0;
    return previewResult.rows.filter((r) => selectedRowIds[r.rowId]).length;
  }, [previewResult, selectedRowIds]);

  const executeBatchCommit = async (rowsToCommit: ArchiveImportPreviewRow[]) => {
    if (!previewResult || rowsToCommit.length === 0 || committing) return;
    setCommitting(true);
    setCommitResults([]);
    setCommitSummary(null);
    setCommitProgress({ done: 0, total: rowsToCommit.length });

    const batchId = lastBatchId || `arch-${Date.now().toString(36)}`;
    const CHUNK_SIZE = 25;
    const allResults: ArchiveImportCommitRowResult[] = [];
    const allCreatedDocs: DocumentRecord[] = [];
    let latestBook: RegistryBook | undefined;

    try {
      for (let i = 0; i < rowsToCommit.length; i += CHUNK_SIZE) {
        const chunk = rowsToCommit.slice(i, i + CHUNK_SIZE);
        const editableByRowId = new Map(
          editableRows.map((er, idx) => [`row-${er.rowNumber}-${idx}`, er])
        );

        const payloadRows = chunk.map((r) => {
          const edited = editableByRowId.get(r.rowId);
          return {
            rowId: r.rowId,
            rowNumber: r.rowNumber,
            docCategory: edited?.docCategory ?? r.docCategory,
            fullDocNumber: edited?.fullDocNumber ?? r.fullDocNumber,
            docName: edited?.docName ?? r.docName,
            issueDate: edited?.issueDate ?? r.issueDate,
            year: edited?.year ?? r.year,
            issuingAgency: edited?.issuingAgency ?? r.issuingAgency,
            signer: edited?.signer ?? r.signer,
            signerRole: edited?.signerRole ?? r.signerRole,
            typeCode: edited?.typeCode ?? r.typeCode,
            parentAgency: edited?.parentAgency ?? r.parentAgency,
            title: edited?.title ?? r.title,
            notes: edited?.notes ?? r.notes,
            matchedDriveFile: r.matchedDriveFile,
            allowDuplicate: Boolean(allowDuplicateMap[r.rowId]),
          };
        });

        const res = await api.commitArchiveImport({
          bookId: previewResult.bookId,
          batchId,
          rows: payloadRows,
        });
        allResults.push(...res.results);
        allCreatedDocs.push(...res.createdDocuments);
        if (res.bookUpdated) latestBook = res.bookUpdated;
        setCommitProgress({
          done: Math.min(rowsToCommit.length, i + chunk.length),
          total: rowsToCommit.length,
        });
      }

      setCommitResults(allResults);
      setCommitSummary({
        batchId,
        imported: allResults.filter((x) => x.status === 'imported').length,
        skipped: allResults.filter((x) => x.status === 'skipped_duplicate').length,
        errors: allResults.filter((x) => x.status === 'error').length,
      });
      if (allCreatedDocs.length > 0) {
        onDocumentsImported(allCreatedDocs, latestBook);
      }
    } catch (err: any) {
      setPreviewError(err?.message || 'Lỗi khi ghi lô hồ sơ nhập kho.');
    } finally {
      setCommitting(false);
    }
  };

  const handleCommitSelected = () => {
    if (!previewResult) return;
    const rowsToCommit = previewResult.rows.filter((r) => selectedRowIds[r.rowId]);
    executeBatchCommit(rowsToCommit);
  };

  const handleRetryFailedRows = () => {
    if (!previewResult) return;
    const failedRowIds = new Set(
      commitResults.filter((r) => r.status === 'error').map((r) => r.rowId)
    );
    const failedRows = previewResult.rows.filter((r) => failedRowIds.has(r.rowId));
    executeBatchCommit(failedRows);
  };

  if (importableBooks.length === 0) {
    return (
      <div className="max-w-4xl mx-auto py-8 px-4">
        <div className="bg-amber-50 border border-amber-200 rounded-xl p-6 text-center space-y-2">
          <AlertTriangle className="w-8 h-8 text-amber-600 mx-auto" />
          <h2 className="text-base font-bold text-amber-950">
            Chưa có quyền Nhập kho văn bản cũ
          </h2>
          <p className="text-xs text-amber-900">
            Liên hệ Quản trị viên để được cấp quyền &ldquo;Nhập kho văn bản cũ&rdquo; trên sổ văn bản tương ứng.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto py-5 sm:py-8 px-3 sm:px-6 lg:px-8 space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900">
            Nhập Kho Văn Bản Cũ (Excel &amp; Google Drive)
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 mt-1">
            Nhập danh mục hồ sơ lịch sử từ Excel, ghép nối trực tiếp tệp đã lưu trên Google Drive (không nhân bản tệp) và bảo toàn số thứ tự của từng dãy.
          </p>
        </div>
        <button
          type="button"
          onClick={handleDownloadTemplate}
          disabled={downloadingTpl}
          className="inline-flex items-center justify-center gap-2 min-h-[42px] px-4 py-2 bg-emerald-700 hover:bg-emerald-800 disabled:opacity-60 text-white text-xs font-semibold rounded-lg shadow-2xs shrink-0"
        >
          {downloadingTpl ? (
            <Loader2 className="w-4 h-4 animate-spin" />
          ) : (
            <Download className="w-4 h-4" />
          )}
          <span>Tải mẫu Excel (.xlsx)</span>
        </button>
      </div>

      {/* Step 1: Target Book, Folder, Upload Excel & Column Mapping */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5 shadow-2xs space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 block">
              1. Chọn sổ văn bản nhận hồ sơ *
            </label>
            <select
              value={targetBookId}
              onChange={(e) => {
                setTargetBookId(e.target.value);
                setPreviewResult(null);
              }}
              className="w-full min-h-[42px] px-3 bg-slate-50 border border-slate-300 rounded-lg text-xs sm:text-sm font-semibold text-slate-900"
            >
              {importableBooks.map((b) => (
                <option key={b.id} value={b.id}>
                  {b.name} ({b.unitCode})
                </option>
              ))}
            </select>
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 block">
              2. Thư mục Google Drive đối chiếu tệp (tùy chọn)
            </label>
            <div className="relative">
              <FolderOpen className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                type="text"
                value={folderUrlOrId}
                onChange={(e) => setFolderUrlOrId(e.target.value)}
                placeholder={
                  activeTargetBook?.googleDriveConfig?.folderId
                    ? 'Để trống để dùng thư mục Drive mặc định của sổ'
                    : 'Dán URL hoặc ID thư mục Google Drive...'
                }
                className="w-full min-h-[42px] pl-9 pr-3 bg-slate-50 border border-slate-300 rounded-lg text-xs text-slate-900"
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700 block">
              3. Tải lên tệp Excel (.xlsx) *
            </label>
            <label className="flex items-center justify-center gap-2 min-h-[42px] px-4 bg-sky-50 hover:bg-sky-100 border border-dashed border-sky-400 rounded-lg text-xs font-semibold text-sky-900 cursor-pointer transition-colors">
              {parsingExcel ? (
                <Loader2 className="w-4 h-4 animate-spin text-sky-700" />
              ) : (
                <Upload className="w-4 h-4 text-sky-700" />
              )}
              <span>
                {parsedWorkbook ? `Đã nạp: ${parsedWorkbook.fileName}` : 'Chọn tệp Excel .xlsx...'}
              </span>
              <input
                type="file"
                accept=".xlsx"
                onChange={handleExcelFileSelect}
                className="sr-only"
              />
            </label>
          </div>
        </div>

        {excelError && (
          <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
            <span>{excelError}</span>
          </div>
        )}

        {/* Sheet Selector & Column Mapping */}
        {parsedWorkbook && selectedSheet && (
          <div className="pt-4 border-t border-slate-100 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2 text-xs">
                <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                <span className="font-bold text-slate-800">Sheet dữ liệu:</span>
                <select
                  value={selectedSheet.name}
                  onChange={(e) => handleChangeSheet(e.target.value)}
                  className="min-h-[36px] px-2.5 bg-white border border-slate-300 rounded-md text-xs font-semibold text-slate-900"
                >
                  {parsedWorkbook.sheets.map((s) => (
                    <option key={s.name} value={s.name}>
                      {s.name} ({s.rowCount} dòng)
                    </option>
                  ))}
                </select>
              </div>

              <button
                type="button"
                disabled={previewLoading || !mapping.fullDocNumber || !mapping.docName}
                onClick={() => runPreview(buildRowsFromSheet())}
                className="inline-flex items-center gap-2 min-h-[40px] px-4 py-2 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg"
              >
                {previewLoading ? (
                  <Loader2 className="w-4 h-4 animate-spin" />
                ) : (
                  <RefreshCw className="w-4 h-4" />
                )}
                <span>Kiểm tra trùng &amp; Ghép tệp Drive</span>
              </button>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2.5">
              {MAPPING_FIELDS.map((field) => (
                <div key={field.key} className="space-y-1">
                  <label className="text-[11px] font-semibold text-slate-600 block truncate">
                    {field.label}
                    {field.required ? ' *' : ''}
                  </label>
                  <select
                    value={mapping[field.key] || ''}
                    onChange={(e) =>
                      setMapping((prev) => ({
                        ...prev,
                        [field.key]: e.target.value || undefined,
                      }))
                    }
                    className="w-full min-h-[36px] px-2 bg-slate-50 border border-slate-200 rounded-md text-xs text-slate-900"
                  >
                    <option value="">-- Không ghép --</option>
                    {selectedSheet.headers.map((h) => (
                      <option key={h} value={h}>
                        {h}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      {previewError && (
        <div className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800 flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
          <span>{previewError}</span>
        </div>
      )}

      {/* Step 2: Interactive Preview Table */}
      {previewResult && (
        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5 shadow-2xs space-y-4">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
            <div className="space-y-1">
              <h2 className="text-base font-bold text-slate-900">
                Xem trước &amp; xử lý từng dòng ({previewResult.summary.total} dòng)
              </h2>
              <p className="text-xs text-slate-600 tabular-nums">
                Sổ đích: <strong>{previewResult.bookName}</strong> · Quét thấy{' '}
                <strong>{previewResult.driveFileCount}</strong> tệp trong thư mục Drive
                {previewResult.driveScanError ? ` (${previewResult.driveScanError})` : ''}
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                disabled={previewLoading}
                onClick={() => runPreview(editableRows, true)}
                className="inline-flex items-center gap-1.5 min-h-[38px] px-3 py-1.5 bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 text-xs font-semibold rounded-lg"
              >
                <RotateCw className={`w-3.5 h-3.5 ${previewLoading ? 'animate-spin' : ''}`} />
                <span>Kiểm tra lại sau chỉnh sửa</span>
              </button>

              <button
                type="button"
                disabled={committing || selectedCount === 0}
                onClick={handleCommitSelected}
                className="inline-flex items-center gap-2 min-h-[38px] px-4 py-1.5 bg-sky-700 hover:bg-sky-800 disabled:opacity-50 text-white text-xs font-semibold rounded-lg"
              >
                {committing ? (
                  <Loader2 className="w-4 h-4 animate-spin" />
                ) : (
                  <CheckCircle2 className="w-4 h-4" />
                )}
                <span>Xác nhận nhập {selectedCount} dòng đã chọn</span>
              </button>
            </div>
          </div>

          {/* Filter Tabs */}
          <div className="flex flex-wrap gap-1.5 border-b border-slate-200 pb-3">
            {[
              { key: 'all' as const, label: `Tất cả (${previewResult.summary.total})` },
              { key: 'ready' as const, label: `Sẵn sàng (${previewResult.summary.ready})` },
              {
                key: 'possible_duplicate' as const,
                label: `Có thể trùng (${previewResult.summary.possibleDuplicate})`,
              },
              {
                key: 'missing_info' as const,
                label: `Thiếu thông tin (${previewResult.summary.missingInfo})`,
              },
              {
                key: 'unmatched_any' as const,
                label: `Chưa ghép / trùng tên tệp (${
                  previewResult.summary.unmatchedFile + previewResult.summary.ambiguousFile
                })`,
              },
            ].map((tab) => (
              <button
                key={tab.key}
                type="button"
                onClick={() => setStatusFilter(tab.key)}
                className={`min-h-[34px] px-3 py-1 rounded-md border text-xs font-semibold tabular-nums transition-colors ${
                  statusFilter === tab.key
                    ? 'bg-sky-800 text-white border-sky-800'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* Batch Commit Progress & Summary */}
          {commitProgress && (
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-xs space-y-1.5 tabular-nums">
              <div className="flex items-center justify-between font-semibold text-slate-800">
                <span>
                  {committing ? 'Đang nhập kho theo lô...' : 'Hoàn tất xử lý lô nhập kho'}
                </span>
                <span>
                  {commitProgress.done}/{commitProgress.total} dòng
                </span>
              </div>
              {commitSummary && (
                <div className="flex flex-wrap items-center justify-between gap-2 pt-1 text-xs">
                  <div className="flex flex-wrap items-center gap-3">
                    <span className="text-emerald-800 font-semibold">
                      Thành công: {commitSummary.imported}
                    </span>
                    <span className="text-slate-600">
                      Bỏ qua (trùng/đã nhập): {commitSummary.skipped}
                    </span>
                    <span className={commitSummary.errors > 0 ? 'text-rose-700 font-bold' : 'text-slate-600'}>
                      Lỗi: {commitSummary.errors}
                    </span>
                  </div>
                  {commitSummary.errors > 0 && (
                    <button
                      type="button"
                      disabled={committing}
                      onClick={handleRetryFailedRows}
                      className="inline-flex items-center gap-1.5 px-3 py-1 bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-md"
                    >
                      <RotateCw className="w-3.5 h-3.5" />
                      <span>Thử lại các dòng lỗi</span>
                    </button>
                  )}
                </div>
              )}
            </div>
          )}

          {/* Rows Table */}
          <div className="overflow-x-auto border border-slate-200 rounded-lg">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                  <th className="p-2.5 w-10 text-center">
                    <input
                      type="checkbox"
                      aria-label="Chọn tất cả dòng đang lọc"
                      checked={
                        filteredPreviewRows.length > 0 &&
                        filteredPreviewRows.every((r) => selectedRowIds[r.rowId])
                      }
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setSelectedRowIds((prev) => {
                          const next = { ...prev };
                          for (const r of filteredPreviewRows) next[r.rowId] = checked;
                          return next;
                        });
                      }}
                    />
                  </th>
                  <th className="p-2.5 w-14">Dòng</th>
                  <th className="p-2.5 min-w-[115px]">Nhóm &amp; Số/Ký hiệu</th>
                  <th className="p-2.5 min-w-[210px]">Tên văn bản &amp; Ngày/Năm</th>
                  <th className="p-2.5 min-w-[165px]">Cơ quan &amp; Người ký</th>
                  <th className="p-2.5 min-w-[230px]">Ghép tệp Google Drive</th>
                  <th className="p-2.5 min-w-[190px]">Trạng thái &amp; Cảnh báo</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredPreviewRows.map((row) => {
                  const rowIdx = previewResult.rows.findIndex((r) => r.rowId === row.rowId);
                  const edited = rowIdx !== -1 ? editableRows[rowIdx] : row;
                  const statusLabel =
                    row.status === 'ready'
                      ? { text: 'Sẵn sàng', cls: 'text-emerald-800 bg-emerald-50 border-emerald-200' }
                      : row.status === 'possible_duplicate'
                      ? { text: 'Có thể trùng', cls: 'text-amber-900 bg-amber-50 border-amber-300' }
                      : row.status === 'missing_info'
                      ? { text: 'Thiếu thông tin', cls: 'text-rose-800 bg-rose-50 border-rose-200' }
                      : row.status === 'ambiguous_file'
                      ? { text: 'Trùng nhiều tệp Drive', cls: 'text-amber-900 bg-amber-50 border-amber-300' }
                      : { text: 'Chưa ghép được tệp', cls: 'text-slate-700 bg-slate-100 border-slate-300' };

                  const commitRes = commitResults.find((cr) => cr.rowId === row.rowId);

                  return (
                    <tr key={row.rowId} className="align-top hover:bg-slate-50/50">
                      <td className="p-2.5 text-center">
                        <input
                          type="checkbox"
                          checked={Boolean(selectedRowIds[row.rowId])}
                          onChange={(e) =>
                            setSelectedRowIds((prev) => ({
                              ...prev,
                              [row.rowId]: e.target.checked,
                            }))
                          }
                        />
                      </td>
                      <td className="p-2.5 font-mono text-slate-500 tabular-nums">#{row.rowNumber}</td>

                      {/* Category + Số/Ký hiệu */}
                      <td className="p-2.5 space-y-1.5">
                        <select
                          value={edited?.docCategory || row.docCategory}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, {
                              docCategory: e.target.value as 'outgoing' | 'incoming',
                            })
                          }
                          className="w-full px-2 py-1 border border-slate-200 rounded bg-white text-xs font-semibold"
                        >
                          <option value="outgoing">Văn bản đi</option>
                          <option value="incoming">Văn bản đến</option>
                        </select>
                        <input
                          type="text"
                          value={edited?.fullDocNumber ?? row.fullDocNumber}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, { fullDocNumber: e.target.value })
                          }
                          placeholder="Số/Ký hiệu..."
                          className="w-full px-2 py-1 border border-slate-300 rounded font-mono font-bold text-xs text-slate-900"
                        />
                        {row.seriesUpdateInfo && (
                          <p className="text-[11px] text-sky-800">{row.seriesUpdateInfo}</p>
                        )}
                      </td>

                      {/* Tên văn bản + Ngày + Năm */}
                      <td className="p-2.5 space-y-1.5">
                        <input
                          type="text"
                          value={edited?.docName ?? row.docName}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, { docName: e.target.value })
                          }
                          placeholder="Tên văn bản..."
                          className="w-full px-2 py-1 border border-slate-300 rounded text-xs font-semibold text-slate-900"
                        />
                        <div className="grid grid-cols-2 gap-1.5">
                          <input
                            type="text"
                            value={edited?.issueDate ?? row.issueDate}
                            onChange={(e) =>
                              rowIdx !== -1 &&
                              updateEditableRowField(rowIdx, { issueDate: e.target.value })
                            }
                            placeholder="dd/mm/yyyy"
                            className="px-2 py-1 border border-slate-200 rounded font-mono text-[11px] tabular-nums"
                          />
                          <input
                            type="text"
                            value={edited?.year ?? row.year}
                            onChange={(e) =>
                              rowIdx !== -1 &&
                              updateEditableRowField(rowIdx, { year: e.target.value })
                            }
                            placeholder="Năm"
                            className="px-2 py-1 border border-slate-200 rounded font-mono text-[11px] tabular-nums"
                          />
                        </div>
                      </td>

                      {/* Cơ quan ban hành + Người ký */}
                      <td className="p-2.5 space-y-1.5">
                        <input
                          type="text"
                          value={edited?.issuingAgency ?? row.issuingAgency}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, { issuingAgency: e.target.value })
                          }
                          placeholder="Cơ quan ban hành"
                          className="w-full px-2 py-1 border border-slate-200 rounded text-[11px]"
                        />
                        <input
                          type="text"
                          value={edited?.signer ?? row.signer}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, { signer: e.target.value })
                          }
                          placeholder="Người ký"
                          className="w-full px-2 py-1 border border-slate-200 rounded text-[11px]"
                        />
                      </td>

                      {/* Drive Matching */}
                      <td className="p-2.5 space-y-1.5">
                        <input
                          type="text"
                          value={edited?.driveLinkOrFileName ?? row.driveLinkOrFileName}
                          onChange={(e) =>
                            rowIdx !== -1 &&
                            updateEditableRowField(rowIdx, {
                              driveLinkOrFileName: e.target.value,
                              driveHyperlink: undefined,
                              selectedDriveFileId: undefined,
                            })
                          }
                          placeholder="Liên kết, ID hoặc tên tệp Drive..."
                          className="w-full px-2 py-1 border border-slate-200 rounded text-[11px]"
                        />

                        {row.matchedDriveFile && (
                          <div className="flex items-center justify-between gap-1.5 bg-emerald-50 border border-emerald-200 rounded px-2 py-1 text-[11px] text-emerald-900">
                            <span className="truncate font-medium">
                              {row.matchedDriveFile.fileName}
                            </span>
                            <a
                              href={row.matchedDriveFile.webViewLink}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="text-emerald-700 hover:text-emerald-950 shrink-0"
                              title="Mở tệp trên Drive"
                            >
                              <ExternalLink className="w-3.5 h-3.5" />
                            </a>
                          </div>
                        )}

                        {row.driveCandidates.length > 1 && (
                          <select
                            value={edited?.selectedDriveFileId || ''}
                            onChange={(e) =>
                              rowIdx !== -1 &&
                              updateEditableRowField(rowIdx, {
                                selectedDriveFileId: e.target.value || undefined,
                              })
                            }
                            className="w-full px-2 py-1 bg-amber-50 border border-amber-300 rounded text-[11px] font-semibold text-amber-950"
                          >
                            <option value="">
                              -- Chọn 1 trong {row.driveCandidates.length} tệp trùng tên --
                            </option>
                            {row.driveCandidates.map((cand) => (
                              <option key={cand.fileId} value={cand.fileId}>
                                {cand.fileName}
                                {cand.folderPath ? ` (${cand.folderPath})` : ''}
                              </option>
                            ))}
                          </select>
                        )}
                      </td>

                      {/* Status & Warnings */}
                      <td className="p-2.5 space-y-1.5">
                        <span
                          className={`inline-block px-2 py-0.5 rounded border text-[11px] font-bold ${statusLabel.cls}`}
                        >
                          {statusLabel.text}
                        </span>

                        {row.errors.map((er, i) => (
                          <p key={i} className="text-[11px] text-rose-700 font-medium">
                            {er}
                          </p>
                        ))}
                        {row.warnings.map((w, i) => (
                          <p key={i} className="text-[11px] text-amber-800">
                            {w}
                          </p>
                        ))}

                        {row.duplicateMatches.length > 0 && (
                          <div className="space-y-1 pt-0.5">
                            {row.duplicateMatches.map((dup) => (
                              <div
                                key={dup.docId}
                                className="flex items-center justify-between gap-1 bg-amber-50/80 border border-amber-200 rounded px-2 py-1 text-[11px]"
                              >
                                <span className="truncate text-amber-950">
                                  Trùng: <strong>{dup.fullDocNumber}</strong>
                                </span>
                                <button
                                  type="button"
                                  onClick={() => onOpenExistingDocument(dup.docId)}
                                  className="inline-flex items-center gap-0.5 text-sky-700 hover:underline font-semibold shrink-0"
                                >
                                  <Eye className="w-3 h-3" />
                                  <span>Xem</span>
                                </button>
                              </div>
                            ))}
                            <label className="inline-flex items-center gap-1.5 text-[11px] text-slate-700 cursor-pointer">
                              <input
                                type="checkbox"
                                checked={Boolean(allowDuplicateMap[row.rowId])}
                                onChange={(e) =>
                                  setAllowDuplicateMap((prev) => ({
                                    ...prev,
                                    [row.rowId]: e.target.checked,
                                  }))
                                }
                              />
                              <span>Vẫn nhập nếu cố ý giữ cả hai</span>
                            </label>
                          </div>
                        )}

                        {commitRes && (
                          <p
                            className={`text-[11px] font-semibold pt-1 ${
                              commitRes.status === 'imported'
                                ? 'text-emerald-700'
                                : commitRes.status === 'error'
                                ? 'text-rose-700'
                                : 'text-slate-600'
                            }`}
                          >
                            {commitRes.message}
                          </p>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
