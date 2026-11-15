import React, { useEffect, useMemo, useState } from 'react';
import {
  Users,
  Shield,
  Cloud,
  Check,
  AlertTriangle,
  Plus,
  Trash2,
  Save,
  CheckCircle2,
  RotateCw,
  Copy,
  Clock,
  UserCheck,
  UserX,
  Lock,
  ShieldCheck,
  HardDrive,
  FolderCheck,
  Upload,
  BookOpen,
  Sliders,
} from 'lucide-react';
import { api } from '../api';
import type {
  BookPermission,
  DocumentTypeConfig,
  GoogleDriveAuthMode,
  RegistryBook,
  SignerConfig,
  SystemAuditLog,
  SystemSettings,
  User,
} from '../types';
import { getTypeColorStyle, DOCUMENT_COLOR_PALETTE } from '../utils/documentColors';
import { DEFAULT_MAX_WORD_MB, extractGoogleDriveFolderId } from '../utils/fileNaming';
import {
  ALL_BOOK_PERMISSIONS,
  BOOK_PERMISSION_META,
  DEFAULT_BOOK_ID,
  getBookEnabledTypes,
  hasBookPermission,
  isSystemAdmin,
  normalizeBookCode,
  normalizeBookPermissionList,
} from '../utils/permissions';

interface SettingsViewProps {
  currentUser: User;
  books: RegistryBook[];
  activeBookId: string;
  onSelectBookId: (bookId: string) => void;
  settings: SystemSettings;
  documentTypes: DocumentTypeConfig[];
  signers: SignerConfig[];
  onSettingsUpdated: (newSettings: SystemSettings) => void;
  onBooksUpdated?: () => void;
  onRefreshData: () => void;
}

const GOOGLE_APPS_SCRIPT_TEMPLATE = `function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (data.bookCode || "DOAN_TRUONG").toString().trim();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    var HEADERS = [
      "STT", "Năm", "Ngày ban hành", "Ngày ký", "Nhóm VB", "Loại văn bản",
      "Số thứ tự", "Số/Ký hiệu", "Tên văn bản", "Trích yếu/Nội dung",
      "Người ký", "Chức vụ", "Người thao tác", "Trạng thái hồ sơ",
      "Trạng thái tệp Drive", "Tên tệp Drive", "ID tệp Drive", "Liên kết Drive",
      "Ghi chú", "Mã hệ thống", "Mã sổ", "Thời điểm cập nhật"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground("#005BAA");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
    }

    var docNameVal = data.docName || data.title || "";
    var titleVal = data.title || "";
    var nowStr = new Date().toLocaleString("vi-VN");

    // Thao tác kiểm tra kết nối theo đúng tab của sổ
    if (data.action === "test") {
      sheet.appendRow([
        0, data.year, data.issueDate, data.signDate || "", data.docCategory || "Văn bản đi",
        data.typeName, 0, data.fullDocNumber, docNameVal, titleVal,
        data.signer, data.signerRole, data.issuedBy, data.status,
        data.driveStatus || "Không đính kèm", data.driveFileName || "", data.driveFileId || "",
        data.driveFileUrl || "", data.notes || "", data.id, sheetName, nowStr
      ]);
      return ContentService.createTextOutput(JSON.stringify({result: "success", action: "test", tab: sheetName}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var docId = data.id;
    var rows = sheet.getDataRange().getValues();
    var headerRow = rows[0] || HEADERS;
    var idColIdx = headerRow.indexOf("Mã hệ thống");
    if (idColIdx === -1) idColIdx = 19;

    var foundRowIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][idColIdx] === docId) {
        foundRowIndex = i + 1;
        break;
      }
    }

    var rowValues = [
      data.stt,
      data.year,
      data.issueDate,
      data.signDate || "",
      data.docCategory || "Văn bản đi",
      data.typeName,
      data.sequenceNumber,
      data.fullDocNumber,
      docNameVal,
      titleVal,
      data.signer,
      data.signerRole,
      data.issuedBy,
      data.status,
      data.driveStatus || "Không đính kèm",
      data.driveFileName || "",
      data.driveFileId || "",
      data.driveFileUrl || "",
      data.notes || "",
      data.id,
      sheetName,
      nowStr
    ];

    if (foundRowIndex > 0) {
      sheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return ContentService.createTextOutput(JSON.stringify({result: "success", tab: sheetName}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result: "error", error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

const GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE = `function getOrCreateSubFolder(parent, name) {
  var trimmed = (name || "").toString().trim();
  if (!trimmed) return parent;
  var folders = parent.getFoldersByName(trimmed);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(trimmed);
}

function resolveHierarchyFolder(rootFolder, data) {
  var target = rootFolder;
  var bookName = (data.bookName || "").toString().trim();
  if (bookName && rootFolder.getName().toLowerCase() !== bookName.toLowerCase()) {
    target = getOrCreateSubFolder(target, bookName);
  }
  if (data.year) {
    target = getOrCreateSubFolder(target, String(data.year));
  }
  if (data.categoryFolder) {
    target = getOrCreateSubFolder(target, String(data.categoryFolder));
  }
  return target;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var folderId = (data.folderId || "").trim();
    if (!folderId) {
      return jsonOut({ result: "error", error: "Thiếu folderId thư mục Google Drive." });
    }

    var rootFolder = DriveApp.getFolderById(folderId);
    var folderName = rootFolder.getName();

    // 1. Kiểm tra kết nối & quyền ghi thực tế vào thư mục
    if (data.action === "test") {
      var probe = rootFolder.createFile(".nct_check_write_" + Date.now() + ".txt", "OK");
      probe.setTrashed(true);
      var accountEmail = "";
      try { accountEmail = Session.getEffectiveUser().getEmail(); } catch (err) {}
      return jsonOut({
        result: "success",
        folderId: folderId,
        folderName: folderName,
        accountEmail: accountEmail
      });
    }

    var targetFolder = resolveHierarchyFolder(rootFolder, data);
    var docId = (data.docId || "").trim();
    var marker = docId ? "[NCT_DOC_ID:" + docId + "]" : "";

    function findInFolder(dir) {
      if (!marker) return null;
      var files = dir.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        if (!f.isTrashed() && (f.getDescription() || "").indexOf(marker) !== -1) {
          return f;
        }
      }
      return null;
    }

    function findExisting() {
      return findInFolder(targetFolder) || (targetFolder.getId() !== rootFolder.getId() ? findInFolder(rootFolder) : null);
    }

    // 2. Đối soát trước khi tải lên
    if (data.action === "find") {
      var existing = findExisting();
      if (existing) {
        if (data.expectedFileName && existing.getName() !== data.expectedFileName) {
          existing.setName(data.expectedFileName);
        }
        return jsonOut({
          result: "success",
          found: true,
          fileId: existing.getId(),
          fileName: existing.getName(),
          webViewLink: existing.getUrl(),
          targetFolderId: targetFolder.getId()
        });
      }
      return jsonOut({ result: "success", found: false });
    }

    // 3. Tải tệp vào đúng cây thư mục Tên sổ -> Năm -> Văn bản đi / Văn bản đến
    if (data.action === "upload") {
      var dup = findExisting();
      if (dup && !data.replaceMode) {
        if (data.fileName && dup.getName() !== data.fileName) {
          dup.setName(data.fileName);
        }
        return jsonOut({
          result: "success",
          deduplicated: true,
          fileId: dup.getId(),
          fileName: dup.getName(),
          webViewLink: dup.getUrl(),
          targetFolderId: targetFolder.getId()
        });
      }

      if (data.replaceMode) {
        if (data.previousFileId) {
          try { DriveApp.getFileById(data.previousFileId).setTrashed(true); } catch (ePrev) {}
        }
        if (dup) {
          try { dup.setTrashed(true); } catch (eDup) {}
        }
      }

      var bytes = Utilities.base64Decode(data.fileBase64);
      var mimeType = data.mimeType || "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      var blob = Utilities.newBlob(bytes, mimeType, data.fileName);
      var created = targetFolder.createFile(blob);
      created.setDescription(
        "Số văn bản: " + (data.fullDocNumber || "") +
        " | Tên văn bản: " + (data.docName || "") +
        (marker ? " | " + marker : "")
      );

      return jsonOut({
        result: "success",
        fileId: created.getId(),
        fileName: created.getName(),
        webViewLink: created.getUrl(),
        targetFolderId: targetFolder.getId()
      });
    }

    return jsonOut({ result: "error", error: "Thao tác không hợp lệ: " + data.action });
  } catch (err) {
    return jsonOut({ result: "error", error: err.toString() });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

const PERMISSION_PRESETS: Array<{ label: string; perms: BookPermission[] }> = [
  { label: 'Chỉ xem', perms: ['view'] },
  { label: 'Chỉ xin số', perms: ['view', 'request_number'] },
  { label: 'Lấy số & Xem sổ', perms: ['view', 'issue_number'] },
  {
    label: 'Cán bộ cấp số',
    perms: ['view', 'request_number', 'issue_number', 'import_incoming', 'edit', 'replace_file', 'export'],
  },
  { label: 'Toàn quyền sổ', perms: [...ALL_BOOK_PERMISSIONS] },
  { label: 'Bỏ hết', perms: [] },
];

export const SettingsView: React.FC<SettingsViewProps> = ({
  currentUser,
  books,
  activeBookId,
  onSelectBookId,
  settings,
  documentTypes,
  signers,
  onSettingsUpdated,
  onBooksUpdated,
  onRefreshData,
}) => {
  const isAdmin = isSystemAdmin(currentUser);
  const currentYear = new Date().getFullYear();

  const managedBooks = useMemo(() => {
    return books.filter((b) => isAdmin || hasBookPermission(currentUser, b.id, 'manage_book'));
  }, [books, isAdmin, currentUser]);

  const [selectedBookId, setSelectedBookId] = useState<string>(() => {
    if (managedBooks.some((b) => b.id === activeBookId)) return activeBookId;
    return managedBooks[0]?.id || books[0]?.id || DEFAULT_BOOK_ID;
  });

  useEffect(() => {
    if (!managedBooks.some((b) => b.id === selectedBookId) && managedBooks.length > 0) {
      setSelectedBookId(
        managedBooks.some((b) => b.id === activeBookId) ? activeBookId : managedBooks[0].id
      );
    }
  }, [managedBooks, selectedBookId, activeBookId]);

  const selectedBook = useMemo(() => {
    return (
      books.find((b) => b.id === selectedBookId) ||
      managedBooks[0] ||
      books[0] || {
        id: DEFAULT_BOOK_ID,
        code: 'DOAN_TRUONG',
        name: 'Đoàn trường',
        unitName: settings.unitName,
        unitCode: settings.unitCode,
        isShared: true,
        numberMode: settings.numberMode,
        numberFormat: settings.numberFormat,
        customStartSequences: settings.customStartSequences || {},
        signers,
        isActive: true,
        createdAt: '',
      }
    );
  }, [books, managedBooks, selectedBookId, settings, signers]);

  const canManageSelectedBook = isAdmin || hasBookPermission(currentUser, selectedBook.id, 'manage_book');

  // Sub-tab inside Settings
  const [activeSubTab, setActiveSubTab] = useState<
    'books' | 'general' | 'numbering' | 'docTypes' | 'signers' | 'googledrive' | 'googlesheets' | 'msgraph' | 'users'
  >('books');

  // Book Management State (Multi-book)
  const [newBookName, setNewBookName] = useState('');
  const [newBookCode, setNewBookCode] = useState('');
  const [newBookCodeEdited, setNewBookCodeEdited] = useState(false);
  const [newBookUnitName, setNewBookUnitName] = useState('');
  const [newBookUnitCode, setNewBookUnitCode] = useState('');
  const [newBookIsShared, setNewBookIsShared] = useState<boolean>(false);
  const [newBookMode, setNewBookMode] = useState<'separate' | 'shared'>('separate');
  const [newBookFormat, setNewBookFormat] = useState('{so}-{loai}/{don_vi}');
  const [newBookEnabledTypes, setNewBookEnabledTypes] = useState<string[]>(() =>
    documentTypes.filter((t) => t.isActive).map((t) => t.code)
  );
  const [bookSaving, setBookSaving] = useState(false);
  const [bookMessage, setBookMessage] = useState<{ success: boolean; text: string } | null>(null);

  // General Settings Form (per selected book)
  const [bookTitleInput, setBookTitleInput] = useState(selectedBook.name);
  const [unitName, setUnitName] = useState(selectedBook.unitName);
  const [unitCode, setUnitCode] = useState(selectedBook.unitCode);
  const [numberFormat, setNumberFormat] = useState(selectedBook.numberFormat);
  const [bookEnabledTypes, setBookEnabledTypes] = useState<string[]>(() =>
    getBookEnabledTypes(selectedBook, documentTypes).map((t) => t.code)
  );
  const [generalSaving, setGeneralSaving] = useState(false);
  const [generalSuccess, setGeneralSuccess] = useState(false);

  // Numbering Mode (per selected book)
  const [targetMode, setTargetMode] = useState<'separate' | 'shared'>(selectedBook.numberMode);
  const [modeConfirmText, setModeConfirmText] = useState('');
  const [modeSaving, setModeSaving] = useState(false);
  const [modeError, setModeError] = useState<string | null>(null);
  const [modeSuccess, setModeSuccess] = useState(false);

  // Custom Starting Sequence
  const [seqYear, setSeqYear] = useState<number>(currentYear);
  const [seqTypeCode, setSeqTypeCode] = useState<string>(
    selectedBook.numberMode === 'shared' ? 'SHARED' : documentTypes[0]?.code || 'KH'
  );
  const [startSequenceVal, setStartSequenceVal] = useState<number>(1);
  const [seqSaving, setSeqSaving] = useState(false);
  const [seqError, setSeqError] = useState<string | null>(null);
  const [seqSuccess, setSeqSuccess] = useState(false);

  // Doc Type Modal / Form
  const [newTypeName, setNewTypeName] = useState('');
  const [newTypeCode, setNewTypeCode] = useState('');
  const [newTypeColor, setNewTypeColor] = useState('blue');
  const [typeSaving, setTypeSaving] = useState(false);
  const [typeError, setTypeError] = useState<string | null>(null);

  // Signer Modal / Form (per selected book)
  const [newSignerName, setNewSignerName] = useState('');
  const [newSignerRole, setNewSignerRole] = useState('');
  const [signerSaving, setSignerSaving] = useState(false);
  const [signerError, setSignerError] = useState<string | null>(null);

  // Google Drive Word/PDF Storage state (strictly isolated per selected book)
  const effectiveBookDriveCfg = useMemo(() => {
    if (selectedBook.id !== DEFAULT_BOOK_ID) {
      return selectedBook.googleDriveConfig;
    }
    return selectedBook.googleDriveConfig || settings.googleDriveConfig;
  }, [selectedBook, settings.googleDriveConfig]);

  const [gdEnabled, setGdEnabled] = useState<boolean>(Boolean(effectiveBookDriveCfg?.enabled));
  const [gdAuthMode, setGdAuthMode] = useState<GoogleDriveAuthMode>(
    effectiveBookDriveCfg?.authMode || 'apps_script'
  );
  const [gdFolderUrl, setGdFolderUrl] = useState<string>(
    effectiveBookDriveCfg?.folderUrl || effectiveBookDriveCfg?.folderId || ''
  );
  const [gdAppsScriptUrl, setGdAppsScriptUrl] = useState<string>('');
  const [gdServiceAccountJson, setGdServiceAccountJson] = useState<string>('');
  const [gdSaving, setGdSaving] = useState(false);
  const [gdSaveMessage, setGdSaveMessage] = useState<{ success: boolean; message: string } | null>(null);
  const [gdTesting, setGdTesting] = useState(false);
  const [gdTestResult, setGdTestResult] = useState<{
    success: boolean;
    message: string;
    folderName?: string;
  } | null>(null);
  const [copiedDriveScript, setCopiedDriveScript] = useState(false);

  // Google Sheets state (Master + per-book optional)
  const [gsEnabled, setGsEnabled] = useState(settings.googleSheetsConfig?.enabled || false);
  const [gsWebhookUrl, setGsWebhookUrl] = useState(settings.googleSheetsConfig?.webhookUrl || '');
  const [gsTestBookId, setGsTestBookId] = useState<string>(selectedBook.id);
  const [gsSaving, setGsSaving] = useState(false);
  const [gsTesting, setGsTesting] = useState(false);
  const [gsTestResult, setGsTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [gsSyncingAll, setGsSyncingAll] = useState(false);
  const [gsSyncResult, setGsSyncResult] = useState<string | null>(null);
  const [copiedScript, setCopiedScript] = useState(false);
  const [bookGsEnabled, setBookGsEnabled] = useState<boolean>(Boolean(selectedBook.googleSheetsConfig?.enabled));
  const [bookGsWebhookUrl, setBookGsWebhookUrl] = useState<string>(selectedBook.googleSheetsConfig?.webhookUrl || '');
  const [bookGsSaving, setBookGsSaving] = useState(false);
  const [bookGsMessage, setBookGsMessage] = useState<string | null>(null);

  // Sync form states when switching selectedBook
  useEffect(() => {
    const enabledForBook = getBookEnabledTypes(selectedBook, documentTypes);
    setBookTitleInput(selectedBook.name);
    setUnitName(selectedBook.unitName);
    setUnitCode(selectedBook.unitCode);
    setNumberFormat(selectedBook.numberFormat);
    setBookEnabledTypes(enabledForBook.map((t) => t.code));
    setTargetMode(selectedBook.numberMode);
    setSeqTypeCode(selectedBook.numberMode === 'shared' ? 'SHARED' : enabledForBook[0]?.code || 'KH');
    setModeConfirmText('');
    setModeError(null);
    setModeSuccess(false);
    setGeneralSuccess(false);

    const cfg =
      selectedBook.id !== DEFAULT_BOOK_ID
        ? selectedBook.googleDriveConfig
        : selectedBook.googleDriveConfig || settings.googleDriveConfig;
    setGdEnabled(Boolean(cfg?.enabled));
    setGdAuthMode(cfg?.authMode || 'apps_script');
    setGdFolderUrl(cfg?.folderUrl || cfg?.folderId || '');
    setGdAppsScriptUrl('');
    setGdServiceAccountJson('');
    setGdSaveMessage(null);
    setGdTestResult(null);

    setBookGsEnabled(Boolean(selectedBook.googleSheetsConfig?.enabled));
    setBookGsWebhookUrl(selectedBook.googleSheetsConfig?.webhookUrl || '');
    setBookGsMessage(null);
    setGsTestBookId(selectedBook.id);
  }, [selectedBook, settings.googleDriveConfig, documentTypes]);

  // Microsoft Graph
  const [msEnabled, setMsEnabled] = useState(settings.msGraphConfig.enabled);
  const [msTenantId, setMsTenantId] = useState(settings.msGraphConfig.tenantId);
  const [msClientId, setMsClientId] = useState(settings.msGraphConfig.clientId);
  const [msClientSecret, setMsClientSecret] = useState(settings.msGraphConfig.clientSecret);
  const [msDriveId, setMsDriveId] = useState(settings.msGraphConfig.driveId);
  const [msFilePath, setMsFilePath] = useState(settings.msGraphConfig.filePath);
  const [msTableName, setMsTableName] = useState(settings.msGraphConfig.tableName);
  const [msTesting, setMsTesting] = useState(false);
  const [msTestResult, setMsTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [msSaving, setMsSaving] = useState(false);
  const [msSyncingAll, setMsSyncingAll] = useState(false);
  const [msSyncResult, setMsSyncResult] = useState<string | null>(null);

  // User Management & Per-Book Permissions
  const [usersList, setUsersList] = useState<User[]>([]);
  const [loadingUsers, setLoadingUsers] = useState(false);
  const [newUsername, setNewUsername] = useState('');
  const [newFullName, setNewFullName] = useState('');
  const [newUserPass, setNewUserPass] = useState('');
  const [newBookPermissions, setNewBookPermissions] = useState<Record<string, BookPermission[]>>({
    [DEFAULT_BOOK_ID]: ['view', 'request_number'],
  });
  const [userSaving, setUserSaving] = useState(false);
  const [userError, setUserError] = useState<string | null>(null);

  const [editingPermUserId, setEditingPermUserId] = useState<string | null>(null);
  const [editingBookPermissions, setEditingBookPermissions] = useState<Record<string, BookPermission[]>>({});
  const [savingPermUserId, setSavingPermUserId] = useState<string | null>(null);
  const [auditLogs, setAuditLogs] = useState<SystemAuditLog[]>([]);
  const [loadingAuditLogs, setLoadingAuditLogs] = useState(false);

  const togglePermissionInMap = (
    map: Record<string, BookPermission[]>,
    bookId: string,
    perm: BookPermission
  ): Record<string, BookPermission[]> => {
    const current = new Set(map[bookId] || []);
    if (perm === 'view' && current.has('issue_number') && current.has('view')) {
      // Quyền Lấy số luôn kèm quyền Chỉ xem trên cùng sổ
      return {
        ...map,
        [bookId]: normalizeBookPermissionList(Array.from(current)),
      };
    }
    if (current.has(perm)) current.delete(perm);
    else current.add(perm);
    return {
      ...map,
      [bookId]: normalizeBookPermissionList(Array.from(current)),
    };
  };

  // Fetch users & audit logs when opening users tab
  const handleOpenUsersTab = async () => {
    setActiveSubTab('users');
    if (!isAdmin) return;
    setLoadingUsers(true);
    setLoadingAuditLogs(true);
    try {
      const [usersRes, logsRes] = await Promise.all([
        api.getUsers(),
        api.getAuditLogs().catch(() => ({ logs: [] as SystemAuditLog[] })),
      ]);
      setUsersList(usersRes.users);
      setAuditLogs(logsRes.logs);
    } catch {
      // ignore
    } finally {
      setLoadingUsers(false);
      setLoadingAuditLogs(false);
    }
  };

  // Create new Registry Book (Admin only)
  const handleCreateBook = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    if (newBookEnabledTypes.length === 0) {
      setBookMessage({
        success: false,
        text: 'Vui lòng chọn ít nhất 1 loại văn bản được phép sử dụng trong sổ mới.',
      });
      return;
    }
    setBookSaving(true);
    setBookMessage(null);
    try {
      const resolvedCode = normalizeBookCode(newBookCode || newBookName);
      const res = await api.createBook({
        code: resolvedCode || undefined,
        name: newBookName.trim(),
        unitName: newBookUnitName.trim(),
        unitCode: newBookUnitCode.trim().toUpperCase(),
        isShared: newBookIsShared,
        numberMode: newBookMode,
        numberFormat: newBookFormat.trim() || '{so}-{loai}/{don_vi}',
        enabledTypeCodes: newBookEnabledTypes,
      });
      setNewBookName('');
      setNewBookCode('');
      setNewBookCodeEdited(false);
      setNewBookUnitName('');
      setNewBookUnitCode('');
      setNewBookIsShared(false);
      setNewBookMode('separate');
      setNewBookFormat('{so}-{loai}/{don_vi}');
      setNewBookEnabledTypes(documentTypes.filter((t) => t.isActive).map((t) => t.code));
      setSelectedBookId(res.book.id);
      onSelectBookId(res.book.id);
      setBookMessage({
        success: true,
        text: `Đã tạo sổ "${res.book.name}" (Mã sổ: ${res.book.code} · Ký hiệu: ${res.book.unitCode}). Hãy phân quyền ở tab Tài khoản & Phân quyền sổ và cấu hình Drive riêng cho sổ này nếu cần.`,
      });
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      setBookMessage({
        success: false,
        text: err.message || 'Lỗi khi tạo sổ văn bản mới',
      });
    } finally {
      setBookSaving(false);
    }
  };

  const handleToggleBookActive = async (book: RegistryBook) => {
    if (!isAdmin || book.id === DEFAULT_BOOK_ID) return;
    try {
      await api.updateBook(book.id, { isActive: !book.isActive });
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi cập nhật trạng thái sổ');
    }
  };

  // 1. Save General Settings for selected book
  const handleSaveGeneral = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook) return;
    if (bookEnabledTypes.length === 0) {
      alert('Vui lòng chọn ít nhất 1 loại văn bản được sử dụng trong sổ này.');
      return;
    }
    setGeneralSaving(true);
    setGeneralSuccess(false);
    try {
      await api.updateBook(selectedBook.id, {
        name: bookTitleInput.trim() || selectedBook.name,
        unitName: unitName.trim(),
        unitCode: unitCode.trim().toUpperCase(),
        numberFormat: numberFormat.trim(),
        enabledTypeCodes: bookEnabledTypes,
      });
      onBooksUpdated?.();
      onRefreshData();
      setGeneralSuccess(true);
      setTimeout(() => setGeneralSuccess(false), 3000);
    } catch (err: any) {
      alert(err.message || 'Lỗi khi lưu cài đặt sổ');
    } finally {
      setGeneralSaving(false);
    }
  };

  // 2. Change Number Mode for selected book
  const handleChangeMode = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook) return;
    setModeSaving(true);
    setModeError(null);
    setModeSuccess(false);
    try {
      const res = await api.updateNumberMode(targetMode, modeConfirmText.trim(), selectedBook.id);
      onSettingsUpdated(res.settings);
      onBooksUpdated?.();
      onRefreshData();
      setModeSuccess(true);
      setModeConfirmText('');
      setTimeout(() => setModeSuccess(false), 3000);
    } catch (err: any) {
      setModeError(err.message || 'Lỗi khi đổi chế độ dãy số');
    } finally {
      setModeSaving(false);
    }
  };

  // 3. Set Start Sequence for selected book
  const handleSetStartSequence = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook) return;
    setSeqSaving(true);
    setSeqError(null);
    setSeqSuccess(false);
    try {
      await api.setStartSequence(seqYear, seqTypeCode, startSequenceVal, selectedBook.id);
      setSeqSuccess(true);
      onBooksUpdated?.();
      onRefreshData();
      setTimeout(() => setSeqSuccess(false), 3000);
    } catch (err: any) {
      setSeqError(err.message || 'Lỗi khi đặt số bắt đầu');
    } finally {
      setSeqSaving(false);
    }
  };

  // 4. Create Document Type (Admin)
  const handleCreateDocType = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setTypeSaving(true);
    setTypeError(null);
    try {
      await api.createDocumentType({
        name: newTypeName.trim(),
        code: newTypeCode.trim().toUpperCase(),
        color: newTypeColor,
      });
      setNewTypeName('');
      setNewTypeCode('');
      setNewTypeColor('blue');
      onRefreshData();
    } catch (err: any) {
      setTypeError(err.message || 'Lỗi khi tạo loại văn bản');
    } finally {
      setTypeSaving(false);
    }
  };

  const handleUpdateDocTypeColor = async (id: string, color: string) => {
    if (!isAdmin) return;
    try {
      await api.updateDocumentType(id, { color });
      onRefreshData();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi cập nhật màu sắc loại văn bản');
    }
  };

  const handleDeleteDocType = async (id: string, code: string) => {
    if (!isAdmin || !confirm(`Xác nhận xóa loại văn bản "${code}"?`)) return;
    try {
      await api.deleteDocumentType(id);
      onRefreshData();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi xóa loại văn bản');
    }
  };

  // 5. Book Signers management (Admin or manage_book)
  const displayedBookSigners = useMemo(() => {
    if (selectedBook.signers && selectedBook.signers.length > 0) {
      return selectedBook.signers;
    }
    return signers;
  }, [selectedBook.signers, signers]);

  const handleCreateBookSigner = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook) return;
    setSignerSaving(true);
    setSignerError(null);
    try {
      const nextSigner: SignerConfig = {
        id: `sig-${Date.now()}`,
        name: newSignerName.trim(),
        role: newSignerRole.trim(),
      };
      const updatedList = [...displayedBookSigners, nextSigner];
      await api.updateBook(selectedBook.id, { signers: updatedList });
      if (isAdmin && selectedBook.id === DEFAULT_BOOK_ID) {
        await api.createSigner({ name: nextSigner.name, role: nextSigner.role }).catch(() => {});
      }
      setNewSignerName('');
      setNewSignerRole('');
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      setSignerError(err.message || 'Lỗi khi thêm người ký');
    } finally {
      setSignerSaving(false);
    }
  };

  const handleDeleteBookSigner = async (signerId: string) => {
    if (!canManageSelectedBook || !confirm('Xác nhận xóa người ký này khỏi sổ?')) return;
    try {
      const updatedList = displayedBookSigners.filter((s) => s.id !== signerId);
      await api.updateBook(selectedBook.id, { signers: updatedList });
      if (isAdmin && selectedBook.id === DEFAULT_BOOK_ID) {
        await api.deleteSigner(signerId).catch(() => {});
      }
      onBooksUpdated?.();
      onRefreshData();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi xóa người ký');
    }
  };

  // 6. Google Sheets Config & Sync Handlers
  const handleSaveGoogleSheets = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setGsSaving(true);
    try {
      const res = await api.saveGoogleSheetsConfig({
        enabled: gsEnabled,
        webhookUrl: gsWebhookUrl.trim(),
      });
      onSettingsUpdated({ ...settings, googleSheetsConfig: res.googleSheetsConfig });
      alert('Đã lưu cấu hình Google Trang tính thành công!');
    } catch (err: any) {
      alert(err.message || 'Lỗi khi lưu cấu hình Google Trang tính');
    } finally {
      setGsSaving(false);
    }
  };

  const handleTestGoogleSheets = async () => {
    setGsTesting(true);
    setGsTestResult(null);
    try {
      const res = await api.testGoogleSheetsConnection(gsWebhookUrl.trim(), gsTestBookId || selectedBook.id);
      setGsTestResult(res);
    } catch (err: any) {
      setGsTestResult({ success: false, message: err.message || 'Lỗi kiểm tra Google Sheets' });
    } finally {
      setGsTesting(false);
    }
  };

  const handleSyncAllGoogleSheets = async (targetBookId?: string) => {
    setGsSyncingAll(true);
    setGsSyncResult(null);
    try {
      const res = await api.syncAllGoogleSheets(targetBookId);
      setGsSyncResult(res.message);
      onRefreshData();
    } catch (err: any) {
      setGsSyncResult('Lỗi: ' + err.message);
    } finally {
      setGsSyncingAll(false);
    }
  };

  const handleSaveBookGoogleSheets = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setBookGsSaving(true);
    setBookGsMessage(null);
    try {
      await api.saveBookGoogleSheetsConfig(selectedBook.id, {
        enabled: bookGsEnabled,
        webhookUrl: bookGsWebhookUrl.trim(),
      });
      onBooksUpdated?.();
      onRefreshData();
      setBookGsMessage(`Đã lưu webhook Google Sheets riêng cho sổ "${selectedBook.name}".`);
    } catch (err: any) {
      setBookGsMessage(`Lỗi: ${err.message || 'Không thể lưu cấu hình Google Sheets của sổ.'}`);
    } finally {
      setBookGsSaving(false);
    }
  };

  // 6b. Google Drive Storage Handlers (per selected book)
  const handleSaveGoogleDrive = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canManageSelectedBook) return;
    setGdSaving(true);
    setGdSaveMessage(null);
    try {
      const res = await api.saveGoogleDriveConfig({
        bookId: selectedBook.id,
        enabled: gdEnabled,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim(),
        appsScriptUrl: gdAppsScriptUrl.trim() || undefined,
        serviceAccountJson: gdServiceAccountJson.trim() || undefined,
        scopeUnitCode: selectedBook.unitCode,
      });
      if (selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      onRefreshData();
      setGdAppsScriptUrl('');
      setGdServiceAccountJson('');
      setGdSaveMessage({
        success: true,
        message: `Đã lưu cấu hình thư mục Google Drive cho "${selectedBook.name}". Hãy bấm "Kiểm tra kết nối" để xác nhận quyền tải tệp.`,
      });
    } catch (err: any) {
      setGdSaveMessage({
        success: false,
        message: err.message || 'Lỗi khi lưu cấu hình Google Drive',
      });
    } finally {
      setGdSaving(false);
    }
  };

  const handleClearDriveCredential = async (type: 'apps_script' | 'service_account') => {
    if (!canManageSelectedBook) return;
    setGdSaving(true);
    setGdSaveMessage(null);
    try {
      const res = await api.saveGoogleDriveConfig({
        bookId: selectedBook.id,
        enabled: gdEnabled,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim(),
        clearAppsScriptUrl: type === 'apps_script',
        clearServiceAccountJson: type === 'service_account',
        scopeUnitCode: selectedBook.unitCode,
      });
      if (selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      if (type === 'apps_script') setGdAppsScriptUrl('');
      else setGdServiceAccountJson('');
      setGdSaveMessage({
        success: true,
        message: 'Đã xóa thông tin xác thực cũ.',
      });
    } catch (err: any) {
      setGdSaveMessage({
        success: false,
        message: err.message || 'Không thể xóa thông tin xác thực',
      });
    } finally {
      setGdSaving(false);
    }
  };

  const handleTestGoogleDrive = async () => {
    if (!canManageSelectedBook) return;
    setGdTesting(true);
    setGdTestResult(null);
    try {
      const res = await api.testGoogleDriveConnection({
        bookId: selectedBook.id,
        authMode: gdAuthMode,
        folderUrl: gdFolderUrl.trim() || undefined,
        appsScriptUrl: gdAppsScriptUrl.trim() || undefined,
        serviceAccountJson: gdServiceAccountJson.trim() || undefined,
      });
      if (res.googleDriveConfig && selectedBook.id === DEFAULT_BOOK_ID) {
        onSettingsUpdated({ ...settings, googleDriveConfig: res.googleDriveConfig });
      }
      onBooksUpdated?.();
      setGdTestResult({
        success: res.success,
        folderName: res.folderName,
        message: res.message,
      });
    } catch (err: any) {
      setGdTestResult({
        success: false,
        message: err.message || 'Kiểm tra kết nối thư mục Google Drive thất bại.',
      });
    } finally {
      setGdTesting(false);
    }
  };

  const handleServiceAccountFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      setGdServiceAccountJson(text);
      setGdSaveMessage(null);
    } catch {
      setGdSaveMessage({
        success: false,
        message: 'Không đọc được nội dung tệp JSON đã chọn.',
      });
    }
  };

  // 7. Save Microsoft Graph Config
  const handleSaveMsGraph = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    setMsSaving(true);
    try {
      const res = await api.saveSyncConfig({
        enabled: msEnabled,
        tenantId: msTenantId,
        clientId: msClientId,
        clientSecret: msClientSecret,
        driveId: msDriveId,
        filePath: msFilePath,
        tableName: msTableName,
      });
      onSettingsUpdated({ ...settings, msGraphConfig: res.msGraphConfig });
      alert('Đã lưu cấu hình Microsoft Graph!');
    } catch (err: any) {
      alert(err.message || 'Lỗi khi lưu cấu hình');
    } finally {
      setMsSaving(false);
    }
  };

  const handleTestMsConnection = async () => {
    setMsTesting(true);
    setMsTestResult(null);
    try {
      const res = await api.testSyncConnection();
      setMsTestResult(res);
    } catch (err: any) {
      setMsTestResult({ success: false, message: err.message || 'Lỗi kết nối' });
    } finally {
      setMsTesting(false);
    }
  };

  const handleSyncAll = async () => {
    setMsSyncingAll(true);
    setMsSyncResult(null);
    try {
      const res = await api.syncAllPending();
      setMsSyncResult(res.message);
      onRefreshData();
    } catch (err: any) {
      setMsSyncResult('Lỗi: ' + err.message);
    } finally {
      setMsSyncingAll(false);
    }
  };

  // 8. User Management Handlers (Admin Only)
  const handleCreateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    if (newUserPass.trim().length < 6) {
      setUserError('Mật khẩu khởi tạo phải có tối thiểu 6 ký tự');
      return;
    }
    setUserSaving(true);
    setUserError(null);
    try {
      await api.createUser({
        username: newUsername.trim(),
        fullName: newFullName.trim(),
        password: newUserPass.trim(),
        role: 'issuer',
        bookPermissions: newBookPermissions,
      });
      setNewUsername('');
      setNewFullName('');
      setNewUserPass('');
      setNewBookPermissions({ [DEFAULT_BOOK_ID]: ['view', 'request_number'] });
      handleOpenUsersTab();
    } catch (err: any) {
      setUserError(err.message || 'Lỗi khi tạo tài khoản');
    } finally {
      setUserSaving(false);
    }
  };

  const handleStartEditPermissions = (u: User) => {
    if (editingPermUserId === u.id) {
      setEditingPermUserId(null);
      return;
    }
    setEditingPermUserId(u.id);
    setEditingBookPermissions({ ...(u.bookPermissions || {}) });
  };

  const handleSaveUserPermissions = async (u: User) => {
    if (!isAdmin) return;
    setSavingPermUserId(u.id);
    try {
      await api.updateUser(u.id, { bookPermissions: editingBookPermissions });
      setEditingPermUserId(null);
      await handleOpenUsersTab();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi lưu phân quyền tài khoản');
    } finally {
      setSavingPermUserId(null);
    }
  };

  const handleToggleUserStatus = async (u: User) => {
    if (u.username === 'admin') {
      alert('Không thể khóa tài khoản Quản trị viên cố định');
      return;
    }
    try {
      await api.updateUser(u.id, { isActive: !u.isActive });
      handleOpenUsersTab();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi cập nhật trạng thái tài khoản');
    }
  };

  const handleDeleteUser = async (u: User) => {
    if (u.username === 'admin') {
      alert('Không thể xóa tài khoản Quản trị viên cố định');
      return;
    }
    if (!confirm(`Xác nhận xóa tài khoản "${u.username}" (${u.fullName}) khỏi hệ thống?`)) {
      return;
    }
    try {
      await api.deleteUser(u.id);
      handleOpenUsersTab();
    } catch (err: any) {
      alert(err.message || 'Lỗi khi xóa tài khoản');
    }
  };

  const handleResetUserPass = async (u: User) => {
    const newPass = prompt(`Nhập mật khẩu mới cho tài khoản "${u.username}" (tối thiểu 6 ký tự):`);
    if (!newPass) return;
    if (newPass.trim().length < 6) {
      alert('Mật khẩu phải có tối thiểu 6 ký tự');
      return;
    }
    try {
      await api.updateUser(u.id, { password: newPass.trim() });
      alert(`Đã đặt lại mật khẩu cho tài khoản ${u.username} thành công!`);
    } catch (err: any) {
      alert(err.message || 'Lỗi khi đổi mật khẩu');
    }
  };

  const renderPermissionMatrix = (
    permMap: Record<string, BookPermission[]>,
    onChangeMap: (next: Record<string, BookPermission[]>) => void
  ) => {
    const targetBooks = books.filter((b) => b.isActive);
    return (
      <div className="space-y-3">
        {targetBooks.map((book) => {
          const currentPerms = normalizeBookPermissionList(permMap[book.id] || []);
          const hasIssuePerm = currentPerms.includes('issue_number');
          return (
            <div key={book.id} className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <BookOpen className="w-3.5 h-3.5 text-sky-700" />
                  <span className="font-bold text-slate-900 text-xs">{book.name}</span>
                  <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-600">
                    {book.unitCode}
                  </span>
                  <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-800">
                    {book.code}
                  </span>
                </div>

                <div className="flex flex-wrap items-center gap-1">
                  {PERMISSION_PRESETS.map((preset) => (
                    <button
                      key={preset.label}
                      type="button"
                      onClick={() =>
                        onChangeMap({
                          ...permMap,
                          [book.id]: normalizeBookPermissionList([...preset.perms]),
                        })
                      }
                      className="px-2 py-0.5 text-[10px] font-semibold bg-white hover:bg-sky-50 text-slate-700 hover:text-sky-800 border border-slate-200 rounded transition-colors"
                    >
                      {preset.label}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {ALL_BOOK_PERMISSIONS.map((perm) => {
                  const meta = BOOK_PERMISSION_META[perm];
                  const checked = currentPerms.includes(perm);
                  const lockedViewByIssue = perm === 'view' && hasIssuePerm;
                  return (
                    <label
                      key={perm}
                      title={
                        lockedViewByIssue
                          ? 'Quyền Lấy số tự động bao gồm quyền Chỉ xem trên sổ này'
                          : meta.description
                      }
                      className={`flex items-start gap-2 p-2 rounded-md border text-[11px] transition-colors ${
                        lockedViewByIssue ? 'cursor-default' : 'cursor-pointer'
                      } ${
                        checked
                          ? 'bg-sky-50/80 border-sky-300 text-sky-950 font-semibold'
                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={lockedViewByIssue}
                        onChange={() => onChangeMap(togglePermissionInMap(permMap, book.id, perm))}
                        className="mt-0.5 rounded text-sky-700 focus:ring-sky-700 disabled:opacity-75"
                      />
                      <span className="leading-tight">
                        {meta.shortLabel}
                        {lockedViewByIssue && (
                          <span className="block text-[9px] font-normal text-sky-700">
                            Kèm theo quyền Lấy số
                          </span>
                        )}
                        {(perm === 'request_number' || perm === 'issue_number') && (
                          <span className="block text-[9px] font-normal text-slate-400">
                            {perm === 'request_number' ? 'Chỉ gửi xin số' : 'Tự động kèm Chỉ xem'}
                          </span>
                        )}
                      </span>
                    </label>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div className="max-w-7xl mx-auto py-8 px-4 sm:px-6 lg:px-8 space-y-6">
      {/* Header + Book Scope Selector */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Cài Đặt Hệ Thống & Quản Trị Sổ
          </h1>
          <p className="text-sm text-slate-600 mt-1">
            Quản lý nhiều sổ văn bản, quy tắc cấp số, người ký, phân quyền theo từng sổ và kết nối Google Drive.
          </p>
        </div>

        {managedBooks.length > 0 && (
          <div className="flex items-center gap-2 bg-white border border-slate-300 rounded-lg px-3 py-2 shadow-2xs text-xs">
            <BookOpen className="w-4 h-4 text-sky-700 shrink-0" />
            <span className="font-semibold text-slate-700 whitespace-nowrap">Sổ đang cấu hình:</span>
            <select
              value={selectedBook.id}
              onChange={(e) => {
                setSelectedBookId(e.target.value);
                onSelectBookId(e.target.value);
              }}
              className="font-bold text-sky-900 bg-transparent border-0 focus:outline-hidden cursor-pointer max-w-[240px]"
            >
              {managedBooks.map((b) => (
                <option key={b.id} value={b.id}>
                  {b.name} ({b.unitCode})
                </option>
              ))}
            </select>
          </div>
        )}
      </div>

      {!isAdmin && (
        <div className="p-3 bg-sky-50 border border-sky-200 rounded-lg text-xs text-sky-900 flex items-center gap-2">
          <Sliders className="w-4 h-4 text-sky-700 shrink-0" />
          <span>
            Bạn đang cấu hình với quyền <strong>Quản trị sổ</strong> cho sổ <strong>{selectedBook.name}</strong>. Các thiết lập toàn hệ thống và phân quyền tài khoản do Quản trị viên quản lý.
          </span>
        </div>
      )}

      {/* Sub Tabs */}
      <div className="flex border-b border-slate-200 overflow-x-auto gap-2">
        <button
          onClick={() => setActiveSubTab('books')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'books'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Sổ văn bản ({books.length})
        </button>

        <button
          onClick={() => setActiveSubTab('general')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'general'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Thông tin sổ & Mẫu số
        </button>

        <button
          onClick={() => setActiveSubTab('numbering')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'numbering'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Quy tắc dãy số & Số bắt đầu
        </button>

        <button
          onClick={() => setActiveSubTab('signers')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'signers'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Người ký theo sổ ({displayedBookSigners.length})
        </button>

        <button
          onClick={() => setActiveSubTab('googledrive')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'googledrive'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Lưu trữ Google Drive
        </button>

        <button
          onClick={() => setActiveSubTab('docTypes')}
          className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
            activeSubTab === 'docTypes'
              ? 'border-sky-700 text-sky-800'
              : 'border-transparent text-slate-500 hover:text-slate-800'
          }`}
        >
          Loại văn bản ({documentTypes.length})
        </button>

        {isAdmin && (
          <>
            <button
              onClick={() => setActiveSubTab('googlesheets')}
              className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
                activeSubTab === 'googlesheets'
                  ? 'border-emerald-700 text-emerald-800'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              Google Trang tính
            </button>

            <button
              onClick={() => setActiveSubTab('msgraph')}
              className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
                activeSubTab === 'msgraph'
                  ? 'border-sky-700 text-sky-800'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              OneDrive / SharePoint
            </button>

            <button
              onClick={handleOpenUsersTab}
              className={`pb-3 px-3 text-xs font-bold border-b-2 whitespace-nowrap transition-colors ${
                activeSubTab === 'users'
                  ? 'border-sky-700 text-sky-800'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              Tài khoản & Phân quyền sổ
            </button>
          </>
        )}
      </div>

      {/* SUB-TAB: QUẢN LÝ NHIỀU SỔ VĂN BẢN */}
      {activeSubTab === 'books' && (
        <div className="space-y-6 max-w-4xl">
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h2 className="text-base font-bold text-slate-900">
                  Danh Sách Sổ Văn Bản Trong Hệ Thống
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Hỗ trợ sổ dùng chung toàn trường và sổ riêng theo từng bộ phận / đơn vị với dãy số độc lập.
                </p>
              </div>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {books.map((book) => {
                const isSelected = book.id === selectedBook.id;
                return (
                  <div
                    key={book.id}
                    className={`p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                      isSelected ? 'bg-sky-50/40' : ''
                    }`}
                  >
                    <div className="space-y-1">
                      <div className="flex items-center gap-2 flex-wrap">
                        <strong className="text-sm font-bold text-slate-900">{book.name}</strong>
                        <span className="font-mono font-bold px-2 py-0.5 rounded bg-sky-100 text-sky-900 text-[11px]">
                          {book.unitCode}
                        </span>
                        <span
                          className="font-mono px-2 py-0.5 rounded bg-slate-100 text-slate-700 border border-slate-200 text-[10px]"
                          title="Mã sổ nội bộ cố định (tên tab Google Sheets)"
                        >
                          Mã sổ: {book.code}
                        </span>
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                            book.googleDriveConfig?.isConfigured
                              ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                              : 'bg-amber-50 text-amber-800 border border-amber-200'
                          }`}
                        >
                          {book.googleDriveConfig?.isConfigured ? 'Đã sẵn sàng Drive' : 'Chưa cấu hình Drive'}
                        </span>
                        {!book.isActive && (
                          <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                            Đang tạm khóa
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-slate-500 flex flex-wrap items-center gap-2">
                        <span>{book.unitName}</span>
                        <span>·</span>
                        <span>Mẫu số: <code className="font-mono text-slate-700">{book.numberFormat}</code></span>
                        <span>·</span>
                        <span>{book.numberMode === 'shared' ? 'Dãy số chung' : 'Dãy số riêng theo loại'}</span>
                        <span>·</span>
                        <span>Loại dùng: <strong className="text-slate-700">{getBookEnabledTypes(book, documentTypes).map((t) => t.code).join(', ')}</strong></span>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      {(isAdmin || hasBookPermission(currentUser, book.id, 'manage_book')) && (
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedBookId(book.id);
                            onSelectBookId(book.id);
                            setActiveSubTab('general');
                          }}
                          className="px-2.5 py-1.5 bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 rounded-md font-semibold transition-colors"
                        >
                          Cấu hình sổ này
                        </button>
                      )}
                      {isAdmin && book.id !== DEFAULT_BOOK_ID && (
                        <button
                          type="button"
                          onClick={() => handleToggleBookActive(book)}
                          className={`px-2.5 py-1.5 rounded-md font-semibold border transition-colors ${
                            book.isActive
                              ? 'bg-amber-50 hover:bg-amber-100 text-amber-800 border-amber-200'
                              : 'bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border-emerald-200'
                          }`}
                        >
                          {book.isActive ? 'Tạm khóa' : 'Kích hoạt'}
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {isAdmin && (
            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-4">
              <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                <Plus className="w-4 h-4 text-sky-700" />
                <span>Tạo Sổ Văn Bản Mới (Phòng ban / Tổ chuyên môn / Sổ dùng chung)</span>
              </h3>

              <form onSubmit={handleCreateBook} className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Tên sổ văn bản *</label>
                  <input
                    type="text"
                    value={newBookName}
                    onChange={(e) => {
                      const val = e.target.value;
                      setNewBookName(val);
                      if (!newBookCodeEdited) {
                        setNewBookCode(normalizeBookCode(val));
                      }
                    }}
                    placeholder="Ví dụ: Tổ Ngữ văn / Công đoàn / Văn thư"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">
                    Mã sổ nội bộ cố định (Tên tab Google Sheets) *
                  </label>
                  <input
                    type="text"
                    value={newBookCode}
                    onChange={(e) => {
                      setNewBookCodeEdited(true);
                      setNewBookCode(normalizeBookCode(e.target.value));
                    }}
                    placeholder="Ví dụ: TO_NGU_VAN"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono uppercase text-sm"
                  />
                  <span className="text-[11px] text-slate-500">Không thể đổi sau khi tạo, dùng để định tuyến tab Sheets.</span>
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Tên đơn vị / bộ phận *</label>
                  <input
                    type="text"
                    value={newBookUnitName}
                    onChange={(e) => setNewBookUnitName(e.target.value)}
                    placeholder="Ví dụ: Tổ Ngữ văn - Trường THPT Nguyễn Chí Thanh"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Ký hiệu đơn vị chính thức *</label>
                  <input
                    type="text"
                    value={newBookUnitCode}
                    onChange={(e) => setNewBookUnitCode(e.target.value)}
                    placeholder="Ví dụ: TNV, CB, THPTNCT..."
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono uppercase text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Mẫu định dạng số</label>
                  <input
                    type="text"
                    value={newBookFormat}
                    onChange={(e) => setNewBookFormat(e.target.value)}
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Phạm vi & Chế độ dãy số</label>
                  <div className="grid grid-cols-2 gap-2">
                    <select
                      value={newBookIsShared ? 'shared' : 'unit'}
                      onChange={(e) => setNewBookIsShared(e.target.value === 'shared')}
                      className="w-full h-10 px-2.5 bg-white border border-slate-300 rounded-lg"
                    >
                      <option value="unit">Sổ riêng đơn vị</option>
                      <option value="shared">Dùng chung trường</option>
                    </select>
                    <select
                      value={newBookMode}
                      onChange={(e) => setNewBookMode(e.target.value === 'shared' ? 'shared' : 'separate')}
                      className="w-full h-10 px-2.5 bg-white border border-slate-300 rounded-lg"
                    >
                      <option value="separate">Dãy riêng theo loại</option>
                      <option value="shared">Dãy chung liên tục</option>
                    </select>
                  </div>
                </div>

                <div className="sm:col-span-2 space-y-2">
                  <label className="font-bold text-slate-700 block">
                    Loại văn bản được sử dụng trong sổ mới *
                  </label>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                    {documentTypes.filter((t) => t.isActive).map((t) => {
                      const checked = newBookEnabledTypes.includes(t.code);
                      return (
                        <label
                          key={t.code}
                          className={`flex items-center gap-2 p-2 rounded-lg border cursor-pointer ${
                            checked ? 'bg-sky-50 border-sky-300 text-sky-950 font-semibold' : 'bg-white border-slate-200 text-slate-600'
                          }`}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() =>
                              setNewBookEnabledTypes((prev) =>
                                prev.includes(t.code) ? prev.filter((c) => c !== t.code) : [...prev, t.code]
                              )
                            }
                            className="rounded text-sky-700"
                          />
                          <span>{t.name} ({t.code})</span>
                        </label>
                      );
                    })}
                  </div>
                </div>

                {bookMessage && (
                  <div
                    className={`sm:col-span-2 p-3 rounded-lg flex items-center gap-2 ${
                      bookMessage.success
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : 'bg-rose-50 text-rose-700 border border-rose-200'
                    }`}
                  >
                    {bookMessage.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                    )}
                    <span>{bookMessage.text}</span>
                  </div>
                )}

                <div className="sm:col-span-2 flex justify-end">
                  <button
                    type="submit"
                    disabled={bookSaving}
                    className="px-5 py-2.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-2"
                  >
                    <Plus className="w-4 h-4" />
                    <span>{bookSaving ? 'Đang tạo sổ...' : 'Tạo sổ văn bản mới'}</span>
                  </button>
                </div>
              </form>
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB 1: THÔNG TIN SỔ & MẪU SỐ */}
      {activeSubTab === 'general' && (
        <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs max-w-2xl">
          <div className="mb-4 pb-3 border-b border-slate-100 flex items-center justify-between">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Thông tin đơn vị & Mẫu số: {selectedBook.name}
              </h2>
              <p className="text-xs text-slate-500">
                {selectedBook.isShared ? 'Sổ dùng chung' : 'Sổ riêng đơn vị'} · Mã sổ cố định:{' '}
                <code className="font-mono font-bold text-sky-800">{selectedBook.code}</code>
              </p>
            </div>
          </div>

          <form onSubmit={handleSaveGeneral} className="space-y-5 text-xs">
            <div className="space-y-1">
              <label className="font-bold text-slate-700">Mã sổ nội bộ cố định (Tab Google Sheets)</label>
              <input
                type="text"
                value={selectedBook.code}
                disabled
                className="w-full h-9 px-3 bg-slate-100 border border-slate-200 rounded-lg font-mono text-xs text-slate-600 cursor-not-allowed"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Tên hiển thị sổ văn bản</label>
              <input
                type="text"
                value={bookTitleInput}
                onChange={(e) => setBookTitleInput(e.target.value)}
                disabled={!canManageSelectedBook || generalSaving}
                className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Tên đơn vị</label>
              <input
                type="text"
                value={unitName}
                onChange={(e) => setUnitName(e.target.value)}
                disabled={!canManageSelectedBook || generalSaving}
                className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Ký hiệu đơn vị chính thức</label>
              <input
                type="text"
                value={unitCode}
                onChange={(e) => setUnitCode(e.target.value)}
                disabled={!canManageSelectedBook || generalSaving}
                className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 font-mono focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
              <span className="text-[11px] text-slate-500">
                Ví dụ: ĐTN (Đoàn trường), CB (Chi bộ), THPTNCT (Nhà trường) — quản trị viên có thể đổi theo đúng ký hiệu đơn vị đang dùng.
              </span>
            </div>

            <div className="space-y-1">
              <label className="font-bold text-slate-700">Mẫu định dạng số văn bản</label>
              <input
                type="text"
                value={numberFormat}
                onChange={(e) => setNumberFormat(e.target.value)}
                disabled={!canManageSelectedBook || generalSaving}
                className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm text-slate-900 font-mono focus:outline-hidden focus:ring-2 focus:ring-sky-600 disabled:bg-slate-50"
              />
              <span className="text-[11px] text-slate-500">
                Các biến thay thế: <code className="bg-slate-100 px-1 py-0.5 rounded">{'{so}'}</code> (số thứ tự),{' '}
                <code className="bg-slate-100 px-1 py-0.5 rounded">{'{loai}'}</code> (loại văn bản),{' '}
                <code className="bg-slate-100 px-1 py-0.5 rounded">{'{don_vi}'}</code> (ký hiệu đơn vị). Ví dụ:{' '}
                <code className="font-bold text-sky-800">01-KH/{unitCode || 'ĐTN'}</code>
              </span>
            </div>

            <div className="space-y-2">
              <label className="font-bold text-slate-700 block">
                Loại văn bản được sử dụng trong sổ "{selectedBook.name}"
              </label>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {documentTypes.filter((t) => t.isActive).map((t) => {
                  const checked = bookEnabledTypes.includes(t.code);
                  return (
                    <label
                      key={t.code}
                      className={`flex items-center gap-2 p-2 rounded-lg border cursor-pointer ${
                        checked ? 'bg-sky-50 border-sky-300 text-sky-950 font-semibold' : 'bg-white border-slate-200 text-slate-600'
                      }`}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        disabled={!canManageSelectedBook || generalSaving}
                        onChange={() =>
                          setBookEnabledTypes((prev) =>
                            prev.includes(t.code) ? prev.filter((c) => c !== t.code) : [...prev, t.code]
                          )
                        }
                        className="rounded text-sky-700"
                      />
                      <span>{t.name} ({t.code})</span>
                    </label>
                  );
                })}
              </div>
            </div>

            {generalSuccess && (
              <div className="p-3 bg-emerald-50 text-emerald-800 rounded-lg flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                <span>Đã lưu thành công cấu hình sổ "{bookTitleInput}"!</span>
              </div>
            )}

            {canManageSelectedBook && (
              <div className="pt-2">
                <button
                  type="submit"
                  disabled={generalSaving}
                  className="px-5 py-2.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  <span>{generalSaving ? 'Đang lưu...' : 'Lưu cài đặt sổ'}</span>
                </button>
              </div>
            )}
          </form>
        </div>
      )}

      {/* SUB-TAB 2: QUY TẮC DÃY SỐ & SỐ BẮT ĐẦU */}
      {activeSubTab === 'numbering' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-4">
            <h2 className="text-base font-bold text-slate-900">
              Chế Độ Dãy Số Trong Năm — {selectedBook.name}
            </h2>
            <p className="text-xs text-slate-600">
              Quy tắc quản lý cách tăng số văn bản đi cho sổ <strong>{selectedBook.name}</strong>. Sang năm mới (múi giờ Asia/Ho_Chi_Minh), dãy số tự động bắt đầu lại từ 01.
            </p>

            <form onSubmit={handleChangeMode} className="space-y-4 text-xs">
              <div className="space-y-3">
                <label className="flex items-start gap-3 p-3.5 border rounded-lg cursor-pointer transition-colors hover:bg-slate-50 border-slate-300">
                  <input
                    type="radio"
                    name="numberMode"
                    value="separate"
                    checked={targetMode === 'separate'}
                    onChange={() => setTargetMode('separate')}
                    disabled={!canManageSelectedBook}
                    className="mt-0.5 text-sky-600 focus:ring-sky-600"
                  />
                  <div>
                    <strong className="block font-semibold text-slate-900">
                      Mỗi loại văn bản có một dãy số riêng trong năm (Khuyên dùng)
                    </strong>
                    <span className="text-slate-500 text-[11px] block mt-0.5">
                      Kế hoạch tăng riêng (01-KH, 02-KH...), Báo cáo tăng riêng (01-BC, 02-BC...).
                    </span>
                  </div>
                </label>

                <label className="flex items-start gap-3 p-3.5 border rounded-lg cursor-pointer transition-colors hover:bg-slate-50 border-slate-300">
                  <input
                    type="radio"
                    name="numberMode"
                    value="shared"
                    checked={targetMode === 'shared'}
                    onChange={() => setTargetMode('shared')}
                    disabled={!canManageSelectedBook}
                    className="mt-0.5 text-sky-600 focus:ring-sky-600"
                  />
                  <div>
                    <strong className="block font-semibold text-slate-900">
                      Dùng một dãy số chung cho tất cả các loại văn bản
                    </strong>
                    <span className="text-slate-500 text-[11px] block mt-0.5">
                      Tất cả văn bản dùng chung 1 thứ tự liên tục từ 01, 02, 03...
                    </span>
                  </div>
                </label>
              </div>

              {targetMode !== selectedBook.numberMode && (
                <div className="p-4 bg-rose-50 border border-rose-300 rounded-lg space-y-2">
                  <div className="flex items-start gap-2 text-rose-800">
                    <AlertTriangle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                    <div>
                      <strong className="block font-semibold">Cảnh báo chuyển đổi chế độ cấp số:</strong>
                      <span>
                        Nếu trong năm sổ này đã phát sinh văn bản, để xác nhận vui lòng nhập{' '}
                        <strong className="text-rose-950">XÁC NHẬN CHUYỂN ĐỔI</strong>:
                      </span>
                    </div>
                  </div>

                  <input
                    type="text"
                    value={modeConfirmText}
                    onChange={(e) => setModeConfirmText(e.target.value)}
                    placeholder="Nhập: XÁC NHẬN CHUYỂN ĐỔI"
                    className="w-full h-9 px-3 bg-white border border-rose-300 rounded-lg text-xs font-bold text-rose-900"
                  />
                </div>
              )}

              {modeError && (
                <div className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs">{modeError}</div>
              )}

              {modeSuccess && (
                <div className="p-3 bg-emerald-50 text-emerald-800 rounded-lg text-xs flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                  <span>Đã cập nhật chế độ cấp số thành công!</span>
                </div>
              )}

              {canManageSelectedBook && targetMode !== selectedBook.numberMode && (
                <button
                  type="submit"
                  disabled={modeSaving}
                  className="px-4 py-2 bg-rose-700 hover:bg-rose-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {modeSaving ? 'Đang cập nhật...' : 'Xác nhận áp dụng chế độ mới'}
                </button>
              )}
            </form>
          </div>

          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-4">
            <h2 className="text-base font-bold text-slate-900">
              Thiết Lập Số Bắt Đầu Tiếp Theo — {selectedBook.name}
            </h2>
            <p className="text-xs text-slate-600">
              Nếu đơn vị đã cấp một số văn bản trên sổ giấy trước khi dùng phần mềm, có thể đặt số bắt đầu tiếp theo để hệ thống tiếp nối chính xác.
            </p>

            <form onSubmit={handleSetStartSequence} className="space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Năm áp dụng</label>
                  <select
                    value={seqYear}
                    onChange={(e) => setSeqYear(parseInt(e.target.value, 10))}
                    disabled={!canManageSelectedBook}
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg"
                  >
                    {[currentYear, currentYear - 1, currentYear + 1].map((y) => (
                      <option key={y} value={y}>
                        Năm {y}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Loại văn bản</label>
                  <select
                    value={seqTypeCode}
                    onChange={(e) => setSeqTypeCode(e.target.value)}
                    disabled={!canManageSelectedBook}
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg"
                  >
                    {selectedBook.numberMode === 'shared' ? (
                      <option value="SHARED">Dãy số chung (SHARED)</option>
                    ) : (
                      getBookEnabledTypes(selectedBook, documentTypes).map((t) => (
                        <option key={t.code} value={t.code}>
                          {t.name} ({t.code})
                        </option>
                      ))
                    )}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Số bắt đầu tiếp theo</label>
                  <input
                    type="number"
                    min={1}
                    value={startSequenceVal}
                    onChange={(e) => setStartSequenceVal(parseInt(e.target.value, 10) || 1)}
                    disabled={!canManageSelectedBook}
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono font-bold"
                  />
                </div>
              </div>

              {Object.keys(selectedBook.customStartSequences || {}).length > 0 && (
                <div className="p-3 bg-sky-50/70 border border-sky-200 rounded-lg text-[11px] text-sky-900">
                  <strong className="block mb-1">Các mốc bắt đầu đang cấu hình trên sổ:</strong>
                  <div className="flex flex-wrap gap-2">
                    {Object.entries(selectedBook.customStartSequences).map(([k, val]) => (
                      <span key={k} className="px-2 py-0.5 bg-white border border-sky-200 rounded font-mono">
                        {k}: bắt đầu từ <strong>{val}</strong>
                      </span>
                    ))}
                  </div>
                </div>
              )}

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-[11px] text-slate-600">
                <strong>Quy tắc bảo vệ:</strong> Số bắt đầu không được nhỏ hơn hoặc bằng số lớn nhất đã cấp trong sổ cho loại văn bản và năm đó.
              </div>

              {seqError && <div className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs">{seqError}</div>}

              {seqSuccess && (
                <div className="p-3 bg-emerald-50 text-emerald-800 rounded-lg text-xs flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                  <span>Đã thiết lập số bắt đầu tiếp theo thành công!</span>
                </div>
              )}

              {canManageSelectedBook && (
                <button
                  type="submit"
                  disabled={seqSaving}
                  className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {seqSaving ? 'Đang lưu...' : 'Lưu số bắt đầu'}
                </button>
              )}
            </form>
          </div>
        </div>
      )}

      {/* SUB-TAB 3: LOẠI VĂN BẢN (GLOBAL) */}
      {activeSubTab === 'docTypes' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <h2 className="text-base font-bold text-slate-900">
                Danh Sách Loại Văn Bản Hệ Thống
              </h2>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {documentTypes.map((t) => {
                const style = getTypeColorStyle(t.code, t.color);
                return (
                  <div key={t.id} className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <span
                        className={`font-mono font-bold px-2.5 py-1 rounded border ${style.badgeText} ${style.badgeBg} ${style.badgeBorder} shadow-2xs`}
                      >
                        <span className={`inline-block w-1.5 h-1.5 rounded-full ${style.dotColor} mr-1.5`} />
                        {t.code}
                      </span>
                      <div>
                        <h4 className="font-semibold text-slate-900">{t.name}</h4>
                        <span className="text-[11px] text-slate-500">
                          Mẫu: 01-{t.code}/{selectedBook.unitCode}
                        </span>
                      </div>
                    </div>

                    <div className="flex items-center gap-3">
                      {isAdmin && (
                        <div className="flex items-center gap-1 bg-slate-50 p-1 rounded-md border border-slate-200">
                          <span className="text-[10px] text-slate-500 font-medium px-1">Màu:</span>
                          {Object.values(DOCUMENT_COLOR_PALETTE).map((colorOpt) => (
                            <button
                              key={colorOpt.key}
                              type="button"
                              onClick={() => handleUpdateDocTypeColor(t.id, colorOpt.key)}
                              className={`w-3.5 h-3.5 rounded-full ${colorOpt.dotColor} transition-transform ${
                                (t.color || 'blue') === colorOpt.key
                                  ? 'ring-2 ring-offset-1 ring-slate-800 scale-125'
                                  : 'opacity-60 hover:opacity-100 hover:scale-110'
                              }`}
                              title={`Đổi màu sang ${colorOpt.label}`}
                            />
                          ))}
                        </div>
                      )}

                      <span className="text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-semibold whitespace-nowrap">
                        Đang sử dụng
                      </span>
                      {isAdmin && (
                        <button
                          onClick={() => handleDeleteDocType(t.id, t.code)}
                          className="p-1 text-slate-400 hover:text-rose-600 hover:bg-slate-100 rounded transition-colors"
                          title="Xóa loại văn bản này"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {isAdmin && (
            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Plus className="w-4 h-4 text-sky-700" />
                <span>Thêm loại văn bản mới</span>
              </h3>

              <form onSubmit={handleCreateDocType} className="grid grid-cols-1 sm:grid-cols-12 gap-3 text-xs">
                <div className="sm:col-span-6 space-y-1">
                  <label className="font-bold text-slate-700">Tên loại văn bản *</label>
                  <input
                    type="text"
                    value={newTypeName}
                    onChange={(e) => setNewTypeName(e.target.value)}
                    placeholder="Ví dụ: Hướng dẫn / Quyết định..."
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg"
                  />
                </div>

                <div className="sm:col-span-4 space-y-1">
                  <label className="font-bold text-slate-700">Ký hiệu viết tắt *</label>
                  <input
                    type="text"
                    value={newTypeCode}
                    onChange={(e) => setNewTypeCode(e.target.value)}
                    placeholder="Ví dụ: HD / QĐ"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg uppercase font-mono"
                  />
                </div>

                <div className="sm:col-span-2 flex items-end">
                  <button
                    type="submit"
                    disabled={typeSaving}
                    className="w-full h-10 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                  >
                    {typeSaving ? '...' : 'Thêm'}
                  </button>
                </div>

                <div className="sm:col-span-12 space-y-1.5 pt-1">
                  <label className="font-bold text-slate-700 block">Chọn màu sắc nhận diện loại văn bản</label>
                  <div className="grid grid-cols-5 sm:grid-cols-10 gap-1.5">
                    {Object.values(DOCUMENT_COLOR_PALETTE).map((colorOpt) => (
                      <button
                        key={colorOpt.key}
                        type="button"
                        onClick={() => setNewTypeColor(colorOpt.key)}
                        className={`flex flex-col items-center gap-1 p-1.5 rounded-lg border text-[10px] transition-all ${
                          newTypeColor === colorOpt.key
                            ? `${colorOpt.tagBg} ${colorOpt.tagBorder} ring-2 ring-sky-600 font-bold shadow-xs`
                            : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-600'
                        }`}
                      >
                        <span className={`w-3.5 h-3.5 rounded-full ${colorOpt.dotColor}`} />
                        <span className="truncate w-full text-center">{colorOpt.label.split(' ')[0]}</span>
                      </button>
                    ))}
                  </div>
                </div>

                {typeError && (
                  <div className="sm:col-span-12 p-2 bg-rose-50 text-rose-700 rounded text-xs">
                    {typeError}
                  </div>
                )}
              </form>
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB 4: NGƯỜI KÝ & CHỨC VỤ THEO SỔ */}
      {activeSubTab === 'signers' && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <h2 className="text-base font-bold text-slate-900">
                Danh Mục Người Ký Gợi Ý — {selectedBook.name}
              </h2>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {displayedBookSigners.map((s) => (
                <div key={s.id} className="p-4 flex items-center justify-between">
                  <div>
                    <h4 className="font-semibold text-slate-900">{s.name}</h4>
                    <span className="text-slate-500 text-[11px]">{s.role}</span>
                  </div>

                  {canManageSelectedBook && (
                    <button
                      onClick={() => handleDeleteBookSigner(s.id)}
                      className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-slate-100 rounded transition-colors"
                      title="Xóa người ký này"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          {canManageSelectedBook && (
            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs">
              <h3 className="text-sm font-bold text-slate-900 mb-3 flex items-center gap-2">
                <Plus className="w-4 h-4 text-sky-700" />
                <span>Thêm người ký mới vào sổ "{selectedBook.name}"</span>
              </h3>

              <form onSubmit={handleCreateBookSigner} className="grid grid-cols-1 sm:grid-cols-12 gap-3 text-xs">
                <div className="sm:col-span-5 space-y-1">
                  <label className="font-bold text-slate-700">Họ và tên *</label>
                  <input
                    type="text"
                    value={newSignerName}
                    onChange={(e) => setNewSignerName(e.target.value)}
                    placeholder="Ví dụ: Võ Văn Hoàng"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg"
                  />
                </div>

                <div className="sm:col-span-5 space-y-1">
                  <label className="font-bold text-slate-700">Chức vụ *</label>
                  <input
                    type="text"
                    value={newSignerRole}
                    onChange={(e) => setNewSignerRole(e.target.value)}
                    placeholder="Ví dụ: Phó Bí thư Đoàn trường"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg"
                  />
                </div>

                <div className="sm:col-span-2 flex items-end">
                  <button
                    type="submit"
                    disabled={signerSaving}
                    className="w-full h-10 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                  >
                    {signerSaving ? '...' : 'Thêm'}
                  </button>
                </div>

                {signerError && (
                  <div className="sm:col-span-12 p-2 bg-rose-50 text-rose-700 rounded text-xs">
                    {signerError}
                  </div>
                )}
              </form>
            </div>
          )}
        </div>
      )}

      {/* SUB-TAB: LƯU TRỮ GOOGLE DRIVE (THEO TỪNG SỔ) */}
      {activeSubTab === 'googledrive' && (() => {
        const parsedPreviewFolderId = extractGoogleDriveFolderId(gdFolderUrl);
        const currentDriveCfg = effectiveBookDriveCfg;
        const maxWordMB = currentDriveCfg?.maxFileSizeMB || DEFAULT_MAX_WORD_MB;

        return (
          <div className="space-y-6 max-w-3xl">
            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-5">
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                    <HardDrive className="w-5 h-5 text-sky-700" />
                    <span>Lưu Trữ Tệp Google Drive — {selectedBook.name}</span>
                  </h2>
                  <p className="text-xs text-slate-600 mt-1">
                    Cấu hình thư mục Google Drive riêng cho sổ <strong>{selectedBook.name} ({selectedBook.unitCode})</strong>. Tệp Word khi cấp số hoặc PDF văn bản đến sẽ tự đổi tên theo mẫu <code className="font-mono">Số.ký hiệu - Tên văn bản.đuôi</code> và lưu vào thư mục này.
                  </p>
                </div>

                <div className="flex flex-col items-end gap-1">
                  <span
                    className={`text-xs px-2.5 py-1 rounded font-semibold ${
                      currentDriveCfg?.isConfigured
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : gdEnabled
                        ? 'bg-amber-50 text-amber-800 border border-amber-200'
                        : 'bg-slate-100 text-slate-600'
                    }`}
                  >
                    {currentDriveCfg?.isConfigured
                      ? 'Đã sẵn sàng lưu tệp'
                      : gdEnabled
                      ? 'Thiếu thông tin cấu hình'
                      : 'Chưa bật'}
                  </span>
                  <span className="text-[11px] text-slate-500">
                    Giới hạn tệp máy chủ: <strong>Tối đa {maxWordMB} MB</strong>
                  </span>
                </div>
              </div>

              <form onSubmit={handleSaveGoogleDrive} className="space-y-5 text-xs">
                <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                  <input
                    type="checkbox"
                    checked={gdEnabled}
                    onChange={(e) => setGdEnabled(e.target.checked)}
                    disabled={!canManageSelectedBook}
                    className="rounded text-sky-700 focus:ring-sky-700"
                  />
                  <span className="font-semibold text-slate-900">
                    Bật lưu trữ tệp Word / PDF vào thư mục Google Drive cho sổ "{selectedBook.name}"
                  </span>
                </label>

                <div className="p-3 bg-sky-50/70 border border-sky-200 rounded-lg text-[11px] text-sky-950 space-y-1">
                  <strong className="block">Cấu trúc thư mục tự động trên Google Drive của sổ này (cách ly hoàn toàn với sổ khác):</strong>
                  <div className="font-mono font-semibold text-sky-900">
                    Thư mục gốc ({selectedBook.code}) → {selectedBook.name} → {currentYear} → Văn bản đi / Văn bản đến
                  </div>
                </div>

                {canManageSelectedBook && (
                  <div className="space-y-1.5">
                    <label className="font-bold text-slate-700 block">
                      URL hoặc ID thư mục Google Drive (Sổ: {selectedBook.unitCode}) *
                    </label>
                    <input
                      type="text"
                      value={gdFolderUrl}
                      onChange={(e) => setGdFolderUrl(e.target.value)}
                      placeholder="https://drive.google.com/drive/folders/1AbCDefGhIjkLMN... hoặc nhập trực tiếp Folder ID"
                      className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs focus:ring-2 focus:ring-sky-600"
                    />
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-500">
                      <span>
                        ID thư mục nhận diện:{' '}
                        {parsedPreviewFolderId ? (
                          <strong className="font-mono text-sky-800">{parsedPreviewFolderId}</strong>
                        ) : (
                          <span className="text-rose-600">Chưa hợp lệ</span>
                        )}
                      </span>
                      {(gdTestResult?.folderName || currentDriveCfg?.folderName) && (
                        <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold">
                          <FolderCheck className="w-3.5 h-3.5" />
                          <span>Thư mục: {gdTestResult?.folderName || currentDriveCfg?.folderName}</span>
                        </span>
                      )}
                    </div>
                  </div>
                )}

                {canManageSelectedBook && (
                  <div className="space-y-2">
                    <label className="font-bold text-slate-700 block">
                      Phương thức kết nối Google Drive (Xác thực phía máy chủ)
                    </label>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <label
                        className={`flex items-start gap-2.5 p-3 rounded-lg border cursor-pointer transition-colors ${
                          gdAuthMode === 'apps_script'
                            ? 'bg-sky-50 border-sky-400'
                            : 'bg-white border-slate-200 hover:bg-slate-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="gdAuthMode"
                          checked={gdAuthMode === 'apps_script'}
                          onChange={() => setGdAuthMode('apps_script')}
                          className="mt-0.5 text-sky-700"
                        />
                        <div>
                          <strong className="block text-slate-900">Google Apps Script (Khuyên dùng)</strong>
                          <span className="text-[11px] text-slate-600">
                            Phù hợp cho <strong>Drive của tôi</strong> & tài khoản Google cá nhân/trường học. Hỗ trợ cả thay tệp cũ.
                          </span>
                        </div>
                      </label>

                      <label
                        className={`flex items-start gap-2.5 p-3 rounded-lg border cursor-pointer transition-colors ${
                          gdAuthMode === 'service_account'
                            ? 'bg-sky-50 border-sky-400'
                            : 'bg-white border-slate-200 hover:bg-slate-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="gdAuthMode"
                          checked={gdAuthMode === 'service_account'}
                          onChange={() => setGdAuthMode('service_account')}
                          className="mt-0.5 text-sky-700"
                        />
                        <div>
                          <strong className="block text-slate-900">Google Drive API (Service Account)</strong>
                          <span className="text-[11px] text-slate-600">
                            Dùng trực tiếp cho <strong>Shared Drive</strong> Google Workspace qua khóa JSON tài khoản dịch vụ.
                          </span>
                        </div>
                      </label>
                    </div>
                  </div>
                )}

                {canManageSelectedBook && gdAuthMode === 'apps_script' && (
                  <div className="space-y-4">
                    <div className="bg-sky-50/70 border border-sky-200 rounded-xl p-4 space-y-3">
                      <h3 className="font-bold text-sky-950 flex items-center gap-2">
                        <CheckCircle2 className="w-4 h-4 text-sky-700" />
                        <span>Hướng dẫn triển khai Apps Script (Hỗ trợ upload, đối soát & thay tệp)</span>
                      </h3>
                      <ol className="list-decimal pl-5 space-y-1 text-slate-700">
                        <li>Truy cập <strong>script.google.com</strong> &rarr; Tạo dự án mới &rarr; Dán đoạn mã bên dưới.</li>
                        <li>Triển khai &rarr; Ứng dụng web (Execute as: <strong>Tôi</strong>, Who has access: <strong>Bất kỳ ai</strong>).</li>
                        <li>Dán <strong>URL ứng dụng web (/exec)</strong> vào ô bên dưới (nếu dùng chung credential mặc định của hệ thống có thể để trống nếu đã cấu hình ở Sổ mặc định).</li>
                      </ol>

                      <div className="relative">
                        <pre className="p-3 bg-slate-900 text-slate-100 rounded-lg text-[11px] font-mono overflow-x-auto max-h-52 leading-relaxed">
{GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE}
                        </pre>
                        <button
                          type="button"
                          onClick={() => {
                            navigator.clipboard.writeText(GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE);
                            setCopiedDriveScript(true);
                            setTimeout(() => setCopiedDriveScript(false), 2500);
                          }}
                          className="absolute top-2 right-2 px-2.5 py-1 text-xs font-semibold bg-sky-700 hover:bg-sky-600 text-white rounded shadow transition-colors flex items-center gap-1"
                        >
                          {copiedDriveScript ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                          <span>{copiedDriveScript ? 'Đã sao chép' : 'Sao chép mã Drive Script'}</span>
                        </button>
                      </div>
                    </div>

                    <div className="space-y-1.5">
                      <div className="flex items-center justify-between">
                        <label className="font-bold text-slate-700">
                          URL Web App Google Apps Script (/exec)
                        </label>
                        {currentDriveCfg?.hasAppsScriptUrl && (
                          <div className="flex items-center gap-2">
                            <span className="text-[11px] text-emerald-700 font-mono bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                              Đã lưu: {currentDriveCfg.appsScriptUrlMasked}
                            </span>
                            <button
                              type="button"
                              onClick={() => handleClearDriveCredential('apps_script')}
                              className="text-[11px] text-rose-600 hover:underline"
                            >
                              Xóa URL
                            </button>
                          </div>
                        )}
                      </div>
                      <input
                        type="url"
                        value={gdAppsScriptUrl}
                        onChange={(e) => setGdAppsScriptUrl(e.target.value)}
                        placeholder={
                          currentDriveCfg?.hasAppsScriptUrl
                            ? 'Để trống nếu giữ nguyên URL đang lưu, hoặc dán URL mới để thay thế...'
                            : 'https://script.google.com/macros/s/AKfycbx.../exec'
                        }
                        className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs focus:ring-2 focus:ring-sky-600"
                      />
                    </div>
                  </div>
                )}

                {canManageSelectedBook && gdAuthMode === 'service_account' && (
                  <div className="space-y-4">
                    <div className="space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <label className="font-bold text-slate-700">
                          Khóa JSON Tài khoản dịch vụ (Service Account Key)
                        </label>
                        {currentDriveCfg?.hasServiceAccount && (
                          <div className="flex items-center gap-2">
                            <span className="text-[11px] text-emerald-700 font-mono bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                              Đã cấu hình: {currentDriveCfg.serviceAccountEmail}
                            </span>
                            <button
                              type="button"
                              onClick={() => handleClearDriveCredential('service_account')}
                              className="text-[11px] text-rose-600 hover:underline"
                            >
                              Xóa khóa
                            </button>
                          </div>
                        )}
                      </div>

                      <div className="flex items-center gap-2">
                        <label className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg font-semibold text-slate-700 cursor-pointer">
                          <Upload className="w-3.5 h-3.5 text-sky-700" />
                          <span>Chọn tệp .json từ máy tính</span>
                          <input
                            type="file"
                            accept=".json,application/json"
                            className="hidden"
                            onChange={handleServiceAccountFileUpload}
                          />
                        </label>
                        {gdServiceAccountJson && (
                          <span className="text-[11px] text-sky-700 font-semibold">
                            Đã nạp nội dung JSON mới (bấm Lưu cấu hình để áp dụng)
                          </span>
                        )}
                      </div>

                      <textarea
                        value={gdServiceAccountJson}
                        onChange={(e) => setGdServiceAccountJson(e.target.value)}
                        rows={4}
                        placeholder={
                          currentDriveCfg?.hasServiceAccount
                            ? 'Để trống nếu giữ nguyên khóa JSON hiện tại, hoặc dán JSON mới để thay thế...'
                            : '{\n  "type": "service_account",\n  "client_email": "...",\n  "private_key": "-----BEGIN PRIVATE KEY-----..."\n}'
                        }
                        className="w-full p-2.5 bg-white border border-slate-300 rounded-lg font-mono text-[11px] focus:ring-2 focus:ring-sky-600"
                      />
                    </div>
                  </div>
                )}

                {currentDriveCfg?.lastCheckedAt && !gdTestResult && (
                  <div className="text-[11px] text-slate-500 flex items-center gap-1.5">
                    <Clock className="w-3.5 h-3.5 text-slate-400" />
                    <span>
                      Kiểm tra gần nhất ({new Date(currentDriveCfg.lastCheckedAt).toLocaleString('vi-VN')}):{' '}
                      <strong className={currentDriveCfg.lastCheckStatus === 'ok' ? 'text-emerald-700' : 'text-rose-700'}>
                        {currentDriveCfg.lastCheckMessage}
                      </strong>
                    </span>
                  </div>
                )}

                {gdSaveMessage && (
                  <div
                    className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                      gdSaveMessage.success
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : 'bg-rose-50 text-rose-800 border border-rose-200'
                    }`}
                  >
                    {gdSaveMessage.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    )}
                    <span>{gdSaveMessage.message}</span>
                  </div>
                )}

                {gdTestResult && (
                  <div
                    className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                      gdTestResult.success
                        ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                        : 'bg-rose-50 text-rose-800 border border-rose-200'
                    }`}
                  >
                    {gdTestResult.success ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    ) : (
                      <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    )}
                    <span>{gdTestResult.message}</span>
                  </div>
                )}

                {canManageSelectedBook && (
                  <div className="flex flex-wrap items-center gap-3 pt-2">
                    <button
                      type="submit"
                      disabled={gdSaving}
                      className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-1.5"
                    >
                      <Save className="w-3.5 h-3.5" />
                      <span>{gdSaving ? 'Đang lưu...' : 'Lưu cấu hình Google Drive'}</span>
                    </button>

                    <button
                      type="button"
                      onClick={handleTestGoogleDrive}
                      disabled={gdTesting || !parsedPreviewFolderId}
                      className="px-4 py-2 bg-slate-100 hover:bg-slate-200 disabled:opacity-50 text-slate-800 font-semibold rounded-lg transition-colors flex items-center gap-1.5"
                    >
                      <RotateCw className={`w-3.5 h-3.5 ${gdTesting ? 'animate-spin' : ''}`} />
                      <span>{gdTesting ? 'Đang kiểm tra quyền ghi...' : 'Kiểm tra kết nối'}</span>
                    </button>
                  </div>
                )}
              </form>
            </div>
          </div>
        );
      })()}

      {/* SUB-TAB: GOOGLE TRANG TÍNH (ADMIN ONLY) */}
      {activeSubTab === 'googlesheets' && isAdmin && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-5">
            <div className="flex items-start justify-between">
              <div>
                <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded bg-emerald-600 text-white flex items-center justify-center text-xs font-bold shadow-2xs">
                    GS
                  </span>
                  <span>Đồng Bộ Vào Google Trang Tính (Google Sheets)</span>
                </h2>
                <p className="text-xs text-slate-600 mt-1">
                  Tự động thêm/cập nhật từng dòng văn bản sang Google Trang tính của trường.
                </p>
              </div>

              <div className="text-right">
                <span
                  className={`text-xs px-2.5 py-1 rounded font-semibold ${
                    gsEnabled ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-600'
                  }`}
                >
                  {gsEnabled ? 'Đang hoạt động' : 'Chưa kích hoạt'}
                </span>
              </div>
            </div>

            <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-4 text-xs space-y-3">
              <h3 className="font-bold text-emerald-950 flex items-center gap-2">
                <CheckCircle2 className="w-4 h-4 text-emerald-700" />
                <span>Cách tạo liên kết đồng bộ đến Google Sheets:</span>
              </h3>
              <ol className="list-decimal pl-5 space-y-1.5 text-slate-700">
                <li>Mở Google Trang tính &rarr; Tiện ích mở rộng &rarr; Apps Script.</li>
                <li>Dán đoạn mã bên dưới, bấm Lưu và Triển khai dạng Ứng dụng web (Execute as: Tôi, Access: Bất kỳ ai).</li>
              </ol>

              <div className="relative">
                <pre className="p-3 bg-slate-900 text-slate-100 rounded-lg text-[11px] font-mono overflow-x-auto max-h-56 leading-relaxed">
{GOOGLE_APPS_SCRIPT_TEMPLATE}
                </pre>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE);
                    setCopiedScript(true);
                    setTimeout(() => setCopiedScript(false), 2500);
                  }}
                  className="absolute top-2 right-2 px-2.5 py-1 text-xs font-semibold bg-emerald-700 hover:bg-emerald-600 text-white rounded shadow transition-colors flex items-center gap-1"
                >
                  {copiedScript ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedScript ? 'Đã sao chép' : 'Sao chép mã Apps Script'}</span>
                </button>
              </div>
            </div>

            <form onSubmit={handleSaveGoogleSheets} className="space-y-4 text-xs">
              <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                <input
                  type="checkbox"
                  checked={gsEnabled}
                  onChange={(e) => setGsEnabled(e.target.checked)}
                  className="rounded text-emerald-700 focus:ring-emerald-700"
                />
                <span className="font-semibold text-slate-900">
                  Bật tự động đồng bộ sang Google Trang tính
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="space-y-1 sm:col-span-2">
                  <label className="font-bold text-slate-700">
                    URL Ứng Dụng Web Google Apps Script Master (/exec) *
                  </label>
                  <input
                    type="url"
                    value={gsWebhookUrl}
                    onChange={(e) => setGsWebhookUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/AKfycbx.../exec"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs focus:ring-2 focus:ring-emerald-600"
                  />
                  <span className="text-[11px] text-slate-500">
                    Tự động chia mỗi sổ thành 1 tab theo Mã sổ cố định: {books.map((b) => b.code).join(', ')}.
                  </span>
                </div>

                <div className="space-y-1 sm:col-span-2">
                  <label className="font-bold text-slate-700">Sổ dùng để kiểm tra / đồng bộ theo sổ</label>
                  <select
                    value={gsTestBookId}
                    onChange={(e) => setGsTestBookId(e.target.value)}
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-xs font-semibold"
                  >
                    {books.map((b) => (
                      <option key={b.id} value={b.id}>
                        {b.name} — Tab: {b.code} ({b.unitCode})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {gsTestResult && (
                <div
                  className={`p-3 rounded-lg text-xs flex items-start gap-2 ${
                    gsTestResult.success
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}
                >
                  <span>{gsTestResult.message}</span>
                </div>
              )}

              {gsSyncResult && (
                <div className="p-3 bg-slate-100 text-slate-800 rounded-lg text-xs">{gsSyncResult}</div>
              )}

              <div className="flex flex-wrap items-center gap-3 pt-2">
                <button
                  type="submit"
                  disabled={gsSaving}
                  className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {gsSaving ? 'Đang lưu...' : 'Lưu cấu hình Master Sheets'}
                </button>

                <button
                  type="button"
                  onClick={handleTestGoogleSheets}
                  disabled={gsTesting || !gsWebhookUrl}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg transition-colors flex items-center gap-1.5"
                >
                  <RotateCw className={`w-3.5 h-3.5 ${gsTesting ? 'animate-spin' : ''}`} />
                  <span>{gsTesting ? 'Đang gửi kiểm tra...' : 'Gửi dòng thử nghiệm lên tab sổ'}</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleSyncAllGoogleSheets()}
                  disabled={gsSyncingAll || !gsEnabled || !gsWebhookUrl}
                  className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {gsSyncingAll ? 'Đang chuyển toàn bộ...' : 'Đồng bộ toàn bộ các sổ'}
                </button>
              </div>
            </form>

            {/* Optional per-book Google Sheets webhook */}
            <div className="pt-5 border-t border-slate-200 space-y-3">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  Google Sheets riêng cho sổ "{selectedBook.name}" (tùy chọn)
                </h3>
                <p className="text-[11px] text-slate-500">
                  Dùng khi cần đẩy thêm bản sao riêng cho bảng tính nội bộ của {selectedBook.name} bên cạnh bảng tính Master của trường.
                </p>
              </div>

              <form onSubmit={handleSaveBookGoogleSheets} className="space-y-3 text-xs">
                <label className="flex items-center gap-2.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={bookGsEnabled}
                    onChange={(e) => setBookGsEnabled(e.target.checked)}
                    className="rounded text-emerald-700"
                  />
                  <span className="font-semibold text-slate-800">
                    Bật đồng bộ thêm sang Google Sheets riêng của sổ "{selectedBook.name}" ({selectedBook.code})
                  </span>
                </label>

                <input
                  type="url"
                  value={bookGsWebhookUrl}
                  onChange={(e) => setBookGsWebhookUrl(e.target.value)}
                  placeholder="https://script.google.com/macros/s/.../exec (để trống nếu chỉ dùng bảng tính Master)"
                  className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                />

                {bookGsMessage && (
                  <div className="p-2.5 rounded-lg bg-slate-100 text-slate-800 text-xs">{bookGsMessage}</div>
                )}

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="submit"
                    disabled={bookGsSaving}
                    className="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-semibold rounded-lg"
                  >
                    {bookGsSaving ? 'Đang lưu...' : `Lưu Sheets riêng của ${selectedBook.code}`}
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSyncAllGoogleSheets(selectedBook.id)}
                    disabled={gsSyncingAll}
                    className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg"
                  >
                    Đồng bộ riêng sổ {selectedBook.name}
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* SUB-TAB: MICROSOFT GRAPH EXCEL (ADMIN ONLY) */}
      {activeSubTab === 'msgraph' && isAdmin && (
        <div className="space-y-6 max-w-3xl">
          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-5">
            <div className="flex items-start justify-between">
              <div>
                <h2 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Cloud className="w-5 h-5 text-sky-700" />
                  <span>Đồng Bộ Tự Động Vào Excel Trực Tuyến (OneDrive / SharePoint)</span>
                </h2>
              </div>
              <span
                className={`text-xs px-2.5 py-1 rounded font-semibold ${
                  msEnabled ? 'bg-emerald-50 text-emerald-800' : 'bg-slate-100 text-slate-600'
                }`}
              >
                {msEnabled ? 'Đã bật đồng bộ' : 'Chưa bật'}
              </span>
            </div>

            <form onSubmit={handleSaveMsGraph} className="space-y-4 text-xs">
              <label className="flex items-center gap-3 p-3 bg-slate-50 border border-slate-200 rounded-lg cursor-pointer">
                <input
                  type="checkbox"
                  checked={msEnabled}
                  onChange={(e) => setMsEnabled(e.target.checked)}
                  className="rounded text-sky-700 focus:ring-sky-700"
                />
                <span className="font-semibold text-slate-900">
                  Bật tính năng tự động đồng bộ vào Microsoft Graph Excel
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Azure Directory (Tenant) ID</label>
                  <input
                    type="text"
                    value={msTenantId}
                    onChange={(e) => setMsTenantId(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Application (Client) ID</label>
                  <input
                    type="text"
                    value={msClientId}
                    onChange={(e) => setMsClientId(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
                <div className="space-y-1 sm:col-span-2">
                  <label className="font-bold text-slate-700">Client Secret</label>
                  <input
                    type="password"
                    value={msClientSecret}
                    onChange={(e) => setMsClientSecret(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Drive ID (Tùy chọn)</label>
                  <input
                    type="text"
                    value={msDriveId}
                    onChange={(e) => setMsDriveId(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Đường dẫn file Excel</label>
                  <input
                    type="text"
                    value={msFilePath}
                    onChange={(e) => setMsFilePath(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
                <div className="space-y-1 sm:col-span-2">
                  <label className="font-bold text-slate-700">Tên bảng dữ liệu (Table Name)</label>
                  <input
                    type="text"
                    value={msTableName}
                    onChange={(e) => setMsTableName(e.target.value)}
                    className="w-full h-9 px-3 bg-white border border-slate-300 rounded-lg font-mono text-xs"
                  />
                </div>
              </div>

              {msTestResult && (
                <div
                  className={`p-3 rounded-lg text-xs ${
                    msTestResult.success
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-rose-50 text-rose-800 border border-rose-200'
                  }`}
                >
                  {msTestResult.message}
                </div>
              )}

              {msSyncResult && (
                <div className="p-3 bg-slate-100 text-slate-800 rounded-lg text-xs">{msSyncResult}</div>
              )}

              <div className="flex flex-wrap items-center gap-3 pt-2">
                <button
                  type="submit"
                  disabled={msSaving}
                  className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {msSaving ? 'Đang lưu...' : 'Lưu cấu hình'}
                </button>
                <button
                  type="button"
                  onClick={handleTestMsConnection}
                  disabled={msTesting || !msTenantId || !msClientId}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-semibold rounded-lg transition-colors flex items-center gap-1.5"
                >
                  <RotateCw className={`w-3.5 h-3.5 ${msTesting ? 'animate-spin' : ''}`} />
                  <span>{msTesting ? 'Đang kiểm tra...' : 'Kiểm tra kết nối'}</span>
                </button>
                <button
                  type="button"
                  onClick={handleSyncAll}
                  disabled={msSyncingAll || !msEnabled}
                  className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-semibold rounded-lg shadow-xs transition-colors"
                >
                  {msSyncingAll ? 'Đang đồng bộ...' : 'Đồng bộ lại tất cả'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* SUB-TAB: QUẢN LÝ TÀI KHOẢN & PHÂN QUYỀN THEO TỪNG SỔ (ADMIN ONLY) */}
      {activeSubTab === 'users' && isAdmin && (
        <div className="space-y-6 max-w-4xl">
          <div className="bg-sky-50 border border-sky-200 rounded-xl p-4 text-xs text-sky-900 space-y-2">
            <div className="flex items-center gap-2 font-bold text-sky-950 text-sm">
              <ShieldCheck className="w-4 h-4 text-sky-700" />
              <span>Chính sách tài khoản & Phân quyền chi tiết theo từng sổ</span>
            </div>
            <ul className="list-disc list-inside space-y-1 text-slate-700 leading-relaxed">
              <li>
                Mặc định <strong>01 tài khoản Quản trị viên (admin)</strong> cố định toàn quyền quản lý hệ thống.
              </li>
              <li>
                Quản trị viên gán <strong>Họ và tên cố định</strong> cho từng tài khoản tạo mới; người dùng không thể tự thay đổi tên này.
              </li>
              <li>
                Mỗi tài khoản có thể được phân chi tiết <strong>9 nhóm quyền theo từng sổ văn bản</strong>:{' '}
                <em>Xem, Nhập văn bản đến, Xin số, Cấp số, Sửa, Thay tệp, Hủy, Xuất sổ, Quản trị sổ</em>.
                Quyền <strong>Xin số</strong> và <strong>Cấp số</strong> được tách biệt hoàn toàn.
              </li>
            </ul>
          </div>

          {/* Account List */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100">
              <h2 className="text-base font-bold text-slate-900">
                Danh Sách Tài Khoản Hệ Thống
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Bấm <strong>Phân quyền sổ</strong> để điều chỉnh 9 quyền chi tiết trên từng sổ văn bản cho mỗi tài khoản.
              </p>
            </div>

            <div className="divide-y divide-slate-100 text-xs">
              {loadingUsers ? (
                <div className="p-6 text-center text-slate-400">Đang tải danh sách tài khoản...</div>
              ) : (
                usersList.map((u) => {
                  const isMainAdmin = isSystemAdmin(u);
                  const isEditingPerms = editingPermUserId === u.id;
                  return (
                    <div key={u.id} className="p-4 space-y-3">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div className="space-y-1.5">
                          <div className="flex items-center gap-2 flex-wrap">
                            <strong className="text-slate-900 font-bold text-sm">{u.fullName}</strong>
                            <span className="font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded text-[11px]">
                              @{u.username}
                            </span>
                            {isMainAdmin ? (
                              <span className="bg-amber-100 text-amber-900 font-bold px-2 py-0.5 rounded text-[10px] flex items-center gap-1">
                                <Shield className="w-3 h-3 text-amber-700" />
                                Quản trị viên cố định (Toàn quyền)
                              </span>
                            ) : (
                              <span className="bg-sky-100 text-sky-900 font-medium px-2 py-0.5 rounded text-[10px]">
                                Tên gán cố định
                              </span>
                            )}
                          </div>

                          {!isMainAdmin && (
                            <div className="flex flex-wrap gap-1.5">
                              {books
                                .filter((b) => b.isActive)
                                .map((book) => {
                                  const perms = u.bookPermissions?.[book.id] || [];
                                  if (perms.length === 0) return null;
                                  return (
                                    <span
                                      key={book.id}
                                      className="inline-flex items-center gap-1 px-2 py-0.5 bg-slate-100 border border-slate-200 rounded text-[10px] text-slate-700"
                                    >
                                      <strong className="text-sky-900">{book.unitCode}:</strong>
                                      <span>
                                        {perms.map((p) => BOOK_PERMISSION_META[p]?.shortLabel || p).join(', ')}
                                      </span>
                                    </span>
                                  );
                                })}
                            </div>
                          )}

                          <div className="text-[11px] text-slate-500">
                            Ngày tạo: {new Date(u.createdAt).toLocaleDateString('vi-VN')}
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-2 shrink-0">
                          <span
                            className={`text-[11px] px-2.5 py-1 rounded-md font-semibold ${
                              u.isActive
                                ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                : 'bg-rose-50 text-rose-800 border border-rose-200'
                            }`}
                          >
                            {u.isActive ? 'Đang hoạt động' : 'Tạm khóa'}
                          </span>

                          {!isMainAdmin && (
                            <button
                              type="button"
                              onClick={() => handleStartEditPermissions(u)}
                              className={`px-2.5 py-1 rounded-md font-semibold text-[11px] border transition-colors flex items-center gap-1 ${
                                isEditingPerms
                                  ? 'bg-sky-700 text-white border-sky-700'
                                  : 'bg-sky-50 hover:bg-sky-100 text-sky-800 border-sky-200'
                              }`}
                            >
                              <Users className="w-3 h-3" />
                              <span>{isEditingPerms ? 'Đóng phân quyền' : 'Phân quyền sổ'}</span>
                            </button>
                          )}

                          <button
                            type="button"
                            onClick={() => handleResetUserPass(u)}
                            className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md font-medium text-[11px] transition-colors flex items-center gap-1"
                          >
                            <Lock className="w-3 h-3 text-slate-500" />
                            <span>Đổi MK</span>
                          </button>

                          {!isMainAdmin && (
                            <>
                              <button
                                type="button"
                                onClick={() => handleToggleUserStatus(u)}
                                className={`px-2.5 py-1 rounded-md font-medium text-[11px] transition-colors flex items-center gap-1 ${
                                  u.isActive
                                    ? 'bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200'
                                    : 'bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200'
                                }`}
                              >
                                {u.isActive ? <UserX className="w-3 h-3" /> : <UserCheck className="w-3 h-3" />}
                                <span>{u.isActive ? 'Khóa' : 'Mở'}</span>
                              </button>

                              <button
                                type="button"
                                onClick={() => handleDeleteUser(u)}
                                className="p-1.5 text-slate-400 hover:text-rose-700 hover:bg-rose-50 rounded-md transition-colors"
                                title="Xóa tài khoản này khỏi hệ thống"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </>
                          )}
                        </div>
                      </div>

                      {!isMainAdmin && isEditingPerms && (
                        <div className="pt-3 border-t border-slate-100 space-y-3">
                          {renderPermissionMatrix(editingBookPermissions, setEditingBookPermissions)}
                          <div className="flex justify-end gap-2">
                            <button
                              type="button"
                              onClick={() => setEditingPermUserId(null)}
                              className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-semibold"
                            >
                              Hủy
                            </button>
                            <button
                              type="button"
                              onClick={() => handleSaveUserPermissions(u)}
                              disabled={savingPermUserId === u.id}
                              className="px-4 py-1.5 bg-sky-700 hover:bg-sky-800 text-white rounded-lg font-semibold flex items-center gap-1.5 disabled:opacity-50"
                            >
                              <Save className="w-3.5 h-3.5" />
                              <span>{savingPermUserId === u.id ? 'Đang lưu...' : 'Lưu phân quyền sổ'}</span>
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          </div>

          {/* Add New User Form */}
          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs">
            <h3 className="text-sm font-bold text-slate-900 mb-1 flex items-center gap-2">
              <Plus className="w-4 h-4 text-sky-700" />
              <span>Thêm tài khoản mới và phân quyền theo từng sổ</span>
            </h3>
            <p className="text-xs text-slate-500 mb-4">
              Họ và tên được gán cố định cho tài khoản. Chọn rõ các quyền trên từng sổ (Xin số tách biệt với Cấp số).
            </p>

            <form onSubmit={handleCreateUser} className="space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Tên đăng nhập *</label>
                  <input
                    type="text"
                    value={newUsername}
                    onChange={(e) => setNewUsername(e.target.value)}
                    placeholder="Ví dụ: chidoan12a"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Họ và tên (Gán cố định) *</label>
                  <input
                    type="text"
                    value={newFullName}
                    onChange={(e) => setNewFullName(e.target.value)}
                    placeholder="Ví dụ: Nguyễn Văn A"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg text-sm"
                  />
                </div>

                <div className="space-y-1">
                  <label className="font-bold text-slate-700">Mật khẩu khởi tạo *</label>
                  <input
                    type="password"
                    value={newUserPass}
                    onChange={(e) => setNewUserPass(e.target.value)}
                    placeholder="Tối thiểu 6 ký tự"
                    required
                    className="w-full h-10 px-3 bg-white border border-slate-300 rounded-lg font-mono text-sm"
                  />
                </div>
              </div>

              <div className="space-y-2">
                <label className="font-bold text-slate-700 block">
                  Phân quyền theo từng sổ văn bản
                </label>
                {renderPermissionMatrix(newBookPermissions, setNewBookPermissions)}
              </div>

              {userError && (
                <div className="p-3 bg-rose-50 text-rose-700 rounded-lg text-xs border border-rose-200 flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>{userError}</span>
                </div>
              )}

              <div className="pt-2 flex items-center justify-end">
                <button
                  type="submit"
                  disabled={userSaving}
                  className="px-6 py-2.5 bg-sky-700 hover:bg-sky-800 text-white font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-2 disabled:opacity-50"
                >
                  <Plus className="w-4 h-4" />
                  <span>{userSaving ? 'Đang tạo tài khoản...' : 'Tạo tài khoản & áp dụng quyền'}</span>
                </button>
              </div>
            </form>
          </div>

          {/* Nhật ký thay đổi phân quyền & quản trị sổ */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-slate-900">
                  Nhật ký thay đổi phân quyền & cấu hình sổ ({auditLogs.length})
                </h3>
                <p className="text-[11px] text-slate-500">
                  Ghi nhận các thao tác gán/thu hồi quyền, khóa tài khoản và tạo/sửa sổ để phục vụ kiểm tra.
                </p>
              </div>
              <button
                type="button"
                onClick={handleOpenUsersTab}
                className="text-xs font-semibold text-sky-700 hover:underline flex items-center gap-1"
              >
                <RotateCw className="w-3.5 h-3.5" />
                <span>Làm mới</span>
              </button>
            </div>

            <div className="divide-y divide-slate-100 text-xs max-h-80 overflow-y-auto">
              {loadingAuditLogs ? (
                <div className="p-4 text-center text-slate-400">Đang tải nhật ký...</div>
              ) : auditLogs.length === 0 ? (
                <div className="p-4 text-center text-slate-400">Chưa có nhật ký thay đổi phân quyền hoặc sổ nào.</div>
              ) : (
                auditLogs.slice(0, 40).map((log) => (
                  <div key={log.id} className="p-3.5 space-y-1">
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px]">
                      <div className="flex items-center gap-2">
                        <span className="px-2 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-900 font-mono font-bold">
                          {log.action}
                        </span>
                        <strong className="text-slate-800">{log.actorName}</strong>
                        {log.bookCode && (
                          <span className="font-mono px-1.5 py-0.5 rounded bg-slate-100 text-slate-700">
                            {log.bookCode}
                          </span>
                        )}
                      </div>
                      <span className="font-mono text-slate-400">
                        {new Date(log.timestamp).toLocaleString('vi-VN')}
                      </span>
                    </div>
                    <p className="text-slate-700">{log.details}</p>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
