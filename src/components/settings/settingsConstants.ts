import type { BookPermission, RegistryBook } from '../../types';
import { ALL_BOOK_PERMISSIONS, DEFAULT_BOOK_ID } from '../../utils/permissions';

export type SettingsSubTab =
  | 'books'
  | 'general'
  | 'numbering'
  | 'signers'
  | 'docTypes'
  | 'database'
  | 'googledrive'
  | 'googlesheets'
  | 'users';

export type SettingsGroupKey =
  | 'books_management'
  | 'numbering_catalog'
  | 'storage_sync'
  | 'accounts_permissions';

export interface SettingsNavItemDef {
  id: SettingsSubTab;
  label: string;
  adminOnly?: boolean;
}

export interface SettingsNavGroupDef {
  key: SettingsGroupKey;
  title: string;
  adminOnly?: boolean;
  items: SettingsNavItemDef[];
}

export const SETTINGS_NAV_GROUPS: SettingsNavGroupDef[] = [
  {
    key: 'books_management',
    title: 'Quản lý sổ',
    items: [
      { id: 'books', label: 'Danh sách sổ' },
      { id: 'general', label: 'Thông tin sổ và mẫu số' },
    ],
  },
  {
    key: 'numbering_catalog',
    title: 'Cấp số & Danh mục',
    items: [
      { id: 'numbering', label: 'Quy tắc dãy số và số bắt đầu' },
      { id: 'signers', label: 'Người ký' },
      { id: 'docTypes', label: 'Loại văn bản' },
    ],
  },
  {
    key: 'storage_sync',
    title: 'Lưu trữ & Đồng bộ',
    items: [
      { id: 'database', label: 'Upstash Redis & Lưu trữ chính', adminOnly: true },
      { id: 'googledrive', label: 'Google Drive' },
      { id: 'googlesheets', label: 'Google Trang tính', adminOnly: true },
    ],
  },
  {
    key: 'accounts_permissions',
    title: 'Tài khoản & Phân quyền',
    adminOnly: true,
    items: [{ id: 'users', label: 'Tài khoản, vai trò và quyền theo từng sổ', adminOnly: true }],
  },
];

export interface BookIdentityStyle {
  key: 'blue' | 'purple' | 'teal' | 'neutral';
  iconBg: string;
  iconText: string;
  badgeBg: string;
  badgeText: string;
  badgeBorder: string;
  dotColor: string;
}

export function getBookIdentityStyle(book: Pick<RegistryBook, 'id' | 'code'>): BookIdentityStyle {
  const code = (book.code || '').trim().toUpperCase();
  if (book.id === DEFAULT_BOOK_ID || code === 'DOAN_TRUONG') {
    return {
      key: 'blue',
      iconBg: 'bg-sky-100',
      iconText: 'text-sky-700',
      badgeBg: 'bg-sky-50',
      badgeText: 'text-sky-800',
      badgeBorder: 'border-sky-200',
      dotColor: 'bg-sky-600',
    };
  }
  if (code === 'CHI_BO') {
    return {
      key: 'purple',
      iconBg: 'bg-purple-100',
      iconText: 'text-purple-700',
      badgeBg: 'bg-purple-50',
      badgeText: 'text-purple-800',
      badgeBorder: 'border-purple-200',
      dotColor: 'bg-purple-600',
    };
  }
  if (code === 'NHA_TRUONG') {
    return {
      key: 'teal',
      iconBg: 'bg-teal-100',
      iconText: 'text-teal-700',
      badgeBg: 'bg-teal-50',
      badgeText: 'text-teal-800',
      badgeBorder: 'border-teal-200',
      dotColor: 'bg-teal-600',
    };
  }
  return {
    key: 'neutral',
    iconBg: 'bg-slate-100',
    iconText: 'text-slate-600',
    badgeBg: 'bg-slate-100',
    badgeText: 'text-slate-700',
    badgeBorder: 'border-slate-200',
    dotColor: 'bg-slate-500',
  };
}

export interface BookDriveStatusBadge {
  kind: 'connected' | 'unconfigured' | 'error' | 'locked';
  label: string;
  className: string;
  dotColor: string;
}

export function getBookDriveStatusBadge(book: RegistryBook): BookDriveStatusBadge {
  if (!book.isActive) {
    return {
      kind: 'locked',
      label: 'Tạm khóa',
      className: 'bg-slate-100 text-slate-700 border-slate-300',
      dotColor: 'bg-slate-500',
    };
  }
  if (book.googleDriveConfig?.lastCheckStatus === 'error') {
    return {
      kind: 'error',
      label: 'Lỗi kết nối',
      className: 'bg-rose-50 text-rose-800 border-rose-200',
      dotColor: 'bg-rose-600',
    };
  }
  if (book.googleDriveConfig?.isConfigured) {
    return {
      kind: 'connected',
      label: 'Drive đã kết nối',
      className: 'bg-emerald-50 text-emerald-800 border-emerald-200',
      dotColor: 'bg-emerald-600',
    };
  }
  return {
    kind: 'unconfigured',
    label: 'Chưa cấu hình Drive',
    className: 'bg-amber-50 text-amber-800 border-amber-200',
    dotColor: 'bg-amber-500',
  };
}

export const PERMISSION_PRESETS: Array<{ label: string; perms: BookPermission[] }> = [
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

export const GOOGLE_APPS_SCRIPT_TEMPLATE = `function getOrCreateSheetWithHeaders(ss, name, headers, headerBg) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(headerBg || "#005BAA");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontWeight("bold");
  }
  return sheet;
}

function findRowBySystemId(sheet, docId, fallbackIdx) {
  var rows = sheet.getDataRange().getValues();
  if (!rows || rows.length <= 1) return -1;
  var headerRow = rows[0] || [];
  var idColIdx = headerRow.indexOf("Mã hệ thống");
  if (idColIdx === -1) idColIdx = fallbackIdx;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idColIdx]) === String(docId)) {
      return i + 1;
    }
  }
  return -1;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (data.bookCode || "DOAN_TRUONG").toString().trim();
    var deletedSheetName = (data.deletedSheetName || "VB đã xóa").toString().trim();

    var MAIN_HEADERS = [
      "STT", "Năm", "Ngày ban hành", "Ngày ký", "Nhóm VB", "Loại văn bản",
      "Số thứ tự", "Số/Ký hiệu", "Tên văn bản", "Trích yếu/Nội dung",
      "Người ký", "Chức vụ", "Người thao tác", "Trạng thái hồ sơ",
      "Trạng thái tệp Drive", "Tên tệp Drive", "ID tệp Drive", "Liên kết Drive",
      "Ghi chú", "Mã hệ thống", "Mã sổ", "Sổ", "Thời điểm cập nhật"
    ];

    var DELETED_HEADERS = [
      "STT", "Năm", "Ngày ban hành", "Ngày ký", "Nhóm VB", "Loại văn bản",
      "Số thứ tự", "Số/Ký hiệu", "Tên văn bản", "Trích yếu/Nội dung",
      "Người ký", "Chức vụ", "Người thao tác", "Trạng thái hồ sơ",
      "Trạng thái tệp Drive", "Tên tệp Drive", "ID tệp Drive", "Liên kết Drive",
      "Ghi chú", "Mã hệ thống", "Mã sổ", "Sổ",
      "Người xóa", "Thời điểm xóa", "Lý do xóa", "Thời điểm đồng bộ"
    ];

    var docNameVal = data.docName || data.title || "";
    var titleVal = data.title || "";
    var bookNameVal = data.bookName || sheetName;
    var nowStr = new Date().toLocaleString("vi-VN", { timeZone: "Asia/Ho_Chi_Minh" });
    var action = (data.action || "update").toString().trim();

    // Bước 1 xóa: Tạo/lấy sheet "VB đã xóa" và ghi dòng theo Mã hệ thống (không trùng lặp), giữ liên kết Drive
    if (action === "delete_append_to_deleted" || action === "delete") {
      var delSheet = getOrCreateSheetWithHeaders(ss, deletedSheetName, DELETED_HEADERS, "#991B1B");
      var delRowValues = [
        data.stt, data.year, data.issueDate || "", data.signDate || "",
        data.docCategory || "Văn bản đi", data.typeName || "", data.sequenceNumber,
        data.fullDocNumber || "", docNameVal, titleVal,
        data.signer || "", data.signerRole || "", data.issuedBy || "",
        "Đã xóa", data.driveStatus || "Không đính kèm",
        data.driveFileName || "", data.driveFileId || "", data.driveFileUrl || "",
        data.notes || "", data.id, sheetName, bookNameVal,
        data.deletedBy || "", data.deletedAtVietnam || nowStr, data.deleteReason || "", nowStr
      ];
      var existingDelRow = findRowBySystemId(delSheet, data.id, 19);
      if (existingDelRow > 0) {
        delSheet.getRange(existingDelRow, 1, 1, delRowValues.length).setValues([delRowValues]);
      } else {
        delSheet.appendRow(delRowValues);
      }
      var confirmedDelRow = findRowBySystemId(delSheet, data.id, 19);
      if (confirmedDelRow <= 0) {
        throw new Error("Không xác nhận được dòng đã ghi trong sheet " + deletedSheetName);
      }
      if (action === "delete_append_to_deleted") {
        return ContentService.createTextOutput(JSON.stringify({
          result: "success", action: action, confirmed: true, deletedSheet: deletedSheetName
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Bước 2 xóa: Chỉ xóa dòng khỏi sheet chính sau khi đã xác nhận có trong sheet "VB đã xóa"
    if (action === "delete_remove_main" || action === "delete") {
      var verifyDelSheet = ss.getSheetByName(deletedSheetName);
      var verifiedRow = verifyDelSheet ? findRowBySystemId(verifyDelSheet, data.id, 19) : -1;
      if (verifiedRow <= 0) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "error", error: "Chưa tìm thấy hồ sơ trong " + deletedSheetName + ", không xóa khỏi sheet chính."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var mainSheetForDelete = ss.getSheetByName(sheetName);
      if (mainSheetForDelete) {
        var rowToRemove = findRowBySystemId(mainSheetForDelete, data.id, 19);
        if (rowToRemove > 0) {
          mainSheetForDelete.deleteRow(rowToRemove);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        result: "success", action: action, removedFromMain: true, tab: sheetName
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getOrCreateSheetWithHeaders(ss, sheetName, MAIN_HEADERS, "#005BAA");

    // Thao tác kiểm tra kết nối theo đúng tab của sổ
    if (action === "test") {
      sheet.appendRow([
        0, data.year, data.issueDate, data.signDate || "", data.docCategory || "Văn bản đi",
        data.typeName, 0, data.fullDocNumber, docNameVal, titleVal,
        data.signer, data.signerRole, data.issuedBy, data.status,
        data.driveStatus || "Không đính kèm", data.driveFileName || "", data.driveFileId || "",
        data.driveFileUrl || "", data.notes || "", data.id, sheetName, bookNameVal, nowStr
      ]);
      return ContentService.createTextOutput(JSON.stringify({result: "success", action: "test", tab: sheetName}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var foundRowIndex = findRowBySystemId(sheet, data.id, 19);
    var rowValues = [
      data.stt, data.year, data.issueDate, data.signDate || "",
      data.docCategory || "Văn bản đi", data.typeName, data.sequenceNumber,
      data.fullDocNumber, docNameVal, titleVal, data.signer, data.signerRole,
      data.issuedBy, data.status, data.driveStatus || "Không đính kèm",
      data.driveFileName || "", data.driveFileId || "", data.driveFileUrl || "",
      data.notes || "", data.id, sheetName, bookNameVal, nowStr
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

export const GOOGLE_DRIVE_APPS_SCRIPT_TEMPLATE = `function getOrCreateSubFolder(parent, name) {
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
