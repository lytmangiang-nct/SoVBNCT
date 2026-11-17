import type {
  AttachmentGroup,
  DocumentAttachment,
  DocumentAttachmentVersion,
  DocumentRecord,
} from '../types.js';

export const ATTACHMENT_GROUP_LABELS: Record<AttachmentGroup, string> = {
  word_draft: 'Word soạn thảo',
  signed_pdf: 'PDF đã ký',
  appendix: 'Phụ lục',
};

export function inferLegacyAttachmentGroup(doc: DocumentRecord): AttachmentGroup {
  const fileName = (doc.driveFileName || doc.driveOriginalFileName || '').toLowerCase();
  if (doc.docCategory === 'incoming' || fileName.endsWith('.pdf')) {
    return 'signed_pdf';
  }
  return 'word_draft';
}

/**
 * Chuẩn hóa danh sách tệp đính kèm theo nhóm cho cả hồ sơ mới và hồ sơ cũ
 * (vốn chỉ lưu 1 liên kết ở root fields driveFileId / driveFileUrl / driveStatus).
 */
export function getEffectiveAttachments(doc: DocumentRecord): DocumentAttachment[] {
  if (Array.isArray(doc.attachments) && doc.attachments.length > 0) {
    return doc.attachments.map((att) => ({
      ...att,
      versions: Array.isArray(att.versions) ? att.versions : [],
    }));
  }

  const hasLegacyRoot =
    Boolean(doc.driveFileId) ||
    Boolean(doc.driveFileUrl) ||
    doc.driveStatus === 'saved' ||
    doc.driveStatus === 'failed' ||
    doc.driveStatus === 'uploading';

  if (!hasLegacyRoot) {
    return [];
  }

  const group = inferLegacyAttachmentGroup(doc);
  const fileName = doc.driveFileName || doc.driveOriginalFileName || `${doc.fullDocNumber || 'van-ban'}`;
  const uploadedAt = doc.driveUploadedAt || doc.createdAt || new Date().toISOString();
  const status = doc.driveStatus || (doc.driveFileUrl ? 'saved' : 'none');

  const legacyVersion: DocumentAttachmentVersion = {
    version: 1,
    fileName,
    originalFileName: doc.driveOriginalFileName,
    mimeType: doc.driveMimeType,
    fileSize: doc.driveFileSize,
    sha256: doc.fileSha256,
    driveFileId: doc.driveFileId,
    driveFileUrl: doc.driveFileUrl,
    driveFolderId: doc.driveFolderId,
    driveFolderPath: doc.driveFolderPath,
    driveStatus: status,
    driveError: doc.driveError,
    uploadedBy: doc.issuedBy || 'Hệ thống',
    uploadedByUsername: doc.issuedByUsername,
    uploadedAt,
  };

  return [
    {
      id: `legacy-${doc.id}`,
      group,
      label: group === 'signed_pdf' ? 'Bản PDF gốc' : 'Bản Word chính',
      isOfficial: true,
      currentVersion: 1,
      fileName,
      originalFileName: doc.driveOriginalFileName,
      mimeType: doc.driveMimeType,
      fileSize: doc.driveFileSize,
      sha256: doc.fileSha256,
      driveFileId: doc.driveFileId,
      driveFileUrl: doc.driveFileUrl,
      driveFolderId: doc.driveFolderId,
      driveFolderPath: doc.driveFolderPath,
      driveStatus: status,
      driveError: doc.driveError,
      uploadedBy: doc.issuedBy,
      uploadedByUsername: doc.issuedByUsername,
      uploadedAt,
      versions: [legacyVersion],
    },
  ];
}

/**
 * Chọn tệp đại diện chính thức cho nút "Mở tệp":
 * Ưu tiên tệp gắn cờ isOfficial đang ở trạng thái saved và có driveFileUrl,
 * kế tiếp là bất kỳ tệp saved nào có driveFileUrl.
 */
export function getOfficialAttachment(doc: DocumentRecord): DocumentAttachment | null {
  const attachments = getEffectiveAttachments(doc);
  if (attachments.length === 0) return null;

  const officialSaved = attachments.find(
    (a) => a.isOfficial && a.driveStatus === 'saved' && Boolean(a.driveFileUrl)
  );
  if (officialSaved) return officialSaved;

  const anySaved = attachments.find((a) => a.driveStatus === 'saved' && Boolean(a.driveFileUrl));
  if (anySaved) return anySaved;

  return attachments.find((a) => a.isOfficial) || attachments[0] || null;
}

/**
 * Đồng bộ ngược các trường cache root (driveStatus, driveFileId, driveFileUrl, driveFileName...)
 * từ danh sách attachments để tương thích 100% với bảng danh sách, Excel xuất ra và Google Sheets.
 */
export function syncRootDriveFieldsFromAttachments(doc: DocumentRecord): void {
  if (!Array.isArray(doc.attachments) || doc.attachments.length === 0) {
    return;
  }

  const official = getOfficialAttachment(doc);
  if (!official) return;

  doc.driveStatus = official.driveStatus;
  doc.driveFileId = official.driveFileId;
  doc.driveFileName = official.fileName;
  doc.driveFileUrl = official.driveFileUrl;
  doc.driveOriginalFileName = official.originalFileName;
  doc.driveMimeType = official.mimeType;
  doc.driveFileSize = official.fileSize;
  doc.driveFolderId = official.driveFolderId;
  doc.driveFolderPath = official.driveFolderPath;
  doc.driveUploadedAt = official.uploadedAt;
  if (official.driveError) {
    doc.driveError = official.driveError;
  } else {
    delete doc.driveError;
  }
}
