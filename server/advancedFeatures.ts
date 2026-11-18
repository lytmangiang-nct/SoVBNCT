import crypto from 'crypto';
import ExcelJS from 'exceljs';
import { GoogleGenAI, Type } from '@google/genai';
import type {
  ArchiveDriveCandidate,
  ArchiveImportColumnMapping,
  ArchiveImportEditableRow,
  ArchiveImportPreviewResponse,
  ArchiveImportPreviewRow,
  ArchiveParseExcelResponse,
  ClosingSnapshotDocument,
  ClosingSummaryCounts,
  DocumentQaCitation,
  DocumentQaResponse,
  DocumentRecord,
  DocumentTypeConfig,
  RegistryBook,
  YearEndClosingSnapshot,
} from '../src/types.js';
import {
  generateSuggestedFileName,
  getDocumentFilterYear,
  getSupportedAttachmentExtension,
  normalizeCompareText,
  normalizeDateOnlyDdMmYyyy,
  normalizeDocNumberForCompare,
  removeVietnameseDiacritics,
} from '../src/utils/fileNaming.js';
import { getEffectiveAttachments, getOfficialAttachment } from '../src/utils/attachments.js';
import { getDocBookId } from '../src/utils/permissions.js';
import {
  extractGoogleDriveFileId,
  getDriveFileMetadata,
  listDriveFilesInFolder,
  type ResolvedGoogleDriveConfig,
} from './googleDrive.js';

// =========================================================
// FEATURE 1: GROUNDED DOCUMENT Q&A (CHAT TRA CỨU CÓ DẪN NGUỒN)
// =========================================================

interface DocTypeIntentRule {
  codes: string[];
  label: string;
  phrases: string[];
}

const DOC_TYPE_INTENT_RULES: DocTypeIntentRule[] = [
  { codes: ['QD', 'QĐ'], label: 'Quyết định', phrases: ['quyet dinh', 'qd'] },
  { codes: ['TB'], label: 'Thông báo', phrases: ['thong bao'] },
  { codes: ['KH'], label: 'Kế hoạch', phrases: ['ke hoach'] },
  { codes: ['BC'], label: 'Báo cáo', phrases: ['bao cao'] },
  { codes: ['CV'], label: 'Công văn', phrases: ['cong van'] },
  { codes: ['DN', 'ĐN'], label: 'Đề nghị', phrases: ['de nghi', 'to trinh de nghi'] },
  { codes: ['TTR'], label: 'Tờ trình', phrases: ['to trinh'] },
  { codes: ['BB'], label: 'Biên bản', phrases: ['bien ban'] },
  { codes: ['HD'], label: 'Hướng dẫn', phrases: ['huong dan'] },
];

const QA_STOP_WORDS = new Set([
  'la',
  'gi',
  'nao',
  'cua',
  've',
  'cho',
  'trong',
  'co',
  'khong',
  'duoc',
  'nhung',
  'cac',
  'mot',
  'ngay',
  'thang',
  'nam',
  'van',
  'ban',
  'so',
  'tai',
  'noi',
  'dung',
  'hoi',
  'tim',
  'xem',
  'hay',
  'voi',
  'va',
  'hoac',
  'nhu',
  'the',
  'ai',
  'ky',
  'nguoi',
  'ban',
  'hanh',
  'co',
  'quan',
  'thuoc',
  'so',
  'xin',
  'chao',
  'biet',
  'tong',
  'hop',
]);

function detectRequestedDocTypeIntent(question: string): DocTypeIntentRule | null {
  const norm = ` ${normalizeCompareText(question)} `;
  // Check explicit type phrases in order of specificity
  for (const rule of DOC_TYPE_INTENT_RULES) {
    for (const phrase of rule.phrases) {
      if (norm.includes(` ${phrase} `)) {
        return rule;
      }
    }
  }
  return null;
}

function docMatchesTypeIntent(doc: DocumentRecord, rule: DocTypeIntentRule): boolean {
  const normTypeCode = removeVietnameseDiacritics(doc.typeCode || '').toUpperCase().trim();
  if (rule.codes.includes(normTypeCode)) return true;

  const normTypeName = normalizeCompareText(doc.typeName || '');
  const normTitle = normalizeCompareText(`${doc.docName || ''} ${doc.title || ''}`);
  const normNum = removeVietnameseDiacritics(doc.fullDocNumber || '').toUpperCase();

  for (const phrase of rule.phrases) {
    if (normTypeName.includes(phrase)) return true;
    if (normTitle.startsWith(phrase) || normTitle.includes(` ${phrase} `)) return true;
  }
  for (const code of rule.codes) {
    if (new RegExp(`[-/]${code}(?:[-/]|$)`).test(normNum)) return true;
  }
  return false;
}

function extractQueryTerms(question: string): {
  tokens: string[];
  numbers: string[];
  normQuestion: string;
} {
  const normQuestion = normalizeCompareText(question);
  const rawParts = normQuestion.split(' ').filter(Boolean);
  const tokens: string[] = [];
  const numbers: string[] = [];

  for (const part of rawParts) {
    if (/^\d+$/.test(part) || /^\d+\/\d+$/.test(part)) {
      numbers.push(part);
    }
    if (part.length >= 2 && !QA_STOP_WORDS.has(part)) {
      tokens.push(part);
    }
  }

  // Also grab slash-delimited doc number fragments from raw question (e.g. 09-KH/ĐTN or 20/11)
  const docNumChunks = question.match(/[0-9A-Za-zÀ-ỹ]+(?:[\/\-][0-9A-Za-zÀ-ỹ]+)+/g) || [];
  for (const chunk of docNumChunks) {
    if (!/\d/.test(chunk)) continue;
    const clean = normalizeDocNumberForCompare(chunk);
    if (clean) numbers.push(clean);
  }

  return {
    tokens: Array.from(new Set(tokens)),
    numbers: Array.from(new Set(numbers)),
    normQuestion,
  };
}

function buildRecordMetadataSummary(doc: DocumentRecord, bookName: string): string {
  const parts: string[] = [];
  parts.push(`${doc.docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi'} (${doc.typeName || doc.typeCode})`);
  parts.push(`Số/Ký hiệu: ${doc.fullDocNumber}`);
  parts.push(`Tên văn bản: ${doc.docName || doc.title}`);
  if (doc.title && doc.title !== doc.docName) {
    parts.push(`Trích yếu: ${doc.title}`);
  }
  if (doc.issueDate) parts.push(`Ngày ban hành: ${doc.issueDate}`);
  if (doc.issuingAgency) parts.push(`Cơ quan ban hành: ${doc.issuingAgency}`);
  if (doc.signer) parts.push(`Người ký: ${doc.signer}${doc.signerRole ? ` (${doc.signerRole})` : ''}`);
  parts.push(`Sổ: ${bookName}`);
  if (doc.notes) parts.push(`Ghi chú: ${doc.notes}`);
  return parts.join(' · ');
}

function extractBestPageExcerpt(pageText: string, tokens: string[], normQuestion: string): {
  score: number;
  matchedCount: number;
  excerpt: string;
} {
  const raw = (pageText || '').replace(/\s+/g, ' ').trim();
  if (!raw) return { score: 0, matchedCount: 0, excerpt: '' };
  const normPage = normalizeCompareText(raw);
  const pageWords = new Set(normPage.split(' ').filter(Boolean));
  const paddedNormPage = ` ${normPage} `;

  let score = 0;

  if (normQuestion.length >= 8 && normPage.includes(normQuestion)) {
    score += 12;
  }

  // Bi-gram phrase bonus
  for (let i = 0; i < tokens.length - 1; i++) {
    const bigram = `${tokens[i]} ${tokens[i + 1]}`;
    if (paddedNormPage.includes(` ${bigram} `)) {
      score += 4.5;
    }
  }

  let matchedTokenCount = 0;
  let firstMatchedToken = '';
  for (const tk of tokens) {
    if (pageWords.has(tk)) {
      matchedTokenCount++;
      if (!firstMatchedToken) firstMatchedToken = tk;
      score += tk.length >= 4 ? 3 : 1.8;
    }
  }

  if (matchedTokenCount === 0) return { score: 0, matchedCount: 0, excerpt: '' };

  // Find actual position in raw text using whole-word position in diacritic-stripped text
  let rawStart = 0;
  if (firstMatchedToken) {
    const strippedLower = removeVietnameseDiacritics(raw).toLowerCase();
    const matchRegex = new RegExp(`(?:^|[^0-9a-z])${firstMatchedToken.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?:$|[^0-9a-z])`);
    const m = matchRegex.exec(strippedLower);
    if (m && typeof m.index === 'number') {
      rawStart = Math.max(0, m.index - 70);
    }
  }
  const rawEnd = Math.min(raw.length, rawStart + 260);
  let excerpt = raw.slice(rawStart, rawEnd).trim();
  if (rawStart > 0) excerpt = `…${excerpt}`;
  if (rawEnd < raw.length) excerpt = `${excerpt}…`;

  return { score, matchedCount: matchedTokenCount, excerpt };
}

interface CandidateDocMatch {
  doc: DocumentRecord;
  book: RegistryBook;
  score: number;
  metadataScore: number;
  pdfScore: number;
  bestPage?: { page: number; excerpt: string; score: number };
  topPages: Array<{ page: number; excerpt: string; score: number }>;
  metadataSummary: string;
  attachmentId?: string;
  hasPdfFile: boolean;
  hasDriveFile: boolean;
  matchesTypeIntent: boolean;
}

export async function answerDocumentQuestion(params: {
  question: string;
  permittedDocs: DocumentRecord[];
  booksById: Map<string, RegistryBook>;
  searchedBooks: Array<{ id: string; code: string; name: string }>;
}): Promise<DocumentQaResponse> {
  const question = (params.question || '').trim();
  const { tokens, numbers, normQuestion } = extractQueryTerms(question);
  const typeIntent = detectRequestedDocTypeIntent(question);
  const intentPhraseSet = new Set(typeIntent ? typeIntent.phrases.flatMap((p) => p.split(' ')) : []);
  const contentTokens = tokens.filter((tk) => !intentPhraseSet.has(tk));
  const evalTokens = contentTokens.length > 0 ? contentTokens : tokens;

  // Filter out deleted unless already scoped by caller
  const docs = params.permittedDocs.filter((d) => d.status !== 'deleted');
  let indexedPdfCount = 0;
  for (const d of docs) {
    if (Array.isArray(d.pdfPageContents) && d.pdfPageContents.some((p) => (p.text || '').trim().length >= 10)) {
      indexedPdfCount++;
    }
  }

  const candidates: CandidateDocMatch[] = [];

  for (const doc of docs) {
    const book = params.booksById.get(getDocBookId(doc));
    if (!book) continue;

    const metaText = [
      doc.fullDocNumber,
      doc.docName,
      doc.title,
      doc.issuingAgency,
      doc.parentAgency,
      doc.signer,
      doc.signerRole,
      doc.typeName,
      doc.typeCode,
      doc.issueDate,
      doc.notes,
    ]
      .filter(Boolean)
      .join(' ');
    const normMeta = normalizeCompareText(metaText);
    const metaWords = new Set(normMeta.split(' ').filter(Boolean));
    const paddedNormMeta = ` ${normMeta} `;
    const normDocNum = normalizeDocNumberForCompare(doc.fullDocNumber);

    let metadataScore = 0;
    let exactNumberMatched = false;
    const matchedDocTokens = new Set<string>();

    for (const numToken of numbers) {
      if (numToken.length >= 2 && (normDocNum.includes(numToken) || metaWords.has(numToken))) {
        exactNumberMatched = true;
        metadataScore += normDocNum === numToken ? 16 : 6;
      }
    }
    for (const tk of evalTokens) {
      if (metaWords.has(tk)) {
        matchedDocTokens.add(tk);
        metadataScore += tk.length >= 4 ? 3 : 1.8;
      }
    }
    for (let i = 0; i < evalTokens.length - 1; i++) {
      if (paddedNormMeta.includes(` ${evalTokens[i]} ${evalTokens[i + 1]} `)) {
        metadataScore += 4.5;
      }
    }

    const pages = Array.isArray(doc.pdfPageContents) ? doc.pdfPageContents : [];
    const pageMatches: Array<{ page: number; excerpt: string; score: number }> = [];
    for (const p of pages) {
      if (!p.text || p.text.trim().length < 5) continue;
      const res = extractBestPageExcerpt(p.text, evalTokens, normQuestion);
      if (res.score > 0) {
        const pWords = new Set(normalizeCompareText(p.text).split(' ').filter(Boolean));
        evalTokens.forEach((tk) => {
          if (pWords.has(tk)) matchedDocTokens.add(tk);
        });
        pageMatches.push({ page: p.page, excerpt: res.excerpt, score: res.score });
      }
    }
    pageMatches.sort((a, b) => b.score - a.score);
    const bestPage = pageMatches[0];
    const pdfScore = bestPage ? bestPage.score : 0;

    // Require minimum meaningful token coverage so single incidental word overlap doesn't trigger false answers
    const minTokensRequired = evalTokens.length >= 3 ? 2 : 1;
    if (!exactNumberMatched && matchedDocTokens.size < minTokensRequired) {
      continue;
    }

    const totalScore = metadataScore + pdfScore;
    if (!exactNumberMatched && totalScore < (evalTokens.length >= 2 ? 5.0 : 2.5)) {
      continue;
    }

    const matchesType = typeIntent ? docMatchesTypeIntent(doc, typeIntent) : true;
    const attachments = getEffectiveAttachments(doc);
    const officialAtt = getOfficialAttachment(doc);
    const pdfAtt =
      attachments.find(
        (a) =>
          a.driveStatus === 'saved' &&
          (a.group === 'signed_pdf' || a.fileName.toLowerCase().endsWith('.pdf'))
      ) ||
      (officialAtt?.driveStatus === 'saved' ? officialAtt : undefined);

    const hasDriveFile =
      (doc.driveStatus === 'saved' && Boolean(doc.driveFileUrl || doc.driveFileId)) ||
      attachments.some((a) => a.driveStatus === 'saved' && Boolean(a.driveFileUrl));
    const hasPdfFile =
      hasDriveFile &&
      (Boolean(pdfAtt) ||
        (doc.driveFileName || doc.driveOriginalFileName || '').toLowerCase().endsWith('.pdf') ||
        doc.docCategory === 'incoming' ||
        pages.length > 0);

    candidates.push({
      doc,
      book,
      score: totalScore + (matchesType ? 3 : 0),
      metadataScore,
      pdfScore,
      bestPage,
      topPages: pageMatches.slice(0, 2),
      metadataSummary: buildRecordMetadataSummary(doc, book.name),
      attachmentId: pdfAtt?.id,
      hasPdfFile,
      hasDriveFile,
      matchesTypeIntent: matchesType,
    });
  }

  candidates.sort((a, b) => b.score - a.score);
  const topCandidates = candidates.slice(0, 10);

  if (topCandidates.length === 0) {
    return {
      question,
      answer: 'Chưa tìm thấy văn bản đủ căn cứ phù hợp với câu hỏi trong các sổ bạn có quyền xem.',
      insufficientEvidence: true,
      citations: [],
      searchedBooks: params.searchedBooks,
      usedAi: false,
      indexedPdfCount,
      totalSearchedDocs: docs.length,
    };
  }

  // Enforce document-type honesty BEFORE & AFTER AI (e.g. question asks for "Quyết định", only "Thông báo" exists)
  const typeMatchedCandidates = typeIntent
    ? topCandidates.filter((c) => c.matchesTypeIntent)
    : topCandidates;

  const hasTypeMismatchOnly =
    Boolean(typeIntent) && typeMatchedCandidates.length === 0 && topCandidates.length > 0;

  // Build citations helper
  const toCitation = (
    c: CandidateDocMatch,
    preferredSourceType?: 'pdf_page' | 'record_metadata',
    preferredPage?: number,
    preferredExcerpt?: string
  ): DocumentQaCitation => {
    const canUsePdf =
      c.topPages.length > 0 && (preferredSourceType === 'pdf_page' || (!preferredSourceType && c.pdfScore >= 2.5));
    if (canUsePdf) {
      const matchedPageObj =
        (preferredPage ? c.topPages.find((p) => p.page === preferredPage) : undefined) || c.bestPage!;
      return {
        docId: c.doc.id,
        bookId: c.book.id,
        bookCode: c.book.code,
        bookName: c.book.name,
        docCategory: c.doc.docCategory,
        typeCode: c.doc.typeCode,
        typeName: c.doc.typeName || c.doc.typeCode,
        fullDocNumber: c.doc.fullDocNumber,
        docName: c.doc.docName || c.doc.title,
        issueDate: c.doc.issueDate,
        issuingAgency: c.doc.issuingAgency,
        sourceType: 'pdf_page',
        page: matchedPageObj.page,
        excerpt: preferredExcerpt?.trim() || matchedPageObj.excerpt,
        attachmentId: c.attachmentId,
        hasPdfFile: c.hasPdfFile,
        hasDriveFile: c.hasDriveFile,
      };
    }

    return {
      docId: c.doc.id,
      bookId: c.book.id,
      bookCode: c.book.code,
      bookName: c.book.name,
      docCategory: c.doc.docCategory,
      typeCode: c.doc.typeCode,
      typeName: c.doc.typeName || c.doc.typeCode,
      fullDocNumber: c.doc.fullDocNumber,
      docName: c.doc.docName || c.doc.title,
      issueDate: c.doc.issueDate,
      issuingAgency: c.doc.issuingAgency,
      sourceType: 'record_metadata',
      excerpt: preferredExcerpt?.trim() || c.metadataSummary,
      attachmentId: c.attachmentId,
      hasPdfFile: c.hasPdfFile,
      hasDriveFile: c.hasDriveFile,
    };
  };

  if (hasTypeMismatchOnly && typeIntent) {
    const closest = topCandidates.slice(0, 3).map((c) => toCitation(c));
    const foundTypes = Array.from(new Set(closest.map((c) => c.typeName))).join(', ');
    return {
      question,
      answer: `Chưa tìm thấy văn bản loại "${typeIntent.label}" đủ căn cứ cho câu hỏi này trong phạm vi được phép xem. Hệ thống chỉ tìm thấy ${closest.length} văn bản liên quan thuộc loại ${foundTypes} để bạn đối chiếu.`,
      insufficientEvidence: true,
      mismatchNote: `Câu hỏi yêu cầu "${typeIntent.label}", nhưng hồ sơ liên quan tìm thấy chỉ là "${foundTypes}" — không phải ${typeIntent.label}.`,
      citations: closest,
      searchedBooks: params.searchedBooks,
      usedAi: false,
      indexedPdfCount,
      totalSearchedDocs: docs.length,
    };
  }

  // Try Gemini API server-side when GEMINI_API_KEY is available
  const apiKey = (process.env.GEMINI_API_KEY || '').trim();
  if (apiKey) {
    try {
      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: { headers: { 'User-Agent': 'aistudio-build' } },
      });

      const evidencePayload = topCandidates.map((c) => ({
        docId: c.doc.id,
        bookName: c.book.name,
        docCategory: c.doc.docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi',
        typeName: c.doc.typeName || c.doc.typeCode,
        fullDocNumber: c.doc.fullDocNumber,
        docName: c.doc.docName || c.doc.title,
        issueDate: c.doc.issueDate || '',
        issuingAgency: c.doc.issuingAgency || '',
        signer: c.doc.signer || '',
        metadataSummary: c.metadataSummary,
        pdfExcerpts: c.topPages.map((p) => ({ page: p.page, text: p.excerpt })),
      }));

      const prompt = [
        'Bạn là trợ lý tra cứu sổ văn bản hành chính của Trường THPT Nguyễn Chí Thanh.',
        'CHỈ ĐƯỢC TRẢ LỜI DỰA TRÊN CÁC HỒ SƠ VÀ ĐOẠN TRÍCH ĐƯỢC CUNG CẤP DƯỚI ĐÂY. Tuyệt đối không suy diễn hoặc bịa thông tin ngoài nguồn.',
        'Quy tắc bắt buộc:',
        '1. Kiểm tra đúng loại văn bản người dùng hỏi (ví dụ: hỏi "Quyết định" thì KHÔNG được khẳng định một "Thông báo" hay "Kế hoạch" là Quyết định; nếu chỉ thấy Thông báo thì đặt insufficientEvidence = true và giải thích rõ sự khác biệt loại văn bản trong mismatchNote).',
        '2. Mỗi căn cứ trích dẫn phải ghi đúng docId, sourceType ("pdf_page" nếu lấy từ nội dung pdfExcerpts có số trang, hoặc "record_metadata" nếu chỉ lấy từ metadataSummary của bảng kê) và trích đoạn chính xác.',
        '3. Nếu không đủ bằng chứng trả lời trực tiếp câu hỏi, đặt insufficientEvidence = true và trả lời rõ "Chưa tìm thấy văn bản đủ căn cứ".',
        `Câu hỏi của người dùng: "${question}"`,
        `Danh sách hồ sơ hợp lệ theo quyền xem:\n${JSON.stringify(evidencePayload, null, 2)}`,
      ].join('\n\n');

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: prompt,
        config: {
          abortSignal: AbortSignal.timeout(8000),
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              answer: { type: Type.STRING },
              insufficientEvidence: { type: Type.BOOLEAN },
              mismatchNote: { type: Type.STRING },
              citedItems: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    docId: { type: Type.STRING },
                    sourceType: { type: Type.STRING },
                    page: { type: Type.INTEGER },
                    excerpt: { type: Type.STRING },
                  },
                  required: ['docId', 'sourceType', 'excerpt'],
                },
              },
            },
            required: ['answer', 'insufficientEvidence', 'citedItems'],
          },
        },
      });

      const rawText = response.text || '';
      const parsed = JSON.parse(rawText);
      const candidateById = new Map(topCandidates.map((c) => [c.doc.id, c]));
      const verifiedCitations: DocumentQaCitation[] = [];
      const seenKeys = new Set<string>();

      if (Array.isArray(parsed.citedItems)) {
        for (const item of parsed.citedItems) {
          const cand = candidateById.get(String(item.docId || ''));
          if (!cand) continue;
          const requestedSource: 'pdf_page' | 'record_metadata' =
            item.sourceType === 'pdf_page' && cand.topPages.length > 0 ? 'pdf_page' : 'record_metadata';
          const citation = toCitation(cand, requestedSource, item.page, item.excerpt);
          const key = `${citation.docId}:${citation.sourceType}:${citation.page || 0}`;
          if (!seenKeys.has(key)) {
            seenKeys.add(key);
            verifiedCitations.push(citation);
          }
        }
      }

      if (verifiedCitations.length === 0 && !parsed.insufficientEvidence) {
        for (const c of typeMatchedCandidates.slice(0, 3)) {
          verifiedCitations.push(toCitation(c));
        }
      }

      return {
        question,
        answer:
          String(parsed.answer || '').trim() ||
          'Chưa tìm thấy văn bản đủ căn cứ phù hợp với câu hỏi.',
        insufficientEvidence: Boolean(parsed.insufficientEvidence),
        mismatchNote: parsed.mismatchNote ? String(parsed.mismatchNote).trim() : undefined,
        citations: verifiedCitations,
        searchedBooks: params.searchedBooks,
        usedAi: true,
        indexedPdfCount,
        totalSearchedDocs: docs.length,
      };
    } catch (aiErr) {
      console.warn('Gemini QA fallback to deterministic grounded answer:', aiErr);
    }
  }

  // Deterministic grounded answer (used when GEMINI_API_KEY is absent or fails)
  const chosen = (typeMatchedCandidates.length > 0 ? typeMatchedCandidates : topCandidates).slice(0, 4);
  const citations = chosen.map((c) => toCitation(c));
  const best = citations[0];
  const hasStrongMatch = chosen[0].score >= 4.5;

  if (!hasStrongMatch) {
    return {
      question,
      answer: 'Chưa tìm thấy văn bản đủ căn cứ trả lời trực tiếp câu hỏi này. Dưới đây là các hồ sơ gần nhất trong sổ để bạn kiểm tra:',
      insufficientEvidence: true,
      citations,
      searchedBooks: params.searchedBooks,
      usedAi: false,
      indexedPdfCount,
      totalSearchedDocs: docs.length,
    };
  }

  const lines = citations.map((c, idx) => {
    const sourceDesc =
      c.sourceType === 'pdf_page'
        ? `nội dung PDF (Trang ${c.page})`
        : 'thông tin hồ sơ (bảng kê)';
    return `${idx + 1}. ${c.typeName} số ${c.fullDocNumber} — "${c.docName}" (${sourceDesc}${c.issueDate ? `, ngày ${c.issueDate}` : ''}).`;
  });

  return {
    question,
    answer: `Tìm thấy ${citations.length} văn bản liên quan trong phạm vi sổ được phép xem (nổi bật nhất: ${best.typeName} số ${best.fullDocNumber}):\n${lines.join('\n')}`,
    insufficientEvidence: false,
    citations,
    searchedBooks: params.searchedBooks,
    usedAi: false,
    indexedPdfCount,
    totalSearchedDocs: docs.length,
  };
}

// =========================================================
// FEATURE 2: HISTORICAL ARCHIVE IMPORT FROM EXCEL & DRIVE
// =========================================================

const TEMPLATE_HEADERS = [
  'Nhóm văn bản',
  'Số/Ký hiệu',
  'Tên văn bản',
  'Ngày ban hành',
  'Năm',
  'Cơ quan ban hành',
  'Người ký',
  'Liên kết hoặc tên tệp Drive',
  'Loại văn bản',
  'Chức vụ người ký',
  'Cơ quan chủ quản',
  'Trích yếu',
  'Ghi chú',
];

export async function buildArchiveTemplateWorkbookBuffer(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Sổ văn bản Trường THPT Nguyễn Chí Thanh';
  wb.created = new Date();

  const sheet = wb.addWorksheet('Danh sách nhập kho', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  sheet.columns = [
    { key: 'docCategory', width: 18 },
    { key: 'fullDocNumber', width: 24, style: { numFmt: '@' } },
    { key: 'docName', width: 42 },
    { key: 'issueDate', width: 16, style: { numFmt: '@' } },
    { key: 'year', width: 10 },
    { key: 'issuingAgency', width: 30 },
    { key: 'signer', width: 24 },
    { key: 'driveLinkOrFileName', width: 42 },
    { key: 'typeCode', width: 15 },
    { key: 'signerRole', width: 22 },
    { key: 'parentAgency', width: 26 },
    { key: 'title', width: 34 },
    { key: 'notes', width: 26 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.values = TEMPLATE_HEADERS;
  headerRow.height = 26;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0369A1' } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });

  const sampleRow1 = sheet.addRow({
    docCategory: 'Văn bản đi',
    fullDocNumber: '007-KH/ĐTN',
    docName: 'Kế hoạch tổ chức hoạt động chào mừng 20/11 năm 2024',
    issueDate: '05/11/2024',
    year: 2024,
    issuingAgency: 'Đoàn Trường THPT Nguyễn Chí Thanh',
    signer: 'Nguyễn Văn Minh',
    driveLinkOrFileName: '007-KH.ĐTN - Kế hoạch tổ chức hoạt động chào mừng 20-11 năm 2024.pdf',
    typeCode: 'KH',
    signerRole: 'Bí thư Đoàn trường',
    parentAgency: '',
    title: 'Kế hoạch hội diễn và thi đua chào mừng Ngày Nhà giáo Việt Nam 20/11',
    notes: 'Hồ sơ lưu trữ năm 2024',
  });
  sampleRow1.getCell('fullDocNumber').numFmt = '@';
  sampleRow1.getCell('issueDate').numFmt = '@';

  const sampleRow2 = sheet.addRow({
    docCategory: 'Văn bản đến',
    fullDocNumber: '0124/SGDĐT-GDTrH',
    docName: 'Công văn hướng dẫn tổ chức kỳ thi học sinh giỏi cấp tỉnh',
    issueDate: '12/10/2024',
    year: 2024,
    issuingAgency: 'Sở GDĐT TP.HCM',
    signer: 'Trần Anh Tuấn',
    driveLinkOrFileName: 'https://drive.google.com/file/d/EXAMPLE_FILE_ID_123456/view',
    typeCode: 'CV',
    signerRole: 'Phó Giám đốc',
    parentAgency: 'UBND TP.HCM',
    title: 'Hướng dẫn công tác coi thi và chấm thi',
    notes: '',
  });
  sampleRow2.getCell('fullDocNumber').numFmt = '@';
  sampleRow2.getCell('issueDate').numFmt = '@';

  const guide = wb.addWorksheet('Hướng dẫn');
  guide.columns = [{ width: 28 }, { width: 72 }];
  guide.addRow(['Cột dữ liệu', 'Quy tắc nhập kho văn bản cũ']).font = { bold: true };
  guide.addRow(['Nhóm văn bản', 'Nhập "Văn bản đi" hoặc "Văn bản đến" (bắt buộc).']);
  guide.addRow(['Số/Ký hiệu', 'Định dạng văn bản (Text) để giữ nguyên số 0 ở đầu như 007/QĐ-...']);
  guide.addRow(['Ngày ban hành', 'Nhập dd/mm/yyyy (ví dụ 05/11/2024). Nếu cột Năm khác năm của Ngày ban hành, hệ thống sẽ cảnh báo.']);
  guide.addRow([
    'Liên kết hoặc tên tệp Drive',
    'Có thể dán URL tệp Google Drive, ID tệp Drive, tên tệp chính xác trong thư mục đã chọn, hoặc để trống để hệ thống tự ghép theo tên chuẩn "<Số ký hiệu> - <Tên văn bản>".',
  ]);

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out as ArrayBuffer);
}

function formatDateDdMmYyyyUtc(d: Date): string {
  const day = String(d.getUTCDate()).padStart(2, '0');
  const month = String(d.getUTCMonth() + 1).padStart(2, '0');
  const year = d.getUTCFullYear();
  return `${day}/${month}/${year}`;
}

function extractExcelCellTextAndLink(cell: ExcelJS.Cell): { value: string; hyperlink?: string } {
  const rawVal = cell.value;
  let hyperlink: string | undefined;

  if (cell.hyperlink && typeof cell.hyperlink === 'string') {
    hyperlink = cell.hyperlink.trim();
  }

  if (rawVal === null || rawVal === undefined) {
    return { value: '', hyperlink };
  }

  if (rawVal instanceof Date) {
    return { value: formatDateDdMmYyyyUtc(rawVal), hyperlink };
  }

  if (typeof rawVal === 'object') {
    const obj: any = rawVal;
    if (typeof obj.hyperlink === 'string') {
      hyperlink = obj.hyperlink.trim();
    }
    if (obj.text !== undefined) {
      if (typeof obj.text === 'object' && Array.isArray(obj.text.richText)) {
        return {
          value: obj.text.richText.map((r: any) => r.text || '').join('').trim(),
          hyperlink,
        };
      }
      return { value: String(obj.text || '').trim(), hyperlink };
    }
    if (Array.isArray(obj.richText)) {
      return {
        value: obj.richText.map((r: any) => r.text || '').join('').trim(),
        hyperlink,
      };
    }
    if (obj.result !== undefined) {
      if (obj.result instanceof Date) {
        return { value: formatDateDdMmYyyyUtc(obj.result), hyperlink };
      }
      return { value: String(obj.result || '').trim(), hyperlink };
    }
  }

  // Also check cell.text if available (preserves formatted zero-padded display when formatted cleanly)
  if (typeof rawVal === 'number') {
    return { value: String(rawVal).trim(), hyperlink };
  }

  return { value: String(rawVal).trim(), hyperlink };
}

function suggestMappingFromHeaders(headers: string[]): ArchiveImportColumnMapping {
  const mapping: ArchiveImportColumnMapping = {};
  const matchHeader = (candidates: string[]): string | undefined => {
    for (const h of headers) {
      const norm = normalizeCompareText(h);
      if (!norm) continue;
      for (const c of candidates) {
        if (norm === c || norm.includes(c)) return h;
      }
    }
    return undefined;
  };

  mapping.docCategory = matchHeader(['nhom van ban', 'nhom vb', 'phan loai', 'chieu van ban', 'loai so']);
  mapping.fullDocNumber = matchHeader(['so ky hieu', 'so hieu', 'so van ban', 'ky hieu']);
  mapping.docName = matchHeader(['ten van ban', 'ten vb', 'tieu de']);
  mapping.issueDate = matchHeader(['ngay ban hanh', 'ngay van ban', 'ngay vb', 'ngay ky']);
  mapping.year = matchHeader(['nam ban hanh', 'nam']);
  mapping.issuingAgency = matchHeader(['co quan ban hanh', 'don vi ban hanh', 'cq ban hanh', 'noi gui']);
  mapping.signer = matchHeader(['nguoi ky']);
  mapping.signerRole = matchHeader(['chuc vu nguoi ky', 'chuc vu']);
  mapping.typeCode = matchHeader(['loai van ban', 'ma loai', 'loai vb']);
  mapping.parentAgency = matchHeader(['co quan chu quan', 'cap tren']);
  mapping.title = matchHeader(['trich yeu', 'noi dung']);
  mapping.notes = matchHeader(['ghi chu']);
  mapping.driveLinkOrFileName = matchHeader([
    'lien ket hoac ten tep drive',
    'lien ket drive',
    'ten tep drive',
    'duong dan drive',
    'tep dinh kem',
    'link drive',
    'file drive',
    'ten tep',
  ]);

  return mapping;
}

export async function parseArchiveExcelBuffer(
  buffer: Buffer,
  fileName: string
): Promise<ArchiveParseExcelResponse> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);

  const sheets: ArchiveParseExcelResponse['sheets'] = [];

  for (const ws of wb.worksheets) {
    if (ws.rowCount === 0) continue;
    // Detect header row (first row with >= 2 non-empty cells)
    let headerRowIdx = 1;
    for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
      const row = ws.getRow(r);
      let nonEmpty = 0;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (extractExcelCellTextAndLink(cell).value) nonEmpty++;
      });
      if (nonEmpty >= 2) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = ws.getRow(headerRowIdx);
    const colHeaderMap = new Map<number, string>();
    const headers: string[] = [];
    const seenHeaderCount = new Map<string, number>();

    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const raw = extractExcelCellTextAndLink(cell).value;
      if (!raw) return;
      const count = seenHeaderCount.get(raw) || 0;
      seenHeaderCount.set(raw, count + 1);
      const uniqueHeader = count === 0 ? raw : `${raw} (${count + 1})`;
      colHeaderMap.set(colNumber, uniqueHeader);
      headers.push(uniqueHeader);
    });

    if (headers.length === 0) continue;

    const rows: Array<Record<string, { value: string; hyperlink?: string }>> = [];
    for (let r = headerRowIdx + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const obj: Record<string, { value: string; hyperlink?: string }> = {};
      let hasAnyValue = false;
      for (const [colIdx, headerName] of colHeaderMap.entries()) {
        const cellVal = extractExcelCellTextAndLink(row.getCell(colIdx));
        if (cellVal.value || cellVal.hyperlink) hasAnyValue = true;
        obj[headerName] = cellVal;
      }
      if (hasAnyValue) {
        rows.push(obj);
      }
    }

    sheets.push({
      name: ws.name,
      rowCount: rows.length,
      headers,
      suggestedMapping: suggestMappingFromHeaders(headers),
      rows: rows.slice(0, 2000),
    });
  }

  return {
    fileName,
    sheets,
  };
}

export function normalizeArchiveCategory(raw: string): 'outgoing' | 'incoming' | null {
  const norm = normalizeCompareText(raw);
  if (!norm) return null;
  if (norm === 'incoming' || norm.includes('van ban den') || norm === 'den' || norm === 'vb den') {
    return 'incoming';
  }
  if (norm === 'outgoing' || norm.includes('van ban di') || norm === 'di' || norm === 'vb di') {
    return 'outgoing';
  }
  return null;
}

export function inferArchiveTypeCode(
  explicitType: string | undefined,
  fullDocNumber: string,
  docName: string,
  documentTypes: DocumentTypeConfig[]
): { code: string; name: string } {
  const cleanExplicit = (explicitType || '').trim();
  if (cleanExplicit) {
    const upper = cleanExplicit.toUpperCase();
    const byCode = documentTypes.find((t) => t.code.toUpperCase() === upper);
    if (byCode) return { code: byCode.code, name: byCode.name };
    const normExp = normalizeCompareText(cleanExplicit);
    const byName = documentTypes.find((t) => normalizeCompareText(t.name) === normExp);
    if (byName) return { code: byName.code, name: byName.name };
    // Try code stripped of diacritics
    const strippedExp = removeVietnameseDiacritics(upper);
    const byStripped = documentTypes.find(
      (t) => removeVietnameseDiacritics(t.code).toUpperCase() === strippedExp
    );
    if (byStripped) return { code: byStripped.code, name: byStripped.name };
  }

  // Try matching code embedded in fullDocNumber (e.g. "007-KH/ĐTN" or "12/QĐ-CB")
  const normNumUpper = removeVietnameseDiacritics(fullDocNumber || '').toUpperCase();
  for (const t of documentTypes) {
    const strippedCode = removeVietnameseDiacritics(t.code).toUpperCase();
    if (new RegExp(`[-/]${strippedCode}(?:[-/]|$)`).test(normNumUpper)) {
      return { code: t.code, name: t.name };
    }
  }

  // Try matching docName prefix (e.g. "Quyết định...", "Thông báo...")
  const normTitle = normalizeCompareText(docName);
  for (const t of documentTypes) {
    const tNameNorm = normalizeCompareText(t.name);
    if (tNameNorm && (normTitle.startsWith(tNameNorm) || normTitle.includes(tNameNorm))) {
      return { code: t.code, name: t.name };
    }
  }

  return {
    code: cleanExplicit ? cleanExplicit.toUpperCase().slice(0, 10) : 'CV',
    name: cleanExplicit || 'Công văn / Khác',
  };
}

export function parseHistoricalOutgoingSequenceNumber(
  fullDocNumber: string,
  _book?: RegistryBook
): number | null {
  const trimmed = (fullDocNumber || '').trim();
  if (!trimmed) return null;
  // Leading digits before '-', '/', '.' (e.g. "007-KH/ĐTN" -> 7, "12/QĐ-CB" -> 12)
  const lead = trimmed.match(/^0*(\d{1,6})(?=\s*[-\/.])/);
  if (lead) {
    const num = parseInt(lead[1], 10);
    if (Number.isFinite(num) && num >= 1 && num <= 999999) return num;
  }
  const pureDigits = trimmed.match(/^0*(\d{1,6})$/);
  if (pureDigits) {
    const num = parseInt(pureDigits[1], 10);
    if (Number.isFinite(num) && num >= 1 && num <= 999999) return num;
  }
  return null;
}

export function computeArchiveRowFingerprint(bookId: string, row: {
  docCategory: 'outgoing' | 'incoming';
  fullDocNumber: string;
  docName: string;
  issueDate?: string;
  year?: string | number;
  issuingAgency?: string;
}): string {
  const payload = [
    bookId.trim(),
    row.docCategory,
    normalizeDocNumberForCompare(row.fullDocNumber),
    normalizeCompareText(row.docName),
    normalizeDateOnlyDdMmYyyy(row.issueDate || ''),
    String(row.year || '').trim(),
    row.docCategory === 'incoming' ? normalizeCompareText(row.issuingAgency || '') : '',
  ].join('|');
  return crypto.createHash('sha256').update(payload).digest('hex').slice(0, 24);
}

function stripExtForMatch(fileName: string): string {
  const ext = getSupportedAttachmentExtension(fileName);
  if (ext) return fileName.slice(0, -ext.length).trim();
  const dotIdx = fileName.lastIndexOf('.');
  if (dotIdx > 0) return fileName.slice(0, dotIdx).trim();
  return fileName.trim();
}

export async function previewArchiveImportRows(params: {
  book: RegistryBook;
  folderIdOverride?: string;
  driveConfig: ResolvedGoogleDriveConfig;
  rows: ArchiveImportEditableRow[];
  existingDocuments: DocumentRecord[];
  documentTypes: DocumentTypeConfig[];
  archiveImportedFingerprints?: Record<string, string>;
}): Promise<ArchiveImportPreviewResponse> {
  const { book, folderIdOverride, driveConfig, rows, existingDocuments, documentTypes } = params;
  const effectiveFolderId = (folderIdOverride || driveConfig.folderId || '').trim();
  const folderSource: 'custom' | 'book' | 'none' = folderIdOverride
    ? 'custom'
    : driveConfig.folderId
    ? 'book'
    : 'none';

  // Step 1: Scan Drive folder once if configured
  let scannedFiles: ArchiveDriveCandidate[] = [];
  let driveScanError: string | undefined;

  if (effectiveFolderId && driveConfig.enabled) {
    const scanRes = await listDriveFilesInFolder(driveConfig, effectiveFolderId, 4);
    scannedFiles = scanRes.files;
    driveScanError = scanRes.error;
  } else if (!driveConfig.enabled) {
    driveScanError = 'Sổ chưa bật Google Drive — chỉ có thể kiểm tra nếu cấu hình Drive hoặc nhập trực tiếp sau.';
  }

  // Index scanned files for fast lookup
  const fileById = new Map<string, ArchiveDriveCandidate>();
  const filesByExactLowerName = new Map<string, ArchiveDriveCandidate[]>();
  const filesByNormStem = new Map<string, ArchiveDriveCandidate[]>();

  for (const f of scannedFiles) {
    fileById.set(f.fileId, f);
    const lowerName = f.fileName.trim().toLowerCase();
    filesByExactLowerName.set(lowerName, [...(filesByExactLowerName.get(lowerName) || []), f]);

    const normStem = normalizeCompareText(stripExtForMatch(f.fileName));
    if (normStem) {
      filesByNormStem.set(normStem, [...(filesByNormStem.get(normStem) || []), f]);
    }
  }

  // Cache per explicit Drive fileId lookup
  const explicitFileCache = new Map<
    string,
    Awaited<ReturnType<typeof getDriveFileMetadata>>
  >();

  // Prepare duplicate lookup in target book
  const activeBookDocs = existingDocuments.filter(
    (d) => getDocBookId(d) === book.id && d.status !== 'deleted'
  );

  // Also detect intra-batch duplicates
  const batchSeenMap = new Map<string, number[]>();
  rows.forEach((r) => {
    const normNum = normalizeDocNumberForCompare(r.fullDocNumber);
    if (!normNum) return;
    const key =
      r.docCategory === 'incoming'
        ? `${r.docCategory}|${normNum}|${normalizeCompareText(r.issuingAgency)}`
        : `${r.docCategory}|${normNum}|${String(r.year || '').trim()}`;
    batchSeenMap.set(key, [...(batchSeenMap.get(key) || []), r.rowNumber]);
  });

  const previewRows: ArchiveImportPreviewRow[] = [];

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];
    const rowId = `row-${row.rowNumber}-${idx}`;
    const errors: string[] = [];
    const warnings: string[] = [];

    const fullDocNumber = String(row.fullDocNumber || '').trim();
    const docName = String(row.docName || '').trim();
    const issueDateRaw = String(row.issueDate || '').trim();
    const normalizedIssueDate = issueDateRaw ? normalizeDateOnlyDdMmYyyy(issueDateRaw) : '';
    const issuingAgency = String(row.issuingAgency || '').trim();
    const signer = String(row.signer || '').trim();

    if (!fullDocNumber) errors.push('Thiếu Số/Ký hiệu');
    if (!docName) errors.push('Thiếu Tên văn bản');

    if (issueDateRaw && !normalizedIssueDate) {
      warnings.push(`Ngày ban hành "${issueDateRaw}" chưa đúng chuẩn dd/mm/yyyy.`);
    }

    let issueDateYear: number | null = null;
    if (normalizedIssueDate) {
      const yp = normalizedIssueDate.split('/')[2];
      const parsedYp = parseInt(yp, 10);
      if (Number.isFinite(parsedYp)) issueDateYear = parsedYp;
    }

    let yearNum: number | null = null;
    const yearStr = String(row.year || '').trim();
    if (yearStr) {
      const parsedY = parseInt(yearStr, 10);
      if (Number.isFinite(parsedY) && parsedY >= 1990 && parsedY <= 2100) {
        yearNum = parsedY;
      } else {
        errors.push(`Năm "${yearStr}" không hợp lệ`);
      }
    } else if (issueDateYear) {
      yearNum = issueDateYear;
    } else {
      errors.push('Thiếu Năm văn bản (hoặc Ngày ban hành hợp lệ)');
    }

    let yearMismatchWarning: string | undefined;
    if (yearNum && issueDateYear && yearNum !== issueDateYear) {
      yearMismatchWarning = `Năm nhập (${yearNum}) khác với năm trong Ngày ban hành (${issueDateYear}).`;
      warnings.push(yearMismatchWarning);
    }

    const rowFingerprint = computeArchiveRowFingerprint(book.id, {
      docCategory: row.docCategory,
      fullDocNumber,
      docName,
      issueDate: normalizedIssueDate || issueDateRaw,
      year: yearNum || yearStr,
      issuingAgency,
    });

    // Idempotency check against previously archived records
    const committedDocId = params.archiveImportedFingerprints?.[rowFingerprint];
    const alreadyCommittedDoc =
      committedDocId && activeBookDocs.find((d) => d.id === committedDocId);

    // Duplicate check in book
    const normNum = normalizeDocNumberForCompare(fullDocNumber);
    const normAgency = normalizeCompareText(issuingAgency);
    const duplicateMatches: ArchiveImportPreviewRow['duplicateMatches'] = [];

    if (normNum) {
      for (const existing of activeBookDocs) {
        if (normalizeDocNumberForCompare(existing.fullDocNumber) !== normNum) continue;
        if (row.docCategory === 'incoming') {
          if (existing.docCategory !== 'incoming') continue;
          const exAgency = normalizeCompareText(existing.issuingAgency || '');
          // Yêu cầu kiểm tra cả số/ký hiệu và cơ quan ban hành cho văn bản đến
          if (normAgency && exAgency) {
            if (normAgency !== exAgency) continue;
          } else if (normAgency !== exAgency) {
            // One side lacks issuingAgency: compare year too before flagging
            const exYear = getDocumentFilterYear(existing);
            if (yearNum && exYear !== 'unknown' && exYear !== yearNum) continue;
          }
        } else {
          if (existing.docCategory === 'incoming') continue;
          const exYear = getDocumentFilterYear(existing);
          if (yearNum && exYear !== 'unknown' && exYear !== yearNum) continue;
        }
        duplicateMatches.push({
          docId: existing.id,
          fullDocNumber: existing.fullDocNumber,
          docName: existing.docName || existing.title,
          issueDate: existing.issueDate,
          issuingAgency: existing.issuingAgency,
          status: existing.status,
        });
      }
    }

    const batchKey =
      row.docCategory === 'incoming'
        ? `${row.docCategory}|${normNum}|${normAgency}`
        : `${row.docCategory}|${normNum}|${String(yearNum || yearStr)}`;
    const batchDuplicateRows = (batchSeenMap.get(batchKey) || []).filter(
      (n) => n !== row.rowNumber
    );
    if (batchDuplicateRows.length > 0) {
      warnings.push(`Trùng với dòng ${batchDuplicateRows.join(', ')} trong cùng tệp Excel.`);
    }

    // Drive Matching Priority:
    // 0) User explicitly resolved candidate via selectedDriveFileId
    // 1) Drive ID / link
    // 2) Exact file name in selected folder
    // 3) Suggested filename from Số/Ký hiệu & Tên văn bản
    let driveMatchStatus: ArchiveImportPreviewRow['driveMatchStatus'] = 'none';
    let matchedDriveFile: ArchiveDriveCandidate | undefined;
    let driveCandidates: ArchiveDriveCandidate[] = [];

    const linkOrName = (row.driveHyperlink || row.driveLinkOrFileName || '').trim();

    if (row.selectedDriveFileId) {
      const fromScan = fileById.get(row.selectedDriveFileId);
      if (fromScan) {
        matchedDriveFile = fromScan;
        driveMatchStatus = 'matched_by_id';
      } else if (driveConfig.enabled) {
        if (!explicitFileCache.has(row.selectedDriveFileId)) {
          explicitFileCache.set(
            row.selectedDriveFileId,
            await getDriveFileMetadata(driveConfig, row.selectedDriveFileId)
          );
        }
        const meta = explicitFileCache.get(row.selectedDriveFileId)!;
        if (meta.status === 'ok' && meta.file) {
          matchedDriveFile = meta.file;
          driveMatchStatus = 'matched_by_id';
        } else if (meta.status === 'access_denied') {
          driveMatchStatus = 'access_denied';
          warnings.push(meta.message || 'Thiếu quyền truy cập tệp Google Drive đã chọn.');
        } else {
          driveMatchStatus = 'not_found';
          warnings.push(meta.message || 'Không tìm thấy tệp Google Drive đã chọn.');
        }
      }
    }

    if (!matchedDriveFile && driveMatchStatus === 'none' && linkOrName) {
      const extractedId = extractGoogleDriveFileId(linkOrName);
      if (extractedId) {
        const fromScan = fileById.get(extractedId);
        if (fromScan) {
          matchedDriveFile = fromScan;
          driveMatchStatus = 'matched_by_id';
        } else if (driveConfig.enabled) {
          if (!explicitFileCache.has(extractedId)) {
            explicitFileCache.set(
              extractedId,
              await getDriveFileMetadata(driveConfig, extractedId)
            );
          }
          const meta = explicitFileCache.get(extractedId)!;
          if (meta.status === 'ok' && meta.file) {
            matchedDriveFile = meta.file;
            driveMatchStatus = 'matched_by_id';
          } else if (meta.status === 'access_denied') {
            driveMatchStatus = 'access_denied';
            warnings.push(meta.message || 'Thiếu quyền truy cập liên kết Google Drive.');
          } else if (meta.status === 'not_found') {
            driveMatchStatus = 'not_found';
            warnings.push(meta.message || 'Không tìm thấy tệp tại liên kết/ID Drive đã nhập.');
          } else {
            driveMatchStatus = 'not_found';
            warnings.push(meta.message || 'Chưa xác minh được liên kết Drive.');
          }
        } else {
          warnings.push('Chưa bật tích hợp Drive để xác minh ID tệp.');
          driveMatchStatus = 'not_found';
        }
      } else {
        // Match exact filename in scanned Drive folder
        const cleanLower = linkOrName.toLowerCase();
        const exactMatches = filesByExactLowerName.get(cleanLower) || [];
        if (exactMatches.length === 1) {
          matchedDriveFile = exactMatches[0];
          driveMatchStatus = 'matched_exact_name';
        } else if (exactMatches.length > 1) {
          driveMatchStatus = 'ambiguous';
          driveCandidates = exactMatches;
          warnings.push(`Tìm thấy ${exactMatches.length} tệp trùng tên "${linkOrName}" trong thư mục Drive — vui lòng chọn đúng tệp.`);
        } else {
          // Try normalized stem match against linkOrName
          const stemMatches =
            filesByNormStem.get(normalizeCompareText(stripExtForMatch(linkOrName))) || [];
          if (stemMatches.length === 1) {
            matchedDriveFile = stemMatches[0];
            driveMatchStatus = 'matched_exact_name';
          } else if (stemMatches.length > 1) {
            driveMatchStatus = 'ambiguous';
            driveCandidates = stemMatches;
            warnings.push(`Tìm thấy ${stemMatches.length} tệp có tên gần giống "${linkOrName}" trên Drive — vui lòng chọn đúng tệp.`);
          }
        }
      }
    }

    // Priority 3: Suggested filename "<Số ký hiệu (đổi / thành .)> - <Tên văn bản>"
    if (
      !matchedDriveFile &&
      driveMatchStatus === 'none' &&
      fullDocNumber &&
      docName &&
      scannedFiles.length > 0
    ) {
      const suggestedStem = generateSuggestedFileName(fullDocNumber, docName);
      const suggestedNorm = normalizeCompareText(suggestedStem);
      const suggestedPdfLower = `${suggestedStem}.pdf`.toLowerCase();
      const suggestedDocxLower = `${suggestedStem}.docx`.toLowerCase();

      let suggestedMatches = [
        ...(filesByExactLowerName.get(suggestedPdfLower) || []),
        ...(filesByExactLowerName.get(suggestedDocxLower) || []),
      ];
      if (suggestedMatches.length === 0 && suggestedNorm) {
        suggestedMatches = filesByNormStem.get(suggestedNorm) || [];
      }

      if (suggestedMatches.length === 1) {
        matchedDriveFile = suggestedMatches[0];
        driveMatchStatus = 'matched_suggested_name';
      } else if (suggestedMatches.length > 1) {
        driveMatchStatus = 'ambiguous';
        driveCandidates = suggestedMatches;
        warnings.push(`Có ${suggestedMatches.length} tệp Drive khớp tên gợi ý cho văn bản này — vui lòng chọn.`);
      } else {
        driveMatchStatus = 'not_found';
      }
    } else if (!matchedDriveFile && driveMatchStatus === 'none') {
      driveMatchStatus = 'not_found';
    }

    // Outgoing series/sequence info
    const inferredType = inferArchiveTypeCode(
      row.typeCode,
      fullDocNumber,
      docName,
      documentTypes
    );
    let parsedSequenceNumber: number | undefined;
    let seriesUpdateInfo: string | undefined;

    if (row.docCategory === 'outgoing' && yearNum) {
      const parsedSeq = parseHistoricalOutgoingSequenceNumber(fullDocNumber, book);
      if (parsedSeq !== null) {
        parsedSequenceNumber = parsedSeq;
        const seriesType = book.numberMode === 'shared' ? 'SHARED' : inferredType.code;
        seriesUpdateInfo = `Giữ số "${fullDocNumber}" (STT ${parsedSeq}) · Dãy ${yearNum}/${seriesType} sẽ cập nhật mốc tối đa >= ${parsedSeq}.`;
      }
    }

    // Compute overall row status
    let status: ArchiveImportPreviewRow['status'] = 'ready';
    if (errors.length > 0) {
      status = 'missing_info';
    } else if (alreadyCommittedDoc || duplicateMatches.length > 0 || batchDuplicateRows.length > 0) {
      status = 'possible_duplicate';
    } else if (driveMatchStatus === 'ambiguous') {
      status = 'ambiguous_file';
    } else if (driveMatchStatus === 'not_found' || driveMatchStatus === 'access_denied') {
      status = 'unmatched_file';
    }

    previewRows.push({
      ...row,
      fullDocNumber,
      docName,
      issueDate: normalizedIssueDate || issueDateRaw,
      year: yearNum ? String(yearNum) : yearStr,
      issuingAgency,
      signer,
      rowId,
      rowFingerprint,
      status,
      errors,
      warnings,
      yearMismatchWarning,
      alreadyCommittedDocId: alreadyCommittedDoc ? alreadyCommittedDoc.id : undefined,
      duplicateMatches,
      batchDuplicateRows,
      driveMatchStatus,
      matchedDriveFile,
      driveCandidates,
      resolvedTypeCode: inferredType.code,
      parsedSequenceNumber,
      seriesUpdateInfo,
    });
  }

  const summary = {
    total: previewRows.length,
    ready: previewRows.filter((r) => r.status === 'ready').length,
    possibleDuplicate: previewRows.filter((r) => r.status === 'possible_duplicate').length,
    missingInfo: previewRows.filter((r) => r.status === 'missing_info').length,
    unmatchedFile: previewRows.filter((r) => r.status === 'unmatched_file').length,
    ambiguousFile: previewRows.filter((r) => r.status === 'ambiguous_file').length,
  };

  return {
    bookId: book.id,
    bookName: book.name,
    bookCode: book.code,
    folderId: effectiveFolderId,
    folderSource,
    driveScanError,
    driveFileCount: scannedFiles.length,
    summary,
    rows: previewRows,
  };
}

// =========================================================
// FEATURE 3: YEAR-END CLOSING & HANDOVER HELPERS
// =========================================================

export function isClosingDocMissingFile(doc: Pick<DocumentRecord, 'status' | 'driveStatus' | 'driveFileId' | 'driveFileUrl'>): boolean {
  if (doc.status === 'cancelled' || doc.status === 'deleted') return false;
  return doc.driveStatus !== 'saved' || (!doc.driveFileId && !doc.driveFileUrl);
}

export function buildClosingSummaryAndDocs(
  allDocs: DocumentRecord[],
  book: RegistryBook,
  year: number
): { summary: ClosingSummaryCounts; documents: ClosingSnapshotDocument[] } {
  const scoped = allDocs
    .filter((d) => {
      if (d.status === 'deleted') return false;
      if (getDocBookId(d) !== book.id) return false;
      return getDocumentFilterYear(d) === year;
    })
    .sort((a, b) => {
      const catA = a.docCategory === 'incoming' ? 1 : 0;
      const catB = b.docCategory === 'incoming' ? 1 : 0;
      if (catA !== catB) return catA - catB;
      if (a.sequenceNumber !== b.sequenceNumber) return a.sequenceNumber - b.sequenceNumber;
      return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
    });

  let outgoing = 0;
  let incoming = 0;
  let cancelled = 0;
  let missingFile = 0;
  let unsyncedSheets = 0;
  let pendingDateConfirmation = 0;

  const frozenDocs: ClosingSnapshotDocument[] = scoped.map((d) => {
    const officialAtt = getOfficialAttachment(d);
    const driveFileId = officialAtt?.driveFileId || d.driveFileId;
    const driveFileName = officialAtt?.fileName || d.driveFileName;
    const driveFileUrl = officialAtt?.driveFileUrl || d.driveFileUrl;
    const driveStatus = officialAtt?.driveStatus || d.driveStatus || 'none';

    if (d.docCategory === 'incoming') incoming++;
    else outgoing++;
    if (d.status === 'cancelled') cancelled++;
    if (d.status !== 'cancelled' && (driveStatus !== 'saved' || (!driveFileId && !driveFileUrl))) {
      missingFile++;
    }
    if (d.syncStatus !== 'synced') {
      unsyncedSheets++;
    }
    if (d.pendingDateConfirmation) {
      pendingDateConfirmation++;
    }

    return {
      docId: d.id,
      bookId: book.id,
      bookCode: book.code,
      bookName: book.name,
      docCategory: d.docCategory || 'outgoing',
      typeCode: d.typeCode,
      typeName: d.typeName || d.typeCode,
      sequenceNumber: d.sequenceNumber,
      fullDocNumber: d.fullDocNumber,
      docName: d.docName || d.title,
      title: d.title || '',
      issueDate: d.issueDate || '',
      signDate: d.signDate,
      year: typeof d.year === 'number' ? d.year : year,
      issuingAgency: d.issuingAgency,
      parentAgency: d.parentAgency,
      signer: d.signer || '',
      signerRole: d.signerRole || '',
      issuedBy: d.issuedBy || '',
      status: d.status,
      cancelReason: d.cancelReason,
      driveStatus,
      driveFileId,
      driveFileName,
      driveFileUrl,
      syncStatus: d.syncStatus || 'not_configured',
      pendingDateConfirmation: Boolean(d.pendingDateConfirmation),
      notes: d.notes,
      createdAt: d.createdAt,
    };
  });

  return {
    summary: {
      totalActiveAndCancelled: frozenDocs.length,
      outgoing,
      incoming,
      cancelled,
      missingFile,
      unsyncedSheets,
      pendingDateConfirmation,
    },
    documents: frozenDocs,
  };
}

export async function buildYearEndHandoverWorkbookBuffer(
  snapshot: YearEndClosingSnapshot,
  options: { canOpenDriveLinks: boolean }
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Sổ văn bản Trường THPT Nguyễn Chí Thanh';
  wb.created = new Date(snapshot.closedAt);

  // Sheet 1: Tổng hợp
  const summarySheet = wb.addWorksheet('Tổng hợp', {
    pageSetup: { orientation: 'portrait', paperSize: 9 },
  });
  summarySheet.columns = [{ width: 34 }, { width: 48 }];

  summarySheet.mergeCells('A1:B1');
  const titleCell = summarySheet.getCell('A1');
  titleCell.value = `BIÊN BẢN BÀN GIAO SỔ VĂN BẢN NĂM ${snapshot.year}`;
  titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF0369A1' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  summarySheet.getRow(1).height = 28;

  summarySheet.mergeCells('A2:B2');
  const subCell = summarySheet.getCell('A2');
  subCell.value = `Đơn vị / Sổ: ${snapshot.bookName} (${snapshot.bookCode}) · Mã lần chốt: ${snapshot.closingCode}`;
  subCell.font = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF475569' } };
  subCell.alignment = { vertical: 'middle', horizontal: 'center' };
  summarySheet.getRow(2).height = 20;

  const infoRows: Array<[string, string | number]> = [
    ['Mã lần chốt bất biến', snapshot.closingCode],
    ['Tên sổ', `${snapshot.bookName} (${snapshot.bookCode})`],
    ['Năm chốt sổ', snapshot.year],
    ['Người chốt sổ', `${snapshot.closedBy} (@${snapshot.closedByUsername})`],
    ['Thời điểm chốt (giờ VN)', snapshot.closedAtVietnam],
    ['Ghi chú khi chốt', snapshot.note || '—'],
    ['', ''],
    ['TỔNG HỢP SỐ LIỆU TẠI THỜI ĐIỂM CHỐT', 'GIÁ TRỊ ĐÃ KHÓA'],
    ['Tổng số hồ sơ hiệu lực & hủy', snapshot.summary.totalActiveAndCancelled],
    ['Số lượng Văn bản đi', snapshot.summary.outgoing],
    ['Số lượng Văn bản đến', snapshot.summary.incoming],
    ['Số văn bản đã hủy', snapshot.summary.cancelled],
    ['Hồ sơ thiếu tệp Google Drive', snapshot.summary.missingFile],
    ['Hồ sơ chưa đồng bộ Google Trang tính', snapshot.summary.unsyncedSheets],
    ['Hồ sơ chờ xác nhận ngày', snapshot.summary.pendingDateConfirmation],
    ['Số lần đính chính sau chốt', snapshot.amendments.length],
  ];

  infoRows.forEach(([label, val], idx) => {
    const row = summarySheet.addRow([label, val]);
    if (idx === 7) {
      row.eachCell((c) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0369A1' } };
        c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      });
    } else if (label) {
      row.getCell(1).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1E293B' } };
      row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: 'FF0F172A' } };
    }
  });

  // Sheet 2: Danh mục văn bản (Frozen Snapshot)
  const listSheet = wb.addWorksheet('Danh mục văn bản', {
    views: [{ state: 'frozen', ySplit: 3 }],
    pageSetup: { orientation: 'landscape', paperSize: 9 },
  });
  listSheet.mergeCells('A1:N1');
  listSheet.getCell('A1').value = `DANH MỤC VĂN BẢN ĐÃ CHỐT — ${snapshot.bookName.toUpperCase()} NĂM ${snapshot.year} (${snapshot.closingCode})`;
  listSheet.getCell('A1').font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF0369A1' } };
  listSheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };
  listSheet.getRow(1).height = 26;

  listSheet.mergeCells('A2:N2');
  listSheet.getCell('A2').value = `Chốt bởi ${snapshot.closedBy} lúc ${snapshot.closedAtVietnam} · Bản chụp bất biến (${snapshot.documents.length} hồ sơ)`;
  listSheet.getCell('A2').font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF475569' } };
  listSheet.getCell('A2').alignment = { vertical: 'middle', horizontal: 'center' };

  listSheet.columns = [
    { key: 'stt', width: 7 },
    { key: 'docId', width: 22 },
    { key: 'docCategory', width: 15 },
    { key: 'fullDocNumber', width: 22 },
    { key: 'issueDate', width: 15 },
    { key: 'year', width: 9 },
    { key: 'typeName', width: 18 },
    { key: 'docName', width: 40 },
    { key: 'agency', width: 28 },
    { key: 'signer', width: 22 },
    { key: 'status', width: 18 },
    { key: 'driveStatus', width: 22 },
    { key: 'driveLink', width: 38 },
    { key: 'syncStatus', width: 22 },
  ];

  const listHeaderRow = listSheet.getRow(3);
  listHeaderRow.values = [
    'STT',
    'Mã hồ sơ',
    'Nhóm VB',
    'Số/Ký hiệu',
    'Ngày ban hành',
    'Năm',
    'Loại văn bản',
    'Tên văn bản',
    'Cơ quan ban hành',
    'Người ký',
    'Trạng thái hồ sơ',
    'Trạng thái tệp Drive',
    'Liên kết tệp Drive',
    'Google Trang tính',
  ];
  listHeaderRow.height = 26;
  listHeaderRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0369A1' } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });

  snapshot.documents.forEach((doc, idx) => {
    const syncLabel =
      doc.syncStatus === 'synced'
        ? 'Đã đồng bộ'
        : doc.syncStatus === 'error'
        ? 'Lỗi đồng bộ'
        : doc.syncStatus === 'pending'
        ? 'Chờ đồng bộ'
        : 'Chưa cấu hình';
    const driveLabel =
      doc.driveStatus === 'saved'
        ? doc.driveFileName || 'Đã lưu'
        : doc.driveStatus === 'failed'
        ? 'Lưu thất bại'
        : 'Thiếu tệp';

    const row = listSheet.addRow({
      stt: idx + 1,
      docId: doc.docId,
      docCategory: doc.docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi',
      fullDocNumber: doc.fullDocNumber,
      issueDate: doc.issueDate || (doc.pendingDateConfirmation ? 'Chờ xác nhận' : ''),
      year: doc.year,
      typeName: doc.typeName,
      docName: doc.docName,
      agency: doc.issuingAgency || (doc.docCategory === 'outgoing' ? snapshot.bookName : ''),
      signer: doc.signer,
      status: doc.status === 'cancelled' ? `Đã hủy (${doc.cancelReason || ''})` : 'Hiệu lực',
      driveStatus: driveLabel,
      driveLink: options.canOpenDriveLinks && doc.driveFileUrl ? doc.driveFileUrl : '',
      syncStatus: syncLabel,
    });

    if (options.canOpenDriveLinks && doc.driveFileUrl) {
      row.getCell('driveLink').value = {
        text: doc.driveFileName || doc.driveFileUrl,
        hyperlink: doc.driveFileUrl,
      };
      row.getCell('driveLink').font = { name: 'Arial', size: 10, color: { argb: 'FF0369A1' }, underline: true };
    }
  });

  // Sheet 3: Mục cần xử lý (Frozen Snapshot)
  const actionSheet = wb.addWorksheet('Mục cần xử lý', {
    views: [{ state: 'frozen', ySplit: 2 }],
    pageSetup: { orientation: 'landscape', paperSize: 9 },
  });
  actionSheet.mergeCells('A1:G1');
  actionSheet.getCell('A1').value = `DANH SÁCH MỤC CẦN XỬ LÝ TẠI THỜI ĐIỂM CHỐT (${snapshot.closingCode})`;
  actionSheet.getCell('A1').font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF991B1B' } };
  actionSheet.getRow(1).height = 24;

  actionSheet.columns = [
    { key: 'stt', width: 7 },
    { key: 'issueType', width: 28 },
    { key: 'fullDocNumber', width: 22 },
    { key: 'docName', width: 40 },
    { key: 'docCategory', width: 15 },
    { key: 'detail', width: 38 },
    { key: 'docId', width: 22 },
  ];
  const actHeader = actionSheet.getRow(2);
  actHeader.values = ['STT', 'Loại mục cần xử lý', 'Số/Ký hiệu', 'Tên văn bản', 'Nhóm VB', 'Chi tiết', 'Mã hồ sơ'];
  actHeader.height = 24;
  actHeader.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB45309' } };
    cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  let actionStt = 0;
  for (const doc of snapshot.documents) {
    const reasons: Array<{ kind: string; detail: string }> = [];
    if (doc.status !== 'cancelled' && (doc.driveStatus !== 'saved' || (!doc.driveFileId && !doc.driveFileUrl))) {
      reasons.push({ kind: 'Thiếu tệp Drive', detail: 'Hồ sơ chưa có tệp lưu hợp lệ trên Google Drive' });
    }
    if (doc.syncStatus !== 'synced') {
      reasons.push({
        kind: 'Chưa đồng bộ Google Trang tính',
        detail: `Trạng thái đồng bộ: ${doc.syncStatus}`,
      });
    }
    if (doc.pendingDateConfirmation) {
      reasons.push({ kind: 'Chờ xác nhận ngày', detail: 'Ngày văn bản cần quản trị viên xác nhận thống nhất' });
    }
    if (doc.status === 'cancelled') {
      reasons.push({ kind: 'Số đã hủy', detail: doc.cancelReason || 'Đã hủy hiệu lực' });
    }
    for (const r of reasons) {
      actionStt++;
      actionSheet.addRow({
        stt: actionStt,
        issueType: r.kind,
        fullDocNumber: doc.fullDocNumber,
        docName: doc.docName,
        docCategory: doc.docCategory === 'incoming' ? 'Văn bản đến' : 'Văn bản đi',
        detail: r.detail,
        docId: doc.docId,
      });
    }
  }

  // Sheet 4: Đính chính sau chốt (nếu có)
  if (snapshot.amendments.length > 0) {
    const amSheet = wb.addWorksheet('Đính chính sau chốt');
    amSheet.columns = [
      { width: 8 },
      { width: 22 },
      { width: 22 },
      { width: 28 },
      { width: 20 },
      { width: 24 },
      { width: 26 },
      { width: 26 },
    ];
    const h = amSheet.getRow(1);
    h.values = [
      'STT',
      'Thời điểm (VN)',
      'Người đính chính',
      'Lý do đính chính',
      'Số/Ký hiệu VB',
      'Trường thông tin',
      'Giá trị trước đính chính',
      'Giá trị sau đính chính',
    ];
    h.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
      c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    });
    let aStt = 0;
    for (const am of snapshot.amendments) {
      for (const ch of am.changes) {
        aStt++;
        amSheet.addRow([
          aStt,
          am.amendedAtVietnam,
          `${am.amendedBy} (@${am.amendedByUsername})`,
          am.reason,
          ch.fullDocNumber || '—',
          ch.fieldLabel,
          ch.beforeValue || '—',
          ch.afterValue || '—',
        ]);
      }
    }
  }

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out as ArrayBuffer);
}
