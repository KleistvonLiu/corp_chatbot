import { createHash, randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";
import path from "node:path";
import type {
  ChatMessage,
  ChatSession,
  Citation,
  CitationImage,
  EntryType,
  FixedKnowledgeSourceStatus,
  ImportJobRecord,
  KnowledgeBaseRecord,
  KnowledgeChunk,
  KnowledgeSource,
  QuestionStatsRecord,
  RetrievalDebugRecord,
  RetrievalEvidenceSelectionStep,
  RetrievalScoredHit,
  UnansweredReason,
  UserDocumentId,
  UserDocumentLink
} from "../../shared/contracts";
import {
  buildRetrievalDebugFileName,
  buildActiveKnowledgeResponse,
  createEmptyQuestionStats,
  createJobRecord,
  ensureStorage,
  loadJob,
  loadKnowledgeBase,
  loadQuestionStats,
  loadSession,
  loadState,
  sanitizeFileName,
  saveJob,
  saveKnowledgeBase,
  saveQuestionStats,
  saveRetrievalDebug,
  saveSession,
  saveState,
  writeUploadFile
} from "./storage";
import { createAnswerProvider, createProvider, cosineSimilarity, extractDirectMatchTerms, extractKeywords } from "./providers";
import { collectCanonicalAttachmentReferences, parseWorkflowWorkbook, resolveCanonicalAttachmentPath } from "./parsers";

interface SearchHit {
  chunk: KnowledgeChunk;
  score: number;
}

interface ScoredSearchHit extends SearchHit {
  breakdown: RetrievalScoredHit;
}

interface ChunkMetadata {
  entryType?: EntryType;
  category?: string;
  aliases: string[];
}

const noEvidenceAnswer =
  "非常抱歉，该问题无法回答。行政相关请咨询Helen（yanhong wang），人事相关请咨询Susie（susie sl huang），如果需要补充问题及答案请联系Kleist（kleist jf liu）。";
const defaultFixedWorkbookPath = "/home/kleist/Downloads/corp-eng-knowledge-merged-20260407-canonical/knowledge.xlsx";
const defaultFixedAttachmentsDir = "/home/kleist/Downloads/corp-eng-knowledge-merged-20260407-canonical/attachments";
const defaultNewStaffGuidePath = "/home/kleist/Downloads/Corp. Eng New Staff Guide Book-20260401.pdf";
const defaultWorkflowSummaryPath = "/home/kleist/Downloads/流程教程汇总_20260209.xlsx";
const topicActionPattern = /怎么|如何|谁|联系|申请|审批|安装|报价|区别|对比|比较|差异|流程|购买|下单|负责/u;
const procurementIntentPattern = /买|购买|采购|报价|下单/u;
const referenceLookupPattern = /供应商|联系人|列表|系统链接|参考/u;
const explicitStepPattern = /(?:^|\n)\s*(?:\d+[.)、]|[-•])/u;
const sequentialActionPattern = /首先|然后|之后|最后|第一|第二|第三|第四|第五/u;
const imageAttachmentKinds = new Set(["png", "jpg", "jpeg", "webp"]);
const retrievalEligibleThreshold = 0.08;
const userDocumentConfigs = [
  {
    id: "new-staff-guide",
    label: "新同事共享册",
    envName: "USER_DOC_NEW_STAFF_PATH",
    defaultPath: defaultNewStaffGuidePath
  },
  {
    id: "workflow-summary",
    label: "流程教程汇总",
    envName: "USER_DOC_WORKFLOW_SUMMARY_PATH",
    defaultPath: defaultWorkflowSummaryPath
  }
] satisfies Array<{
  id: UserDocumentId;
  label: string;
  envName: string;
  defaultPath: string;
}>;

function isRetrievalDebugEnabled() {
  const raw = process.env.KEYWORD_RETRIEVAL_DEBUG_ENABLED?.trim().toLowerCase();
  return raw === "1" || raw === "true" || raw === "yes" || raw === "on";
}

function normalizeSpace(text: string) {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeSearchText(text: string) {
  return normalizeSpace(text).toLowerCase();
}

function stripUrls(text: string) {
  return text.replace(/https?:\/\/\S+/g, " ");
}

function isShortAlphanumericToken(term: string) {
  return /^[a-z0-9]+$/i.test(term) && term.length <= 4;
}

function clip(text: string, limit = 220) {
  const normalized = normalizeSpace(text);
  if (normalized.length <= limit) {
    return normalized;
  }
  return `${normalized.slice(0, limit - 1)}…`;
}

async function pathExists(filePath: string) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function createRetrievalDebugRecord(params: {
  traceId: string;
  createdAt: string;
  sessionId: string;
  turnIndex: number;
  knowledgeBaseId: string;
  originalQuestion: string;
  resolvedQuestion: string;
  providerMode: string;
}): RetrievalDebugRecord {
  const { traceId, createdAt, sessionId, turnIndex, knowledgeBaseId, originalQuestion, resolvedQuestion, providerMode } = params;
  return {
    traceId,
    fileName: "",
    createdAt,
    sessionId,
    turnIndex,
    knowledgeBaseId,
    originalQuestion,
    resolvedQuestion,
    providerMode,
    retrievalDebugEnabled: true,
    queryKeywords: [],
    directTerms: [],
    queryPhrases: [],
    directCandidateRows: [],
    candidateChunkCount: 0,
    totalChunkCount: 0,
    scoredHits: [],
    sortedHitOrder: [],
    eligibleThreshold: retrievalEligibleThreshold,
    eligibleHitIds: [],
    topicLookup: false,
    compareMode: false,
    dominantTopRow: false,
    evidenceSelectionSteps: [],
    selectedEvidenceChunkIds: [],
    hasStrongEvidence: false,
    summaryOnlyEvidence: false,
    shouldCallModel: false,
    answered: false,
    unansweredReason: null,
    citationRowNumbers: []
  };
}

function readConfiguredFixedSource(): FixedKnowledgeSourceStatus {
  const workbookPath = process.env.KNOWLEDGE_SOURCE_WORKBOOK_PATH?.trim() || defaultFixedWorkbookPath;
  const attachmentsDir = process.env.KNOWLEDGE_SOURCE_ATTACHMENTS_DIR?.trim() || defaultFixedAttachmentsDir;
  return {
    configured: Boolean(workbookPath),
    workbookPath: workbookPath || undefined,
    attachmentsDir: attachmentsDir || undefined
  };
}

function readConfiguredUserDocumentPath(envName: string, defaultPath: string) {
  return process.env[envName]?.trim() || defaultPath;
}

async function listUserDocumentLinks(): Promise<UserDocumentLink[]> {
  return Promise.all(
    userDocumentConfigs.map(async (config) => ({
      id: config.id,
      label: config.label,
      url: `/api/docs/${config.id}`,
      available: await pathExists(readConfiguredUserDocumentPath(config.envName, config.defaultPath))
    }))
  );
}

function fingerprintBuffer(hash: ReturnType<typeof createHash>, label: string, buffer: Buffer) {
  hash.update(label);
  hash.update("\0");
  hash.update(buffer);
  hash.update("\0");
}

async function computeFixedSourceFingerprint(workbookPath: string, attachmentsDir: string) {
  const workbookBuffer = await fs.readFile(workbookPath);
  const attachmentRefs = await collectCanonicalAttachmentReferences(workbookBuffer);
  const hash = createHash("sha256");
  fingerprintBuffer(hash, "workbook", workbookBuffer);

  for (const relativePath of [...new Set(attachmentRefs)].sort()) {
    const attachmentPath = resolveCanonicalAttachmentPath(attachmentsDir, relativePath);
    const attachmentBuffer = await fs.readFile(attachmentPath);
    fingerprintBuffer(hash, relativePath, attachmentBuffer);
  }

  return {
    fingerprint: hash.digest("hex"),
    workbookBuffer
  };
}

function parseDelimitedList(text: string) {
  return normalizeSpace(text)
    .split(/[\n;,；，]+/g)
    .map((item) => normalizeSpace(item))
    .filter(Boolean);
}

function parseEntryType(rawValue?: string): EntryType | undefined {
  if (
    rawValue === "流程" ||
    rawValue === "联系人" ||
    rawValue === "供应商" ||
    rawValue === "参考" ||
    rawValue === "系统链接"
  ) {
    return rawValue;
  }

  return undefined;
}

function isImageAttachmentKind(kind?: KnowledgeSource["attachmentKind"]) {
  return Boolean(kind && imageAttachmentKinds.has(kind));
}

function readMetadataLine(text: string, label: string) {
  const match = new RegExp(`(?:^|\\n)${label}:\\s*([^\\n]+)`, "u").exec(text);
  return match?.[1] ? normalizeSpace(match[1]) : "";
}

function getChunkMetadata(chunk: KnowledgeChunk): ChunkMetadata {
  return {
    entryType: chunk.entryType ?? parseEntryType(readMetadataLine(chunk.text, "条目类型")),
    category: chunk.category ?? (readMetadataLine(chunk.text, "一级分类") || undefined),
    aliases: chunk.aliases?.length ? chunk.aliases : parseDelimitedList(readMetadataLine(chunk.text, "关键词/别名"))
  };
}

function buildChunkTopicText(chunk: KnowledgeChunk) {
  const metadata = getChunkMetadata(chunk);
  return normalizeSpace([chunk.title, metadata.aliases.join("；"), chunk.text].filter(Boolean).join("\n"));
}

function isTopicLookupQuestion(question: string) {
  const compact = normalizeSearchText(question).replace(/[()（）【】[\]{}“”"'`·、，,.;；:：!?？!_\-\s]/g, "");
  return compact.length >= 2 && compact.length <= 24 && !topicActionPattern.test(question);
}

function prefersReferenceEntry(question: string) {
  return referenceLookupPattern.test(question);
}

function shouldPreferFlowCompanion(question: string) {
  return !prefersReferenceEntry(question) && (isTopicLookupQuestion(question) || procurementIntentPattern.test(question));
}

function topicCoverage(questionTerms: string[], chunk: KnowledgeChunk) {
  return exactTermCoverage(questionTerms, buildChunkTopicText(chunk));
}

function prioritizeTopicHits(question: string, hits: SearchHit[]) {
  if (!hits.length || !shouldPreferFlowCompanion(question)) {
    return hits;
  }

  const topMetadata = getChunkMetadata(hits[0].chunk);
  if (topMetadata.entryType !== "供应商" && topMetadata.entryType !== "参考") {
    return hits;
  }

  const questionTerms = extractDirectMatchTerms(question);
  const fallbackTerms = questionTerms.length ? questionTerms : extractKeywords(question).filter((term) => term.length >= 2);
  const topCoverage = topicCoverage(fallbackTerms, hits[0].chunk);
  const topCandidates = hits.slice(0, 8);
  const flowCandidates = topCandidates
    .map((hit) => ({
      hit,
      coverage: topicCoverage(fallbackTerms, hit.chunk),
      metadata: getChunkMetadata(hit.chunk)
    }))
    .filter(
      (candidate) =>
        candidate.metadata.entryType === "流程" &&
        candidate.coverage >= Math.max(0.45, topCoverage) &&
        (!topMetadata.category || candidate.metadata.category === topMetadata.category)
    )
    .sort((left, right) => right.coverage - left.coverage || right.hit.score - left.hit.score);

  const promoted = flowCandidates[0]?.hit;
  if (!promoted) {
    return hits;
  }

  return [promoted, ...hits.filter((hit) => hit.chunk.chunkId !== promoted.chunk.chunkId)];
}

function hasExplicitSteps(text: string) {
  return explicitStepPattern.test(text) || sequentialActionPattern.test(text);
}

function isSummaryOnlyEvidence(evidence: KnowledgeChunk[]) {
  if (!evidence.length) {
    return false;
  }

  const hasFlowRow = evidence.some((chunk) => chunk.sourceType === "row" && getChunkMetadata(chunk).entryType === "流程");
  if (hasFlowRow) {
    return false;
  }

  return !evidence.some((chunk) => hasExplicitSteps(chunk.text));
}

function extractOptionLabels(text: string) {
  const options = new Map<number, string>();
  for (const line of text.split(/\n+/)) {
    const match = /^\s*(\d+)[.)、]\s*(.+)$/u.exec(line.trim());
    if (!match) {
      continue;
    }

    const optionNumber = Number(match[1]);
    let label = normalizeSpace(match[2].split(/[：:]/u)[0] ?? "");
    label = label.replace(/\*\*/g, "").replace(/`/g, "").trim();
    if (label) {
      options.set(optionNumber, label);
    }
  }

  return options;
}

function resolveFollowUpQuestion(message: string, history: ChatMessage[]) {
  const normalized = normalizeSpace(message);
  const optionMatch = /^(?:选项|option)\s*(\d+)$/iu.exec(normalized);
  if (!optionMatch) {
    return normalized;
  }

  const optionNumber = Number(optionMatch[1]);
  for (let index = history.length - 1; index >= 0; index -= 1) {
    const historyMessage = history[index];
    if (historyMessage.role !== "assistant") {
      continue;
    }

    const label = extractOptionLabels(historyMessage.content).get(optionNumber);
    if (label) {
      return label;
    }
  }

  return normalized;
}

function detectUnansweredReason(answer: string, hasStrongEvidence: boolean): UnansweredReason | null {
  if (!hasStrongEvidence) {
    return "insufficient_evidence";
  }

  const normalized = normalizeSpace(answer);
  if (!normalized) {
    return "model_declined";
  }

  return /未找到明确依据|没有找到足够依据|没有足够依据|未在.+找到依据|当前知识库里没有找到足够依据|非常抱歉，该问题无法回答。行政相关请咨询Helen（yanhong wang），人事相关请咨询Susie（susie sl huang），如果需要补充问题及答案请联系Kleist（kleist jf liu）。/u.test(
    normalized
  )
    ? "model_declined"
    : null;
}

async function recordQuestionStats(params: {
  knowledgeBaseId: string;
  question: string;
  sessionId: string;
  unansweredReason: UnansweredReason | null;
}) {
  const { knowledgeBaseId, question, sessionId, unansweredReason } = params;
  const stats = (await loadQuestionStats(knowledgeBaseId)) ?? createEmptyQuestionStats(knowledgeBaseId);
  stats.totalQuestions += 1;

  if (unansweredReason) {
    stats.unansweredCount += 1;
    stats.recentUnanswered = [
      {
        question,
        sessionId,
        createdAt: new Date().toISOString(),
        reason: unansweredReason
      },
      ...stats.recentUnanswered
    ].slice(0, 20);
  }

  await saveQuestionStats(stats);
  return stats;
}

function chunkSegments(text: string) {
  const normalized = normalizeSpace(text);
  if (!normalized) {
    return [];
  }

  const splitBySteps = normalized
    .split(/(?=\s*\d+[.)、])/g)
    .map((segment) => normalizeSpace(segment))
    .filter(Boolean);
  const baseSegments = splitBySteps.length > 1 ? splitBySteps : normalized.split(/\n+/).map((segment) => normalizeSpace(segment));
  const segments = baseSegments.filter(Boolean);

  if (!segments.length) {
    return [];
  }

  const chunks: string[] = [];
  let buffer = "";

  for (const segment of segments) {
    if (!buffer) {
      buffer = segment;
      continue;
    }

    if ((buffer + "\n" + segment).length <= 520) {
      buffer = `${buffer}\n${segment}`;
      continue;
    }

    chunks.push(buffer);
    buffer = segment;
  }

  if (buffer) {
    chunks.push(buffer);
  }

  return chunks.flatMap((chunk) =>
    chunk.length > 650
      ? chunk.match(/.{1,480}(?:\s|$)/g)?.map((item) => normalizeSpace(item)).filter(Boolean) ?? [chunk]
      : [chunk]
  );
}

async function buildKnowledgeChunks(
  knowledgeBaseId: string,
  sources: KnowledgeSource[],
  providerMode: string
): Promise<KnowledgeChunk[]> {
  const provider = createProvider(providerMode);
  const drafts: Array<Omit<KnowledgeChunk, "embedding">> = [];

  for (const source of sources) {
    if (source.sourceType === "attachment" && isImageAttachmentKind(source.attachmentKind)) {
      continue;
    }

    const prefix = [`流程 ${source.rowNumber}`, source.title, source.attachmentName].filter(Boolean).join(" · ");
    const baseText =
      source.sourceType === "link" && source.url
        ? normalizeSpace(`${source.text}\n链接地址: ${source.url}`)
        : source.text;

    if (!baseText) {
      continue;
    }

    const segments = chunkSegments(baseText);
    if (!segments.length) {
      continue;
    }

    for (let index = 0; index < segments.length; index += 1) {
      const text = normalizeSpace(`${prefix}\n${segments[index]}`);
      drafts.push({
        chunkId: `${source.sourceId}-chunk-${index + 1}`,
        knowledgeBaseId,
        sourceId: source.sourceId,
        rowNumber: source.rowNumber,
        sourceType: source.sourceType,
        title: source.title,
        entryType: source.entryType,
        category: source.category,
        aliases: source.aliases,
        attachmentName: source.attachmentName,
        url: source.url,
        text,
        keywords: extractKeywords(`${source.title}\n${source.aliases?.join("\n") ?? ""}\n${segments[index]}`)
      });
    }
  }

  const embeddings = await provider.embedMany(drafts.map((draft) => draft.text));
  return drafts.map((draft, index) => ({
    ...draft,
    embedding: embeddings[index] ?? []
  }));
}

async function saveActivatedKnowledgeBase(
  knowledgeBaseId: string,
  originalFileName: string,
  storedFileName: string,
  parsed: Awaited<ReturnType<typeof parseWorkflowWorkbook>>,
  providerMode = createProvider().mode,
  canonicalAttachmentsDir?: string
) {
  const chunks = await buildKnowledgeChunks(knowledgeBaseId, parsed.sources, providerMode);
  const knowledgeBase: KnowledgeBaseRecord = {
    knowledgeBaseId,
    originalFileName,
    storedFileName,
    importedAt: new Date().toISOString(),
    providerMode,
    sourceCount: parsed.sources.length,
    chunkCount: chunks.length,
    canonicalAttachmentsDir,
    versionNotes: parsed.versionNotes,
    sheets: parsed.sheets,
    sources: parsed.sources,
    chunks,
    warnings: parsed.warnings
  };

  await saveKnowledgeBase(knowledgeBase);
  return knowledgeBase;
}

export async function enqueueKnowledgeImport(fileName: string, buffer: Buffer) {
  await ensureStorage();
  const knowledgeBaseId = randomUUID();
  const jobId = randomUUID();
  const storedFileName = `${knowledgeBaseId}-${sanitizeFileName(fileName)}`;
  const job = createJobRecord(jobId, knowledgeBaseId, fileName);

  await writeUploadFile(storedFileName, buffer);
  await saveJob(job);
  const state = await loadState();
  await saveState({ ...state, latestJobId: jobId });

  void runKnowledgeImport({
    job,
    storedFileName,
    originalFileName: fileName,
    buffer
  });

  return job;
}

async function runKnowledgeImport(params: {
  job: ImportJobRecord;
  storedFileName: string;
  originalFileName: string;
  buffer: Buffer;
}) {
  const { job, storedFileName, originalFileName, buffer } = params;
  try {
    job.status = "running";
    await saveJob(job);

    const parsed = await parseWorkflowWorkbook(buffer, job.knowledgeBaseId, originalFileName);
    const provider = createProvider();
    const knowledgeBase = await saveActivatedKnowledgeBase(
      job.knowledgeBaseId,
      originalFileName,
      storedFileName,
      parsed,
      provider.mode
    );
    job.status = "completed";
    job.sourceCount = knowledgeBase.sourceCount;
    job.chunkCount = knowledgeBase.chunkCount;
    job.warnings = knowledgeBase.warnings;
    await saveJob(job);

    const state = await loadState();
    await saveState({
      ...state,
      activeKnowledgeBaseId: job.knowledgeBaseId,
      latestJobId: job.jobId
    });
  } catch (error) {
    job.status = "failed";
    job.error = error instanceof Error ? error.message : "导入知识库时发生未知错误";
    await saveJob(job);
  }
}

export async function syncConfiguredKnowledgeBase() {
  await ensureStorage();
  const fixedSource = readConfiguredFixedSource();
  const state = await loadState();

  if (!fixedSource.configured || !fixedSource.workbookPath || !fixedSource.attachmentsDir) {
    if (state.fixedSource?.configured) {
      await saveState({
        ...state,
        fixedSource
      });
    }

    return {
      status: "disabled" as const
    };
  }

  try {
    const currentKnowledgeBase = state.activeKnowledgeBaseId ? await loadKnowledgeBase(state.activeKnowledgeBaseId) : null;
    const { fingerprint, workbookBuffer } = await computeFixedSourceFingerprint(
      fixedSource.workbookPath,
      fixedSource.attachmentsDir
    );

    if (
      state.fixedSource?.lastSyncFingerprint === fingerprint &&
      state.activeKnowledgeBaseId &&
      !knowledgeBaseNeedsAssetRefresh(currentKnowledgeBase)
    ) {
      await saveState({
        ...state,
        fixedSource: {
          ...fixedSource,
          lastSyncFingerprint: fingerprint,
          lastSyncAt: state.fixedSource?.lastSyncAt,
          syncError: undefined
        }
      });

      return {
        status: "unchanged" as const,
        knowledgeBaseId: state.activeKnowledgeBaseId
      };
    }

    const knowledgeBaseId = randomUUID();
    const parsed = await parseWorkflowWorkbook(
      workbookBuffer,
      knowledgeBaseId,
      path.basename(fixedSource.workbookPath),
      {
        canonicalAttachmentsDir: fixedSource.attachmentsDir
      }
    );
    const knowledgeBase = await saveActivatedKnowledgeBase(
      knowledgeBaseId,
      path.basename(fixedSource.workbookPath),
      `fixed-${path.basename(fixedSource.workbookPath)}`,
      parsed,
      createProvider().mode,
      fixedSource.attachmentsDir
    );

    await saveState({
      ...state,
      activeKnowledgeBaseId: knowledgeBase.knowledgeBaseId,
      fixedSource: {
        ...fixedSource,
        lastSyncFingerprint: fingerprint,
        lastSyncAt: new Date().toISOString(),
        syncError: undefined
      }
    });

    return {
      status: "synced" as const,
      knowledgeBaseId: knowledgeBase.knowledgeBaseId,
      sourceCount: knowledgeBase.sourceCount,
      chunkCount: knowledgeBase.chunkCount
    };
  } catch (error) {
    const syncError = error instanceof Error ? error.message : "固定知识库同步失败";
    await saveState({
      ...state,
      fixedSource: {
        ...fixedSource,
        syncError,
        lastSyncAt: new Date().toISOString()
      }
    });

    return {
      status: "failed" as const,
      error: syncError
    };
  }
}

function keywordOverlapScore(queryKeywords: string[], chunkKeywords: string[]) {
  if (!queryKeywords.length || !chunkKeywords.length) {
    return 0;
  }

  const lookup = new Set(chunkKeywords);
  const overlap = queryKeywords.filter((keyword) => lookup.has(keyword)).length;
  return overlap / Math.sqrt(queryKeywords.length * chunkKeywords.length);
}

function extractAsciiPhrases(question: string) {
  const phrases = new Set<string>();
  for (const match of normalizeSearchText(question).match(/[a-z][a-z0-9/-]*(?:\s+[a-z0-9/-]+)*/g) ?? []) {
    const phrase = match.trim();
    if (phrase.length > 1) {
      phrases.add(phrase);
    }
  }
  return [...phrases];
}

function substringScore(terms: string[], text: string) {
  if (!terms.length) {
    return 0;
  }

  const normalizedText = normalizeSearchText(text);
  let matched = 0;
  let total = 0;

  for (const term of terms) {
    if (term.length < 2) {
      continue;
    }
    const weight = Math.min(term.length, 8);
    total += weight;
    if (normalizedText.includes(term.toLowerCase())) {
      matched += weight;
    }
  }

  return total ? matched / total : 0;
}

function exactTermCoverage(terms: string[], text: string) {
  if (!terms.length) {
    return 0;
  }

  const normalizedText = normalizeSearchText(text);
  let matched = 0;
  let total = 0;

  for (const term of terms) {
    if (term.length < 2) {
      continue;
    }
    const weight = Math.min(term.length, 6);
    total += weight;
    if (normalizedText.includes(term.toLowerCase())) {
      matched += weight;
    }
  }

  return total ? matched / total : 0;
}

function exactTermCoverageForDirectPrefilter(terms: string[], text: string) {
  if (!terms.length) {
    return 0;
  }

  const normalizedText = normalizeSearchText(text);
  const normalizedTextWithoutUrls = stripUrls(normalizedText);
  let matched = 0;
  let total = 0;

  for (const term of terms) {
    if (term.length < 2) {
      continue;
    }
    const weight = Math.min(term.length, 6);
    total += weight;
    const haystack = isShortAlphanumericToken(term) ? normalizedTextWithoutUrls : normalizedText;
    if (haystack.includes(term.toLowerCase())) {
      matched += weight;
    }
  }

  return total ? matched / total : 0;
}

function hasAnswerableEvidence(question: string, hits: SearchHit[], evidence: KnowledgeChunk[]) {
  if (!evidence.length) {
    return false;
  }

  const topScore = hits[0]?.score ?? 0;
  const directTerms = extractDirectMatchTerms(question);
  const keywordTerms = extractKeywords(question).filter((term) => term.length >= 2);
  const combinedEvidence = evidence.map((chunk) => `${chunk.title}\n${chunk.text}`).join("\n");
  const normalizedEvidence = normalizeSearchText(combinedEvidence);
  const matchedDirectTerms = directTerms.filter((term) => normalizedEvidence.includes(term.toLowerCase()));
  const keywordCoverage = substringScore(keywordTerms, combinedEvidence);

  if (directTerms.length) {
    const hasStrongDirectHit = matchedDirectTerms.some((term) => term.length >= 3);
    const hasWeakDirectHit = matchedDirectTerms.some((term) => term.length >= 2);
    return hasStrongDirectHit || hasWeakDirectHit || (topScore >= 0.22 && keywordCoverage >= 0.16);
  }

  return topScore >= 0.18 && keywordCoverage >= 0.18;
}

function isCompareQuestion(question: string) {
  return /区别|不同|对比|比较|差异|vs/i.test(question);
}

function findDirectCandidateRows(knowledgeBase: KnowledgeBaseRecord, directTerms: string[]) {
  if (!directTerms.length) {
    return new Set<number>();
  }

  const rows = knowledgeBase.sources
    .filter((source) => source.sourceType === "row")
    .map((source) => ({
      rowNumber: source.rowNumber,
      titleCoverage: exactTermCoverage(directTerms, source.title),
      textCoverage: exactTermCoverageForDirectPrefilter(directTerms, source.text)
    }));

  const titleMatches = rows.filter((row) => row.titleCoverage > 0);
  if (titleMatches.length) {
    return new Set(titleMatches.map((row) => row.rowNumber));
  }

  const textMatches = rows.filter((row) => row.textCoverage > 0);
  return new Set(textMatches.map((row) => row.rowNumber));
}

async function searchKnowledgeBase(
  knowledgeBase: KnowledgeBaseRecord,
  question: string,
  retrievalDebug?: RetrievalDebugRecord
): Promise<SearchHit[]> {
  const provider = createProvider(knowledgeBase.providerMode);
  const [queryEmbedding] = await provider.embedMany([question]);
  const queryKeywords = extractKeywords(question);
  const directTerms = extractDirectMatchTerms(question);
  const queryPhrases = extractAsciiPhrases(question);
  const directCandidateRows = findDirectCandidateRows(knowledgeBase, directTerms);
  const candidateChunks = directCandidateRows.size
    ? knowledgeBase.chunks.filter((chunk) => directCandidateRows.has(chunk.rowNumber))
    : knowledgeBase.chunks;

  if (retrievalDebug) {
    retrievalDebug.queryKeywords = queryKeywords;
    retrievalDebug.directTerms = directTerms;
    retrievalDebug.queryPhrases = queryPhrases;
    retrievalDebug.directCandidateRows = [...directCandidateRows].sort((left, right) => left - right);
    retrievalDebug.candidateChunkCount = candidateChunks.length;
    retrievalDebug.totalChunkCount = knowledgeBase.chunks.length;
  }

  const hits: ScoredSearchHit[] = candidateChunks.map((chunk) => {
    const semanticScore = cosineSimilarity(queryEmbedding, chunk.embedding);
    const keywordScore = keywordOverlapScore(queryKeywords, chunk.keywords);
    const titleScore = keywordOverlapScore(queryKeywords, extractKeywords(chunk.title));
    const substringMatch = substringScore(queryKeywords, `${chunk.title}\n${chunk.text}`);
    const phraseScore = substringScore(queryPhrases, `${chunk.title}\n${chunk.text}`);
    const exactTitleScore = exactTermCoverage(directTerms, chunk.title);
    const exactBodyScore = exactTermCoverage(directTerms, chunk.text);
    const urlBonus = chunk.url && queryKeywords.some((keyword) => chunk.url?.toLowerCase().includes(keyword)) ? 0.15 : 0;
    const lexicalPresence = Math.max(keywordScore, titleScore, substringMatch, phraseScore);
    const baseScore =
      knowledgeBase.providerMode === "offline"
        ? semanticScore * (lexicalPresence > 0 ? 0.25 : 0.03) +
          keywordScore * 0.34 +
          titleScore * 0.18 +
          substringMatch * 0.15 +
          phraseScore * 0.08 +
          urlBonus
        : semanticScore * 0.7 + keywordScore * 0.16 + titleScore * 0.08 + substringMatch * 0.06 + phraseScore * 0.04 + urlBonus;
    const exactBonus = exactTitleScore * 0.55 + exactBodyScore * 0.28;
    const sourceTypeBonus = chunk.sourceType === "row" ? 0.12 : chunk.sourceType === "attachment" ? -0.02 : -0.04;
    const score = baseScore + exactBonus + sourceTypeBonus;

    return {
      chunk,
      score,
      breakdown: {
        chunkId: chunk.chunkId,
        sourceId: chunk.sourceId,
        rowNumber: chunk.rowNumber,
        title: chunk.title,
        sourceType: chunk.sourceType,
        semanticScore,
        keywordScore,
        titleScore,
        substringMatch,
        phraseScore,
        exactTitleScore,
        exactBodyScore,
        urlBonus,
        sourceTypeBonus,
        baseScore,
        exactBonus,
        finalScore: score
      }
    };
  });

  if (retrievalDebug) {
    retrievalDebug.scoredHits = hits.map((hit) => hit.breakdown);
  }

  const sortedHits = prioritizeTopicHits(
    question,
    [...hits].sort((left, right) => right.score - left.score)
  );

  if (retrievalDebug) {
    retrievalDebug.sortedHitOrder = sortedHits.map((hit) => hit.chunk.chunkId);
  }

  return sortedHits.map((hit) => ({
    chunk: hit.chunk,
    score: hit.score
  }));
}

function selectEvidence(question: string, hits: SearchHit[], retrievalDebug?: RetrievalDebugRecord) {
  const eligibleHits = hits.filter((hit) => hit.score >= retrievalEligibleThreshold);
  const compareMode = isCompareQuestion(question);
  const topicLookup = shouldPreferFlowCompanion(question);

  if (retrievalDebug) {
    retrievalDebug.eligibleHitIds = eligibleHits.map((hit) => hit.chunk.chunkId);
    retrievalDebug.compareMode = compareMode;
    retrievalDebug.topicLookup = topicLookup;
    retrievalDebug.topRowNumber = eligibleHits[0]?.chunk.rowNumber;
    retrievalDebug.dominantTopRow = eligibleHits.length
      ? eligibleHits[0].score >= (eligibleHits[1]?.score ?? 0) + 0.05
      : false;
  }

  if (!eligibleHits.length) {
    return [];
  }

  const evidence: KnowledgeChunk[] = [];
  const seenChunks = new Set<string>();
  const seenRows = new Set<number>();
  const scoreByChunkId = new Map(eligibleHits.map((hit) => [hit.chunk.chunkId, hit.score]));
  const topRowNumber = eligibleHits[0].chunk.rowNumber;
  const evidenceLimit = compareMode ? 3 : 4;
  const dominantTopRow = eligibleHits[0].score >= (eligibleHits[1]?.score ?? 0) + 0.05;

  const finalizeEvidence = () => {
    const selectedEvidence = evidence.slice(0, evidenceLimit);
    if (retrievalDebug) {
      retrievalDebug.selectedEvidenceChunkIds = selectedEvidence.map((chunk) => chunk.chunkId);
    }
    return selectedEvidence;
  };

  const pushChunk = (chunk: KnowledgeChunk | undefined, reason: string) => {
    if (!chunk || seenChunks.has(chunk.chunkId)) {
      return;
    }
    evidence.push(chunk);
    seenChunks.add(chunk.chunkId);
    seenRows.add(chunk.rowNumber);
    if (retrievalDebug) {
      retrievalDebug.evidenceSelectionSteps.push({
        reason,
        chunkId: chunk.chunkId,
        sourceId: chunk.sourceId,
        rowNumber: chunk.rowNumber,
        title: chunk.title,
        sourceType: chunk.sourceType,
        score: scoreByChunkId.get(chunk.chunkId) ?? 0
      } satisfies RetrievalEvidenceSelectionStep);
    }
  };

  pushChunk(
    eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "row")?.chunk ??
      eligibleHits[0]?.chunk,
    "initial-top-row"
  );

  if (topicLookup) {
    const topMetadata = getChunkMetadata(evidence[0] ?? eligibleHits[0].chunk);
    if (topMetadata.entryType === "流程") {
      const topicTerms = extractDirectMatchTerms(question);
      const fallbackTerms = topicTerms.length ? topicTerms : extractKeywords(question).filter((term) => term.length >= 2);
      const topCoverage = topicCoverage(fallbackTerms, evidence[0] ?? eligibleHits[0].chunk);
      const companionRow = eligibleHits.find((hit) => {
        if (hit.chunk.rowNumber === topRowNumber || hit.chunk.sourceType !== "row") {
          return false;
        }

        const metadata = getChunkMetadata(hit.chunk);
        if (metadata.entryType !== "供应商" && metadata.entryType !== "参考" && metadata.entryType !== "流程") {
          return false;
        }

        return (
          topicCoverage(fallbackTerms, hit.chunk) >= Math.max(0.45, topCoverage * 0.6) &&
          (!topMetadata.category || metadata.category === topMetadata.category)
        );
      });

      pushChunk(companionRow?.chunk, "topic-lookup-companion-row");
    }

    pushChunk(eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk, "topic-lookup-top-row-attachment");
    return finalizeEvidence();
  }

  if (!compareMode && !topicLookup) {
    pushChunk(
      eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk,
      "top-row-attachment"
    );
  }

  if (compareMode) {
    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (hit.chunk.sourceType !== "row" || seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk, "compare-add-row");
    }

    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (seenChunks.has(hit.chunk.chunkId) || !seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk, "compare-add-related-chunk");
    }
  } else if (!dominantTopRow) {
    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (hit.chunk.sourceType !== "row" || seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk, "balanced-add-row");
    }

    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (seenChunks.has(hit.chunk.chunkId)) {
        continue;
      }
      pushChunk(hit.chunk, "balanced-add-chunk");
    }
  }

  if (!compareMode && topicLookup) {
    pushChunk(
      eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk,
      "topic-lookup-top-row-attachment"
    );
  }

  return finalizeEvidence();
}

function buildCitationImages(knowledgeBase: KnowledgeBaseRecord, rowNumber: number, limit = 3): CitationImage[] {
  return knowledgeBase.sources
    .filter(
      (source) =>
        source.rowNumber === rowNumber &&
        source.sourceType === "attachment" &&
        isImageAttachmentKind(source.attachmentKind) &&
        source.attachmentRelativePath
    )
    .sort((left, right) =>
      (left.attachmentName ?? "").localeCompare(right.attachmentName ?? "", "zh-Hans-CN", { numeric: true })
    )
    .slice(0, limit)
    .map((source) => ({
      sourceId: source.sourceId,
      attachmentName: source.attachmentName,
      label: source.attachmentDescription || source.attachmentName || `流程 ${rowNumber} 图片`,
      url: `/api/knowledge/assets/${encodeURIComponent(knowledgeBase.knowledgeBaseId)}/${encodeURIComponent(source.sourceId)}`
    }));
}

function buildCitations(chunks: KnowledgeChunk[], knowledgeBase: KnowledgeBaseRecord) {
  const citations: Citation[] = [];
  const seenSources = new Set<string>();
  const seenRows = new Set<number>();

  for (const chunk of chunks) {
    if (chunk.sourceType !== "row") {
      continue;
    }
    if (seenRows.has(chunk.rowNumber) || seenSources.has(chunk.sourceId)) {
      continue;
    }

    seenRows.add(chunk.rowNumber);
    seenSources.add(chunk.sourceId);
    citations.push({
      sourceId: chunk.sourceId,
      rowNumber: chunk.rowNumber,
      title: chunk.title,
      attachmentName: chunk.attachmentName,
      url: chunk.url,
      snippet: clip(chunk.text),
      images: buildCitationImages(knowledgeBase, chunk.rowNumber)
    });

    if (citations.length >= 3) {
      break;
    }
  }

  for (const chunk of chunks) {
    if (seenSources.has(chunk.sourceId)) {
      continue;
    }

    seenSources.add(chunk.sourceId);
    citations.push({
      sourceId: chunk.sourceId,
      rowNumber: chunk.rowNumber,
      title: chunk.title,
      attachmentName: chunk.attachmentName,
      url: chunk.url,
      snippet: clip(chunk.text),
      images: buildCitationImages(knowledgeBase, chunk.rowNumber)
    });

    if (citations.length >= 4) {
      break;
    }
  }

  return citations;
}

export async function answerQuestion(message: string, sessionId?: string) {
  const state = await loadState();

  let session: ChatSession | null = sessionId ? await loadSession(sessionId) : null;
  if (sessionId && !session) {
    throw new Error("会话不存在或已失效，请刷新页面后重试。");
  }

  const knowledgeBaseId = session?.knowledgeBaseId ?? state.activeKnowledgeBaseId;
  if (!knowledgeBaseId) {
    throw new Error("当前没有活动知识库，请检查固定知识源配置。");
  }

  const knowledgeBase = await loadKnowledgeBase(knowledgeBaseId);
  if (!knowledgeBase) {
    throw new Error("活动知识库文件不存在，请检查固定知识源同步。");
  }

  if (!session) {
    session = {
      sessionId: randomUUID(),
      knowledgeBaseId,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: []
    };
  }

  const resolvedQuestion = resolveFollowUpQuestion(message, session.messages);
  const retrievalDebugEnabled = isRetrievalDebugEnabled();
  const retrievalDebugCreatedAt = new Date().toISOString();
  const turnIndex = session.messages.filter((item) => item.role === "user").length + 1;
  const retrievalDebug = retrievalDebugEnabled
    ? createRetrievalDebugRecord({
        traceId: randomUUID(),
        createdAt: retrievalDebugCreatedAt,
        sessionId: session.sessionId,
        turnIndex,
        knowledgeBaseId,
        originalQuestion: message,
        resolvedQuestion,
        providerMode: knowledgeBase.providerMode
      })
    : undefined;
  const answerProvider = createAnswerProvider(knowledgeBase.providerMode);

  const hits = await searchKnowledgeBase(knowledgeBase, resolvedQuestion, retrievalDebug);
  const evidence = selectEvidence(resolvedQuestion, hits, retrievalDebug);
  const citations = buildCitations(evidence, knowledgeBase);
  const hasStrongEvidence = hasAnswerableEvidence(resolvedQuestion, hits, evidence);
  const summaryOnlyEvidence = isSummaryOnlyEvidence(evidence);
  const shouldCallModel = hasStrongEvidence && !summaryOnlyEvidence;

  const answerResult = shouldCallModel
    ? await answerProvider.answer({
        question: resolvedQuestion,
        history: session?.messages ?? [],
        evidence,
        citations
      })
    : { answer: noEvidenceAnswer };
  let answer = answerResult.answer;
  const unansweredReason = detectUnansweredReason(answer, shouldCallModel);
  if (unansweredReason) {
    answer = noEvidenceAnswer;
  }
  const answered = !unansweredReason;

  let retrievalDebugRef: ChatMessage["retrievalDebug"];
  if (retrievalDebug) {
    retrievalDebug.hasStrongEvidence = hasStrongEvidence;
    retrievalDebug.summaryOnlyEvidence = summaryOnlyEvidence;
    retrievalDebug.shouldCallModel = shouldCallModel;
    retrievalDebug.modelRequest = answerResult.modelRequest;
    retrievalDebug.answered = answered;
    retrievalDebug.unansweredReason = unansweredReason;
    retrievalDebug.citationRowNumbers = citations.map((citation) => citation.rowNumber);
    retrievalDebug.fileName = buildRetrievalDebugFileName({
      createdAt: retrievalDebug.createdAt,
      sessionId: retrievalDebug.sessionId,
      turnIndex: retrievalDebug.turnIndex,
      traceId: retrievalDebug.traceId
    });
    await saveRetrievalDebug(retrievalDebug);
    retrievalDebugRef = {
      traceId: retrievalDebug.traceId,
      fileName: retrievalDebug.fileName,
      createdAt: retrievalDebug.createdAt
    };
  }

  const nextMessages: ChatMessage[] = [
    ...session.messages,
    {
      role: "user" as const,
      content: message,
      createdAt: new Date().toISOString()
    },
    {
      role: "assistant" as const,
      content: answer,
      createdAt: new Date().toISOString(),
      citations,
      retrievalDebug: retrievalDebugRef
    }
  ].slice(-12);
  session.messages = nextMessages;
  await saveSession(session);
  const questionStats: QuestionStatsRecord = await recordQuestionStats({
    knowledgeBaseId,
    question: message,
    sessionId: session.sessionId,
    unansweredReason
  });

  return {
    sessionId: session.sessionId,
    knowledgeBaseId,
    answer,
    citations,
    providerMode: answerProvider.mode,
    answered,
    questionStats,
    modelRequest: answerResult.modelRequest
  };
}

function knowledgeBaseNeedsAssetRefresh(knowledgeBase: KnowledgeBaseRecord | null) {
  if (!knowledgeBase) {
    return true;
  }

  const imageSources = knowledgeBase.sources.filter((source) => isImageAttachmentKind(source.attachmentKind));
  if (!imageSources.length) {
    return false;
  }

  const imageSourceIds = new Set(imageSources.map((source) => source.sourceId));
  return (
    !knowledgeBase.canonicalAttachmentsDir ||
    imageSources.some((source) => !source.attachmentRelativePath) ||
    knowledgeBase.chunks.some((chunk) => imageSourceIds.has(chunk.sourceId))
  );
}

export async function resolveKnowledgeAsset(knowledgeBaseId: string, sourceId: string) {
  const knowledgeBase = await loadKnowledgeBase(knowledgeBaseId);
  if (!knowledgeBase) {
    return null;
  }

  const source = knowledgeBase.sources.find((item) => item.sourceId === sourceId);
  if (
    !source ||
    source.sourceType !== "attachment" ||
    !isImageAttachmentKind(source.attachmentKind) ||
    !source.attachmentRelativePath
  ) {
    return null;
  }

  const state = await loadState();
  const attachmentsDir =
    knowledgeBase.canonicalAttachmentsDir ||
    (state.activeKnowledgeBaseId === knowledgeBaseId ? state.fixedSource?.attachmentsDir : undefined);
  if (!attachmentsDir) {
    return null;
  }

  const filePath = resolveCanonicalAttachmentPath(attachmentsDir, source.attachmentRelativePath);
  try {
    await fs.access(filePath);
  } catch {
    return null;
  }

  return {
    filePath,
    fileName: source.attachmentName ?? path.basename(filePath)
  };
}

export async function resolveUserDocument(docId: string) {
  const config = userDocumentConfigs.find((item) => item.id === docId);
  if (!config) {
    return null;
  }

  const filePath = readConfiguredUserDocumentPath(config.envName, config.defaultPath);
  if (!(await pathExists(filePath))) {
    return null;
  }

  return {
    filePath,
    fileName: path.basename(filePath)
  };
}

export async function getActiveKnowledgeResponse() {
  const response = await buildActiveKnowledgeResponse();
  return {
    ...response,
    documentLinks: await listUserDocumentLinks()
  };
}

export async function getImportJob(jobId: string) {
  return loadJob(jobId);
}
