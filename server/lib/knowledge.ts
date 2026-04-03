import { randomUUID } from "node:crypto";
import type {
  ChatMessage,
  ChatSession,
  Citation,
  EntryType,
  ImportJobRecord,
  KnowledgeBaseRecord,
  KnowledgeChunk,
  KnowledgeSource,
  QuestionStatsRecord,
  UnansweredReason
} from "../../shared/contracts";
import {
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
  saveSession,
  saveState,
  writeUploadFile
} from "./storage";
import { createAnswerProvider, createProvider, cosineSimilarity, extractDirectMatchTerms, extractKeywords } from "./providers";
import { parseWorkflowWorkbook } from "./parsers";

interface SearchHit {
  chunk: KnowledgeChunk;
  score: number;
}

interface ChunkMetadata {
  entryType?: EntryType;
  category?: string;
  aliases: string[];
}

const noEvidenceAnswer =
  "当前知识库里没有找到足够依据来回答这个问题。请换一个更具体的问法，或者先补充相关流程文档。";
const topicActionPattern = /怎么|如何|谁|联系|申请|审批|安装|报价|区别|对比|比较|差异|流程|购买|下单|负责/u;
const procurementIntentPattern = /买|购买|采购|报价|下单/u;
const referenceLookupPattern = /供应商|联系人|列表|系统链接|参考/u;
const explicitStepPattern = /(?:^|\n)\s*(?:\d+[.)、]|[-•])/u;
const sequentialActionPattern = /首先|然后|之后|最后|第一|第二|第三|第四|第五/u;

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

function clip(text: string, limit = 220) {
  const normalized = normalizeSpace(text);
  if (normalized.length <= limit) {
    return normalized;
  }
  return `${normalized.slice(0, limit - 1)}…`;
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

  return /未找到明确依据|没有找到足够依据|没有足够依据|未在.+找到依据|当前知识库里没有找到足够依据/u.test(
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

    const provider = createProvider();
    const parsed = await parseWorkflowWorkbook(buffer, job.knowledgeBaseId, originalFileName);
    const chunks = await buildKnowledgeChunks(job.knowledgeBaseId, parsed.sources, provider.mode);

    const knowledgeBase: KnowledgeBaseRecord = {
      knowledgeBaseId: job.knowledgeBaseId,
      originalFileName,
      storedFileName,
      importedAt: new Date().toISOString(),
      providerMode: provider.mode,
      sourceCount: parsed.sources.length,
      chunkCount: chunks.length,
      versionNotes: parsed.versionNotes,
      sheets: parsed.sheets,
      sources: parsed.sources,
      chunks,
      warnings: parsed.warnings
    };

    await saveKnowledgeBase(knowledgeBase);
    job.status = "completed";
    job.sourceCount = parsed.sources.length;
    job.chunkCount = chunks.length;
    job.warnings = parsed.warnings;
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
      textCoverage: exactTermCoverage(directTerms, source.text)
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
  question: string
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

  const hits = candidateChunks.map((chunk) => {
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
      score
    };
  });

  return prioritizeTopicHits(
    question,
    hits.sort((left, right) => right.score - left.score)
  );
}

function selectEvidence(question: string, hits: SearchHit[]) {
  const eligibleHits = hits.filter((hit) => hit.score >= 0.08);
  if (!eligibleHits.length) {
    return [];
  }

  const evidence: KnowledgeChunk[] = [];
  const seenChunks = new Set<string>();
  const seenRows = new Set<number>();
  const topRowNumber = eligibleHits[0].chunk.rowNumber;
  const compareMode = isCompareQuestion(question);
  const topicLookup = shouldPreferFlowCompanion(question);
  const evidenceLimit = compareMode ? 3 : 4;
  const dominantTopRow = eligibleHits[0].score >= (eligibleHits[1]?.score ?? 0) + 0.05;

  const pushChunk = (chunk?: KnowledgeChunk) => {
    if (!chunk || seenChunks.has(chunk.chunkId)) {
      return;
    }
    evidence.push(chunk);
    seenChunks.add(chunk.chunkId);
    seenRows.add(chunk.rowNumber);
  };

  pushChunk(
    eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "row")?.chunk ??
      eligibleHits[0]?.chunk
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

      pushChunk(companionRow?.chunk);
    }

    pushChunk(eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk);
    return evidence.slice(0, evidenceLimit);
  }

  if (!compareMode && !topicLookup) {
    pushChunk(eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk);
  }

  if (compareMode) {
    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (hit.chunk.sourceType !== "row" || seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk);
    }

    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (seenChunks.has(hit.chunk.chunkId) || !seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk);
    }
  } else if (!dominantTopRow) {
    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (hit.chunk.sourceType !== "row" || seenRows.has(hit.chunk.rowNumber)) {
        continue;
      }
      pushChunk(hit.chunk);
    }

    for (const hit of eligibleHits) {
      if (evidence.length >= evidenceLimit) {
        break;
      }
      if (seenChunks.has(hit.chunk.chunkId)) {
        continue;
      }
      pushChunk(hit.chunk);
    }
  }

  if (!compareMode && topicLookup) {
    pushChunk(eligibleHits.find((hit) => hit.chunk.rowNumber === topRowNumber && hit.chunk.sourceType === "attachment")?.chunk);
  }

  return evidence.slice(0, evidenceLimit);
}

function buildCitations(chunks: KnowledgeChunk[]) {
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
      snippet: clip(chunk.text)
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
      snippet: clip(chunk.text)
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
    throw new Error("当前没有活动知识库，请先上传并导入 Excel 文件。");
  }

  const knowledgeBase = await loadKnowledgeBase(knowledgeBaseId);
  if (!knowledgeBase) {
    throw new Error("活动知识库文件不存在，请重新导入。");
  }

  const answerProvider = createAnswerProvider(knowledgeBase.providerMode);
  const resolvedQuestion = resolveFollowUpQuestion(message, session?.messages ?? []);

  const hits = await searchKnowledgeBase(knowledgeBase, resolvedQuestion);
  const evidence = selectEvidence(resolvedQuestion, hits);
  const citations = buildCitations(evidence);
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
  const answer = answerResult.answer;
  const unansweredReason = detectUnansweredReason(answer, shouldCallModel);

  if (!session) {
    session = {
      sessionId: randomUUID(),
      knowledgeBaseId,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: []
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
      citations
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
    answered: !unansweredReason,
    questionStats,
    modelRequest: answerResult.modelRequest
  };
}

export async function getActiveKnowledgeResponse() {
  return buildActiveKnowledgeResponse();
}

export async function getImportJob(jobId: string) {
  return loadJob(jobId);
}
