export type SourceType = "row" | "attachment" | "link";
export type AttachmentKind = "docx" | "pptx" | "xlsx" | "doc" | "png" | "jpg" | "jpeg" | "webp" | "unknown";
export type JobStatus = "queued" | "running" | "completed" | "failed";
export type EntryType = "流程" | "联系人" | "供应商" | "参考" | "系统链接";

export interface KnowledgeSource {
  sourceId: string;
  knowledgeBaseId: string;
  rowNumber: number;
  sourceType: SourceType;
  title: string;
  entryType?: EntryType;
  category?: string;
  aliases?: string[];
  relatedForm?: string;
  attachmentName?: string;
  attachmentKind?: AttachmentKind;
  attachmentRelativePath?: string;
  attachmentDescription?: string;
  url?: string;
  text: string;
  parentSourceId?: string;
  parseWarning?: string;
}

export interface KnowledgeChunk {
  chunkId: string;
  knowledgeBaseId: string;
  sourceId: string;
  rowNumber: number;
  sourceType: SourceType;
  title: string;
  entryType?: EntryType;
  category?: string;
  aliases?: string[];
  attachmentName?: string;
  url?: string;
  text: string;
  embedding: number[];
  keywords: string[];
}

export interface Citation {
  sourceId: string;
  rowNumber: number;
  title: string;
  attachmentName?: string;
  url?: string;
  snippet: string;
  images?: CitationImage[];
}

export interface CitationImage {
  sourceId: string;
  attachmentName?: string;
  label: string;
  url: string;
}

export interface RetrievalDebugRef {
  traceId: string;
  fileName: string;
  createdAt: string;
}

export interface RetrievalScoredHit {
  chunkId: string;
  sourceId: string;
  rowNumber: number;
  title: string;
  sourceType: SourceType;
  semanticScore: number;
  keywordScore: number;
  titleScore: number;
  substringMatch: number;
  phraseScore: number;
  exactTitleScore: number;
  exactBodyScore: number;
  urlBonus: number;
  sourceTypeBonus: number;
  baseScore: number;
  exactBonus: number;
  finalScore: number;
}

export interface RetrievalEvidenceSelectionStep {
  reason: string;
  chunkId: string;
  sourceId: string;
  rowNumber: number;
  title: string;
  sourceType: SourceType;
  score: number;
}

export interface RetrievalDebugRecord {
  traceId: string;
  fileName: string;
  createdAt: string;
  sessionId: string;
  turnIndex: number;
  knowledgeBaseId: string;
  originalQuestion: string;
  resolvedQuestion: string;
  providerMode: string;
  retrievalDebugEnabled: boolean;
  queryKeywords: string[];
  directTerms: string[];
  queryPhrases: string[];
  directCandidateRows: number[];
  candidateChunkCount: number;
  totalChunkCount: number;
  scoredHits: RetrievalScoredHit[];
  sortedHitOrder: string[];
  eligibleThreshold: number;
  eligibleHitIds: string[];
  topicLookup: boolean;
  compareMode: boolean;
  topRowNumber?: number;
  dominantTopRow: boolean;
  evidenceSelectionSteps: RetrievalEvidenceSelectionStep[];
  selectedEvidenceChunkIds: string[];
  hasStrongEvidence: boolean;
  summaryOnlyEvidence: boolean;
  shouldCallModel: boolean;
  modelRequest?: ModelRequestDebug;
  answered: boolean;
  unansweredReason?: UnansweredReason | null;
  citationRowNumbers: number[];
}

export interface KnowledgeBaseRecord {
  knowledgeBaseId: string;
  originalFileName: string;
  storedFileName: string;
  importedAt: string;
  providerMode: string;
  sourceCount: number;
  chunkCount: number;
  canonicalAttachmentsDir?: string;
  versionNotes?: string;
  sheets: string[];
  sources: KnowledgeSource[];
  chunks: KnowledgeChunk[];
  warnings: string[];
}

export interface ImportJobRecord {
  jobId: string;
  knowledgeBaseId: string;
  fileName: string;
  status: JobStatus;
  createdAt: string;
  updatedAt: string;
  sourceCount: number;
  chunkCount: number;
  warnings: string[];
  error?: string;
}

export interface ActiveKnowledgeResponse {
  activeKnowledgeBaseId?: string;
  knowledgeBase?: Omit<KnowledgeBaseRecord, "sources" | "chunks">;
  latestJob?: ImportJobRecord;
  questionStats?: QuestionStatsRecord;
  fixedSource?: FixedKnowledgeSourceStatus;
}

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  createdAt: string;
  citations?: Citation[];
  retrievalDebug?: RetrievalDebugRef;
}

export interface ChatSession {
  sessionId: string;
  knowledgeBaseId: string;
  createdAt: string;
  updatedAt: string;
  messages: ChatMessage[];
}

export type UnansweredReason = "insufficient_evidence" | "model_declined";

export interface UnansweredQuestionRecord {
  question: string;
  sessionId: string;
  createdAt: string;
  reason: UnansweredReason;
}

export interface QuestionStatsRecord {
  knowledgeBaseId: string;
  totalQuestions: number;
  unansweredCount: number;
  recentUnanswered: UnansweredQuestionRecord[];
}

export interface ModelRequestDebug {
  providerMode: string;
  endpoint: string;
  body: string;
}

export interface AuthStatusResponse {
  enabled: boolean;
  authenticated: boolean;
}

export interface FixedKnowledgeSourceStatus {
  configured: boolean;
  workbookPath?: string;
  attachmentsDir?: string;
  lastSyncAt?: string;
  lastSyncFingerprint?: string;
  syncError?: string;
}

export interface ImportResponse {
  jobId: string;
  knowledgeBaseId: string;
  status: JobStatus;
}

export interface ChatRequestBody {
  sessionId?: string;
  message: string;
}

export interface ChatResponseBody {
  sessionId: string;
  knowledgeBaseId: string;
  answer: string;
  citations: Citation[];
  providerMode: string;
  answered: boolean;
  questionStats: QuestionStatsRecord;
  modelRequest?: ModelRequestDebug;
}
