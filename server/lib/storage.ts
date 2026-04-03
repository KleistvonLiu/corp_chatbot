import { promises as fs } from "node:fs";
import path from "node:path";
import type {
  ActiveKnowledgeResponse,
  ChatSession,
  FixedKnowledgeSourceStatus,
  ImportJobRecord,
  KnowledgeBaseRecord,
  QuestionStatsRecord
} from "../../shared/contracts";

interface AppState {
  activeKnowledgeBaseId?: string;
  latestJobId?: string;
  fixedSource?: FixedKnowledgeSourceStatus;
}

const rootDir = process.cwd();
const dataDir = path.join(rootDir, "data");
const uploadsDir = path.join(dataDir, "uploads");
const knowledgeDir = path.join(dataDir, "knowledge-bases");
const jobsDir = path.join(dataDir, "jobs");
const sessionsDir = path.join(dataDir, "sessions");
const questionStatsDir = path.join(dataDir, "question-stats");
const stateFile = path.join(dataDir, "state.json");

function nowIso() {
  return new Date().toISOString();
}

function formatSessionTimestamp(timestamp: string) {
  const date = new Date(timestamp);
  if (Number.isNaN(date.getTime())) {
    return "unknown-time";
  }

  const iso = date.toISOString();
  return iso.replace(/[:-]/g, "").replace("T", "-").replace(/\.\d{3}Z$/, (match) => match.replace(".", "-"));
}

export function sanitizeFileName(fileName: string): string {
  return fileName.replace(/[^a-zA-Z0-9._-]+/g, "_");
}

async function ensureDir(dirPath: string) {
  await fs.mkdir(dirPath, { recursive: true });
}

function isLegacySessionFileName(fileName: string) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.json$/i.test(fileName);
}

async function migrateLegacySessionFiles() {
  await ensureDir(sessionsDir);
  const fileNames = await fs.readdir(sessionsDir);

  for (const fileName of fileNames) {
    if (!isLegacySessionFileName(fileName)) {
      continue;
    }

    const legacyPath = path.join(sessionsDir, fileName);
    const session = await readJson<ChatSession>(legacyPath);
    if (!session) {
      continue;
    }

    const targetPath = timestampedSessionPath(session);
    if (targetPath === legacyPath) {
      continue;
    }

    try {
      await fs.access(targetPath);
      await fs.rm(legacyPath, { force: true });
    } catch {
      await fs.rename(legacyPath, targetPath);
    }
  }
}

export async function ensureStorage() {
  await Promise.all([
    ensureDir(dataDir),
    ensureDir(uploadsDir),
    ensureDir(knowledgeDir),
    ensureDir(jobsDir),
    ensureDir(sessionsDir),
    ensureDir(questionStatsDir)
  ]);
  await migrateLegacySessionFiles();
}

async function readJson<T>(filePath: string): Promise<T | null> {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw) as T;
  } catch (error) {
    const isMissing = error instanceof Error && "code" in error && error.code === "ENOENT";
    if (isMissing) {
      return null;
    }

    throw error;
  }
}

async function writeJson(filePath: string, payload: unknown) {
  await ensureDir(path.dirname(filePath));
  const tempPath = `${filePath}.${process.pid}.tmp`;
  await fs.writeFile(tempPath, JSON.stringify(payload, null, 2), "utf8");
  await fs.rename(tempPath, filePath);
}

export async function writeUploadFile(storedFileName: string, buffer: Buffer) {
  await ensureStorage();
  await fs.writeFile(path.join(uploadsDir, storedFileName), buffer);
}

export function knowledgeBasePath(knowledgeBaseId: string) {
  return path.join(knowledgeDir, `${knowledgeBaseId}.json`);
}

export async function loadKnowledgeBase(knowledgeBaseId: string) {
  return readJson<KnowledgeBaseRecord>(knowledgeBasePath(knowledgeBaseId));
}

export async function saveKnowledgeBase(record: KnowledgeBaseRecord) {
  await writeJson(knowledgeBasePath(record.knowledgeBaseId), record);
}

export function jobPath(jobId: string) {
  return path.join(jobsDir, `${jobId}.json`);
}

export async function loadJob(jobId: string) {
  return readJson<ImportJobRecord>(jobPath(jobId));
}

export async function saveJob(job: ImportJobRecord) {
  job.updatedAt = nowIso();
  await writeJson(jobPath(job.jobId), job);
}

export function sessionPath(sessionId: string) {
  return path.join(sessionsDir, `${sessionId}.json`);
}

function timestampedSessionPath(session: Pick<ChatSession, "sessionId" | "createdAt">) {
  const fileName = `${formatSessionTimestamp(session.createdAt)}-${session.sessionId}.json`;
  return path.join(sessionsDir, fileName);
}

async function findTimestampedSessionPath(sessionId: string) {
  await ensureDir(sessionsDir);
  const fileNames = await fs.readdir(sessionsDir);
  const matched = fileNames
    .filter((fileName) => fileName.endsWith(`-${sessionId}.json`))
    .sort();

  return matched.length ? path.join(sessionsDir, matched[0]) : null;
}

export async function loadSession(sessionId: string) {
  const timestampedPath = await findTimestampedSessionPath(sessionId);
  if (timestampedPath) {
    return readJson<ChatSession>(timestampedPath);
  }

  return readJson<ChatSession>(sessionPath(sessionId));
}

export async function saveSession(session: ChatSession) {
  session.updatedAt = nowIso();
  const desiredPath = timestampedSessionPath(session);
  const legacyPath = sessionPath(session.sessionId);
  const existingTimestampedPath = await findTimestampedSessionPath(session.sessionId);
  const currentPath = existingTimestampedPath ?? legacyPath;

  await writeJson(desiredPath, session);

  if (currentPath !== desiredPath) {
    await fs.rm(currentPath, { force: true });
  }
}

export function questionStatsPath(knowledgeBaseId: string) {
  return path.join(questionStatsDir, `${knowledgeBaseId}.json`);
}

export function createEmptyQuestionStats(knowledgeBaseId: string): QuestionStatsRecord {
  return {
    knowledgeBaseId,
    totalQuestions: 0,
    unansweredCount: 0,
    recentUnanswered: []
  };
}

export async function loadQuestionStats(knowledgeBaseId: string) {
  return (await readJson<QuestionStatsRecord>(questionStatsPath(knowledgeBaseId))) ?? createEmptyQuestionStats(knowledgeBaseId);
}

export async function saveQuestionStats(stats: QuestionStatsRecord) {
  await writeJson(questionStatsPath(stats.knowledgeBaseId), stats);
}

export async function loadState() {
  return (await readJson<AppState>(stateFile)) ?? {};
}

export async function saveState(next: AppState) {
  await writeJson(stateFile, next);
}

export function createJobRecord(jobId: string, knowledgeBaseId: string, fileName: string): ImportJobRecord {
  const timestamp = nowIso();
  return {
    jobId,
    knowledgeBaseId,
    fileName,
    status: "queued",
    createdAt: timestamp,
    updatedAt: timestamp,
    sourceCount: 0,
    chunkCount: 0,
    warnings: []
  };
}

export async function buildActiveKnowledgeResponse(): Promise<ActiveKnowledgeResponse> {
  const state = await loadState();
  const knowledgeBase = state.activeKnowledgeBaseId
    ? await loadKnowledgeBase(state.activeKnowledgeBaseId)
    : null;
  const latestJob = state.latestJobId ? await loadJob(state.latestJobId) : null;

  return {
    activeKnowledgeBaseId: knowledgeBase?.knowledgeBaseId,
    knowledgeBase: knowledgeBase
      ? {
          knowledgeBaseId: knowledgeBase.knowledgeBaseId,
          originalFileName: knowledgeBase.originalFileName,
          storedFileName: knowledgeBase.storedFileName,
          importedAt: knowledgeBase.importedAt,
          providerMode: knowledgeBase.providerMode,
          sourceCount: knowledgeBase.sourceCount,
          chunkCount: knowledgeBase.chunkCount,
          versionNotes: knowledgeBase.versionNotes,
          sheets: knowledgeBase.sheets,
          warnings: knowledgeBase.warnings
        }
      : undefined,
    latestJob: latestJob ?? undefined,
    questionStats: knowledgeBase ? await loadQuestionStats(knowledgeBase.knowledgeBaseId) : undefined,
    fixedSource: state.fixedSource
  };
}
