import { promises as fs } from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";
import type { ChatMessage, ChatSession } from "../shared/contracts";

interface CliOptions {
  fromDate: string;
  outputPath: string;
  sessionsDir: string;
}

interface ExportRow {
  sessionFileName: string;
  sessionId: string;
  knowledgeBaseId: string;
  sessionCreatedAt: string;
  sessionUpdatedAt: string;
  turnIndex: number;
  questionCreatedAt: string;
  question: string;
  answerCreatedAt: string;
  answer: string;
}

function printUsage() {
  console.error(
    [
      "用法:",
      "  npm run export:sessions -- <开始日期YYYYMMDD> <输出路径.xlsx|.csv> [--sessions-dir <目录>]",
      "",
      "示例:",
      "  npm run export:sessions -- 20260420 /tmp/session-qa-20260420.xlsx",
      "  npm run export:sessions -- 20260420 ./exports/session-qa-20260420.csv"
    ].join("\n")
  );
}

function resolveOutputPath(rawPath: string) {
  const resolved = path.resolve(rawPath);
  return path.extname(resolved) ? resolved : `${resolved}.xlsx`;
}

function parseCliArgs(argv: string[]): CliOptions {
  const positional: string[] = [];
  let sessionsDir = path.join(process.cwd(), "data", "sessions");

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--sessions-dir") {
      const value = argv[index + 1];
      if (!value) {
        throw new Error("--sessions-dir 缺少目录参数。");
      }
      sessionsDir = value;
      index += 1;
      continue;
    }

    positional.push(arg);
  }

  if (positional.length < 2) {
    throw new Error("参数不足。");
  }

  const [fromDate, outputPath] = positional;
  if (!/^\d{8}$/.test(fromDate)) {
    throw new Error(`开始日期格式错误: ${fromDate}，需要 YYYYMMDD。`);
  }

  return {
    fromDate,
    outputPath: resolveOutputPath(outputPath),
    sessionsDir: path.resolve(sessionsDir)
  };
}

async function readSession(filePath: string) {
  const raw = await fs.readFile(filePath, "utf8");
  return JSON.parse(raw) as ChatSession;
}

function buildExportRows(sessionFileName: string, session: ChatSession): ExportRow[] {
  const rows: ExportRow[] = [];
  let turnIndex = 0;
  let pendingUser: ChatMessage | null = null;
  let pendingAssistants: ChatMessage[] = [];

  const flushTurn = () => {
    if (!pendingUser && pendingAssistants.length === 0) {
      return;
    }

    turnIndex += 1;
    rows.push({
      sessionFileName,
      sessionId: session.sessionId,
      knowledgeBaseId: session.knowledgeBaseId,
      sessionCreatedAt: session.createdAt,
      sessionUpdatedAt: session.updatedAt,
      turnIndex,
      questionCreatedAt: pendingUser?.createdAt ?? "",
      question: pendingUser?.content ?? "",
      answerCreatedAt: pendingAssistants[0]?.createdAt ?? "",
      answer: pendingAssistants.map((message) => message.content).join("\n\n")
    });

    pendingUser = null;
    pendingAssistants = [];
  };

  for (const message of session.messages) {
    if (message.role === "user") {
      flushTurn();
      pendingUser = message;
      continue;
    }

    pendingAssistants.push(message);
  }

  flushTurn();
  return rows;
}

async function listMatchedSessionFiles(sessionsDir: string, fromDate: string) {
  const fileNames = await fs.readdir(sessionsDir);
  return fileNames
    .filter((fileName) => /^\d{8}-.*\.json$/i.test(fileName))
    .filter((fileName) => fileName.slice(0, 8) >= fromDate)
    .sort();
}

function buildWorksheet(rows: ExportRow[]) {
  const localizedRows = rows.map((row) => ({
    session文件名: row.sessionFileName,
    sessionId: row.sessionId,
    knowledgeBaseId: row.knowledgeBaseId,
    session创建时间: row.sessionCreatedAt,
    session更新时间: row.sessionUpdatedAt,
    轮次: row.turnIndex,
    问题时间: row.questionCreatedAt,
    问题: row.question,
    回答时间: row.answerCreatedAt,
    回答: row.answer
  }));

  const worksheet = XLSX.utils.json_to_sheet(localizedRows);
  worksheet["!cols"] = [
    { wch: 42 },
    { wch: 38 },
    { wch: 38 },
    { wch: 24 },
    { wch: 24 },
    { wch: 8 },
    { wch: 24 },
    { wch: 50 },
    { wch: 24 },
    { wch: 80 }
  ];

  return worksheet;
}

async function writeOutput(outputPath: string, worksheet: XLSX.WorkSheet) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  const extension = path.extname(outputPath).toLowerCase();
  if (extension === ".csv") {
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    await fs.writeFile(outputPath, `\ufeff${csv}`, "utf8");
    return;
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "问答导出");
  XLSX.writeFile(workbook, outputPath, { compression: true });
}

async function main() {
  const options = parseCliArgs(process.argv.slice(2));
  const sessionFiles = await listMatchedSessionFiles(options.sessionsDir, options.fromDate);
  const rows: ExportRow[] = [];

  for (const sessionFileName of sessionFiles) {
    const filePath = path.join(options.sessionsDir, sessionFileName);
    const session = await readSession(filePath);
    rows.push(...buildExportRows(sessionFileName, session));
  }

  const worksheet = buildWorksheet(rows);
  await writeOutput(options.outputPath, worksheet);

  console.log(
    [
      `已匹配 session 文件: ${sessionFiles.length}`,
      `已导出问答行数: ${rows.length}`,
      `输出文件: ${options.outputPath}`
    ].join("\n")
  );
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`导出失败: ${message}`);
  printUsage();
  process.exitCode = 1;
});
