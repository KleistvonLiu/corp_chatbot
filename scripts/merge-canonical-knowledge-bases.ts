import { promises as fs } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import { parseWorkflowWorkbook } from "../server/lib/parsers";

type EntryType = "流程" | "联系人" | "供应商" | "参考" | "系统链接";

type SourceName = "old" | "new";

interface SourceFlowRow {
  source: SourceName;
  sourceRowNumber: number;
  category: string;
  entryType: EntryType;
  title: string;
  relatedForm: string;
  contacts: string;
  body: string;
  url: string;
  keywords: string;
}

interface SourceAttachmentRow {
  source: SourceName;
  sourceRowNumber: number;
  attachmentName: string;
  relativePath: string;
  attachmentType: string;
  description: string;
}

interface FinalFlowRow {
  rowNumber: number;
  category: string;
  entryType: EntryType;
  title: string;
  relatedForm: string;
  contacts: string;
  body: string;
  url: string;
  keywords: string;
  imageLinks: string[];
}

interface FinalAttachmentRow {
  rowNumber: number;
  attachmentName: string;
  relativePath: string;
  attachmentType: string;
  description: string;
  sourceAbsolutePath: string;
}

const oldPackageDir = "/home/kleist/Downloads/流程教程知识库规范包_20260402(1)";
const newPackageDir = "/home/kleist/Downloads/corp-eng-new-staff-guide-20260401-canonical";
const outputDir = "/home/kleist/Downloads/corp-eng-knowledge-merged-20260407-canonical";
const zipPath = "/home/kleist/Downloads/corp-eng-knowledge-merged-20260407-canonical.zip";

const duplicateTitlesToReplace = new Set(["加班", "小车申请"]);

function toText(value: unknown) {
  return String(value ?? "").trim();
}

function splitDelimited(value: string) {
  return value
    .split("；")
    .map((item) => item.trim())
    .filter(Boolean);
}

function mergeDelimited(existing: string, incoming: string) {
  const merged = [...splitDelimited(existing), ...splitDelimited(incoming)];
  return [...new Set(merged)].join("；");
}

function pad2(value: number) {
  return String(value).padStart(2, "0");
}

function makeFileName(rowNumber: number, sourceAttachmentName: string) {
  const match = sourceAttachmentName.match(/^row-\d+-(.+)$/);
  const suffix = match?.[1] ?? sourceAttachmentName;
  return `row-${pad2(rowNumber)}-${suffix}`;
}

function assertEntryType(value: string): EntryType {
  if (value === "流程" || value === "联系人" || value === "供应商" || value === "参考" || value === "系统链接") {
    return value;
  }
  throw new Error(`不支持的条目类型：${value || "空值"}`);
}

async function loadWorkbook(workbookPath: string) {
  const buffer = await fs.readFile(workbookPath);
  return XLSX.read(buffer, { type: "buffer" });
}

async function readFlowRows(workbookPath: string, source: SourceName) {
  const workbook = await loadWorkbook(workbookPath);
  const rows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["流程汇总"], {
    header: 1,
    blankrows: false,
    defval: ""
  });

  return rows.slice(1).flatMap((row) => {
    const sourceRowNumber = Number(row[0] ?? "");
    if (!Number.isInteger(sourceRowNumber) || sourceRowNumber <= 0) {
      return [];
    }

    return [
      {
        source,
        sourceRowNumber,
        category: toText(row[1]),
        entryType: assertEntryType(toText(row[2])),
        title: toText(row[3]),
        relatedForm: toText(row[4]),
        contacts: toText(row[5]),
        body: toText(row[6]),
        url: toText(row[7]),
        keywords: toText(row[8])
      } satisfies SourceFlowRow
    ];
  });
}

async function readAttachmentRows(workbookPath: string, source: SourceName) {
  const workbook = await loadWorkbook(workbookPath);
  const rows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["附件清单"], {
    header: 1,
    blankrows: false,
    defval: ""
  });

  return rows.slice(1).flatMap((row) => {
    const sourceRowNumber = Number(row[0] ?? "");
    if (!Number.isInteger(sourceRowNumber) || sourceRowNumber <= 0) {
      return [];
    }

    return [
      {
        source,
        sourceRowNumber,
        attachmentName: toText(row[1]),
        relativePath: toText(row[2]),
        attachmentType: toText(row[3]),
        description: toText(row[4])
      } satisfies SourceAttachmentRow
    ];
  });
}

function buildMergedRows(oldRows: SourceFlowRow[], newRows: SourceFlowRow[]) {
  const oldByTitle = new Map(oldRows.map((row) => [row.title, row]));
  const filteredOldRows = oldRows.filter((row) => !duplicateTitlesToReplace.has(row.title));

  const mergedNewRows = newRows.map((row) => {
    if (row.title !== "小车申请") {
      return row;
    }

    const oldCarRow = oldByTitle.get("小车申请");
    return {
      ...row,
      url: row.url || oldCarRow?.url || "",
      keywords: mergeDelimited(row.keywords, oldCarRow?.keywords ?? "")
    };
  });

  const mergedRows = [...filteredOldRows, ...mergedNewRows];
  const rowNumberBySourceKey = new Map<string, number>();

  const finalRows = mergedRows.map((row, index) => {
    const rowNumber = index + 1;
    rowNumberBySourceKey.set(`${row.source}:${row.sourceRowNumber}`, rowNumber);
    return {
      rowNumber,
      category: row.category,
      entryType: row.entryType,
      title: row.title,
      relatedForm: row.relatedForm,
      contacts: row.contacts,
      body: row.body,
      url: row.url,
      keywords: row.keywords,
      imageLinks: []
    } satisfies FinalFlowRow;
  });

  return { finalRows, rowNumberBySourceKey };
}

function buildFinalAttachments(
  allAttachments: SourceAttachmentRow[],
  rowNumberBySourceKey: Map<string, number>
) {
  const attachmentsByRow = new Map<number, string[]>();
  const finalAttachments: FinalAttachmentRow[] = [];

  for (const attachment of allAttachments) {
    const finalRowNumber = rowNumberBySourceKey.get(`${attachment.source}:${attachment.sourceRowNumber}`);
    if (finalRowNumber == null) {
      continue;
    }

    const fileName = makeFileName(finalRowNumber, attachment.attachmentName);
    const relativePath = `attachments/${fileName}`;
    const packageDir = attachment.source === "old" ? oldPackageDir : newPackageDir;
    const sourceAbsolutePath = path.join(packageDir, attachment.relativePath);

    finalAttachments.push({
      rowNumber: finalRowNumber,
      attachmentName: fileName,
      relativePath,
      attachmentType: attachment.attachmentType,
      description: attachment.description,
      sourceAbsolutePath
    });

    if (attachment.attachmentType === "png" || attachment.attachmentType === "jpg" || attachment.attachmentType === "jpeg" || attachment.attachmentType === "webp") {
      const current = attachmentsByRow.get(finalRowNumber) ?? [];
      current.push(relativePath);
      attachmentsByRow.set(finalRowNumber, current);
    }
  }

  finalAttachments.sort((left, right) => {
    if (left.rowNumber !== right.rowNumber) {
      return left.rowNumber - right.rowNumber;
    }
    return left.attachmentName.localeCompare(right.attachmentName, "zh-Hans-CN");
  });

  return { finalAttachments, attachmentsByRow };
}

function applyImageLinks(finalRows: FinalFlowRow[], attachmentsByRow: Map<number, string[]>) {
  return finalRows.map((row) => ({
    ...row,
    imageLinks: attachmentsByRow.get(row.rowNumber) ?? []
  }));
}

function createWorkbookBuffer(finalRows: FinalFlowRow[], finalAttachments: FinalAttachmentRow[]) {
  const workbook = XLSX.utils.book_new();

  const versionSheet = XLSX.utils.aoa_to_sheet([
    ["版本", "日期", "说明", "作者"],
    ["2026.04.07", "2026.04.07", "合并旧知识库与 Corp. Eng 新同事共享册", "Codex"],
    ["Ver1.0", "2026.02.06", "来源：流程教程知识库规范包_20260402(1)", "付军/李凯旺/兰善财"],
    ["2026.04.01", "2026.04.01", "来源：Corp. Eng New Staff Guide Book PDF 转制", "集团工程部行政组"]
  ]);

  const flowSheet = XLSX.utils.aoa_to_sheet([
    ["编号", "一级分类", "条目类型", "标题", "相关单据", "联系人/责任人", "正文", "外部链接", "关键词/别名", "图片链接"],
    ...finalRows.map((row) => [
      row.rowNumber,
      row.category,
      row.entryType,
      row.title,
      row.relatedForm,
      row.contacts,
      row.body,
      row.url,
      row.keywords,
      row.imageLinks.join("；")
    ])
  ]);

  const attachmentSheet = XLSX.utils.aoa_to_sheet([
    ["条目编号", "附件文件名", "相对路径", "附件类型", "附件说明"],
    ...finalAttachments.map((attachment) => [
      attachment.rowNumber,
      attachment.attachmentName,
      attachment.relativePath,
      attachment.attachmentType,
      attachment.description
    ])
  ]);

  XLSX.utils.book_append_sheet(workbook, versionSheet, "版本说明");
  XLSX.utils.book_append_sheet(workbook, flowSheet, "流程汇总");
  XLSX.utils.book_append_sheet(workbook, attachmentSheet, "附件清单");

  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

async function materializeOutput(workbookBuffer: Buffer, finalAttachments: FinalAttachmentRow[]) {
  await fs.rm(outputDir, { recursive: true, force: true });
  await fs.mkdir(path.join(outputDir, "attachments"), { recursive: true });

  for (const attachment of finalAttachments) {
    await fs.copyFile(attachment.sourceAbsolutePath, path.join(outputDir, attachment.relativePath));
  }

  await fs.writeFile(path.join(outputDir, "knowledge.xlsx"), workbookBuffer);

  const zip = new JSZip();
  zip.file("knowledge.xlsx", workbookBuffer);
  for (const attachment of finalAttachments) {
    const buffer = await fs.readFile(path.join(outputDir, attachment.relativePath));
    zip.file(attachment.relativePath, buffer);
  }
  const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(zipPath, zipBuffer);
}

async function smokeTest(workbookBuffer: Buffer, finalRows: FinalFlowRow[], finalAttachments: FinalAttachmentRow[]) {
  const workbook = XLSX.read(workbookBuffer, { type: "buffer" });
  const flowSheetRows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["流程汇总"], {
    header: 1,
    blankrows: false
  });
  const attachmentSheetRows = XLSX.utils.sheet_to_json<(string | number)[]>(workbook.Sheets["附件清单"], {
    header: 1,
    blankrows: false
  });

  if (flowSheetRows[0]?.[9] !== "图片链接") {
    throw new Error("流程汇总 sheet 缺少 J 列“图片链接”。");
  }
  if (flowSheetRows.length - 1 !== finalRows.length) {
    throw new Error(`流程汇总记录数异常：期望 ${finalRows.length}，实际 ${flowSheetRows.length - 1}`);
  }
  if (attachmentSheetRows.length - 1 !== finalAttachments.length) {
    throw new Error(`附件清单记录数异常：期望 ${finalAttachments.length}，实际 ${attachmentSheetRows.length - 1}`);
  }

  for (const attachment of finalAttachments) {
    await fs.access(path.join(outputDir, attachment.relativePath));
  }

  const requiredTitles = ["P200厂区平面图", "江门厂区平面图", "职员线路车时刻表", "来往厂穿梭巴时刻表", "6S定义"];
  for (const title of requiredTitles) {
    const row = finalRows.find((item) => item.title === title);
    if (!row?.imageLinks.length) {
      throw new Error(`条目缺少图片链接：${title}`);
    }
  }

  const overtimeRows = finalRows.filter((row) => row.title === "加班");
  if (overtimeRows.length !== 1) {
    throw new Error(`“加班”条目数量异常：${overtimeRows.length}`);
  }

  const carRow = finalRows.find((row) => row.title === "小车申请");
  if (!carRow?.url.includes("forms.office.com")) {
    throw new Error("“小车申请”未保留旧知识库中的表单入口 URL。");
  }

  const parsed = await parseWorkflowWorkbook(workbookBuffer, "merged-smoke-test", "knowledge.xlsx", {
    canonicalAttachmentsDir: path.join(outputDir, "attachments")
  });

  if (parsed.warnings.length > 0) {
    throw new Error(`解析出现 warning：${parsed.warnings.join(" | ")}`);
  }

  return {
    sourceCount: parsed.sources.length,
    warningCount: parsed.warnings.length
  };
}

async function main() {
  const oldWorkbookPath = path.join(oldPackageDir, "knowledge.xlsx");
  const newWorkbookPath = path.join(newPackageDir, "knowledge.xlsx");

  const oldRows = await readFlowRows(oldWorkbookPath, "old");
  const newRows = await readFlowRows(newWorkbookPath, "new");
  const oldAttachments = await readAttachmentRows(oldWorkbookPath, "old");
  const newAttachments = await readAttachmentRows(newWorkbookPath, "new");

  const { finalRows: mergedRows, rowNumberBySourceKey } = buildMergedRows(oldRows, newRows);
  const { finalAttachments, attachmentsByRow } = buildFinalAttachments([...oldAttachments, ...newAttachments], rowNumberBySourceKey);
  const finalRows = applyImageLinks(mergedRows, attachmentsByRow);
  const workbookBuffer = createWorkbookBuffer(finalRows, finalAttachments);

  await materializeOutput(workbookBuffer, finalAttachments);
  const smoke = await smokeTest(workbookBuffer, finalRows, finalAttachments);

  console.log(`Merged rows: ${finalRows.length}`);
  console.log(`Merged attachments: ${finalAttachments.length}`);
  console.log(`Workbook: ${path.join(outputDir, "knowledge.xlsx")}`);
  console.log(`Zip: ${zipPath}`);
  console.log(`Parsed sources: ${smoke.sourceCount}`);
  console.log(`Warnings: ${smoke.warningCount}`);
}

void main().catch((error) => {
  console.error(error instanceof Error ? error.stack : error);
  process.exitCode = 1;
});
