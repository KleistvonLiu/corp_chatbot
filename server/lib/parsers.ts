import { promises as fs } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type {
  AttachmentKind,
  EntryType,
  KnowledgeSource,
  SourceType
} from "../../shared/contracts";

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  parseTagValue: false,
  trimValues: false
});

interface SheetRow {
  sheetRowIndex: number;
  cells: Map<string, string>;
}

interface WorkbookSheet {
  name: string;
  target: string;
  rows: SheetRow[];
  hyperlinks: Map<string, string>;
}

interface ParsedWorkbook {
  sheetNames: string[];
  sheets: WorkbookSheet[];
}

interface ParsedAttachment {
  text: string;
  kind: AttachmentKind;
  warning?: string;
}

interface ParsedWorkflowWorkbook {
  sheets: string[];
  versionNotes?: string;
  sources: KnowledgeSource[];
  warnings: string[];
}

interface ParsedCanonicalWorkbookParts {
  workbook: ParsedWorkbook;
  flowSheet: WorkbookSheet;
  versionSheet?: WorkbookSheet;
  attachmentSheet: WorkbookSheet;
}

type RelationshipMap = Map<string, { type: string; target: string }>;

function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (value == null) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

function parseXmlText<T>(text: string) {
  return xmlParser.parse(text) as T;
}

function normalizeSpace(text: string) {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeDelimitedValue(text: string, separator = "；") {
  const parts = text
    .replace(/\r/g, "")
    .split(/[\n;,；]+/g)
    .map((item) => normalizeSpace(item))
    .filter(Boolean);

  return parts.join(separator);
}

function normalizeSingleLine(text: string) {
  return normalizeSpace(text.replace(/\n+/g, " "));
}

function isZipFileName(fileName: string) {
  return fileName.toLowerCase().endsWith(".zip");
}

function resolveZipTarget(baseDir: string, target: string) {
  return path.posix.normalize(path.posix.join(baseDir, target));
}

function inferAttachmentKind(fileName: string): AttachmentKind {
  const extension = path.extname(fileName).toLowerCase();
  if (extension === ".docx") {
    return "docx";
  }
  if (extension === ".pptx") {
    return "pptx";
  }
  if (extension === ".xlsx") {
    return "xlsx";
  }
  if (extension === ".doc") {
    return "doc";
  }

  return "unknown";
}

function cellReferenceToColumn(ref: string) {
  return ref.replace(/\d+/g, "");
}

function readRichText(node: unknown): string {
  const parts: string[] = [];

  function walk(value: unknown) {
    if (value == null) {
      return;
    }
    if (typeof value === "string") {
      if (value.trim()) {
        parts.push(value);
      }
      return;
    }
    if (Array.isArray(value)) {
      value.forEach(walk);
      return;
    }
    if (typeof value === "object") {
      for (const [key, nested] of Object.entries(value)) {
        if (key.startsWith("@_")) {
          continue;
        }
        walk(nested);
      }
    }
  }

  walk(node);
  return normalizeSpace(parts.join(""));
}

function decodeCellValue(cell: Record<string, unknown>, sharedStrings: string[]) {
  const type = typeof cell["@_t"] === "string" ? cell["@_t"] : "";
  if (type === "s") {
    const sharedIndex = Number(cell.v ?? 0);
    return normalizeSpace(sharedStrings[sharedIndex] ?? "");
  }

  if (type === "inlineStr") {
    return readRichText(cell.is);
  }

  if (typeof cell.v === "string") {
    return normalizeSpace(cell.v);
  }

  if (typeof cell.v === "number") {
    return String(cell.v);
  }

  return readRichText(cell);
}

async function readXmlFromZip(zip: JSZip, filePath: string) {
  const entry = zip.file(filePath);
  if (!entry) {
    throw new Error(`缺少 XML 文件: ${filePath}`);
  }
  return entry.async("string");
}

async function readRelationships(zip: JSZip, relsPath: string): Promise<RelationshipMap> {
  const entry = zip.file(relsPath);
  if (!entry) {
    return new Map();
  }
  const raw = await entry.async("string");
  const parsed = parseXmlText<{
    Relationships?: {
      Relationship?: Array<{ "@_Id": string; "@_Type": string; "@_Target": string }> | { "@_Id": string; "@_Type": string; "@_Target": string };
    };
  }>(raw);

  const rels = new Map<string, { type: string; target: string }>();
  for (const relationship of asArray(parsed.Relationships?.Relationship)) {
    rels.set(relationship["@_Id"], {
      type: relationship["@_Type"],
      target: relationship["@_Target"]
    });
  }
  return rels;
}

async function parseWorkbook(zip: JSZip): Promise<ParsedWorkbook> {
  const workbookXml = await readXmlFromZip(zip, "xl/workbook.xml");
  const workbookRels = await readRelationships(zip, "xl/_rels/workbook.xml.rels");
  const sharedStringsEntry = zip.file("xl/sharedStrings.xml");
  const sharedStrings: string[] = [];

  if (sharedStringsEntry) {
    const sharedXml = parseXmlText<{
      sst?: {
        si?: unknown[] | unknown;
      };
    }>(await sharedStringsEntry.async("string"));

    for (const item of asArray(sharedXml.sst?.si)) {
      sharedStrings.push(readRichText(item));
    }
  }

  const workbook = parseXmlText<{
    workbook?: {
      sheets?: {
        sheet?: Array<{ "@_name": string; "@_id": string }> | { "@_name": string; "@_id": string };
      };
    };
  }>(workbookXml);

  const sheets = [];
  for (const sheet of asArray(workbook.workbook?.sheets?.sheet)) {
    const relationship = workbookRels.get(sheet["@_id"]);
    if (!relationship) {
      continue;
    }

    const target = resolveZipTarget("xl", relationship.target);
    const sheetXml = parseXmlText<{
      worksheet?: {
        sheetData?: {
          row?: Array<Record<string, unknown>> | Record<string, unknown>;
        };
        hyperlinks?: {
          hyperlink?:
            | Array<{ "@_ref": string; "@_id": string }>
            | { "@_ref": string; "@_id": string };
        };
      };
    }>(await readXmlFromZip(zip, target));

    const rows: SheetRow[] = [];
    for (const row of asArray(sheetXml.worksheet?.sheetData?.row)) {
      const rowIndex = Number(row["@_r"] ?? 0);
      const rowCells = new Map<string, string>();
      for (const cell of asArray((row.c ?? []) as Array<Record<string, unknown>>)) {
        const ref = typeof cell["@_r"] === "string" ? cell["@_r"] : "";
        const column = ref ? cellReferenceToColumn(ref) : "";
        if (!column) {
          continue;
        }
        rowCells.set(column, decodeCellValue(cell, sharedStrings));
      }

      rows.push({
        sheetRowIndex: rowIndex,
        cells: rowCells
      });
    }

    const relsPath = `${path.posix.dirname(target)}/_rels/${path.posix.basename(target)}.rels`;
    const sheetRelationships = await readRelationships(zip, relsPath);
    const hyperlinks = new Map<string, string>();

    for (const hyperlink of asArray(sheetXml.worksheet?.hyperlinks?.hyperlink)) {
      const relationship = sheetRelationships.get(hyperlink["@_id"]);
      if (!relationship) {
        continue;
      }
      hyperlinks.set(hyperlink["@_ref"], relationship.target.replace(/&amp;/g, "&"));
    }

    sheets.push({
      name: sheet["@_name"],
      target,
      rows,
      hyperlinks
    });
  }

  return {
    sheetNames: sheets.map((sheet) => sheet.name),
    sheets
  };
}

function deriveTitle(processNumber: number, typeCell: string, flowCell: string) {
  if (typeCell.trim()) {
    return typeCell.trim();
  }

  const normalizedFlow = normalizeSpace(flowCell);
  const asciiLead = normalizedFlow.match(/^([A-Za-z][A-Za-z0-9 +()/._-]{0,40})(?=\s*[\u4e00-\u9fff]|$)/u)?.[1]?.trim();
  if (asciiLead) {
    return asciiLead;
  }

  const preview = normalizedFlow.split(/[。；;!?]/)[0]?.trim();
  return preview?.slice(0, 36) || `流程 ${processNumber}`;
}

async function parseEmbeddedAttachment(buffer: Buffer, fileName: string): Promise<ParsedAttachment> {
  const kind = inferAttachmentKind(fileName);
  if (kind === "doc") {
    return {
      kind,
      text: "",
      warning: `${fileName} 是旧版 .doc 附件，当前版本未自动解析。`
    };
  }

  if (kind === "unknown") {
    return {
      kind,
      text: "",
      warning: `${fileName} 的格式暂不支持自动解析。`
    };
  }

  if (kind === "xlsx") {
    const workbook = await parseWorkbook(await JSZip.loadAsync(buffer));
    const sections = workbook.sheets.flatMap((sheet) => {
      const lines = sheet.rows
        .map((row) => ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"].map((col) => row.cells.get(col) ?? "").join(" | "))
        .map((line) => normalizeSpace(line))
        .filter(Boolean);
      if (!lines.length) {
        return [];
      }
      return [`Sheet: ${sheet.name}`, ...lines.slice(0, 80)];
    });

    return {
      kind,
      text: normalizeSpace(sections.join("\n"))
    };
  }

  const zip = await JSZip.loadAsync(buffer);
  if (kind === "docx") {
    const documentXml = await readXmlFromZip(zip, "word/document.xml");
    const parsed = parseXmlText<{ document?: { body?: { p?: unknown[] | unknown } } }>(documentXml);
    const paragraphs = asArray(parsed.document?.body?.p)
      .map((paragraph) => readRichText(paragraph))
      .filter(Boolean);

    return {
      kind,
      text: normalizeSpace(paragraphs.join("\n"))
    };
  }

  if (kind === "pptx") {
    const slideEntries = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));

    const slides: string[] = [];
    for (let index = 0; index < slideEntries.length; index += 1) {
      const slideXml = await readXmlFromZip(zip, slideEntries[index]);
      const parsed = parseXmlText<Record<string, unknown>>(slideXml);
      const text = readRichText(parsed);
      if (text) {
        slides.push(`Slide ${index + 1}: ${text}`);
      }
    }

    return {
      kind,
      text: normalizeSpace(slides.join("\n"))
    };
  }

  return {
    kind,
    text: ""
  };
}

function createSourceId(knowledgeBaseId: string, sourceType: SourceType, suffix: string) {
  return `${knowledgeBaseId}-${sourceType}-${suffix}`;
}

function readVersionNotes(versionSheet?: WorkbookSheet) {
  return versionSheet?.rows.find((row) => row.sheetRowIndex === 2)
    ? ["A", "B", "C", "D"]
        .map((column) => versionSheet.rows.find((row) => row.sheetRowIndex === 2)?.cells.get(column) ?? "")
        .join(" / ")
        .trim()
    : undefined;
}

async function parseLegacyWorkflowWorkbook(
  buffer: Buffer,
  knowledgeBaseId: string
): Promise<ParsedWorkflowWorkbook> {
  const zip = await JSZip.loadAsync(buffer);
  const workbook = await parseWorkbook(zip);
  const flowSheet = workbook.sheets.find((sheet) => sheet.name === "流程汇总");
  const versionSheet = workbook.sheets.find((sheet) => sheet.name === "版本说明");

  if (!flowSheet) {
    throw new Error("Excel 中缺少名为“流程汇总”的 sheet。");
  }

  const warnings: string[] = [];
  const sources: KnowledgeSource[] = [];
  const sheetRowToProcessNumber = new Map<number, number>();
  const parentRowSourceIds = new Map<number, string>();
  let previousProcessNumber = 0;

  for (const row of flowSheet.rows) {
    const rawSequence = Number(row.cells.get("A") ?? "");
    const hasContent = ["B", "C", "D", "E"].some((column) => Boolean(normalizeSpace(row.cells.get(column) ?? "")));
    const processNumber =
      Number.isFinite(rawSequence) && rawSequence > 0
        ? rawSequence
        : hasContent && previousProcessNumber > 0
          ? previousProcessNumber + 1
          : 0;

    if (!processNumber) {
      continue;
    }

    previousProcessNumber = processNumber;
    const typeCell = row.cells.get("B") ?? "";
    const relatedForm = normalizeSpace(row.cells.get("C") ?? "");
    const flowText = normalizeSpace(row.cells.get("D") ?? "");
    const title = deriveTitle(processNumber, typeCell, flowText);
    const rowSourceId = createSourceId(knowledgeBaseId, "row", String(processNumber));
    const rowUrl = flowSheet.hyperlinks.get(`D${row.sheetRowIndex}`);

    sheetRowToProcessNumber.set(row.sheetRowIndex, processNumber);
    parentRowSourceIds.set(processNumber, rowSourceId);

    sources.push({
      sourceId: rowSourceId,
      knowledgeBaseId,
      rowNumber: processNumber,
      sourceType: "row",
      title,
      relatedForm: relatedForm || undefined,
      url: rowUrl,
      text: normalizeSpace(
        [`流程标题: ${title}`, relatedForm ? `相关的单: ${relatedForm}` : "", `流程内容: ${flowText}`]
          .filter(Boolean)
          .join("\n")
      )
    });

    if (rowUrl) {
      sources.push({
        sourceId: createSourceId(knowledgeBaseId, "link", String(processNumber)),
        knowledgeBaseId,
        rowNumber: processNumber,
        sourceType: "link",
        title,
        relatedForm: relatedForm || undefined,
        url: rowUrl,
        parentSourceId: rowSourceId,
        text: normalizeSpace([`流程标题: ${title}`, `外部链接: ${rowUrl}`, flowText].join("\n"))
      });
    }
  }

  const versionNotes = readVersionNotes(versionSheet);

  const sheetRelationships = await readRelationships(
    zip,
    `${path.posix.dirname(flowSheet.target)}/_rels/${path.posix.basename(flowSheet.target)}.rels`
  );
  const flowSheetXml = parseXmlText<{
    worksheet?: {
      oleObjects?: {
        AlternateContent?: Array<{
          Choice?: {
            oleObject?: {
              "@_id"?: string;
              objectPr?: {
                anchor?: {
                  from?: {
                    row?: string;
                  };
                };
              };
            };
          };
        }> | {
          Choice?: {
            oleObject?: {
              "@_id"?: string;
              objectPr?: {
                anchor?: {
                  from?: {
                    row?: string;
                  };
                };
              };
            };
          };
        };
      };
    };
  }>(await readXmlFromZip(zip, flowSheet.target));

  for (const item of asArray(flowSheetXml.worksheet?.oleObjects?.AlternateContent)) {
    const oleObject = item.Choice?.oleObject;
    const relationshipId = oleObject?.["@_id"];
    const zeroBasedRow = Number(oleObject?.objectPr?.anchor?.from?.row ?? "");
    if (!relationshipId || !Number.isFinite(zeroBasedRow)) {
      continue;
    }

    const relationship = sheetRelationships.get(relationshipId);
    if (!relationship) {
      continue;
    }

    const sheetRowIndex = zeroBasedRow + 1;
    const processNumber = sheetRowToProcessNumber.get(sheetRowIndex);
    if (!processNumber) {
      warnings.push(`有一个嵌入附件挂在 sheet 行 ${sheetRowIndex}，但未能映射到流程编号。`);
      continue;
    }

    const parentSourceId = parentRowSourceIds.get(processNumber);
    const attachmentPath = resolveZipTarget(path.posix.dirname(flowSheet.target), relationship.target);
    const attachmentFile = zip.file(attachmentPath);

    if (!attachmentFile) {
      warnings.push(`附件 ${attachmentPath} 在压缩包中不存在。`);
      continue;
    }

    const attachmentName = path.posix.basename(attachmentPath);
    const sourceTitle = sources.find((source) => source.sourceId === parentSourceId)?.title ?? `流程 ${processNumber}`;
    const attachment = await parseEmbeddedAttachment(await attachmentFile.async("nodebuffer"), attachmentName);

    if (attachment.warning) {
      warnings.push(attachment.warning);
    }

    sources.push({
      sourceId: createSourceId(knowledgeBaseId, "attachment", `${processNumber}-${attachmentName}`),
      knowledgeBaseId,
      rowNumber: processNumber,
      sourceType: "attachment",
      title: sourceTitle,
      attachmentName,
      attachmentKind: attachment.kind,
      parentSourceId,
      parseWarning: attachment.warning,
      text: attachment.text
    });
  }

  return {
    sheets: workbook.sheetNames,
    versionNotes: versionNotes ? normalizeSpace(versionNotes) : undefined,
    sources,
    warnings
  };
}

function parseAliases(rawValue: string) {
  return normalizeDelimitedValue(rawValue)
    .split("；")
    .map((item) => normalizeSpace(item))
    .filter(Boolean);
}

function readCanonicalEntryType(rawValue: string, sheetRowIndex: number, warnings?: string[]): EntryType {
  const value = normalizeSingleLine(rawValue);
  if (value === "流程" || value === "联系人" || value === "供应商" || value === "参考" || value === "系统链接") {
    return value;
  }

  if (warnings) {
    warnings.push(`流程汇总 sheet 第 ${sheetRowIndex} 行的“条目类型”无效：${value || "空值"}，已按“流程”处理。`);
    return "流程";
  }

  throw new Error(`流程汇总 sheet 第 ${sheetRowIndex} 行的“条目类型”无效：${value || "空值"}`);
}

function extractAttachmentRelativePaths(attachmentSheet: WorkbookSheet) {
  const paths: string[] = [];

  for (const row of attachmentSheet.rows) {
    if (row.sheetRowIndex === 1) {
      continue;
    }

    const hasContent = ["A", "B", "C", "D", "E"].some((column) => Boolean(normalizeSpace(row.cells.get(column) ?? "")));
    if (!hasContent) {
      continue;
    }

    const relativePath = normalizeSingleLine((row.cells.get("C") ?? "").replace(/\\/g, "/"));
    if (relativePath) {
      paths.push(relativePath);
    }
  }

  return paths;
}

function normalizeAttachmentRelativePath(relativePath: string) {
  return normalizeSingleLine(relativePath.replace(/\\/g, "/"));
}

export function resolveCanonicalAttachmentPath(attachmentsDir: string, relativePath: string) {
  const normalized = normalizeAttachmentRelativePath(relativePath);
  const segments = normalized.split("/").filter(Boolean);
  const trimmedSegments =
    path.basename(attachmentsDir).toLowerCase() === "attachments" && segments[0]?.toLowerCase() === "attachments"
      ? segments.slice(1)
      : segments;

  return path.join(attachmentsDir, ...trimmedSegments);
}

async function parseCanonicalWorkbookParts(workbook: ParsedWorkbook): Promise<ParsedCanonicalWorkbookParts> {
  const flowSheet = workbook.sheets.find((sheet) => sheet.name === "流程汇总");
  const versionSheet = workbook.sheets.find((sheet) => sheet.name === "版本说明");
  const attachmentSheet = workbook.sheets.find((sheet) => sheet.name === "附件清单");

  if (!flowSheet) {
    throw new Error("规范化工作簿中缺少名为“流程汇总”的 sheet。");
  }
  if (!attachmentSheet) {
    throw new Error("规范化工作簿中缺少名为“附件清单”的 sheet。");
  }

  return {
    workbook,
    flowSheet,
    versionSheet,
    attachmentSheet
  };
}

async function buildCanonicalSources(
  knowledgeBaseId: string,
  canonical: ParsedCanonicalWorkbookParts,
  readAttachment: (relativePath: string, attachmentName: string) => Promise<Buffer>
): Promise<ParsedWorkflowWorkbook> {
  const { workbook, flowSheet, versionSheet, attachmentSheet } = canonical;
  const warnings: string[] = [];
  const sources: KnowledgeSource[] = [];
  const parentRowSourceIds = new Map<number, string>();
  const rowTitles = new Map<number, string>();
  const rowMetadata = new Map<number, { category: string; entryType: EntryType; aliases: string[] }>();
  const seenRowNumbers = new Set<number>();

  for (const row of flowSheet.rows) {
    if (row.sheetRowIndex === 1) {
      continue;
    }

    const hasContent = ["A", "B", "C", "D", "E", "F", "G", "H", "I"].some((column) =>
      Boolean(normalizeSpace(row.cells.get(column) ?? ""))
    );
    if (!hasContent) {
      continue;
    }

    const processNumber = Number(normalizeSpace(row.cells.get("A") ?? ""));
    if (!Number.isInteger(processNumber) || processNumber <= 0) {
      throw new Error(`流程汇总 sheet 第 ${row.sheetRowIndex} 行的“编号”必须是正整数。`);
    }
    if (seenRowNumbers.has(processNumber)) {
      throw new Error(`流程汇总 sheet 中的编号 ${processNumber} 重复。`);
    }
    seenRowNumbers.add(processNumber);

    const category = normalizeSingleLine(row.cells.get("B") ?? "");
    if (!category) {
      throw new Error(`流程汇总 sheet 第 ${row.sheetRowIndex} 行缺少“一级分类”。`);
    }

    const entryType = readCanonicalEntryType(row.cells.get("C") ?? "", row.sheetRowIndex, warnings);
    const title = normalizeSingleLine(row.cells.get("D") ?? "");
    if (!title) {
      throw new Error(`流程汇总 sheet 第 ${row.sheetRowIndex} 行缺少“标题”。`);
    }

    const relatedForm = normalizeDelimitedValue(row.cells.get("E") ?? "");
    const contacts = normalizeDelimitedValue(row.cells.get("F") ?? "");
    const body = normalizeSpace(row.cells.get("G") ?? "");
    const url = normalizeSpace(row.cells.get("H") ?? "");
    const keywords = normalizeDelimitedValue(row.cells.get("I") ?? "");
    const aliases = parseAliases(row.cells.get("I") ?? "");

    if (!body) {
      throw new Error(`流程汇总 sheet 第 ${row.sheetRowIndex} 行缺少“正文”。`);
    }

    const rowSourceId = createSourceId(knowledgeBaseId, "row", String(processNumber));
    const rowText = normalizeSpace(
      [
        `一级分类: ${category}`,
        `条目类型: ${entryType}`,
        `条目标题: ${title}`,
        relatedForm ? `相关的单: ${relatedForm}` : "",
        contacts ? `联系人/责任人: ${contacts}` : "",
        keywords ? `关键词/别名: ${keywords}` : "",
        `正文: ${body}`,
        url ? `相关链接: ${url}` : ""
      ]
        .filter(Boolean)
        .join("\n")
    );

    sources.push({
      sourceId: rowSourceId,
      knowledgeBaseId,
      rowNumber: processNumber,
      sourceType: "row",
      title,
      entryType,
      category,
      aliases,
      relatedForm: relatedForm || undefined,
      url: url || undefined,
      text: rowText
    });

    parentRowSourceIds.set(processNumber, rowSourceId);
    rowTitles.set(processNumber, title);
    rowMetadata.set(processNumber, { category, entryType, aliases });

    if (url) {
      sources.push({
        sourceId: createSourceId(knowledgeBaseId, "link", String(processNumber)),
        knowledgeBaseId,
        rowNumber: processNumber,
        sourceType: "link",
        title,
        entryType,
        category,
        aliases,
        relatedForm: relatedForm || undefined,
        url,
        parentSourceId: rowSourceId,
        text: normalizeSpace(
          [
            `一级分类: ${category}`,
            `条目类型: ${entryType}`,
            `条目标题: ${title}`,
            contacts ? `联系人/责任人: ${contacts}` : "",
            keywords ? `关键词/别名: ${keywords}` : "",
            `外部链接: ${url}`,
            `正文: ${body}`
          ]
            .filter(Boolean)
            .join("\n")
        )
      });
    }
  }

  const seenAttachmentPaths = new Set<string>();
  for (const row of attachmentSheet.rows) {
    if (row.sheetRowIndex === 1) {
      continue;
    }

    const hasContent = ["A", "B", "C", "D", "E"].some((column) => Boolean(normalizeSpace(row.cells.get(column) ?? "")));
    if (!hasContent) {
      continue;
    }

    const processNumber = Number(normalizeSpace(row.cells.get("A") ?? ""));
    if (!Number.isInteger(processNumber) || processNumber <= 0) {
      throw new Error(`附件清单 sheet 第 ${row.sheetRowIndex} 行的“条目编号”必须是正整数。`);
    }
    if (!parentRowSourceIds.has(processNumber)) {
      throw new Error(`附件清单 sheet 第 ${row.sheetRowIndex} 行引用了不存在的条目编号 ${processNumber}。`);
    }

    const attachmentName = normalizeSingleLine(row.cells.get("B") ?? "");
    const relativePath = normalizeAttachmentRelativePath(row.cells.get("C") ?? "");
    const attachmentDescription = normalizeSpace(row.cells.get("E") ?? "");
    if (!attachmentName) {
      throw new Error(`附件清单 sheet 第 ${row.sheetRowIndex} 行缺少“附件文件名”。`);
    }
    if (!relativePath) {
      throw new Error(`附件清单 sheet 第 ${row.sheetRowIndex} 行缺少“相对路径”。`);
    }
    if (seenAttachmentPaths.has(relativePath)) {
      throw new Error(`附件清单中存在重复的附件路径：${relativePath}`);
    }
    seenAttachmentPaths.add(relativePath);

    const attachmentBuffer = await readAttachment(relativePath, attachmentName);
    const attachment = await parseEmbeddedAttachment(attachmentBuffer, attachmentName);
    if (attachment.warning) {
      warnings.push(attachment.warning);
    }

    sources.push({
      sourceId: createSourceId(knowledgeBaseId, "attachment", `${processNumber}-${attachmentName}`),
      knowledgeBaseId,
      rowNumber: processNumber,
      sourceType: "attachment",
      title: rowTitles.get(processNumber) ?? `流程 ${processNumber}`,
      entryType: rowMetadata.get(processNumber)?.entryType,
      category: rowMetadata.get(processNumber)?.category,
      aliases: rowMetadata.get(processNumber)?.aliases,
      attachmentName,
      attachmentKind: attachment.kind,
      parentSourceId: parentRowSourceIds.get(processNumber),
      parseWarning: attachment.warning,
      text: normalizeSpace([attachmentDescription ? `附件说明: ${attachmentDescription}` : "", attachment.text].filter(Boolean).join("\n"))
    });
  }

  const versionNotes = readVersionNotes(versionSheet);

  return {
    sheets: workbook.sheetNames,
    versionNotes: versionNotes ? normalizeSpace(versionNotes) : undefined,
    sources,
    warnings
  };
}

export async function collectCanonicalAttachmentReferences(buffer: Buffer) {
  const workbook = await parseWorkbook(await JSZip.loadAsync(buffer));
  const canonical = await parseCanonicalWorkbookParts(workbook);
  return extractAttachmentRelativePaths(canonical.attachmentSheet);
}

async function parseCanonicalWorkflowPackage(
  buffer: Buffer,
  knowledgeBaseId: string
): Promise<ParsedWorkflowWorkbook> {
  const packageZip = await JSZip.loadAsync(buffer);
  const workbookEntry = packageZip.file("knowledge.xlsx");
  if (!workbookEntry) {
    throw new Error("规范化 zip 包根目录缺少 knowledge.xlsx。");
  }

  const workbookBuffer = await workbookEntry.async("nodebuffer");
  const workbook = await parseWorkbook(await JSZip.loadAsync(workbookBuffer));
  const canonical = await parseCanonicalWorkbookParts(workbook);
  return buildCanonicalSources(knowledgeBaseId, canonical, async (relativePath) => {
    const attachmentEntry = packageZip.file(relativePath);
    if (!attachmentEntry) {
      throw new Error(`附件清单引用的文件不存在：${relativePath}`);
    }
    return attachmentEntry.async("nodebuffer");
  });
}

async function parseCanonicalWorkflowWorkbookFromXlsx(
  buffer: Buffer,
  knowledgeBaseId: string,
  attachmentsDir: string
): Promise<ParsedWorkflowWorkbook> {
  const workbook = await parseWorkbook(await JSZip.loadAsync(buffer));
  const canonical = await parseCanonicalWorkbookParts(workbook);
  return buildCanonicalSources(knowledgeBaseId, canonical, async (relativePath, attachmentName) => {
    const attachmentPath = resolveCanonicalAttachmentPath(attachmentsDir, relativePath);
    try {
      return await fs.readFile(attachmentPath);
    } catch (error) {
      const isMissing = error instanceof Error && "code" in error && error.code === "ENOENT";
      if (isMissing) {
        throw new Error(`附件清单引用的文件不存在：${relativePath} -> ${attachmentPath}`);
      }
      throw new Error(`读取附件失败：${attachmentName} -> ${attachmentPath}`);
    }
  });
}

export async function parseWorkflowWorkbook(
  buffer: Buffer,
  knowledgeBaseId: string,
  fileName = "workflow.xlsx",
  options?: {
    canonicalAttachmentsDir?: string;
  }
): Promise<ParsedWorkflowWorkbook> {
  if (isZipFileName(fileName)) {
    return parseCanonicalWorkflowPackage(buffer, knowledgeBaseId);
  }

  if (options?.canonicalAttachmentsDir) {
    return parseCanonicalWorkflowWorkbookFromXlsx(buffer, knowledgeBaseId, options.canonicalAttachmentsDir);
  }

  return parseLegacyWorkflowWorkbook(buffer, knowledgeBaseId);
}
