import { promises as fs } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import type { KnowledgeSource } from "../shared/contracts";
import { parseWorkflowWorkbook } from "../server/lib/parsers";

type CanonicalEntryType = "流程" | "联系人" | "供应商" | "参考" | "系统链接";

interface CanonicalRow {
  rowNumber: number;
  category: string;
  entryType: CanonicalEntryType;
  title: string;
  relatedForm: string;
  contacts: string;
  body: string;
  url: string;
  keywords: string;
}

interface CanonicalAttachment {
  rowNumber: number;
  fileName: string;
  relativePath: string;
  attachmentKind: string;
  description: string;
  buffer: Buffer;
}

const titleOverrides = new Map<number, string>([
  [19, "非直接物料采购（指定供应商）"],
  [23, "物料购买 - 直接物料（大批量购买）"],
  [25, "德昌 PCB 供应商"],
  [27, "嘉立创/捷多邦快速样板供应商"],
  [28, "德昌供应商（小批量 PCBA）"],
  [29, "江门 PCBA 产线（批量生产）"],
  [35, "Delia dx Huang"],
  [42, "楼下机加工房（周海强）"],
  [43, "楼下机加工房（Terrance Au）"],
  [44, "江门机加工房（Terrance）"],
  [45, "一楼样板房（Lisa Zhang Team）"],
  [46, "一楼样板房（Denie Zeng Team）"]
]);

const bodyOverrides = new Map<number, string>([
  [4, "Remote Desktop 官方安装教程入口。打开链接后按文档步骤安装客户端，并按公司要求连接云端环境。"],
  [13, "江门临时宿舍申请表单入口。打开链接填写入住信息并提交审批。"],
  [14, "小车申请表单入口。打开链接填写用车信息并提交。"],
  [17, "会议室预定应用入口。打开链接后按需选择会议室和时间段。"],
  [26, "内部贴片机，一般用于前期验证的小批量样板，目前由袁杨负责贴片。贴片文件格式需要统一。繁琐性--，速度+++。"]
]);

const keywordsMap = new Map<number, string[]>([
  [1, ["电脑", "显示器", "键盘", "鼠标", "硬件", "ITPM"]],
  [2, ["软件安装", "ITPM", "SD", "Allan Kwan", "Jie Li"]],
  [3, ["ECAD", "Altium Designer", "AD"]],
  [4, ["Remote Desktop", "远程桌面"]],
  [12, ["班车", "Shuttle", "JE In Motion"]],
  [13, ["临时宿舍", "Forms"]],
  [14, ["小车", "用车", "Forms"]],
  [15, ["出差", "TravelApprovalApp", "Power Apps"]],
  [18, ["非直接物料", "淘宝", "京东", "拼多多", "闲鱼"]],
  [19, ["非直接物料", "供应商联系方式"]],
  [20, ["P-CARD", "Master card", "信用卡"]],
  [21, ["Petty Cash", "报销", "增值税专票"]],
  [22, ["Spot Buy", "直接物料采购"]],
  [23, ["Normal Buy", "QQ", "Quick Quotation", "PR", "PO"]],
  [24, ["PCB", "嘉立创", "捷多邦"]],
  [25, ["PCB", "QQ", "EQ"]],
  [26, ["PCBA", "贴片", "袁杨"]],
  [27, ["嘉立创", "捷多邦", "快速样板"]],
  [28, ["PCBA", "德昌供应商"]],
  [29, ["PCBA", "江门产线"]],
  [30, ["元器件", "主被动料", "供应商列表"]],
  [32, ["电气测试", "load dump", "瞬断", "EJR"]],
  [33, ["EMC", "RE", "CE", "ESD", "EJR"]],
  [34, ["采购职责", "采购人员", "Contacts"]],
  [42, ["机加工", "周海强", "EJR"]],
  [43, ["机加工", "Terrance Au", "EJR"]],
  [44, ["机加工", "Terrance", "JR"]],
  [45, ["样板房", "Lisa Zhang Team", "SOF"]],
  [46, ["样板房", "Denie Zeng Team", "benchmarking"]]
]);

const contactsMap = new Map<number, string[]>([
  [1, ["Allan Kwan", "Jie Li"]],
  [2, ["Allan Kwan", "Jie Li"]],
  [3, ["Jie Li"]],
  [6, ["陈霞", "项目经理"]],
  [8, ["陈霞", "项目经理", "Jialian Li"]],
  [9, ["Jialian Li"]],
  [10, ["Jialian Li"]],
  [11, ["Jialian Li"]],
  [18, ["Sindy wp Peng", "Yuling Ye", "项目经理"]],
  [19, ["Sindy wp Peng"]],
  [22, ["直接物料采购"]],
  [23, ["Yuling Ye", "项目经理"]],
  [26, ["袁杨"]],
  [32, ["敖显学", "宋康"]],
  [33, ["敖显学", "宋康"]],
  [35, ["Delia dx Huang"]],
  [36, ["Rosie jj Luo"]],
  [37, ["Ivy yh Xie"]],
  [38, ["Sarah xq Lin"]],
  [39, ["Jane xz Wang"]],
  [40, ["Amy sh Ye"]],
  [41, ["Sindy wp Peng"]],
  [42, ["周海强"]],
  [43, ["Terrance Au"]],
  [44, ["Terrance"]],
  [45, ["Lisa Zhang Team"]],
  [46, ["Denie Zeng Team"]]
]);

function normalizeSpace(text: string) {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizeDelimitedValue(text: string) {
  return text
    .replace(/\r/g, "")
    .split(/[\n;,；]+/g)
    .map((item) => normalizeSpace(item))
    .filter(Boolean)
    .join("；");
}

function normalizeSingleLine(text: string) {
  return normalizeSpace(text.replace(/\n+/g, " "));
}

function normalizeComparable(text: string) {
  return normalizeSingleLine(text).replace(/[^\p{L}\p{N}\u4e00-\u9fff]+/gu, "").toLowerCase();
}

function getCategory(rowNumber: number) {
  if (rowNumber >= 1 && rowNumber <= 5) {
    return "IT";
  }
  if (rowNumber >= 6 && rowNumber <= 12) {
    return "行政";
  }
  if (rowNumber >= 13 && rowNumber <= 17) {
    return "出差/后勤";
  }
  if ((rowNumber >= 18 && rowNumber <= 31) || (rowNumber >= 34 && rowNumber <= 41)) {
    return "采购";
  }
  if (rowNumber >= 32 && rowNumber <= 33) {
    return "测试";
  }
  if (rowNumber >= 42 && rowNumber <= 46) {
    return "机加工";
  }
  return "参考";
}

function getEntryType(rowNumber: number): CanonicalEntryType {
  if ([4, 13, 14, 15, 17].includes(rowNumber)) {
    return "系统链接";
  }
  if (rowNumber >= 35 && rowNumber <= 41) {
    return "联系人";
  }
  if ([25, 27, 28, 29, 42, 43, 44, 45, 46].includes(rowNumber)) {
    return "供应商";
  }
  if ([30, 34].includes(rowNumber)) {
    return "参考";
  }
  return "流程";
}

function extractLegacyBody(source: KnowledgeSource) {
  const relatedFormTokens = new Set(
    normalizeDelimitedValue(source.relatedForm ?? "")
      .split("；")
      .map((item) => normalizeSpace(item))
      .filter(Boolean)
  );
  const lines = source.text
    .split("\n")
    .map((line) => normalizeSpace(line))
    .filter(Boolean);
  const bodyLines: string[] = [];

  for (const line of lines) {
    if (line.startsWith("流程标题:") || line.startsWith("条目标题:")) {
      continue;
    }
    if (line.startsWith("相关的单:")) {
      continue;
    }
    if (line.startsWith("流程内容:")) {
      const value = normalizeSpace(line.slice("流程内容:".length));
      if (value) {
        bodyLines.push(value);
      }
      continue;
    }
    if (line.startsWith("正文:")) {
      const value = normalizeSpace(line.slice("正文:".length));
      if (value) {
        bodyLines.push(value);
      }
      continue;
    }
    if (relatedFormTokens.has(line)) {
      continue;
    }
    bodyLines.push(line);
  }

  return normalizeSpace(bodyLines.join("\n"));
}

function stripRedundantLeadLines(body: string, title: string, url: string) {
  const lines = body
    .split("\n")
    .map((line) => normalizeSpace(line))
    .filter(Boolean);
  const comparableTitle = normalizeComparable(title);
  const comparableUrl = normalizeSpace(url).toLowerCase();

  while (lines.length) {
    const candidate = lines[0];
    const comparableCandidate = normalizeComparable(candidate);
    if (comparableUrl && candidate.toLowerCase() === comparableUrl) {
      lines.shift();
      continue;
    }
    if (
      comparableCandidate &&
      comparableTitle &&
      (comparableTitle.includes(comparableCandidate) || comparableCandidate.includes(comparableTitle))
    ) {
      lines.shift();
      continue;
    }
    break;
  }

  return normalizeSpace(lines.join("\n"));
}

function uniqueJoin(values: string[]) {
  return [...new Set(values.map((item) => normalizeSpace(item)).filter(Boolean))].join("；");
}

function getCanonicalRows(rowSources: KnowledgeSource[]) {
  return rowSources
    .slice()
    .sort((left, right) => left.rowNumber - right.rowNumber)
    .map<CanonicalRow>((source) => {
      const title = titleOverrides.get(source.rowNumber) ?? normalizeSingleLine(source.title);
      const url = normalizeSpace(source.url ?? "");
      const relatedForm = normalizeDelimitedValue(source.relatedForm ?? "");
      let body = bodyOverrides.get(source.rowNumber) ?? stripRedundantLeadLines(extractLegacyBody(source), title, url);
      if (!body) {
        body = url ? `打开外部链接查看“${title}”相关入口。` : `请补充“${title}”的详细说明。`;
      }

      const contacts = uniqueJoin(contactsMap.get(source.rowNumber) ?? []);
      const keywords = uniqueJoin(keywordsMap.get(source.rowNumber) ?? []);

      return {
        rowNumber: source.rowNumber,
        category: getCategory(source.rowNumber),
        entryType: getEntryType(source.rowNumber),
        title,
        relatedForm,
        contacts,
        body,
        url,
        keywords
      };
    });
}

function xmlEscape(text: string) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function createDocxFromText(text: string) {
  const zip = new JSZip();
  const paragraphs = text
    .split(/\n+/)
    .map((line) => normalizeSpace(line))
    .filter(Boolean)
    .map((line) => `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(line)}</w:t></w:r></w:p>`)
    .join("");

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.folder("_rels")?.file(
    ".rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.folder("word")?.file(
    "document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 wp14">
  <w:body>
    ${paragraphs || "<w:p/>"}
    <w:sectPr>
      <w:pgSz w:w=\"11906\" w:h=\"16838\"/>
      <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

function buildWorkbookBuffer(canonicalRows: CanonicalRow[], attachments: CanonicalAttachment[], versionNotes?: string) {
  const workbook = XLSX.utils.book_new();
  const rawVersionParts = (versionNotes ?? "").split(/\s*\/\s*/);
  const versionParts = [
    rawVersionParts[0] ?? "",
    rawVersionParts[1] ?? "",
    rawVersionParts[2] ?? "",
    rawVersionParts.slice(3).join("/")
  ];

  const versionSheet = XLSX.utils.aoa_to_sheet([
    ["版本", "日期", "说明", "作者"],
    versionParts
  ]);
  versionSheet["!cols"] = [{ wch: 14 }, { wch: 14 }, { wch: 32 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(workbook, versionSheet, "版本说明");

  const flowSheet = XLSX.utils.aoa_to_sheet([
    ["编号", "一级分类", "条目类型", "标题", "相关单据", "联系人/责任人", "正文", "外部链接", "关键词/别名"],
    ...canonicalRows.map((row) => [
      row.rowNumber,
      row.category,
      row.entryType,
      row.title,
      row.relatedForm,
      row.contacts,
      row.body,
      row.url,
      row.keywords
    ])
  ]);
  flowSheet["!cols"] = [
    { wch: 8 },
    { wch: 12 },
    { wch: 12 },
    { wch: 34 },
    { wch: 24 },
    { wch: 24 },
    { wch: 88 },
    { wch: 48 },
    { wch: 30 }
  ];
  XLSX.utils.book_append_sheet(workbook, flowSheet, "流程汇总");

  const attachmentSheet = XLSX.utils.aoa_to_sheet([
    ["条目编号", "附件文件名", "相对路径", "附件类型", "附件说明"],
    ...attachments.map((attachment) => [
      attachment.rowNumber,
      attachment.fileName,
      attachment.relativePath,
      attachment.attachmentKind,
      attachment.description
    ])
  ]);
  attachmentSheet["!cols"] = [{ wch: 10 }, { wch: 26 }, { wch: 32 }, { wch: 12 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(workbook, attachmentSheet, "附件清单");

  return Buffer.from(XLSX.write(workbook, { bookType: "xlsx", type: "buffer", compression: true }));
}

function buildGeneratedDocxSummary(row: CanonicalRow) {
  return [
    `${row.title} 教程摘要`,
    "",
    "原始内嵌附件为旧版 .doc，迁移时未保留该二进制格式。",
    "当前规范化知识库包保留以下可读摘要：",
    row.relatedForm ? `相关单据：${row.relatedForm}` : "",
    row.contacts ? `联系人/责任人：${row.contacts}` : "",
    row.body,
    "",
    "如需补全详细教程，请将原始文件另存为 .docx 后替换当前附件。"
  ]
    .filter(Boolean)
    .join("\n");
}

async function getCanonicalAttachments(
  legacyBuffer: Buffer,
  canonicalRows: CanonicalRow[],
  attachmentSources: KnowledgeSource[]
) {
  const legacyZip = await JSZip.loadAsync(legacyBuffer);
  const embeddedEntries = new Map<string, JSZip.JSZipObject>();
  for (const entryName of Object.keys(legacyZip.files)) {
    if (!entryName.startsWith("xl/embeddings/")) {
      continue;
    }
    embeddedEntries.set(path.posix.basename(entryName), legacyZip.files[entryName]);
  }

  const rowIndex = new Map<number, CanonicalRow>(canonicalRows.map((row) => [row.rowNumber, row]));
  const perRowCount = new Map<number, number>();
  const attachments: CanonicalAttachment[] = [];

  for (const source of attachmentSources.slice().sort((left, right) => left.rowNumber - right.rowNumber || left.sourceId.localeCompare(right.sourceId))) {
    const rowCount = (perRowCount.get(source.rowNumber) ?? 0) + 1;
    perRowCount.set(source.rowNumber, rowCount);

    const extension = source.attachmentKind === "doc" ? ".docx" : path.extname(source.attachmentName ?? "") || ".bin";
    const fileName = `row-${String(source.rowNumber).padStart(2, "0")}-${String(rowCount).padStart(2, "0")}${extension}`;
    const relativePath = `attachments/${fileName}`;
    const row = rowIndex.get(source.rowNumber);
    if (!row) {
      throw new Error(`附件 ${source.attachmentName ?? source.sourceId} 找不到对应条目 ${source.rowNumber}。`);
    }

    let buffer: Buffer;
    let attachmentKind = source.attachmentKind ?? extension.replace(/^\./, "");
    let description = `关联条目：${row.title}`;

    if (source.attachmentKind === "doc") {
      buffer = await createDocxFromText(buildGeneratedDocxSummary(row));
      attachmentKind = "docx";
      description = `由旧版 .doc 迁移生成的摘要附件；原始附件名：${source.attachmentName ?? "未知附件"}`;
    } else {
      const embeddedEntry = source.attachmentName ? embeddedEntries.get(source.attachmentName) : undefined;
      if (!embeddedEntry) {
        throw new Error(`原始工作簿中找不到附件 ${source.attachmentName ?? source.sourceId}。`);
      }
      buffer = await embeddedEntry.async("nodebuffer");
      description = `原始附件名：${source.attachmentName ?? fileName}；关联条目：${row.title}`;
    }

    attachments.push({
      rowNumber: source.rowNumber,
      fileName,
      relativePath,
      attachmentKind,
      description,
      buffer
    });
  }

  return attachments;
}

async function main() {
  const [inputFile, outputFileArg] = process.argv.slice(2);
  if (!inputFile) {
    throw new Error("用法: npm run normalize:legacy -- <legacy.xlsx> [output.zip]");
  }

  if (!inputFile.toLowerCase().endsWith(".xlsx")) {
    throw new Error("当前迁移脚本只接受 legacy .xlsx 文件作为输入。");
  }

  const inputBuffer = await fs.readFile(inputFile);
  const parsed = await parseWorkflowWorkbook(inputBuffer, "normalize-legacy", path.basename(inputFile));
  const rowSources = parsed.sources.filter((source) => source.sourceType === "row");
  const attachmentSources = parsed.sources.filter((source) => source.sourceType === "attachment");

  if (rowSources.length !== 46) {
    throw new Error(`legacy 工作簿条目数异常，期望 46，实际 ${rowSources.length}`);
  }
  if (attachmentSources.length !== 11) {
    throw new Error(`legacy 工作簿附件数异常，期望 11，实际 ${attachmentSources.length}`);
  }

  const canonicalRows = getCanonicalRows(rowSources);
  const canonicalAttachments = await getCanonicalAttachments(inputBuffer, canonicalRows, attachmentSources);
  const workbookBuffer = buildWorkbookBuffer(canonicalRows, canonicalAttachments, parsed.versionNotes);

  const outputFile =
    outputFileArg ||
    path.join(
      path.dirname(inputFile),
      `${path.basename(inputFile, path.extname(inputFile))}_knowledge_package.zip`
    );

  const outputZip = new JSZip();
  outputZip.file("knowledge.xlsx", workbookBuffer);
  for (const attachment of canonicalAttachments) {
    outputZip.file(attachment.relativePath, attachment.buffer);
  }

  await fs.mkdir(path.dirname(outputFile), { recursive: true });
  await fs.writeFile(outputFile, await outputZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));

  console.log(`已生成规范化知识库包: ${outputFile}`);
  console.log(`条目数: ${canonicalRows.length}`);
  console.log(`附件数: ${canonicalAttachments.length}`);
}

void main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
