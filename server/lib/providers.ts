import type { ChatMessage, Citation, KnowledgeChunk, ModelRequestDebug } from "../../shared/contracts";

const stopwords = new Set([
  "的",
  "了",
  "和",
  "与",
  "或",
  "是",
  "要",
  "在",
  "我",
  "你",
  "他",
  "她",
  "它",
  "吗",
  "呢",
  "啊",
  "请",
  "一个",
  "这个",
  "那个",
  "需要",
  "什么",
  "怎么",
  "如何",
  "可以",
  "一下",
  "流程",
  "教程"
]);

const segmenter = new Intl.Segmenter("zh-CN", { granularity: "word" });

export interface AnswerInput {
  question: string;
  history: ChatMessage[];
  evidence: KnowledgeChunk[];
  citations: Citation[];
}

export interface AnswerResult {
  answer: string;
  modelRequest?: ModelRequestDebug;
}

export interface ModelProvider {
  mode: string;
  embedMany(texts: string[]): Promise<number[][]>;
  answer(input: AnswerInput): Promise<AnswerResult>;
}

interface ChatProvider {
  mode: string;
  answer(input: AnswerInput): Promise<AnswerResult>;
}

interface EmbeddingProvider {
  mode: string;
  embedMany(texts: string[]): Promise<number[][]>;
}

interface OpenAiCompatibleConfig {
  mode: string;
  baseUrl: string;
  apiKey: string;
  chatModel?: string;
  embeddingModel?: string;
  chatMaxTokens: number;
  enableThinking: boolean;
  stripThinkOutput: boolean;
}

const directMatchStopwords = new Set([
  "我",
  "你",
  "他",
  "她",
  "它",
  "我想",
  "我想要",
  "想",
  "想要",
  "请问",
  "问下",
  "问",
  "一下",
  "那个",
  "这个",
  "该",
  "找",
  "找谁",
  "该找谁",
  "联系谁",
  "谁负责",
  "负责",
  "谁",
  "怎么",
  "如何",
  "可以",
  "需要",
  "公司",
  "部门",
  "同事",
  "申请",
  "审批",
  "安装",
  "流程",
  "教程",
  "什么",
  "什么单",
  "走什么单",
  "怎么办",
  "咋办",
  "几点",
  "开门",
  "买",
  "购买"
]);

function normalizeBaseUrl(baseUrl: string) {
  return baseUrl.replace(/\/+$/, "");
}

function parseBooleanEnv(raw: string | undefined, defaultValue: boolean) {
  if (!raw) {
    return defaultValue;
  }

  const normalized = raw.trim().toLowerCase();
  if (["1", "true", "yes", "on"].includes(normalized)) {
    return true;
  }
  if (["0", "false", "no", "off"].includes(normalized)) {
    return false;
  }

  return defaultValue;
}

function parsePositiveInt(raw: string | undefined, defaultValue: number) {
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return defaultValue;
  }
  return Math.floor(parsed);
}

function buildOpenAiHeaders(apiKey?: string) {
  return {
    "Content-Type": "application/json",
    ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {})
  };
}

function failIfMissing(value: string | undefined, envName: string, mode: string) {
  if (!value) {
    throw new Error(`${mode} 配置不完整，缺少 ${envName}。`);
  }
  return value;
}

export type SimpleProviderMode = "offline" | "openai" | "vllm";

function parseSimpleMode(raw: string): SimpleProviderMode {
  if (raw === "offline" || raw === "openai" || raw === "vllm") {
    return raw;
  }
  throw new Error(`未知 provider mode: ${raw}`);
}

function getCurrentChatMode(): SimpleProviderMode {
  return parseSimpleMode((process.env.CHATBOT_PROVIDER ?? "openai").toLowerCase());
}

function parseCompositeMode(modeOverride?: string) {
  if (!modeOverride) {
    const chatMode = getCurrentChatMode();
    const embeddingMode = parseSimpleMode((process.env.EMBEDDING_PROVIDER ?? defaultEmbeddingMode(chatMode)).toLowerCase());
    return {
      chatMode,
      embeddingMode,
      compositeMode: `chat:${chatMode}|embed:${embeddingMode}`
    };
  }

  if (modeOverride.startsWith("chat:") && modeOverride.includes("|embed:")) {
    const [chatPart, embeddingPart] = modeOverride.split("|");
    return {
      chatMode: parseSimpleMode(chatPart.replace("chat:", "")),
      embeddingMode: parseSimpleMode(embeddingPart.replace("embed:", "")),
      compositeMode: modeOverride
    };
  }

  const simpleMode = parseSimpleMode(modeOverride.toLowerCase());
  return {
    chatMode: simpleMode,
    embeddingMode: simpleMode === "vllm" ? "offline" : simpleMode,
    compositeMode: `chat:${simpleMode}|embed:${simpleMode === "vllm" ? "offline" : simpleMode}`
  };
}

function defaultEmbeddingMode(chatMode: SimpleProviderMode): SimpleProviderMode {
  if (chatMode === "vllm") {
    return "offline";
  }
  return chatMode;
}

function getOpenAiConfig(): OpenAiCompatibleConfig {
  return {
    mode: "openai",
    baseUrl: normalizeBaseUrl(process.env.OPENAI_BASE_URL ?? "https://api.openai.com/v1"),
    apiKey: process.env.OPENAI_API_KEY ?? "",
    chatModel: process.env.OPENAI_CHAT_MODEL,
    embeddingModel: process.env.OPENAI_EMBEDDING_MODEL,
    chatMaxTokens: parsePositiveInt(process.env.CHAT_MAX_TOKENS, 300),
    enableThinking: true,
    stripThinkOutput: false
  };
}

function getVllmConfig(): OpenAiCompatibleConfig {
  return {
    mode: "vllm",
    baseUrl: normalizeBaseUrl(process.env.VLLM_BASE_URL ?? "http://localhost:8000/v1"),
    apiKey: process.env.VLLM_API_KEY ?? "",
    chatModel: process.env.VLLM_CHAT_MODEL,
    embeddingModel: process.env.VLLM_EMBEDDING_MODEL,
    chatMaxTokens: parsePositiveInt(process.env.CHAT_MAX_TOKENS, 300),
    enableThinking: parseBooleanEnv(process.env.VLLM_ENABLE_THINKING, false),
    stripThinkOutput: parseBooleanEnv(process.env.VLLM_STRIP_THINK_OUTPUT, true)
  };
}

function extractHanPhrases(text: string) {
  const phrases = new Set<string>();
  for (const block of text.match(/[\p{Script=Han}]{2,}/gu) ?? []) {
    const maxLength = Math.min(6, block.length);
    for (let length = 2; length <= maxLength; length += 1) {
      for (let start = 0; start <= block.length - length; start += 1) {
        const phrase = block.slice(start, start + length).trim();
        if (!phrase || stopwords.has(phrase)) {
          continue;
        }
        phrases.add(phrase);
      }
    }
  }
  return phrases;
}

function isGenericDirectTerm(term: string) {
  if (directMatchStopwords.has(term)) {
    return true;
  }

  for (const stopword of directMatchStopwords) {
    if ((term.startsWith(stopword) || term.endsWith(stopword)) && term.length <= stopword.length + 2) {
      return true;
    }
  }

  return false;
}

function stripThinkBlocks(text: string) {
  let cleaned = text.replace(/<think>[\s\S]*?<\/think>/gi, "").trim();

  if (!/^\s*(thinking process|thought process|reasoning|思考过程|推理过程)\s*:/iu.test(cleaned)) {
    return cleaned;
  }

  const marker = /\n\s*\n(?=根据|未找到|结论|答复|可以|可按|建议|先|如果|1\.|\d+\.|- |• )/u.exec(cleaned);
  if (marker?.index != null) {
    cleaned = cleaned.slice(marker.index).trim();
  }

  if (/^\s*(thinking process|thought process|reasoning|思考过程|推理过程)\s*:/iu.test(cleaned)) {
    const sections = cleaned
      .split(/\n\s*\n+/)
      .map((section) => section.trim())
      .filter(Boolean);
    const answerSectionIndex = sections.findIndex(
      (section, index) =>
        index > 0 &&
        !/^(thinking process|thought process|reasoning|思考过程|推理过程)\s*:/iu.test(section) &&
        /[\p{Script=Han}\p{Letter}\p{Number}]/u.test(section)
    );
    if (answerSectionIndex > 0) {
      cleaned = sections.slice(answerSectionIndex).join("\n\n").trim();
    }
  }

  return cleaned;
}

function normalizeSpace(text: string) {
  return text
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function extractKeywords(text: string): string[] {
  const normalized = text.toLowerCase().replace(/\r/g, " ").replace(/\n/g, " ");
  const tokens = new Set<string>();

  for (const piece of normalized.match(/[a-z0-9][a-z0-9._/-]*/g) ?? []) {
    if (piece.length > 1 && !stopwords.has(piece)) {
      tokens.add(piece);
    }
  }

  for (const token of segmenter.segment(normalized)) {
    const segment = token.segment.trim();
    if (!segment) {
      continue;
    }

    if (!/[\p{Script=Han}\p{Letter}\p{Number}]/u.test(segment)) {
      continue;
    }

    if (segment.length === 1 && !/[\p{Script=Han}]/u.test(segment)) {
      continue;
    }

    if (stopwords.has(segment)) {
      continue;
    }

    tokens.add(segment);
  }

  for (const phrase of extractHanPhrases(normalized)) {
    tokens.add(phrase);
  }

  return [...tokens];
}

export function extractDirectMatchTerms(text: string): string[] {
  return extractKeywords(text)
    .filter((term) => term.length >= 2 && !isGenericDirectTerm(term))
    .sort((left, right) => right.length - left.length || left.localeCompare(right));
}

function isNoEvidenceAnswer(text: string) {
  return /未找到明确依据|没有找到足够依据|当前知识库里没有找到足够依据/u.test(normalizeSpace(text));
}

function summarizeAssistantHistory(message: ChatMessage) {
  if (message.citations?.length) {
    const seenRows = new Set<number>();
    const citedFlows = message.citations
      .filter((citation) => {
        if (seenRows.has(citation.rowNumber)) {
          return false;
        }
        seenRows.add(citation.rowNumber);
        return true;
      })
      .map((citation) => `流程 ${citation.rowNumber}《${citation.title}》`);

    if (citedFlows.length) {
      return `助手已引用: ${citedFlows.join("；")}`;
    }
  }

  return isNoEvidenceAnswer(message.content) ? "助手: 未找到明确依据" : "助手: 已回答";
}

function compactHistoryForPrompt(history: ChatMessage[]) {
  if (!history.length) {
    return "无";
  }

  let userCount = 0;
  let startIndex = 0;
  for (let index = history.length - 1; index >= 0; index -= 1) {
    if (history[index].role === "user") {
      userCount += 1;
      if (userCount === 4) {
        startIndex = index;
        break;
      }
    }
  }

  const recentHistory = history.slice(userCount >= 4 ? startIndex : 0);
  return (
    recentHistory
      .map((message) =>
        message.role === "user" ? `用户: ${message.content}` : summarizeAssistantHistory(message)
      )
      .join("\n") || "无"
  );
}

function cosineSimilarity(left: number[], right: number[]) {
  if (!left.length || left.length !== right.length) {
    return 0;
  }

  let dot = 0;
  let leftNorm = 0;
  let rightNorm = 0;

  for (let index = 0; index < left.length; index += 1) {
    dot += left[index] * right[index];
    leftNorm += left[index] * left[index];
    rightNorm += right[index] * right[index];
  }

  if (!leftNorm || !rightNorm) {
    return 0;
  }

  return dot / (Math.sqrt(leftNorm) * Math.sqrt(rightNorm));
}

function hashToken(token: string) {
  let hash = 2166136261;
  for (const character of token) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }

  return hash >>> 0;
}

function buildOfflineEmbedding(text: string) {
  const vectorLength = 256;
  const vector = new Array<number>(vectorLength).fill(0);
  const tokens = extractKeywords(text);

  for (const token of tokens) {
    const hash = hashToken(token);
    const index = hash % vectorLength;
    const sign = hash % 2 === 0 ? 1 : -1;
    vector[index] += sign;
  }

  const norm = Math.sqrt(vector.reduce((sum, value) => sum + value * value, 0)) || 1;
  return vector.map((value) => value / norm);
}

class OfflineEmbeddingProvider implements EmbeddingProvider {
  mode = "offline";

  async embedMany(texts: string[]) {
    return texts.map((text) => buildOfflineEmbedding(text));
  }
}

class OfflineChatProvider implements ChatProvider {
  mode = "offline";

  async answer(input: AnswerInput) {
    if (!input.evidence.length || !input.citations.length) {
      return {
        answer: "当前知识库里没有找到足够依据来回答这个问题。请换一种问法，或者先补充文档内容。"
      };
    }

    const compareMode = /区别|不同|对比|比较|差异|vs/i.test(input.question);
    const uniqueCitations = input.citations.slice(0, compareMode ? 4 : 3);
    const lines = uniqueCitations.map((citation) => {
      const prefix = `流程 ${citation.rowNumber}《${citation.title}》`;
      const attachment = citation.attachmentName ? `，附件 ${citation.attachmentName}` : "";
      return `${prefix}${attachment}：${citation.snippet}`;
    });

    return {
      answer: compareMode
        ? `根据检索到的证据，相关流程如下：\n${lines.join("\n")}`
        : `根据当前文档，最相关的依据是：\n${lines.join("\n")}`
    };
  }
}

class OpenAiCompatibleEmbeddingProvider implements EmbeddingProvider {
  readonly mode: string;
  private readonly baseUrl: string;
  private readonly apiKey: string;
  private readonly embeddingModel: string;

  constructor(config: OpenAiCompatibleConfig) {
    this.mode = config.mode;
    this.baseUrl = config.baseUrl;
    this.apiKey = config.apiKey;
    this.embeddingModel = failIfMissing(
      config.embeddingModel,
      config.mode === "vllm" ? "VLLM_EMBEDDING_MODEL" : "OPENAI_EMBEDDING_MODEL",
      config.mode
    );
  }

  async embedMany(texts: string[]) {
    const results: number[][] = [];

    for (let index = 0; index < texts.length; index += 64) {
      const batch = texts.slice(index, index + 64);
      const response = await fetch(`${this.baseUrl}/embeddings`, {
        method: "POST",
        headers: buildOpenAiHeaders(this.apiKey),
        body: JSON.stringify({
          model: this.embeddingModel,
          input: batch
        })
      });

      if (!response.ok) {
        const body = await response.text();
        throw new Error(`嵌入模型调用失败: ${response.status} ${body}`);
      }

      const payload = (await response.json()) as {
        data?: Array<{ embedding?: number[] }>;
      };
      const batchEmbeddings = payload.data?.map((item) => item.embedding ?? []);
      if (!batchEmbeddings?.length) {
        throw new Error("嵌入模型没有返回有效向量");
      }

      results.push(...batchEmbeddings);
    }

    return results;
  }
}

class OpenAiCompatibleChatProvider implements ChatProvider {
  readonly mode: string;
  private readonly baseUrl: string;
  private readonly apiKey: string;
  private readonly chatModel: string;
  private readonly chatMaxTokens: number;
  private readonly enableThinking: boolean;
  private readonly stripThinkOutput: boolean;

  constructor(config: OpenAiCompatibleConfig) {
    this.mode = config.mode;
    this.baseUrl = config.baseUrl;
    this.apiKey = config.apiKey;
    this.chatModel = failIfMissing(
      config.chatModel,
      config.mode === "vllm" ? "VLLM_CHAT_MODEL" : "OPENAI_CHAT_MODEL",
      config.mode
    );
    this.chatMaxTokens = config.chatMaxTokens;
    this.enableThinking = config.enableThinking;
    this.stripThinkOutput = config.stripThinkOutput;
  }

  async answer(input: AnswerInput) {
    const history = compactHistoryForPrompt(input.history);
    const evidence = input.evidence
      .slice(0, 4)
      .map(
        (chunk, index) =>
          `[证据 ${index + 1}] 流程 ${chunk.rowNumber}《${chunk.title}》${
            chunk.attachmentName ? ` / ${chunk.attachmentName}` : ""
          }\n${chunk.text}`
      )
      .join("\n\n");

    const systemPrompt =
      "你是公司内部流程助手。只能依据本轮提供的证据回答；不要编造，也不要复用历史对话里未在本轮证据出现的步骤。直接给结论和必要步骤，不要输出分析、草稿、思考过程、Thinking Process、<think> 标签或引用编号。若问题对应明确流程，优先回答该流程中的联系人与办理步骤。只有当前证据里明确存在的步骤才能写入答案。若当前证据不包含完整办理步骤，就明确说明“当前证据不包含完整办理步骤”，不要自行补全。若证据不足，只需明确说明“未找到明确依据”，并用一句话说明最接近的流程。回答默认使用中文，保留文档中的英文专有名词。请尽量简短：普通问题控制在 3-6 句，对比问题只总结最关键的 2-3 点。";
    const userPrompt = `历史对话:\n${history || "无"}\n\n问题:\n${input.question}\n\n证据:\n${evidence}`;

    const requestBody: Record<string, unknown> = {
      model: this.chatModel,
      temperature: 0.1,
      max_tokens: this.chatMaxTokens,
      messages: [
        {
          role: "system",
          content: systemPrompt
        },
        {
          role: "user",
          content: userPrompt
        }
      ]
    };

    if (this.mode === "vllm" && !this.enableThinking) {
      requestBody.chat_template_kwargs = { enable_thinking: false };
    }

    const response = await fetch(`${this.baseUrl}/chat/completions`, {
      method: "POST",
      headers: buildOpenAiHeaders(this.apiKey),
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`聊天模型调用失败: ${response.status} ${body}`);
    }

    const payload = (await response.json()) as {
      choices?: Array<{ message?: { content?: string } }>;
    };

    const rawAnswer = payload.choices?.[0]?.message?.content?.trim();
    const answer = this.stripThinkOutput && rawAnswer ? stripThinkBlocks(rawAnswer) : rawAnswer;
    if (!answer) {
      throw new Error("聊天模型没有返回有效答案");
    }

    return {
      answer,
      modelRequest: {
        providerMode: this.mode,
        endpoint: `${this.baseUrl}/chat/completions`,
        body: JSON.stringify(requestBody, null, 2)
      }
    };
  }
}

class CompositeProvider implements ModelProvider {
  readonly mode: string;
  private readonly chatProvider: ChatProvider;
  private readonly embeddingProvider: EmbeddingProvider;

  constructor(mode: string, chatProvider: ChatProvider, embeddingProvider: EmbeddingProvider) {
    this.mode = mode;
    this.chatProvider = chatProvider;
    this.embeddingProvider = embeddingProvider;
  }

  embedMany(texts: string[]) {
    return this.embeddingProvider.embedMany(texts);
  }

  answer(input: AnswerInput) {
    return this.chatProvider.answer(input);
  }
}

function createChatProviderByMode(mode: SimpleProviderMode): ChatProvider {
  if (mode === "offline") {
    return new OfflineChatProvider();
  }

  if (mode === "openai") {
    return new OpenAiCompatibleChatProvider(getOpenAiConfig());
  }

  return new OpenAiCompatibleChatProvider(getVllmConfig());
}

function createEmbeddingProviderByMode(mode: SimpleProviderMode): EmbeddingProvider {
  if (mode === "offline") {
    return new OfflineEmbeddingProvider();
  }

  if (mode === "openai") {
    return new OpenAiCompatibleEmbeddingProvider(getOpenAiConfig());
  }

  return new OpenAiCompatibleEmbeddingProvider(getVllmConfig());
}

export function createProvider(modeOverride?: string): ModelProvider {
  const resolved = parseCompositeMode(modeOverride);
  return new CompositeProvider(
    resolved.compositeMode,
    createChatProviderByMode(resolved.chatMode),
    createEmbeddingProviderByMode(resolved.embeddingMode)
  );
}

export function createAnswerProvider(knowledgeBaseProviderMode?: string): ModelProvider {
  const runtimeChatMode = getCurrentChatMode();
  const knowledgeBaseModes = parseCompositeMode(knowledgeBaseProviderMode);
  return new CompositeProvider(
    `chat:${runtimeChatMode}|embed:${knowledgeBaseModes.embeddingMode}`,
    createChatProviderByMode(runtimeChatMode),
    createEmbeddingProviderByMode(knowledgeBaseModes.embeddingMode)
  );
}

export { cosineSimilarity };
