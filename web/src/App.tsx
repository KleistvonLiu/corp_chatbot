import { FormEvent, useEffect, useMemo, useState } from "react";
import type { Citation, CitationImage, UnansweredQuestionRecord } from "../../shared/contracts";
import { ApiError, fetchActiveKnowledge, fetchAuthStatus, loginWithPassword, logout, sendChat } from "./api";
import type { ActiveKnowledgeResponse, AuthStatusResponse } from "./types";

interface UiMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  citations?: Citation[];
}

interface ImagePreview {
  image: CitationImage;
  title: string;
  rowNumber: number;
}

const welcomeMessage: UiMessage = {
  id: "welcome",
  role: "assistant",
  content: "我是 E小助，会基于固定流程知识库和附件证据回答问题，并附上引用来源。"
};

function formatUnansweredReason(reason: UnansweredQuestionRecord["reason"]) {
  return reason === "insufficient_evidence" ? "未命中证据" : "模型拒答";
}

function formatErrorMessage(error: unknown, fallback: string) {
  return error instanceof Error ? error.message : fallback;
}

function toDisplayName(filePath?: string) {
  if (!filePath) {
    return "未配置";
  }

  return filePath.split(/[\\/]/).filter(Boolean).pop() ?? filePath;
}

export default function App() {
  const [auth, setAuth] = useState<AuthStatusResponse | null>(null);
  const [password, setPassword] = useState("");
  const [authSubmitting, setAuthSubmitting] = useState(false);
  const [active, setActive] = useState<ActiveKnowledgeResponse | null>(null);
  const [messages, setMessages] = useState<UiMessage[]>([welcomeMessage]);
  const [draft, setDraft] = useState("");
  const [sending, setSending] = useState(false);
  const [sessionId, setSessionId] = useState<string | undefined>(undefined);
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<ImagePreview | null>(null);

  useEffect(() => {
    void initializeApp();
  }, []);

  function resetConversation() {
    setSessionId(undefined);
    setMessages([welcomeMessage]);
  }

  function markUnauthenticated() {
    setAuth((current) => ({
      enabled: current?.enabled ?? true,
      authenticated: false
    }));
    setActive(null);
    resetConversation();
  }

  function handleApiError(error: unknown, fallback: string) {
    if (error instanceof ApiError && error.status === 401) {
      markUnauthenticated();
      setError("请输入访问密码。");
      return true;
    }

    setError(formatErrorMessage(error, fallback));
    return false;
  }

  async function initializeApp() {
    try {
      const status = await fetchAuthStatus();
      setAuth(status);
      if (!status.enabled || status.authenticated) {
        await refreshActive();
      }
    } catch (authError) {
      setAuth({
        enabled: false,
        authenticated: true
      });
      setError(formatErrorMessage(authError, "读取认证状态失败"));
    }
  }

  async function refreshActive() {
    try {
      const payload = await fetchActiveKnowledge();
      setActive(payload);
    } catch (activeError) {
      handleApiError(activeError, "读取知识库状态失败");
    }
  }

  async function handleLoginSubmit(event: FormEvent) {
    event.preventDefault();
    if (!password.trim()) {
      setError("请输入访问密码。");
      return;
    }

    setError(null);
    setAuthSubmitting(true);

    try {
      const status = await loginWithPassword(password);
      setAuth(status);
      setPassword("");
      await refreshActive();
    } catch (loginError) {
      setError(formatErrorMessage(loginError, "登录失败"));
    } finally {
      setAuthSubmitting(false);
    }
  }

  async function handleLogout() {
    setError(null);

    try {
      const status = await logout();
      setAuth(status);
      setActive(null);
      setPassword("");
      resetConversation();
    } catch (logoutError) {
      setError(formatErrorMessage(logoutError, "退出失败"));
    }
  }

  async function handleChatSubmit(event: FormEvent) {
    event.preventDefault();
    const message = draft.trim();
    if (!message) {
      return;
    }

    setDraft("");
    setError(null);
    const userMessage: UiMessage = {
      id: crypto.randomUUID(),
      role: "user",
      content: message
    };
    setMessages((current) => [...current, userMessage]);
    setSending(true);

    try {
      const response = await sendChat(message, sessionId);
      setSessionId(response.sessionId);
      setActive((current) =>
        current
          ? {
              ...current,
              questionStats: response.questionStats
            }
          : current
      );
      setMessages((current) => [
        ...current,
        {
          id: crypto.randomUUID(),
          role: "assistant",
          content: response.answer,
          citations: response.citations
        }
      ]);
    } catch (chatError) {
      if (!handleApiError(chatError, "发送消息失败")) {
        setMessages((current) => current.filter((item) => item.id !== userMessage.id || item.role !== "user"));
        setDraft(message);
      }
    } finally {
      setSending(false);
    }
  }

  const knowledgeCard = useMemo(() => {
    if (!active?.knowledgeBase) {
      return {
        headline: "当前没有活动知识库",
        subline: "请检查固定知识源配置和启动日志。"
      };
    }

    return {
      headline: active.knowledgeBase.originalFileName,
      subline: `版本时间 ${new Date(active.knowledgeBase.importedAt).toLocaleString("zh-CN")} · ${active.knowledgeBase.sourceCount} 条来源 · ${active.knowledgeBase.chunkCount} 个检索块`
    };
  }, [active]);

  const sourceCard = useMemo(() => {
    const fixedSource = active?.fixedSource;
    if (!fixedSource?.configured) {
      return {
        title: "固定知识源未配置",
        subline: "请在 .env 中设置 KNOWLEDGE_SOURCE_WORKBOOK_PATH 和 KNOWLEDGE_SOURCE_ATTACHMENTS_DIR。",
        error: undefined as string | undefined
      };
    }

    return {
      title: toDisplayName(fixedSource.workbookPath),
      subline: fixedSource.lastSyncAt
        ? `最近同步 ${new Date(fixedSource.lastSyncAt).toLocaleString("zh-CN")} · 附件目录 ${toDisplayName(fixedSource.attachmentsDir)}`
        : `附件目录 ${toDisplayName(fixedSource.attachmentsDir)}`,
      error: fixedSource.syncError
    };
  }, [active]);

  const statsCard = useMemo(() => {
    const stats = active?.questionStats;
    if (!stats) {
      return {
        headline: "未答问题统计尚未生成",
        subline: "开始提问后，这里会累计显示未回答问题的数量。"
      };
    }

    const rate = stats.totalQuestions ? `${Math.round((stats.unansweredCount / stats.totalQuestions) * 100)}%` : "0%";
    return {
      headline: `累计 ${stats.unansweredCount} 个问题未回答`,
      subline: `总提问 ${stats.totalQuestions} 次 · 未回答占比 ${rate}`
    };
  }, [active]);

  if (!auth) {
    return (
      <main className="auth-shell">
        <section className="auth-card">
          <p className="eyebrow">E小助</p>
          <h1>正在加载</h1>
          <p className="hero-text">正在读取访问状态和固定知识库。</p>
        </section>
      </main>
    );
  }

  if (auth.enabled && !auth.authenticated) {
    return (
      <main className="auth-shell">
        <section className="auth-card">
          <p className="eyebrow">E小助</p>
          <h1>输入访问密码</h1>
          <p className="hero-text">E小助 已加密码，仅供内部小范围使用。</p>

          <form className="password-form" onSubmit={handleLoginSubmit}>
            <input
              type="password"
              placeholder="访问密码"
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              autoComplete="current-password"
              disabled={authSubmitting}
              autoFocus
            />
            <button className="primary-button" disabled={authSubmitting || !password.trim()}>
              {authSubmitting ? "验证中..." : "进入 E小助"}
            </button>
          </form>

          {error ? <p className="error-text">{error}</p> : null}
        </section>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <section className="hero-panel">
        <div className="hero-copy">
          <p className="eyebrow">E小助</p>
          <h1>固定流程知识库问答助手</h1>
          <p className="hero-text">E小助 会在服务启动时自动同步固定知识源，并基于主表和附件证据回答问题。</p>
        </div>

        <div className="active-card">
          <div className="active-card-top">
            <div>
              <h2>活动知识库</h2>
              <p className="active-title">{knowledgeCard.headline}</p>
              <p className="active-subline">{knowledgeCard.subline}</p>
            </div>
            {auth.enabled ? (
              <button type="button" className="ghost-button" onClick={handleLogout}>
                退出
              </button>
            ) : null}
          </div>
          <div className="stats-block">
            <p className="stats-title">固定知识源</p>
            <p className="stats-subline">{sourceCard.title}</p>
            <p className="stats-subline">{sourceCard.subline}</p>
            {sourceCard.error ? <p className="error-text">{sourceCard.error}</p> : null}
          </div>
          <div className="stats-block">
            <p className="stats-title">{statsCard.headline}</p>
            <p className="stats-subline">{statsCard.subline}</p>
          </div>
        </div>
      </section>

      <section className="workspace-grid single-pane">
        <aside className="import-panel">
          <div className="panel-header">
            <h2>系统说明</h2>
            <p>知识库更新方式已固定。请替换服务端配置的 Excel 和附件目录后重启服务，不再支持网页上传。</p>
          </div>

          {active?.questionStats?.recentUnanswered.length ? (
            <div className="warning-box">
              <h3>最近未回答问题</h3>
              {active.questionStats.recentUnanswered.slice(0, 5).map((item, index) => (
                <p key={`${item.createdAt}-${index}`}>
                  {item.question}
                  <br />
                  <small>
                    {new Date(item.createdAt).toLocaleString("zh-CN")} · {formatUnansweredReason(item.reason)}
                  </small>
                </p>
              ))}
            </div>
          ) : null}

          {error ? <p className="error-text">{error}</p> : null}
        </aside>

        <section className="chat-panel">
          <div className="panel-header">
            <h2>E小助 对话窗口</h2>
            <p>例如“安装新软件要走什么单？”、“软件测试找谁？”</p>
          </div>

          <div className="message-list">
            {messages.map((message) => (
              <article key={message.id} className={`message-card ${message.role}`}>
                <p className="message-role">{message.role === "assistant" ? "E小助" : "你"}</p>
                <p className="message-content">{message.content}</p>

                {message.citations?.length ? (
                  <div className="citation-list">
                    {message.citations.map((citation) => (
                      <div key={`${message.id}-${citation.sourceId}`} className="citation-card">
                        <p className="citation-title">
                          流程 {citation.rowNumber}: {citation.title}
                          {citation.attachmentName ? ` · ${citation.attachmentName}` : ""}
                        </p>
                        <p className="citation-snippet">{citation.snippet}</p>
                        {citation.images?.length ? (
                          <div className="citation-image-grid">
                            {citation.images.map((image) => (
                              <button
                                key={image.sourceId}
                                type="button"
                                className="citation-image-button"
                                onClick={() =>
                                  setPreview({
                                    image,
                                    title: citation.title,
                                    rowNumber: citation.rowNumber
                                  })
                                }
                              >
                                <img src={image.url} alt={image.label} loading="lazy" className="citation-image" />
                                <span className="citation-image-label">{image.label}</span>
                              </button>
                            ))}
                          </div>
                        ) : null}
                        {citation.url ? (
                          <a href={citation.url} target="_blank" rel="noreferrer">
                            打开外链
                          </a>
                        ) : null}
                      </div>
                    ))}
                  </div>
                ) : null}

              </article>
            ))}
          </div>

          <form className="composer" onSubmit={handleChatSubmit}>
            <textarea
              placeholder="输入你的问题"
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              disabled={sending}
            />
            <button className="primary-button" disabled={sending || !draft.trim()}>
              {sending ? "回答中..." : "发送"}
            </button>
          </form>
        </section>
      </section>

      {preview ? (
        <div className="image-lightbox" role="dialog" aria-modal="true" onClick={() => setPreview(null)}>
          <div className="image-lightbox-card" onClick={(event) => event.stopPropagation()}>
            <button type="button" className="image-lightbox-close" onClick={() => setPreview(null)}>
              关闭
            </button>
            <p className="image-lightbox-title">
              流程 {preview.rowNumber}: {preview.title}
            </p>
            <p className="image-lightbox-label">{preview.image.label}</p>
            <img src={preview.image.url} alt={preview.image.label} className="image-lightbox-image" />
          </div>
        </div>
      ) : null}
    </main>
  );
}
