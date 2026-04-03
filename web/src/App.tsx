import { FormEvent, useEffect, useMemo, useState } from "react";
import type { Citation, ImportJobRecord, ModelRequestDebug, UnansweredQuestionRecord } from "../../shared/contracts";
import { ApiError, fetchActiveKnowledge, fetchAuthStatus, fetchJob, loginWithPassword, logout, sendChat, uploadKnowledge } from "./api";
import type { ActiveKnowledgeResponse, AuthStatusResponse } from "./types";

interface UiMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  citations?: Citation[];
  modelRequest?: ModelRequestDebug;
  providerMode?: string;
}

const welcomeMessage: UiMessage = {
  id: "welcome",
  role: "assistant",
  content:
    "上传流程 Excel 或规范化知识库 zip 后，我会解析主表和附件，再根据检索到的证据回答问题。每次回答都会附上引用来源。"
};

function formatJobStatus(job?: ImportJobRecord): string {
  if (!job) {
    return "尚未导入知识库";
  }

  if (job.status === "completed") {
    return `导入完成，已解析 ${job.sourceCount} 条来源 / ${job.chunkCount} 个检索块`;
  }

  if (job.status === "failed") {
    return `导入失败：${job.error ?? "未知错误"}`;
  }

  return "知识库正在导入中";
}

function formatUnansweredReason(reason: UnansweredQuestionRecord["reason"]) {
  return reason === "insufficient_evidence" ? "未命中证据" : "模型拒答";
}

function formatErrorMessage(error: unknown, fallback: string) {
  return error instanceof Error ? error.message : fallback;
}

export default function App() {
  const [auth, setAuth] = useState<AuthStatusResponse | null>(null);
  const [password, setPassword] = useState("");
  const [authSubmitting, setAuthSubmitting] = useState(false);
  const [active, setActive] = useState<ActiveKnowledgeResponse | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [job, setJob] = useState<ImportJobRecord | undefined>(undefined);
  const [messages, setMessages] = useState<UiMessage[]>([welcomeMessage]);
  const [draft, setDraft] = useState("");
  const [sending, setSending] = useState(false);
  const [sessionId, setSessionId] = useState<string | undefined>(undefined);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void initializeApp();
  }, []);

  useEffect(() => {
    if (auth?.enabled && !auth.authenticated) {
      return;
    }

    if (!job || job.status === "completed" || job.status === "failed") {
      return;
    }

    const timer = window.setInterval(() => {
      void fetchJob(job.jobId)
        .then((next) => {
          setJob(next);
          if (next.status === "completed") {
            window.clearInterval(timer);
            resetConversation();
            void refreshActive();
          }
          if (next.status === "failed") {
            window.clearInterval(timer);
          }
        })
        .catch((jobError) => {
          if (handleApiError(jobError, "轮询导入任务失败")) {
            window.clearInterval(timer);
            return;
          }

          window.clearInterval(timer);
        });
    }, 1500);

    return () => window.clearInterval(timer);
  }, [auth, job]);

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
    setJob(undefined);
    setSelectedFile(null);
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
      setJob(payload.latestJob);
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
      setJob(undefined);
      setSelectedFile(null);
      setPassword("");
      resetConversation();
    } catch (logoutError) {
      setError(formatErrorMessage(logoutError, "退出失败"));
    }
  }

  async function handleUpload(event: FormEvent) {
    event.preventDefault();
    if (!selectedFile) {
      setError("请先选择 Excel 文件");
      return;
    }

    setError(null);
    setUploading(true);

    try {
      const result = await uploadKnowledge(selectedFile);
      const nextJob = await fetchJob(result.jobId);
      setJob(nextJob);
    } catch (uploadError) {
      handleApiError(uploadError, "上传失败");
    } finally {
      setUploading(false);
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
          citations: response.citations,
          modelRequest: response.modelRequest,
          providerMode: response.providerMode
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

  const kbCard = useMemo(() => {
    if (!active?.knowledgeBase) {
      return {
        headline: "当前没有活动知识库",
        subline: "上传 Excel 后会在这里显示版本信息。"
      };
    }

    return {
      headline: active.knowledgeBase.originalFileName,
      subline: `版本时间 ${new Date(active.knowledgeBase.importedAt).toLocaleString("zh-CN")} · ${active.knowledgeBase.sourceCount} 条来源 · ${active.knowledgeBase.chunkCount} 个检索块`
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
          <p className="eyebrow">Corp Workflow Assistant</p>
          <h1>正在加载</h1>
          <p className="hero-text">正在读取访问状态和活动知识库。</p>
        </section>
      </main>
    );
  }

  if (auth.enabled && !auth.authenticated) {
    return (
      <main className="auth-shell">
        <section className="auth-card">
          <p className="eyebrow">Protected Access</p>
          <h1>输入访问密码</h1>
          <p className="hero-text">这个页面已加密码，仅供内部小范围使用。</p>

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
              {authSubmitting ? "验证中..." : "进入系统"}
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
          <p className="eyebrow">Corp Workflow Assistant</p>
          <h1>把流程汇总表变成可追溯的本地问答机器人</h1>
          <p className="hero-text">
            上传新版 Excel 后，系统会解析主表、内嵌 Word / PowerPoint / Excel 附件，并基于证据回答问题。
          </p>
        </div>

        <div className="active-card">
          <div className="active-card-top">
            <div>
              <h2>活动知识库</h2>
              <p className="active-title">{kbCard.headline}</p>
              <p className="active-subline">{kbCard.subline}</p>
            </div>
            {auth.enabled ? (
              <button type="button" className="ghost-button" onClick={handleLogout}>
                退出
              </button>
            ) : null}
          </div>
          <p className="status-pill">{formatJobStatus(job)}</p>
          <div className="stats-block">
            <p className="stats-title">{statsCard.headline}</p>
            <p className="stats-subline">{statsCard.subline}</p>
          </div>
        </div>
      </section>

      <section className="workspace-grid">
        <aside className="import-panel">
          <div className="panel-header">
            <h2>导入知识库</h2>
            <p>支持上传旧版 Excel 或规范化知识库 zip。导入成功后，新对话会自动绑定到最新版本。</p>
          </div>

          <form className="upload-form" onSubmit={handleUpload}>
            <label className="file-dropzone">
              <span>选择 `.xlsx` 或 `.zip` 文件</span>
              <input
                type="file"
                accept=".xlsx,.zip"
                onChange={(event) => setSelectedFile(event.target.files?.[0] ?? null)}
              />
              <strong>{selectedFile?.name ?? "未选择文件"}</strong>
            </label>

            <button className="primary-button" disabled={uploading}>
              {uploading ? "上传中..." : "开始导入"}
            </button>
          </form>

          {job?.warnings.length ? (
            <div className="warning-box">
              <h3>导入提示</h3>
              {job.warnings.map((warning) => (
                <p key={warning}>{warning}</p>
              ))}
            </div>
          ) : null}

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
            <h2>聊天窗口</h2>
            <p>问法可以自然一些，例如“安装新软件要走什么单？”</p>
          </div>

          <div className="message-list">
            {messages.map((message) => (
              <article key={message.id} className={`message-card ${message.role}`}>
                <p className="message-role">{message.role === "assistant" ? "助手" : "你"}</p>
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
                        {citation.url ? (
                          <a href={citation.url} target="_blank" rel="noreferrer">
                            打开外链
                          </a>
                        ) : null}
                      </div>
                    ))}
                  </div>
                ) : null}

                {message.role === "assistant" && message.id !== welcomeMessage.id ? (
                  <details className="model-request-card">
                    <summary>查看发送给模型的内容</summary>
                    {message.modelRequest ? (
                      <>
                        <p className="model-request-meta">
                          {message.modelRequest.providerMode} · {message.modelRequest.endpoint}
                        </p>
                        <pre className="model-request-body">{message.modelRequest.body}</pre>
                      </>
                    ) : (
                      <p className="model-request-meta">
                        本次未调用模型。
                        {message.providerMode?.startsWith("chat:offline")
                          ? " 当前聊天 provider 为 offline。"
                          : " 后端已直接返回拒答文案，没有发送模型请求。"}
                      </p>
                    )}
                  </details>
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
    </main>
  );
}
