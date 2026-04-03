import type {
  ActiveKnowledgeResponse,
  AuthStatusResponse,
  ChatResponseBody,
  ImportJobRecord,
  ImportResponse
} from "./types";

export class ApiError extends Error {
  status: number;

  constructor(status: number, message: string) {
    super(message);
    this.name = "ApiError";
    this.status = status;
  }
}

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const payload = (await response.json().catch(() => null)) as { error?: string } | null;
    throw new ApiError(response.status, payload?.error ?? `请求失败: ${response.status}`);
  }

  return (await response.json()) as T;
}

export async function fetchAuthStatus(): Promise<AuthStatusResponse> {
  return parseJson<AuthStatusResponse>(await fetch("/api/auth/status"));
}

export async function loginWithPassword(password: string): Promise<AuthStatusResponse> {
  return parseJson<AuthStatusResponse>(
    await fetch("/api/auth/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ password })
    })
  );
}

export async function logout(): Promise<AuthStatusResponse> {
  return parseJson<AuthStatusResponse>(
    await fetch("/api/auth/logout", {
      method: "POST"
    })
  );
}

export async function fetchActiveKnowledge(): Promise<ActiveKnowledgeResponse> {
  return parseJson<ActiveKnowledgeResponse>(await fetch("/api/knowledge/active"));
}

export async function uploadKnowledge(file: File): Promise<ImportResponse> {
  const formData = new FormData();
  formData.append("file", file);

  return parseJson<ImportResponse>(
    await fetch("/api/knowledge/import", {
      method: "POST",
      body: formData
    })
  );
}

export async function fetchJob(jobId: string): Promise<ImportJobRecord> {
  return parseJson<ImportJobRecord>(await fetch(`/api/knowledge/jobs/${jobId}`));
}

export async function sendChat(message: string, sessionId?: string): Promise<ChatResponseBody> {
  return parseJson<ChatResponseBody>(
    await fetch("/api/chat", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ message, sessionId })
    })
  );
}
