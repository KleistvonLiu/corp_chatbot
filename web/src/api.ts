import type {
  ActiveKnowledgeResponse,
  ChatResponseBody,
  ImportJobRecord,
  ImportResponse
} from "./types";

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    const payload = (await response.json().catch(() => null)) as { error?: string } | null;
    throw new Error(payload?.error ?? `请求失败: ${response.status}`);
  }

  return (await response.json()) as T;
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
