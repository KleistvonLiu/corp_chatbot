import { promises as fs } from "node:fs";
import path from "node:path";
import express from "express";
import { answerQuestion, getActiveKnowledgeResponse, getImportJob } from "./lib/knowledge";
import { getAuthStatus, isAuthEnabled, isAuthenticated, loginWithPassword, logout, requireAuth, sendLoginPage } from "./lib/auth";
import { ensureStorage } from "./lib/storage";

export async function createApp() {
  await ensureStorage();
  const app = express();
  app.use(express.json({ limit: "1mb" }));

  app.get("/api/auth/status", (request, response) => {
    response.json(getAuthStatus(request));
  });

  app.post("/api/auth/login", (request, response) => {
    const result = loginWithPassword(request, response);
    if (!result.authenticated && result.enabled) {
      response.json({ ...result, error: "访问密码错误。" });
      return;
    }

    response.json(result);
  });

  app.post("/api/auth/logout", (_request, response) => {
    response.json(logout(response));
  });

  app.use((request, response, next) => {
    if (!isAuthEnabled() || isAuthenticated(request)) {
      next();
      return;
    }

    if (request.path.startsWith("/api/auth")) {
      next();
      return;
    }

    if (request.path.startsWith("/api")) {
      response.status(401).json({ error: "请输入访问密码。" });
      return;
    }

    sendLoginPage(response);
  });

  app.use("/api", requireAuth);

  app.get("/api/knowledge/active", async (_request, response, next) => {
    try {
      response.json(await getActiveKnowledgeResponse());
    } catch (error) {
      next(error);
    }
  });

  app.get("/api/knowledge/jobs/:jobId", async (request, response, next) => {
    try {
      const job = await getImportJob(request.params.jobId);
      if (!job) {
        response.status(404).json({ error: "导入任务不存在" });
        return;
      }

      response.json(job);
    } catch (error) {
      next(error);
    }
  });

  app.post("/api/knowledge/import", (_request, response) => {
    response.status(410).json({ error: "当前版本不支持网页上传知识库，请替换固定知识源文件后重启服务。" });
  });

  app.post("/api/chat", async (request, response, next) => {
    try {
      const body = request.body as { message?: string; sessionId?: string };
      const message = body.message?.trim();
      if (!message) {
        response.status(400).json({ error: "message 不能为空。" });
        return;
      }

      response.json(await answerQuestion(message, body.sessionId));
    } catch (error) {
      next(error);
    }
  });

  const webDistDir = path.join(process.cwd(), "dist", "web");
  try {
    await fs.access(webDistDir);
    app.use(express.static(webDistDir));
    app.get("*", (request, response, next) => {
      if (request.path.startsWith("/api")) {
        next();
        return;
      }
      response.sendFile(path.join(webDistDir, "index.html"));
    });
  } catch {
    // Dev mode: the Vite dev server serves the frontend separately.
  }

  app.use(
    (error: unknown, _request: express.Request, response: express.Response, _next: express.NextFunction) => {
      const message = error instanceof Error ? error.message : "服务器内部错误";
      response.status(500).json({ error: message });
    }
  );

  return app;
}
