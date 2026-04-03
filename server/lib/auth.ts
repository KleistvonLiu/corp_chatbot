import { createHash, timingSafeEqual } from "node:crypto";
import type { NextFunction, Request, Response } from "express";

const cookieName = "corp_chatbot_auth";

export interface AuthStatusResponse {
  enabled: boolean;
  authenticated: boolean;
}

function readPassword() {
  const password = process.env.APP_PASSWORD?.trim();
  return password ? password : undefined;
}

function buildToken(password: string) {
  return createHash("sha256").update(`corp-chatbot:${password}`).digest("base64url");
}

function parseCookies(cookieHeader?: string) {
  if (!cookieHeader) {
    return new Map<string, string>();
  }

  return new Map(
    cookieHeader
      .split(";")
      .map((segment) => segment.trim())
      .filter(Boolean)
      .map((segment) => {
        const separatorIndex = segment.indexOf("=");
        if (separatorIndex < 0) {
          return [segment, ""] as const;
        }

        return [segment.slice(0, separatorIndex), decodeURIComponent(segment.slice(separatorIndex + 1))] as const;
      })
  );
}

function safeEqual(left: string, right: string) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  if (leftBuffer.length !== rightBuffer.length) {
    return false;
  }

  return timingSafeEqual(leftBuffer, rightBuffer);
}

export function isAuthEnabled() {
  return Boolean(readPassword());
}

export function isAuthenticated(request: Request) {
  const password = readPassword();
  if (!password) {
    return true;
  }

  const cookies = parseCookies(request.headers.cookie);
  const actual = cookies.get(cookieName);
  if (!actual) {
    return false;
  }

  return safeEqual(actual, buildToken(password));
}

function setAuthCookie(response: Response, password: string) {
  response.cookie(cookieName, buildToken(password), {
    httpOnly: true,
    sameSite: "lax",
    path: "/"
  });
}

export function clearAuthCookie(response: Response) {
  response.clearCookie(cookieName, {
    httpOnly: true,
    sameSite: "lax",
    path: "/"
  });
}

export function getAuthStatus(request: Request): AuthStatusResponse {
  return {
    enabled: isAuthEnabled(),
    authenticated: isAuthenticated(request)
  };
}

export function loginWithPassword(request: Request, response: Response) {
  const password = readPassword();
  if (!password) {
    return {
      enabled: false,
      authenticated: true
    } satisfies AuthStatusResponse;
  }

  const body = request.body as { password?: string };
  if (body.password !== password) {
    clearAuthCookie(response);
    response.status(401);
    return {
      enabled: true,
      authenticated: false
    } satisfies AuthStatusResponse;
  }

  setAuthCookie(response, password);
  return {
    enabled: true,
    authenticated: true
  } satisfies AuthStatusResponse;
}

export function logout(response: Response) {
  clearAuthCookie(response);
  return {
    enabled: isAuthEnabled(),
    authenticated: false
  } satisfies AuthStatusResponse;
}

export function requireAuth(request: Request, response: Response, next: NextFunction) {
  if (isAuthenticated(request)) {
    next();
    return;
  }

  response.status(401).json({ error: "请输入访问密码。" });
}

export function sendLoginPage(response: Response) {
  response
    .status(401)
    .setHeader("Cache-Control", "no-store")
    .type("html")
    .send(`<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>输入访问密码</title>
    <style>
      :root {
        color-scheme: light;
        font-family: "IBM Plex Sans", sans-serif;
        --bg: #f4efe7;
        --panel: rgba(255, 251, 245, 0.92);
        --line: rgba(90, 59, 29, 0.15);
        --ink: #20150f;
        --muted: #6d5647;
        --accent: #9c3f19;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        padding: 24px;
        background:
          radial-gradient(circle at top left, rgba(219, 152, 92, 0.22), transparent 28%),
          radial-gradient(circle at bottom right, rgba(49, 116, 107, 0.18), transparent 30%),
          linear-gradient(160deg, #efe4d3 0%, #f7f2ea 54%, #ede3d1 100%);
        color: var(--ink);
      }
      .card {
        width: min(460px, 100%);
        padding: 34px;
        border-radius: 28px;
        border: 1px solid var(--line);
        background: var(--panel);
        box-shadow: 0 22px 50px rgba(73, 48, 17, 0.08);
      }
      .eyebrow {
        margin: 0 0 10px;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--accent);
        font-size: 12px;
        font-weight: 700;
      }
      h1 { margin: 0; font-size: 34px; }
      p { color: var(--muted); }
      form { display: grid; gap: 14px; margin-top: 20px; }
      input {
        border-radius: 16px;
        border: 1px solid var(--line);
        background: rgba(255, 251, 245, 0.92);
        padding: 14px 16px;
        font: inherit;
        color: var(--ink);
      }
      button {
        border: none;
        border-radius: 16px;
        padding: 13px 18px;
        background: linear-gradient(135deg, #ba4c1c, #7b2a17);
        color: white;
        cursor: pointer;
        font: inherit;
        font-weight: 700;
      }
      button:disabled { opacity: 0.65; cursor: default; }
      .error { min-height: 24px; margin: 12px 0 0; color: #9f3120; }
    </style>
  </head>
  <body>
    <section class="card">
      <p class="eyebrow">Protected Access</p>
      <h1>输入访问密码</h1>
      <p>这个页面已加密码，仅供内部小范围使用。</p>
      <form id="login-form">
        <input id="password" type="password" placeholder="访问密码" autocomplete="current-password" autofocus />
        <button id="submit" type="submit">进入系统</button>
      </form>
      <p id="error" class="error"></p>
    </section>
    <script>
      const form = document.getElementById("login-form");
      const passwordInput = document.getElementById("password");
      const submitButton = document.getElementById("submit");
      const errorNode = document.getElementById("error");
      form.addEventListener("submit", async (event) => {
        event.preventDefault();
        const password = passwordInput.value.trim();
        if (!password) {
          errorNode.textContent = "请输入访问密码。";
          return;
        }
        errorNode.textContent = "";
        submitButton.disabled = true;
        submitButton.textContent = "验证中...";
        try {
          const response = await fetch("/api/auth/login", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ password })
          });
          const payload = await response.json().catch(() => null);
          if (!response.ok || !payload?.authenticated) {
            errorNode.textContent = payload?.error || "访问密码错误。";
            submitButton.disabled = false;
            submitButton.textContent = "进入系统";
            return;
          }
          window.location.replace("/");
        } catch {
          errorNode.textContent = "登录失败，请稍后重试。";
          submitButton.disabled = false;
          submitButton.textContent = "进入系统";
        }
      });
    </script>
  </body>
</html>`);
}
