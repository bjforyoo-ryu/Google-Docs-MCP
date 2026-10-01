import { Env, OAuthTokenResponse } from "./types";
import { BASE_SECURITY_HEADERS } from "./utils";

// ── AES-GCM helpers ────────────────────────────────────────────────────────────

async function importAesKey(base64Key: string): Promise<CryptoKey> {
  const raw = Uint8Array.from(atob(base64Key), c => c.charCodeAt(0));
  return crypto.subtle.importKey("raw", raw, { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);
}

async function encryptData(plaintext: string, base64Key: string): Promise<string> {
  const key = await importAesKey(base64Key);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv },
    key,
    new TextEncoder().encode(plaintext)
  );
  return JSON.stringify({
    v: 1,
    iv: btoa(String.fromCharCode(...iv)),
    ct: btoa(String.fromCharCode(...new Uint8Array(ciphertext))),
  });
}

async function decryptData(stored: string, base64Key: string): Promise<string> {
  const { v, iv: ivB64, ct: ctB64 } = JSON.parse(stored);
  if (v !== 1) throw new Error("Unsupported encryption version");
  const key = await importAesKey(base64Key);
  const iv = Uint8Array.from(atob(ivB64), c => c.charCodeAt(0));
  const ct = Uint8Array.from(atob(ctB64), c => c.charCodeAt(0));
  const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv }, key, ct);
  return new TextDecoder().decode(plain);
}

// ── OAuth constants ────────────────────────────────────────────────────────────

const REQUIRED_SCOPES = [
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
];

const TOKEN_LIFETIME_MS = 90 * 24 * 60 * 60 * 1000; // 90 days

// ── HTML escaping ──────────────────────────────────────────────────────────────

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export async function handleAuthLogin(request: Request, env: Env): Promise<Response> {
  const state = crypto.randomUUID();
  const ip = request.headers.get("CF-Connecting-IP") ?? "unknown";

  // Bind state to client IP to prevent fixation and cross-IP replay
  await env.TOKENS.put(`state:${state}`, ip, { expirationTtl: 300 });
  
  const googleUrl = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  googleUrl.searchParams.set("client_id", env.GOOGLE_CLIENT_ID);
  googleUrl.searchParams.set("redirect_uri", env.GOOGLE_REDIRECT_URI);
  googleUrl.searchParams.set("response_type", "code");
  googleUrl.searchParams.set("scope", "https://www.googleapis.com/auth/documents https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets");
  googleUrl.searchParams.set("access_type", "offline");
  googleUrl.searchParams.set("prompt", "consent");
  googleUrl.searchParams.set("state", state);

  return Response.redirect(googleUrl.toString(), 302);
}

export async function handleAuthCallback(request: Request, env: Env): Promise<Response> {
  const url = new URL(request.url);
  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");

  if (!code || !state) {
    return new Response("Missing code or state", { status: 400 });
  }

  // F11: Verify state and IP binding
  const storedIp = await env.TOKENS.get(`state:${state}`);
  if (!storedIp) {
    return new Response("Invalid or expired state (CSRF protection)", { status: 403 });
  }
  const requestIp = request.headers.get("CF-Connecting-IP") ?? "unknown";
  if (storedIp !== requestIp) {
    await env.TOKENS.delete(`state:${state}`);
    return new Response("State IP mismatch (CSRF protection)", { status: 403 });
  }
  // F11: Delete state BEFORE token exchange to close the replay race window
  await env.TOKENS.delete(`state:${state}`);

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: env.GOOGLE_CLIENT_ID,
      client_secret: env.GOOGLE_CLIENT_SECRET,
      redirect_uri: env.GOOGLE_REDIRECT_URI,
      grant_type: "authorization_code",
    }),
  });

  // F12: Don't surface raw OAuth error details to the browser
  let tokens: OAuthTokenResponse;
  try {
    tokens = await response.json() as OAuthTokenResponse;
  } catch {
    console.error("Token exchange: failed to parse JSON response");
    return new Response("Authentication failed. Please try signing in again.", { status: 502 });
  }
  if (tokens.error) {
    console.error("Token exchange error:", tokens.error, tokens.error_description);
    return new Response("Authentication failed. Please try signing in again.", { status: 400 });
  }

  // F13: Verify all required scopes were granted
  if (tokens.scope) {
    const granted = tokens.scope.split(/\s+/);
    const missing = REQUIRED_SCOPES.filter(s => !granted.includes(s));
    if (missing.length > 0) {
      console.error("Insufficient OAuth scopes granted. Missing:", missing);
      return new Response(`<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>Insufficient Permissions</title>
    <style>body { font-family: system-ui, sans-serif; max-width: 600px; margin: 40px auto; padding: 20px; }</style>
  </head>
  <body>
    <h1>Permissions Required</h1>
    <p>The required Google permissions were not granted. Please sign in again and approve <strong>all</strong> requested permissions.</p>
    <p><a href="/auth/login">Try again</a></p>
  </body>
</html>`, {
        status: 403,
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Content-Security-Policy": "default-src 'none'; style-src 'unsafe-inline'",
          ...BASE_SECURITY_HEADERS,
        },
      });
    }
  }

  // F14: Record access token expiry (Google) and our own 90-day session lifetime
  tokens.expiry_date = Date.now() + ((tokens.expires_in ?? 3600) * 1000);
  tokens.createdAt = Date.now();
  tokens.expiresAt = Date.now() + TOKEN_LIFETIME_MS;

  const userToken = crypto.randomUUID();
  await env.TOKENS.put(userToken, await encryptData(JSON.stringify(tokens), env.TOKEN_ENCRYPTION_KEY));

  const pageOrigin = new URL(request.url).origin;
  const mcpUrl = `${pageOrigin}/mcp?token=${userToken}`;
  const mcpUrlLegacy = `${pageOrigin}/mcp/sse?token=${userToken}`;

  const nonce = crypto.randomUUID().replace(/-/g, "");
  const safeMcpUrl = escapeHtml(mcpUrl);
  const safeMcpUrlLegacy = escapeHtml(mcpUrlLegacy);

  return new Response(`<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>Authentication Successful</title>
    <style>
      body { font-family: system-ui, sans-serif; max-width: 600px; margin: 40px auto; padding: 20px; line-height: 1.5; }
      .token-box { display: flex; align-items: center; background: #f4f4f4; padding: 10px; border: 1px solid #ddd; border-radius: 4px; margin: 20px 0; }
      code { flex-grow: 1; word-break: break-all; margin-right: 10px; }
      button { cursor: pointer; padding: 6px 12px; background: #0066cc; color: white; border: none; border-radius: 4px; }
      button:hover { background: #0052a3; }
    </style>
  </head>
  <body>
    <h1>Authentication Successful!</h1>
    <p>Your unique MCP connection URL (Streamable HTTP — recommended):</p>
    <div class="token-box">
      <code id="mcp-url">${safeMcpUrl}</code>
      <button id="copy-btn">Copy URL</button>
    </div>
    <p>Legacy SSE URL (only if required by your client):</p>
    <div class="token-box">
      <code id="mcp-url-legacy">${safeMcpUrlLegacy}</code>
      <button id="copy-btn-legacy">Copy URL</button>
    </div>
    <p>Copy the Streamable HTTP URL into your Claude.ai Web MCP configuration.</p>
    <p><strong>Security Warning:</strong> This URL contains your private access token. Do not share it. If leaked, anyone can edit your documents.</p>
    <p><a href="/">Return to Home</a></p>

    <script nonce="${nonce}">
      function copyToClipboard(elId, btnId) {
        const urlText = document.getElementById(elId).innerText;
        navigator.clipboard.writeText(urlText).then(function() {
          const btn = document.getElementById(btnId);
          btn.innerText = 'Copied!';
          setTimeout(function() { btn.innerText = 'Copy URL'; }, 2000);
        }).catch(function(err) {
          console.error('Failed to copy:', err);
        });
      }
      document.getElementById('copy-btn').addEventListener('click', function() { copyToClipboard('mcp-url', 'copy-btn'); });
      document.getElementById('copy-btn-legacy').addEventListener('click', function() { copyToClipboard('mcp-url-legacy', 'copy-btn-legacy'); });
    </script>
  </body>
</html>`, {
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "Content-Security-Policy": `default-src 'none'; style-src 'unsafe-inline'; script-src 'nonce-${nonce}'; form-action 'self'`,
      ...BASE_SECURITY_HEADERS,
    }
  });
}

export async function getAccessToken(userToken: string, env: Env): Promise<string | null> {
  const data = await env.TOKENS.get(userToken);
  if (!data) return null;

  let tokens: OAuthTokenResponse;
  try {
    tokens = JSON.parse(await decryptData(data, env.TOKEN_ENCRYPTION_KEY)) as OAuthTokenResponse;
  } catch {
    // Decryption failed — plaintext legacy entry or corrupted; require re-authentication
    return null;
  }

  // F14: Enforce 90-day session lifetime
  if (tokens.expiresAt && Date.now() > tokens.expiresAt) {
    await env.TOKENS.delete(userToken);
    return null;
  }

  // If token is expired or expires in the next 5 minutes, refresh it
  // Using 300,000ms (5 mins) buffer for serverless reliability
  if (tokens.expiry_date && Date.now() > tokens.expiry_date - 300000) {
    if (tokens.refresh_token) {
      const response = await fetch("https://oauth2.googleapis.com/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          refresh_token: tokens.refresh_token,
          client_id: env.GOOGLE_CLIENT_ID,
          client_secret: env.GOOGLE_CLIENT_SECRET,
          grant_type: "refresh_token",
        }),
      });

      const newTokens = await response.json() as OAuthTokenResponse;
      if (newTokens.error || !newTokens.access_token) {
        return null;
      }
      tokens.access_token = newTokens.access_token;
      if (newTokens.expires_in) {
        tokens.expiry_date = Date.now() + (newTokens.expires_in * 1000);
      }
      await env.TOKENS.put(userToken, await encryptData(JSON.stringify(tokens), env.TOKEN_ENCRYPTION_KEY));
    }
  }

  return tokens.access_token ?? null;
}
