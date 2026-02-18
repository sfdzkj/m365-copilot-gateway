import express from "express";
import session from "express-session";
import crypto from "crypto";
import Redis from "ioredis";
import { ConfidentialClientApplication } from "@azure/msal-node";

// =======================
// Logging (stdout JSON)
// =======================
const LOG_LEVEL = (process.env.LOG_LEVEL || "info").toLowerCase();
const LOG_REQUESTS = process.env.LOG_REQUESTS !== "0";
const LOG_BODY = process.env.LOG_BODY === "1";

const levelRank = { debug: 10, info: 20, warn: 30, error: 40 };
const minRank = levelRank[LOG_LEVEL] ?? 20;

function safeString(v, maxLen = 400) {
  if (v == null) return v;
  const s = typeof v === "string" ? v : JSON.stringify(v);
  return s.length > maxLen ? s.slice(0, maxLen) + "…" : s;
}

function redactHeaders(headers = {}) {
  const h = { ...headers };
  for (const k of Object.keys(h)) {
    const lk = k.toLowerCase();
    if (lk === "authorization" || lk === "cookie" || lk === "set-cookie") h[k] = "[REDACTED]";
  }
  return h;
}

function log(level, msg, fields = {}) {
  const r = levelRank[level] ?? 20;
  if (r < minRank) return;
  const line = {
    ts: new Date().toISOString(),
    level,
    msg,
    ...fields,
  };
  // error -> stderr, others -> stdout
  const out = level === "error" ? process.stderr : process.stdout;
  out.write(JSON.stringify(line) + "\n");
}

// process-level safety nets
process.on("uncaughtException", (err) => {
  log("error", "uncaughtException", { err: { name: err.name, message: err.message, stack: err.stack } });
  // let container restart
  process.exit(1);
});
process.on("unhandledRejection", (reason) => {
  log("error", "unhandledRejection", { reason: safeString(reason), stack: reason?.stack });
});

// =======================
// Env & Config
// =======================
const env = process.env;
const TENANT_ID = env.TENANT_ID;
const CLIENT_ID = env.CLIENT_ID;
const CLIENT_SECRET = env.CLIENT_SECRET;
const PUBLIC_BASE_URL = (env.PUBLIC_BASE_URL || "http://localhost:8080").replace(/\/$/, "");
const API_BEARER_TOKEN = env.API_BEARER_TOKEN || "";
const SESSION_SECRET = env.SESSION_SECRET;
const REDIS_URL = env.REDIS_URL || "redis://redis:6379";
const TIME_ZONE = env.TIME_ZONE || "Asia/Shanghai";
const COUNTRY_OR_REGION = env.COUNTRY_OR_REGION || "US";
const TRUST_PROXY = env.TRUST_PROXY === "1";
const PORT = parseInt(env.PORT || "8080", 10);
const SESSION_TTL = parseInt(env.SESSION_TTL || "86400", 10);
const EMPTY_RESPONSE_HINT = env.EMPTY_RESPONSE_HINT !== "0";
const DEBUG_EVENT_LIMIT = parseInt(env.DEBUG_EVENT_LIMIT || "50", 10);

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  throw new Error("Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET");
}
if (!SESSION_SECRET) {
  throw new Error("Missing SESSION_SECRET");
}

const REDIRECT_URI = `${PUBLIC_BASE_URL}/auth/callback`;

log("info", "config", {
  port: PORT,
  publicBaseUrl: PUBLIC_BASE_URL,
  redirectUri: REDIRECT_URI,
  trustProxy: TRUST_PROXY,
  logLevel: LOG_LEVEL,
  logRequests: LOG_REQUESTS,
});

// =======================
// Redis
// =======================
const redis = new Redis(REDIS_URL);
redis.on("connect", () => log("info", "redis.connect"));
redis.on("error", (e) => log("error", "redis.error", { err: safeString(e?.message || e) }));

class RedisSessionStore extends session.Store {
  constructor(client, opts = {}) {
    super();
    this.client = client;
    this.prefix = opts.prefix || "sess:";
    this.ttl = opts.ttl || 86400;
  }
  _key(sid) { return `${this.prefix}${sid}`; }
  async get(sid, cb) {
    try {
      const data = await this.client.get(this._key(sid));
      cb(null, data ? JSON.parse(data) : null);
    } catch (e) { cb(e); }
  }
  async set(sid, sess, cb) {
    try {
      const ttl = this._getTTL(sess);
      await this.client.set(this._key(sid), JSON.stringify(sess), "EX", ttl);
      cb(null);
    } catch (e) { cb(e); }
  }
  async destroy(sid, cb) {
    try { await this.client.del(this._key(sid)); cb(null); }
    catch (e) { cb(e); }
  }
  async touch(sid, sess, cb) {
    try {
      const ttl = this._getTTL(sess);
      await this.client.expire(this._key(sid), ttl);
      cb(null);
    } catch (e) { cb(e); }
  }
  _getTTL(sess) {
    const cookie = sess?.cookie;
    if (cookie?.expires) {
      const exp = new Date(cookie.expires).getTime();
      const ttl = Math.ceil((exp - Date.now()) / 1000);
      if (ttl > 0) return ttl;
    }
    if (typeof cookie?.maxAge === "number") {
      const ttl = Math.ceil(cookie.maxAge / 1000);
      if (ttl > 0) return ttl;
    }
    return this.ttl;
  }
}

async function pushDebugEvent(obj) {
  try {
    const key = "debug:last_events";
    await redis.lpush(key, JSON.stringify(obj));
    await redis.ltrim(key, 0, Math.max(0, DEBUG_EVENT_LIMIT - 1));
    await redis.expire(key, 3600);
  } catch {}
}

async function redisGetJson(key) {
  const v = await redis.get(key);
  return v ? JSON.parse(v) : null;
}
async function redisSetJson(key, obj, ttlSeconds = null) {
  const v = JSON.stringify(obj);
  if (ttlSeconds) await redis.set(key, v, "EX", ttlSeconds);
  else await redis.set(key, v);
}

// =======================
// MSAL
// =======================
const msal = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  },
});

const SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "Sites.Read.All",
  "Mail.Read",
  "People.Read.All",
  "OnlineMeetingTranscript.Read.All",
  "Chat.Read",
  "ChannelMessage.Read.All",
  "ExternalItem.Read.All",
];

async function loadMsalCache(homeAccountId) { return await redis.get(`msal:${homeAccountId}`); }
async function saveMsalCache(homeAccountId, cache, ttlSeconds = 60 * 60 * 24 * 7) {
  await redis.set(`msal:${homeAccountId}`, cache, "EX", ttlSeconds);
}
async function getAccountByUserKey(userKey) {
  if (!userKey) return null;
  const home = await redis.get(`userkey:${userKey}`);
  if (!home) return null;
  return await redisGetJson(`account:${home}`);
}

async function acquireAccessToken({ account, requestId }) {
  const cached = await loadMsalCache(account.homeAccountId);
  if (cached) msal.getTokenCache().deserialize(cached);
  const t0 = Date.now();
  const result = await msal.acquireTokenSilent({ account, scopes: SCOPES });
  await saveMsalCache(account.homeAccountId, msal.getTokenCache().serialize());
  log("debug", "msal.acquireTokenSilent", { requestId, ms: Date.now() - t0 });
  return result.accessToken;
}

// =======================
// Express + RequestId
// =======================
const app = express();
if (TRUST_PROXY) app.set("trust proxy", 1);

app.use(express.json({ limit: "2mb" }));

// request id middleware
app.use((req, res, next) => {
  const incoming = req.headers["x-request-id"];
  const requestId = (typeof incoming === "string" && incoming.length <= 200) ? incoming : crypto.randomUUID();
  req.requestId = requestId;
  res.setHeader("x-request-id", requestId);
  next();
});

// request logging middleware
if (LOG_REQUESTS) {
  app.use((req, res, next) => {
    const start = Date.now();
    const reqMeta = {
      requestId: req.requestId,
      method: req.method,
      path: req.originalUrl,
      ip: req.ip,
      ua: req.headers["user-agent"],
    };
    if (LOG_BODY && req.body) {
      reqMeta.body = safeString(req.body, 800);
    }
    log("info", "http.request", reqMeta);

    res.on("finish", () => {
      log("info", "http.response", {
        requestId: req.requestId,
        method: req.method,
        path: req.originalUrl,
        status: res.statusCode,
        ms: Date.now() - start,
      });
    });
    next();
  });
}

app.use(
  session({
    store: new RedisSessionStore(redis, { prefix: "sess:", ttl: SESSION_TTL }),
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: PUBLIC_BASE_URL.startsWith("https://"),
      maxAge: SESSION_TTL * 1000,
    },
  })
);

// =======================
// Utility
// =======================
function nowUnix() { return Math.floor(Date.now() / 1000); }

function requireGatewayToken(req, res) {
  if (!API_BEARER_TOKEN) return true;
  const h = req.headers.authorization || "";
  if (h === `Bearer ${API_BEARER_TOKEN}`) return true;
  log("warn", "auth.invalidGatewayToken", { requestId: req.requestId, auth: "[REDACTED]" });
  res.status(401).json({ error: { message: "Invalid gateway token" } });
  return false;
}

function normalizeMode(model) {
  const m = (model || "auto").toLowerCase();
  if (m === "gpt-5.2-fast") return "fast";
  if (m === "gpt-5.2-deep") return "deep";
  if (["auto", "fast", "deep"].includes(m)) return m;
  return "auto";
}

function buildPrompt(messages, mode) {
  const hints = {
    auto: "你是企业办公助手。根据问题复杂度自动选择简洁或深入的回答方式。",
    fast: "请快速、简洁回答，优先给结论和要点，避免长篇铺垫。",
    deep: "请深入分析，分步骤给出思考与可执行建议，必要时列出风险与注意事项。",
  };
  const hint = hints[mode] || hints.auto;
  const lines = [];
  lines.push(`SYSTEM: ${hint}`);
  for (const m of messages || []) {
    const role = (m.role || "user").toUpperCase();
    const content = typeof m.content === "string" ? m.content : JSON.stringify(m.content);
    lines.push(`${role}: ${content}`);
  }
  lines.push("ASSISTANT:");
  return lines.join("\n");
}

function extractTextFromMessage(msg) {
  if (!msg || typeof msg !== "object") return null;
  if (typeof msg.text === "string") return msg.text;
  if (typeof msg.content === "string") return msg.content;
  if (msg.message && typeof msg.message.text === "string") return msg.message.text;
  return null;
}

function extractBestTextFromEvent(obj, promptEcho) {
  const msgs = obj?.messages;
  if (Array.isArray(msgs) && msgs.length) {
    for (let i = msgs.length - 1; i >= 0; i--) {
      const t = extractTextFromMessage(msgs[i]);
      if (!t) continue;
      if (promptEcho && t.trim() === promptEcho.trim()) continue;
      return t;
    }
    for (let i = msgs.length - 1; i >= 0; i--) {
      const t = extractTextFromMessage(msgs[i]);
      if (t) return t;
    }
  }
  const wrapped = obj?.value?.messages;
  if (Array.isArray(wrapped) && wrapped.length) {
    for (let i = wrapped.length - 1; i >= 0; i--) {
      const t = extractTextFromMessage(wrapped[i]);
      if (t) return t;
    }
  }
  return null;
}

// =======================
// Graph upstream wrapper
// =======================
async function fetchGraph(requestId, url, options) {
  const t0 = Date.now();
  const safeOpts = {
    method: options?.method,
    headers: redactHeaders(options?.headers || {}),
  };
  log("debug", "graph.request", { requestId, url, options: safeOpts });

  const res = await fetch(url, options);
  const ms = Date.now() - t0;
  const graphReqId = res.headers.get("request-id") || res.headers.get("client-request-id");
  log("info", "graph.response", { requestId, url, status: res.status, ms, graphReqId });
  return res;
}

async function createCopilotConversation(accessToken, requestId) {
  const res = await fetchGraph(
    requestId,
    "https://graph.microsoft.com/beta/copilot/conversations",
    {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({}),
    }
  );
  if (!res.ok) {
    const txt = await res.text();
    log("error", "graph.createConversation.failed", { requestId, status: res.status, body: safeString(txt, 1200) });
    throw new Error(`Create conversation failed: ${res.status} ${txt}`);
  }
  return await res.json();
}

async function copilotChat(accessToken, conversationId, prompt, requestId) {
  const res = await fetchGraph(
    requestId,
    `https://graph.microsoft.com/beta/copilot/conversations/${conversationId}/chat`,
    {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: { text: prompt },
        locationHint: { timeZone: TIME_ZONE, countryOrRegion: COUNTRY_OR_REGION },
      }),
    }
  );
  if (!res.ok) {
    const txt = await res.text();
    log("error", "graph.chat.failed", { requestId, conversationId, status: res.status, body: safeString(txt, 1800) });
    throw new Error(`Chat failed: ${res.status} ${txt}`);
  }
  return await res.json();
}

async function copilotChatOverStream(accessToken, conversationId, prompt, requestId) {
  const res = await fetchGraph(
    requestId,
    `https://graph.microsoft.com/beta/copilot/conversations/${conversationId}/chatOverStream`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        Accept: "text/event-stream",
      },
      body: JSON.stringify({
        message: { text: prompt },
        locationHint: { timeZone: TIME_ZONE, countryOrRegion: COUNTRY_OR_REGION },
      }),
    }
  );
  if (!res.ok) {
    const txt = await res.text();
    log("error", "graph.chatOverStream.failed", { requestId, conversationId, status: res.status, body: safeString(txt, 1800) });
    throw new Error(`ChatOverStream failed: ${res.status} ${txt}`);
  }
  return res;
}

// =======================
// Routes
// =======================
app.get("/v1/models", (req, res) => {
  if (!requireGatewayToken(req, res)) return;
  res.json({
    object: "list",
    data: [
      { id: "auto", object: "model", created: 0, owned_by: "gateway" },
      { id: "fast", object: "model", created: 0, owned_by: "gateway" },
      { id: "deep", object: "model", created: 0, owned_by: "gateway" },
      { id: "gpt-5.2-fast", object: "model", created: 0, owned_by: "gateway" },
      { id: "gpt-5.2-deep", object: "model", created: 0, owned_by: "gateway" },
    ],
  });
});

function getUserContext(req) {
  const userKey = req.headers["x-user-key"];
  if (typeof userKey === "string" && userKey.length > 10) return { type: "userkey", userKey };
  if (req.session?.account?.homeAccountId) return { type: "session", account: req.session.account };
  return null;
}

// Device code start/status (kept from earlier versions)
app.post("/auth/device/start", async (req, res) => {
  const requestId = req.requestId;
  const txId = crypto.randomUUID();
  const createdAt = Date.now();
  let infoResolve;
  const infoPromise = new Promise((resolve) => (infoResolve = resolve));

  (async () => {
    try {
      const result = await msal.acquireTokenByDeviceCode({
        scopes: SCOPES,
        deviceCodeCallback: async (response) => {
          const deviceInfo = {
            txId,
            status: "pending",
            createdAt,
            user_code: response.userCode,
            verification_uri: response.verificationUri,
            message: response.message,
            expires_in: response.expiresIn,
            interval: response.interval,
          };
          await redisSetJson(`devtx:${txId}`, deviceInfo, response.expiresIn);
          infoResolve(deviceInfo);
          log("info", "auth.device.pending", { requestId, txId, expiresIn: response.expiresIn });
        },
      });

      const account = result.account;
      if (!account?.homeAccountId) throw new Error("No account in device code result");

      await saveMsalCache(account.homeAccountId, msal.getTokenCache().serialize());
      await redisSetJson(`account:${account.homeAccountId}`, account);

      const userKey = base64url(crypto.randomBytes(24));
      await redis.set(`userkey:${userKey}`, account.homeAccountId, "EX", 60 * 60 * 24 * 30);

      const done = (await redisGetJson(`devtx:${txId}`)) || { txId };
      done.status = "complete";
      done.user_key = userKey;
      done.completedAt = Date.now();
      await redisSetJson(`devtx:${txId}`, done, 60 * 60 * 24);

      log("info", "auth.device.complete", { requestId, txId, homeAccountId: account.homeAccountId });
    } catch (e) {
      const cur = (await redisGetJson(`devtx:${txId}`)) || { txId };
      cur.status = "error";
      cur.error = String(e?.message || e);
      await redisSetJson(`devtx:${txId}`, cur, 60 * 60);
      log("error", "auth.device.error", { requestId, txId, err: safeString(e?.message || e) });
    }
  })();

  const info = await infoPromise;
  res.json(info);
});

app.get("/auth/device/status/:txId", async (req, res) => {
  const info = await redisGetJson(`devtx:${req.params.txId}`);
  if (!info) return res.status(404).json({ error: { message: "txId not found" } });
  res.json(info);
});

// health endpoint (redis ping + config readiness)
app.get("/healthz", async (req, res) => {
  const requestId = req.requestId;
  let redisOk = false;
  try {
    const pong = await redis.ping();
    redisOk = pong === "PONG";
  } catch (e) {
    log("warn", "health.redis.ping.failed", { requestId, err: safeString(e?.message || e) });
  }
  res.json({
    ok: true,
    time: new Date().toISOString(),
    redis: redisOk,
    config: {
      hasTenantId: !!TENANT_ID,
      hasClientId: !!CLIENT_ID,
      hasClientSecret: !!CLIENT_SECRET,
      hasSessionSecret: !!SESSION_SECRET,
      hasGatewayToken: !!API_BEARER_TOKEN,
      redirectUri: REDIRECT_URI,
    },
  });
});

// Debug: last events
app.get("/debug/last-events", async (req, res) => {
  const key = "debug:last_events";
  const arr = await redis.lrange(key, 0, Math.max(0, DEBUG_EVENT_LIMIT - 1));
  const parsed = arr.map((s) => { try { return JSON.parse(s); } catch { return { raw: s }; } });
  res.json({ count: parsed.length, events: parsed });
});

// main OpenAI endpoint
app.post("/v1/chat/completions", async (req, res, next) => {
  try {
    if (!requireGatewayToken(req, res)) return;

    const requestId = req.requestId;
    const { model = "auto", messages = [], stream = false } = req.body || {};
    const mode = normalizeMode(model);

    const ctx = getUserContext(req);
    let account = null;

    if (!ctx) {
      log("warn", "auth.noUserContext", { requestId });
      return res.status(401).json({ error: { message: "No user context. Use X-User-Key." } });
    }

    if (ctx.type === "session") account = ctx.account;
    else account = await getAccountByUserKey(ctx.userKey);

    if (!account) {
      log("warn", "auth.invalidUserKey", { requestId });
      return res.status(401).json({ error: { message: "Invalid X-User-Key or expired." } });
    }

    const accessToken = await acquireAccessToken({ account, requestId });
    const prompt = buildPrompt(messages, mode);

    const conversation = await createCopilotConversation(accessToken, requestId);
    const conversationId = conversation.id;

    if (!stream) {
      const data = await copilotChat(accessToken, conversationId, prompt, requestId);
      const text = extractBestTextFromEvent(data, prompt) || "";
      return res.json({
        id: `chatcmpl_${crypto.randomUUID()}`,
        object: "chat.completion",
        created: nowUnix(),
        model,
        choices: [{ index: 0, message: { role: "assistant", content: text }, finish_reason: "stop" }],
      });
    }

    // stream mode
    const upstream = await copilotChatOverStream(accessToken, conversationId, prompt, requestId);

    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache",
      Connection: "keep-alive",
    });

    const reader = upstream.body.getReader();
    const dec = new TextDecoder();
    let carry = "";
    let fullText = "";
    let blocks = 0;
    let jsonFail = 0;

    const sendDelta = (delta) => {
      const chunk = {
        id: `chatcmpl_${conversationId}`,
        object: "chat.completion.chunk",
        created: nowUnix(),
        model,
        choices: [{ index: 0, delta: { content: delta }, finish_reason: null }],
      };
      res.write(`data: ${JSON.stringify(chunk)}\n\n`);
    };

    const sendStop = () => {
      res.write(
        `data: ${JSON.stringify({
          id: `chatcmpl_${conversationId}`,
          object: "chat.completion.chunk",
          created: nowUnix(),
          model,
          choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
        })}\n\n`
      );
      res.write("data: [DONE]\n\n");
      res.end();
    };

    log("info", "stream.start", { requestId, conversationId });

    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;

        carry += dec.decode(value, { stream: true });

        // parse SSE blocks separated by blank line
        while (true) {
          const idx = carry.indexOf("\n\n");
          if (idx == -1) break;
          const block = carry.slice(0, idx);
          carry = carry.slice(idx + 2);
          blocks++;

          const lines = block.split("\n");
          const dataLines = lines
            .map((l) => l.trimEnd())
            .filter((l) => l.startsWith("data:"))
            .map((l) => l.slice(5).trimStart());

          if (!dataLines.length) continue;
          const payload = dataLines.join("\n").trim();
          if (!payload) continue;

          let obj;
          try {
            obj = JSON.parse(payload);
          } catch {
            jsonFail++;
            continue;
          }

          await pushDebugEvent({ ts: Date.now(), requestId, conversationId, sample: payload.slice(0, 400) });

          if (obj?.error) {
            log("error", "graph.stream.errorObject", { requestId, conversationId, error: obj.error });
          }

          const text = extractBestTextFromEvent(obj, prompt);
          if (typeof text === "string" && text.length > fullText.length) {
            const delta = text.slice(fullText.length);
            fullText = text;
            if (delta) sendDelta(delta);
          }
        }
      }
    } catch (e) {
      log("error", "stream.exception", { requestId, conversationId, err: safeString(e?.message || e) });
    }

    if (!fullText && EMPTY_RESPONSE_HINT) {
      sendDelta("\n（提示：上游未返回可解析的文本内容。请检查 Copilot 许可/权限，或访问 /debug/last-events 查看原始事件片段。）\n");
    }

    log("info", "stream.end", { requestId, conversationId, blocks, jsonFail, chars: fullText.length });
    sendStop();
  } catch (err) {
    next(err);
  }
});

// =======================
// Global error handler
// =======================
app.use((err, req, res, next) => {
  const requestId = req?.requestId;
  log("error", "http.error", {
    requestId,
    err: {
      name: err?.name,
      message: err?.message,
      stack: err?.stack,
    },
  });
  if (res.headersSent) return next(err);
  res.status(500).json({ error: { message: err?.message || "Internal error", requestId } });
});

app.listen(PORT, "0.0.0.0", () => {
  log("info", "server.listen", { port: PORT });
});

// ===== helpers =====
function base64url(buf) {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
