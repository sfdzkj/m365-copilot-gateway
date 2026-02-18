# M365 Copilot → OpenAI Compatible Gateway (v1.0.6)

将 **Microsoft 365 Copilot Chat API（Microsoft Graph /beta）** 封装为 **OpenAI 标准接口**，提供统一的 `/v1/chat/completions` 调用入口（支持 `stream=true` SSE），并支持 **多授权账户（多 `X-User-Key`）** 的新增、区分（label）、删除与轮换。

> 适用场景：你已有大量系统/工具基于 OpenAI 标准接口调用大模型，希望无缝接入 Microsoft 365 Copilot 能力，并通过网关统一鉴权、运维、审计与多账号切换。

---

## 1. 核心特性

### 1.1 OpenAI 兼容接口

- `GET /v1/models`
- `POST /v1/chat/completions`
  - `stream=false`：一次性返回结果
  - `stream=true`：SSE 流式输出（OpenAI `chat.completion.chunk` 结构）

### 1.2 多授权账户（调用方维护多个 `X-User-Key`）

> Copilot Chat API 采用 **Delegated（委托）** 模式：每个授权用户必须本人完成一次登录授权后，网关才能以该用户身份调用 Copilot Chat API；**不支持 Application（应用）权限**。

- 设备码授权获取 `X-User-Key`：
  - `POST /auth/device/start`（可携带 `label`）
  - `GET /auth/device/status/:txId`
- Key 管理（管理员接口）：
  - `GET /auth/keys`（列出所有 key，含 label、创建时间、脱敏显示）
  - `GET /auth/keys/:userKey`（查看 key 元数据）
  - `POST /auth/keys/label`（修改 label）
  - `DELETE /auth/keys/:userKey`（删除/吊销 key）
  - `POST /auth/keys/:userKey/rotate`（轮换 key：新 key 生效、旧 key 失效）

### 1.3 可观测性（适配 Docker logs）

- JSON 结构化日志输出到 stdout/stderr
- 自动生成/透传 `X-Request-Id`，全链路日志携带 `requestId`
- Graph 上游调用：记录 URL、状态码、耗时、Graph request-id
- SSE 流式：记录 stream start/end、事件块数量、JSON 解析失败次数、输出字符数
- `GET /healthz`：健康检查（Redis ping + 关键配置就绪）
- `GET /debug/last-events`：查看最近上游 SSE 事件片段（管理员接口）

---

## 2. 前置条件与限制

1) **Graph /beta**：Copilot Chat API 位于 Microsoft Graph `/beta`，接口字段与行为可能变化，生产使用请做好监控与兼容。

2) **Delegated-only**：Copilot Chat API 仅支持 Delegated（委托）权限，不支持 Application（应用）权限，因此必须由用户完成授权。

3) **许可（license）**：若调用时返回 “It looks like you don’t have a valid license ...”，通常表示该用户缺少可用许可（与网页版 Copilot Chat 的可用性可能不同）。

---

## 3. 安装与部署（Docker Compose）

> 推荐使用 Docker Compose，一键启动网关与 Redis。

### 3.1 获取项目并准备 `.env`

```bash
# 解压后进入目录
cd m365-copilot-openai-gateway

# 复制环境变量模板
cp .env.example .env
```

编辑 `.env` 至少填写：

- `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET`
- `PUBLIC_BASE_URL`（**必须与 Entra 应用 Redirect URI 完全一致**）
- `API_BEARER_TOKEN`（访问 `/v1/*` 的网关鉴权 token）
- `SESSION_SECRET`

可选但推荐：

- `ADMIN_BEARER_TOKEN`（管理接口独立 token，避免误操作）

### 3.2 启动

```bash
docker compose up -d --build
```

查看日志：

```bash
docker compose logs -f copilot-gw
```

健康检查：

```bash
curl http://localhost:8080/healthz
```

---

## 4. Entra（Azure AD）应用注册与权限配置

### 4.1 注册应用

在 Entra 管理中心注册应用（推荐）：

- 平台类型：**Web**
- Redirect URI：`{PUBLIC_BASE_URL}/auth/callback`

示例：

- 本机/内网：`http://localhost:8080/auth/callback`
- 公网 HTTPS：`https://your-domain.com/auth/callback`

> Redirect URI 必须与 `.env` 的 `PUBLIC_BASE_URL` 严格匹配（协议、域名/IP、端口、路径全部一致）。

### 4.2 启用设备码（可选但推荐）

在 **Authentication** 中开启：

- Allow public client flows = Yes

这会让 device code 授权更稳定。

### 4.3 配置 Graph Delegated 权限并管理员同意

为 Microsoft Graph 添加 Delegated 权限（并执行 Grant admin consent）：

- `Sites.Read.All`
- `Mail.Read`
- `People.Read.All`
- `OnlineMeetingTranscript.Read.All`
- `Chat.Read`
- `ChannelMessage.Read.All`
- `ExternalItem.Read.All`

---

## 5. 环境变量说明

`.env.example` 已包含完整说明，常用如下：

- **Entra**
  - `TENANT_ID`
  - `CLIENT_ID`
  - `CLIENT_SECRET`
  - `PUBLIC_BASE_URL`

- **网关鉴权**
  - `API_BEARER_TOKEN`：访问 `/v1/*` 必须携带 `Authorization: Bearer <token>`
  - `ADMIN_BEARER_TOKEN`：可选，管理接口（`/auth/keys*`、`/debug/*`）使用不同 token

- **会话与 Redis**
  - `SESSION_SECRET`
  - `SESSION_TTL`
  - `REDIS_URL`

- **运行与日志**
  - `PORT`
  - `LOG_LEVEL=debug|info|warn|error`
  - `LOG_REQUESTS=1|0`
  - `LOG_BODY=1|0`（生产建议 0）

---

## 6. 接口一览

### 6.1 OpenAI 兼容

- `GET /v1/models`
- `POST /v1/chat/completions`

### 6.2 授权（Device Code）

- `POST /auth/device/start`
  - body: `{ "label": "Alice-财务" }`（可选）
- `GET /auth/device/status/:txId`

### 6.3 Key 管理（管理员）

- `GET /auth/keys`
- `GET /auth/keys/:userKey`
- `POST /auth/keys/label`
- `DELETE /auth/keys/:userKey`
- `POST /auth/keys/:userKey/rotate`

### 6.4 运维与调试

- `GET /healthz`
- `GET /debug/last-events`（管理员）

---

## 7. 使用示例

### 7.1 新增授权用户（带 label）

```bash
curl -s http://localhost:8080/auth/device/start   -H 'Content-Type: application/json'   -d '{"label":"Alice-财务"}'
```

按返回 `message` 在浏览器完成登录后轮询：

```bash
curl -s http://localhost:8080/auth/device/status/<txId>
```

当 `status=complete`，返回 `user_key`。

### 7.2 列出所有已授权 key（管理员）

```bash
curl -s http://localhost:8080/auth/keys   -H 'Authorization: Bearer <ADMIN或API_TOKEN>'
```

### 7.3 调用 OpenAI 接口（非流式）

```bash
curl http://localhost:8080/v1/chat/completions   -H 'Authorization: Bearer <API_BEARER_TOKEN>'   -H 'X-User-Key: <user_key>'   -H 'Content-Type: application/json'   -d '{
    "model": "deep",
    "stream": false,
    "messages": [{"role":"user","content":"写一个项目复盘模板"}]
  }'
```

### 7.4 调用 OpenAI 接口（流式 SSE）

```bash
curl http://localhost:8080/v1/chat/completions   -H 'Authorization: Bearer <API_BEARER_TOKEN>'   -H 'X-User-Key: <user_key>'   -H 'Content-Type: application/json'   -d '{
    "model": "fast",
    "stream": true,
    "messages": [{"role":"user","content":"用三点总结今天的待办"}]
  }'
```

---

## 8. 多 Key 治理（删除/轮换）

### 8.1 删除（吊销）key

```bash
curl -s -X DELETE http://localhost:8080/auth/keys/<user_key>   -H 'Authorization: Bearer <ADMIN或API_TOKEN>'
```

### 8.2 轮换（rotate）key

```bash
curl -s -X POST http://localhost:8080/auth/keys/<user_key>/rotate   -H 'Authorization: Bearer <ADMIN或API_TOKEN>'
```

---

## 9. 日志与排障

### 9.1 查看 docker 日志

```bash
docker compose logs -f copilot-gw
```

### 9.2 查看最近上游事件（管理员）

```bash
curl -s http://localhost:8080/debug/last-events   -H 'Authorization: Bearer <ADMIN或API_TOKEN>'
```

### 9.3 常见错误

- `Invalid gateway token`：检查 `API_BEARER_TOKEN` 与请求头 `Authorization: Bearer` 是否一致。
- `Invalid X-User-Key or expired`：该 key 已删除/过期，请重新授权获取新 key。
- `no valid license`：用户缺少 Copilot Chat API 所需许可，请更换有许可的用户重新授权。

---

## 10. 安全建议

- `CLIENT_SECRET` 属于敏感信息：泄露后请立刻吊销并重新生成。
- 强烈建议将 `ADMIN_BEARER_TOKEN` 与 `API_BEARER_TOKEN` 分离。
- `X-User-Key` 等同“用户会话密钥”：建议定期 rotate，泄露立即 delete + rotate。
