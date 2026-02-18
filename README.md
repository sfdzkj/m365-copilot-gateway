# M365 Copilot → OpenAI Compatible Gateway (Multi-user)

## 你要的：日志可观测性（Docker logs 里能看到状态/错误）

本版本（v1.0.4）新增：
- 结构化日志（JSON，默认输出到 stdout，适配 `docker logs`）
- 每个请求自动生成/透传 `X-Request-Id`，日志里全链路带 `requestId`
- 统一错误处理中间件 + 进程级别异常捕获（uncaughtException/unhandledRejection）
- 上游（Graph Copilot API）调用耗时、HTTP 状态、Graph `request-id`/`client-request-id` 记录
- SSE 流式请求：记录开始/结束、收到的事件块数量、解析失败次数
- `/healthz` 增强：包含 redis 连通性与关键配置是否就绪（不泄露敏感信息）

## 运行

```bash
cp .env.example .env
# 填写 TENANT_ID/CLIENT_ID/CLIENT_SECRET/PUBLIC_BASE_URL/API_BEARER_TOKEN/SESSION_SECRET

docker compose up -d --build
```

## 日志环境变量

- `LOG_LEVEL=info|debug|warn|error`（默认 info）
- `LOG_REQUESTS=1` 开启请求日志（默认 1）
- `LOG_BODY=0/1` 是否记录请求体摘要（默认 0，建议生产 0）

## 健康检查

```bash
curl http://localhost:8080/healthz
```
