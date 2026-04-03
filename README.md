# E小助

本项目会把固定的企业流程 Excel 和附件目录转成本地知识库，并提供一个网页聊天界面做检索式问答。

## 功能

- 服务启动时自动同步固定知识源 workbook 和附件目录。
- 自动提取内嵌的 `docx / pptx / xlsx` 附件文本。
- 基于本地持久化的 chunks 和 embeddings 做 RAG 检索。
- 每次回答都返回来源引用，包含流程编号、标题、附件名或外链。
- 自动统计每个知识库版本下“未能回答”的问题数量，并保留最近未回答问题列表。
- 网页端可展开查看每次真正发送给模型的请求体（不含 API key）。
- 固定知识源变更后会生成新的知识库版本；旧会话继续绑定旧版本。

## 运行

1. 安装依赖

```bash
npm install
```

2. 选择 provider

复制配置模板：

```bash
cp .env.example .env
```

服务端现在会自动读取项目根目录下的 `.env`。

如果你只是内部小范围使用，建议直接设置一个访问密码，并固定知识源路径：

```env
APP_PASSWORD=请改成你自己的密码
KNOWLEDGE_SOURCE_WORKBOOK_PATH=/home/kleist/Downloads/Data.xlsx
KNOWLEDGE_SOURCE_ATTACHMENTS_DIR=/home/kleist/Downloads/流程教程知识库规范包_20260402(1)/attachments
```

设置后，E小助 会在启动时自动同步固定知识源。网页会先显示一个密码页；只有输入正确密码后，才能继续访问知识库和发起聊天。认证通过后，服务端会写入一个 `HttpOnly` cookie。

推荐的本地 `vLLM + Qwen3.5-35B-A3B-FP8` 组合：

```env
CHATBOT_PROVIDER=vllm
EMBEDDING_PROVIDER=offline
VLLM_BASE_URL=http://localhost:8000/v1
VLLM_CHAT_MODEL=qwen35
VLLM_ENABLE_THINKING=false
VLLM_STRIP_THINK_OUTPUT=true
CHAT_MAX_TOKENS=300
```

这里的 `VLLM_CHAT_MODEL` 建议直接填你启动 vLLM 时暴露出来的 `model` 名称，例如你用了 `--served-model-name qwen35`，那就填 `qwen35`。这个组合只要求你启动一个 vLLM 聊天模型服务，检索向量使用本项目内置的离线哈希 embedding。

- `VLLM_ENABLE_THINKING=false`: 默认关闭 Qwen3.5 的 thinking，减少延迟并避免把思考过程展示给用户。
- `VLLM_STRIP_THINK_OUTPUT=true`: 如果模型仍返回 `<think>` 或 `Thinking Process`，服务端会自动清洗掉。
- `CHAT_MAX_TOKENS=300`: 限制单次回答长度，减少冗长输出。

如果你要继续使用 OpenAI-compatible 云服务：

- `CHATBOT_PROVIDER=openai`
- `EMBEDDING_PROVIDER=openai`
- `OPENAI_API_KEY`
- `OPENAI_CHAT_MODEL`
- `OPENAI_EMBEDDING_MODEL`

如果只想本地验证流程，可切换到离线模式：

```bash
CHATBOT_PROVIDER=offline npm run dev
```

3. 启动开发环境

```bash
npm run dev
```

- 前端: `http://localhost:5173`
- 后端: `http://localhost:3001`

如果修改了 `APP_PASSWORD` 或其他 `.env` 配置，需要重启服务。

## 规范化知识库包

推荐的后续维护格式是一个 `.zip` 包，根目录包含：

- `knowledge.xlsx`
- `attachments/`

其中 `knowledge.xlsx` 需要至少包含 3 个 sheet：

- `版本说明`
- `流程汇总`
- `附件清单`

`流程汇总` 的固定列为：

- `编号`
- `一级分类`
- `条目类型`
- `标题`
- `相关单据`
- `联系人/责任人`
- `正文`
- `外部链接`
- `关键词/别名`

`附件清单` 的固定列为：

- `条目编号`
- `附件文件名`
- `相对路径`
- `附件类型`
- `附件说明`

当前版本不再支持网页上传知识库。更新方式是替换固定 workbook 或附件目录，然后重启服务。

## 构建

```bash
npm run build
npm start
```

## 数据目录

运行后会在 `data/` 下生成：

- `knowledge-bases/`: 知识库 JSON
- `jobs/`: 导入任务状态
- `sessions/`: 聊天会话
- `state.json`: 当前活动知识库

## 已知限制

- 旧版 `.doc` 内嵌附件当前只记录元数据，不做正文解析。
- 规范化迁移脚本会把旧版 `.doc` 降级为 `.docx` 摘要附件，但不会自动还原原始排版。
- 外部链接只保存 URL，不抓取网页正文。
- 当 `CHATBOT_PROVIDER=vllm` 且 `EMBEDDING_PROVIDER=offline` 时，回答生成走 vLLM，本地检索向量不依赖额外 embedding 模型。
- 如果修改了 `.env`、固定 workbook 或附件目录，需要重启服务才会重新同步知识库。
