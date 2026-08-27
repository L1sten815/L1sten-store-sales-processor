# 项目长期记忆 — FBA需求分配核查工具

## 技术栈与部署
- 单文件后端 `server.py`（基于 `http.server.ThreadingHTTPServer` 的微型 HTTP 服务），单文件前端 `store-sales-processor.html`。
- 部署在 Render，绑定 GitHub `main` 分支自动部署；免费实例 512MB 内存，需警惕大 Excel 处理 OOM。
- Python 依赖：`openpyxl`、`python-calamine`；优先用 calamine 读大表，openpyxl 写结果。

## 关键约定
- 前后端文件槽位名必须保持一致。当前约定：
  - 页签①/②：`sales`、`ops`、`fba`、`mapping`。
  - 页签③ 总可用库存核对：`fba`、`stock`。
- 下载文件名若含中文，服务端构造 `Content-Disposition` 时必须使用 ASCII-safe 的 `filename` fallback，并把 UTF-8 名称做 RFC 5987 百分号编码放到 `filename*`；不能直接把中文字符写入 `filename="..."`，否则会触发 `http.server` 的 latin-1 编码错误。
- 前端状态持久化用 `localStorage`，关键开关（如 rules 模式）需要保存并恢复，否则刷新后后端会按默认模式处理。
