# FBA需求分配核查工具 — 项目迁移包

> 接收方：这是从 WorkBuddy 另一个账号打包过来的完整项目，按下面的"快速上手"3 步即可运行。

## 这是什么

一个**本地 + Render 部署**的销量/FBA 数据处理工具，4 个页签：
1. **店铺销量处理** — 销量快照 → 加权日均销量（含按运营级别×备货规则配置）
2. **FBA需求差异分析** — 销量/运营/FBA需求 三表合一比对，按 (店铺ASIN, 天数) 逐天核对
3. **多余欧洲区销量核对** — 欧洲区多余销量按 MSKU 累计
4. **FBA分配核对** — 重算 FBA 每周期分配量，与实际分配量比对

## 快速上手（3 步）

### 1. 安装 Python 依赖

```bash
# 推荐 Python 3.10+（与 Render 一致）
pip install -r requirements.txt
```

主要依赖：`openpyxl`、`python-calamine`、`flask`-free（用的是 `http.server`）。

### 2. 启动本地服务

**Windows：**
```bash
# 双击或在 cmd 里执行
start_server.bat
```

**Mac / Linux：**
```bash
python server.py 8000
```

启动后浏览器打开 <http://localhost:8000>，看到「已连接（服务器模式）」即成功。

### 3. 测试一下

跑回归测试（确保代码完整）：

```bash
python _test_fba_rules.py     # 页签② 单元测试
python _e2e_fba_http.py       # 页签② HTTP 端到端
```

两条都应输出"全部断言通过 ✅"。

## 文件清单

```
.
├── server.py                        # 后端核心（http.server + calamine 流式读写）
├── store-sales-processor.html       # 前端单页应用（4 个页签）
├── process_sales.py                 # 早期 CLI 脚本（可独立跑销量处理）
├── start_server.bat                 # Windows 一键启动
├── requirements.txt                 # Python 依赖
├── runtime.txt                      # Render Python 版本
├── Procfile                         # Render 进程定义
├── render.yaml                      # Render 部署配置
├── _test_fba_rules.py               # 回归测试：页签② rules 模式
├── _e2e_fba_http.py                 # 回归测试：HTTP 端到端
├── _test_weighted_rules.py          # 回归测试：加权日均规则
├── .gitignore                       # 排除临时文件
└── .workbuddy-ai/
    ├── memory/
    │   ├── MEMORY.md                # 项目长期约定（核心算法/部署/列名）
    │   └── YYYY-MM-DD.md            # 每日工作日志
    └── skills/                      # （如有）项目级技能
```

## Render 部署（可选）

如果新账号也想用线上版本（`https://xxx.onrender.com`）：

1. 新账号登录 Render：<https://render.com>
2. **New +** → **Web Service** → 连接同一个 GitHub 仓库（或 fork 到新账号的 GitHub）
3. 关键设置：
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python server.py $PORT`
   - **Instance Type**: Free（注意会休眠）
4. 部署完成后访问 `https://<service-name>.onrender.com/api/ping` 应返回 `{"ok":true,"name":"store-sales-processor"}`

> ⚠️ Render free 实例 15 分钟无请求会休眠；线上模式建议 15 分钟内完成下载，或直接用本地模式。

## 数据文件（需自行准备）

- 销量快照 .xlsx（列：店铺/ASIN/三天/七天/十四天/三十天/60天/90天日均销量 等）
- 运营级别 .xlsx（列：店铺/ASIN/销售状态/配送类型/运营级别/店铺ASIN）
- FBA 需求 .xlsx（列：店铺ASIN 或 店铺+ASIN/备货周期 或 天数/需求量/日均销量(加权)）
- 运营级别映射表（可选，仅在「开启备货规则」模式）：列 [店铺, ASIN, 运营级别, 补货方式]

## 核心算法速记

- **A 列(店铺ASIN)** = 店铺 + ASIN 拼接，无分隔符
- **M 列(日均销量加权)** = 组合窗权重归一化（默认 0.6×3天 + 0.3×7天 + 0.1×14天；rules 模式可按运营级别配置 6 窗口权重或固定窗）
- **页签② 比对键** = (店铺ASIN, 天数)，当 ops 含「天数」列且 FBA需求含「备货周期/天数」列时启用
- **页签④ 分配规则**：周期0需求全分满 → 周期>0 按销量占比去尾封顶 → 剩余按预计分配量最大补；周期0「追加」最后才分

详见 `.workbuddy-ai/memory/MEMORY.md`（项目长期约定）和 `YYYY-MM-DD.md`（开发日志）。

## WorkBuddy 内的迁移

把这个 zip 解压到新账号的 workspace 目录（如 `D:\workBuddy\WB\新项目名`），然后在 WorkBuddy 里"打开项目"选这个目录即可。

`.workbuddy-ai/memory/` 里的内容会**自动继承**——新账号打开后 WorkBuddy 会读到这些记忆，算法约定/部署信息/坑点提示都还在。

如果新账号想从头积累记忆，把 `.workbuddy-ai/memory/MEMORY.md` 删掉即可。

## 已知坑（从开发日志提炼）

1. **Render 免费实例会休眠** — 15 分钟无请求后下次首次访问冷启动约 30-50 秒；处理大文件（10万行+）需要 2-3 分钟，期间不要刷新页面
2. **真实文件 80MB / 57万行** — 浏览器 SheetJS 解析不了，必须走 server.py（calamine + 流式写最小 xlsx）
3. **页签② 运营级别列名** — 当填了"符合补货条件的运营级别"白名单时，运营级别文件**必须**含 `运营级别` 列；不填白名单则不校验
4. **FBA 需求列名自适应** — 「店铺ASIN」或「店铺+ASIN」任一即可；「备货周期/天数/周期/补货周期/FBA备货周期」任一即可
5. **venv 路径** — Windows venv 在 `C:\Users\71721\.workbuddy\binaries\python\envs\default\Scripts\python.exe`（不是 bin/），新账号需重新建 venv

## 联系方式 / 交接说明

> [在这里写明交接背景、谁接手、什么时间点、有什么未完成事项]
