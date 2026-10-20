# 项目管理工作台（pm_workbench）

本地优先的项目管理助手：开机自动读取各项目 Excel 计划表，桌面常驻显示本周/下周任务并红字标注延期风险，每日定时把摘要推送到手机（WorkBuddy 文件卡片 + Agent Mail 邮件）。已打包为单文件 exe，双击即可打开。

## 架构

```
Excel 计划表（real_plans/，只读）
   │  run_daily.py（每日自动化 / 小组件后台每 10 分钟触发）
   ▼
parser 列映射解析 → SQLite(data/workbench.db) → risk_engine 风险标注
   → output/digest.json（桌面小组件） + output/digest.md（推送正文）
   → Project Pacer.exe 常驻小窗（红字预警） ＋ WorkBuddy 每日自动化（手机推送）
```

## 快速开始

- **双击打开**：`Project Pacer.exe`（想放桌面：右键 exe → 发送到 → 桌面快捷方式）
- 源码方式运行（调试用）：

```bash
PY="C:/Users/Administrator/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
"$PY" run_daily.py             # 手动跑一次，生成摘要
"$PY" widget.pyw               # 打开桌面小组件
```

## 迁移到另一台电脑

**不能整包照搬**，两步走：只复制必要文件 + 目标机上改 `config.json` 里的绝对路径。
完整步骤见同目录 **`迁移指南.md`**。要点：

- **要带**：`Project Pacer.exe`、`config.json`、`real_plans/`；想保留已勾选状态再加 `data/`
- **别带**：`build_exe*/`、`__pycache__/`、`*.pyc`、`*.spec`（纯中间产物）
- **必须改**：`config.json` 的 `plans_dirs` 现在写的是本机绝对路径
  `D:/workspace/项目管理工作台/pm_workbench/real_plans`
- **建议改成相对路径** `"real_plans"`——`common.resolve()` 会自动相对 exe 所在目录解析，
  以后再换机器也不用改配置（2026-10-19 已在全新临时目录实测：成功解析 147 条任务）
- 目标机若窗口空白 → 缺 WebView2 Runtime，装「Edge WebView2 Runtime」

## exe 打包与部署

- 位置：`pm_workbench/Project Pacer.exe`（约 37 MB，PyInstaller onefile 单文件）
- **必须与 `config.json` 保持同目录**——exe 按自身所在目录解析配置和数据（ROOT 逻辑见 `common.py` 的 frozen 适配）。整个 `pm_workbench` 文件夹一起移动，不要单独拷走 exe
- 首次启动有 3~5 秒解压等待（onefile 正常现象）
- 依赖系统 WebView2 运行时（Win11 自带；若窗口空白，安装「Edge WebView2 Runtime」即可）
- 重新打包（改了 .py 之后）：

```bash
# ① 先停掉正在运行的小组件（exe 被占用会导致 PyInstaller 删旧文件失败）
Stop-ScheduledTask -TaskName 'PM_Workbench_Widget'
Stop-Process -Name '项目管理工作台' -Force

# ② 打包（约 45 秒）
cd "D:/workspace/项目管理工作台/pm_workbench"
"C:/Users/Administrator/.workbuddy/binaries/python/envs/default/Scripts/python.exe" -m PyInstaller `
  --onefile --windowed --clean --name 项目管理工作台 `
  --distpath . --workpath build_exe --specpath . `
  --add-data "config.json;." --icon "logo.ico" widget.pyw

# ③ 重新启动
Start-ScheduledTask -TaskName 'PM_Workbench_Widget'
```

> 上面是 PowerShell 语法（反引号 `` ` `` 换行）。用 CMD 的话把 `` ` `` 换成 `^`。
> 打包成功的标志是最后一行 `Building EXE from EXE-00.toc completed successfully.`，
> 产物覆盖同目录的 `Project Pacer.exe`（约 38 MB）。

> ⚠️ **exe 里固化的是打包那一刻的代码和配置**。改了 `.py` 后脚本方式（`run_daily.py`/每日自动化）立即生效，但**小组件 exe 必须重新打包才会用上新逻辑**——2026-10-14 实测事故：升级解析器后忘了重打包，旧 exe 拿旧代码读新配置，表头行探测失效，「干净地」解析出 0 条任务，界面全空且不报错。**只改 `config.json` 不用重打包**（外部配置优先，重启组件即生效），**改 `.py` 必须重打包**。现在解析 0 条任务时小组件会显示红色警告横幅，不会再有这种「静默空视图」

**怎么判断要不要重打包**：只动 `config.json` 里的数值 → 不用打，重启即可；
往 HTML 模板里**新增**了 `__XXX__` 占位符、或改了任何 `.py` 逻辑 → 必须重打。
（本次新增 `titlebar_blur_px`/`titlebar_dim` 就属于后者，所以打了包。）

- `build_exe/`、`build_exe_dbg/`、`*.spec` 是打包中间产物，可随时删除
- 已知坑：从带运行时插桩的进程链（如某些自动化 shell）启动会触发 .NET「不允许附加探查器」错误导致闪退；**Explorer 双击、计划任务拉起均正常**，遇到闪退先换干净上下文验证

## 接入真实计划表（当前配置）

已接入 `real_plans/`：AR882（Sobi & Arthrosi）与 VASQ352 两个项目、147 条任务。两表同模板（英文表头，`Project Plan` 工作表）：

| 配置项 | 值 | 说明 |
|---|---|---|
| `plans_dirs` | `real_plans` | 计划表目录，只读 |
| `sheet_name` | `Project Plan` | 计划数据所在工作表 |
| `header_row` | `0` | 0 = 自动探测表头行（AR882 在第 8 行、VASQ 在第 10 行，自动识别） |
| `project_id_cell` | `B1` | 从该单元格取项目名 → AR882 / VASQ352 |
| `require_task` | `true` | Task ID 为空的行视为阶段汇总行，跳过 |
| `columns` | Task ID→任务、Task Description→子任务、Owner→负责人、Start Date→计划开始、Finish Date→计划结束、Status→状态 | 英文表头映射 |
| `status_map` | Completed/Ongoing/To be started/N/A 等英文状态 | Completed→done、Ongoing→doing、其余→todo |
| `status_write_labels` | doing→`Ongoing`、todo→`To be start`、done→`Completed` | 回写 Excel 时写入的文案，可按表内习惯改 |
| `risk` | `deadline_warn_days: 0`（临期提醒，0=关）、`overdue_grace_days: 1`（逾期宽限天数） | 风险口径阈值 |
| `updated_dir` | `data/updated_plans` | 回写新版本的存放目录 |
| `write_back_excel` | `true` | 关掉则只记本地勾选、不生成新版本 Excel |
| `flush_interval_minutes` | `180` | 待回写队列的批量落盘间隔（分钟）；退出时也会立即落盘 |

- 日期列为文本（如 `TBD`）的行会计入「跳过」计数并在运行时打印——VSQ-032（Potential Mock Audit）即属此类，在表里补上日期后自动出现
- 新增同模板项目：把 .xlsx 丢进 `real_plans/` 即可，无需改配置
- **示例模板（中文，sample_plans/）**：8 列契约 项目名称/任务名称/子任务/负责人/计划开始/计划结束/状态/前置任务；想切回示例数据需同时改回 `sheet_name=null`、`header_row=1`、清空 `project_id_cell`、`require_task=false` 及中文列映射

限制：仅支持 `.xlsx`（`.xls`/`.et` 请先另存为 `.xlsx`）；`~$` 开头的 Excel 临时文件自动跳过；每个文件只读配置的那张工作表。

## 视图口径（2026-10-15 起，按项目分组 + 不重复计算）

三个进行中类目**互不重叠**，每个任务只归一类；每个类目下**按项目分组**，默认全部展开，**点项目名折叠/展开**该组：

| 类目 | 收录规则 | 标注 |
|---|---|---|
| **已延期** | 计划开始日与计划结束日**都早于本周一**（窗口整体在上周及更早）且仍未完成 | 强红色卡片 |
| **本周 10/12 - 10/18** | 未完成且计划窗口覆盖本周：本周启动、或**上周启动本周仍未结束（跨周进行中）** | 带风险标记的用**浅红色**（🔴延期风险角标） |
| **下周 10/19 - 10/25** | 下周**新启动**的任务（跨周任务已在本周视图，不重复收） | 普通卡片 |
| **已完成** | Excel 标记完成 + 手动勾选；同样按项目分组 | 绿色删除线卡片 |

### 风险标记口径（卡片角标）

- 🔴 **已延期**：计划开始日与计划结束日都早于本周一，且状态不是 Completed
- 🔴 **延期风险**（本周类目内浅红标注，两种情况）：
  1. 计划开始日与计划结束日都在本周，状态却仍是 **To be start**（该启动没启动）；
     或计划开始日已经过去却仍未启动
  2. 计划结束日在本周**前半周（周一/周二）**，而今天已经到周三及以后，状态仍 **Ongoing**；
     一般化口径：结束日已过且逾期 ≥ `overdue_grace_days + 1` 天（默认 2 天）仍 Ongoing
- **明确不算风险**：开始在上周、结束在本周、仍 Ongoing = 正常进行中；结束日还没到 = 正常
- 已完成的任务不计风险

> 旧口径「距计划结束日不足 N 天即预警」已默认关闭（`risk.deadline_warn_days = 0`），
> 需要的话把它改回 3 即可叠加这层提醒。

## 桌面小组件：交互（勾选 + 状态按钮 + Excel 回写）

- **左下角切换**：「进行中」主 sheet = 已延期 + 本周 + 下周；「已完成」次 sheet = 已完成任务
- **勾选框**：每项任务左侧 □/☑，点 □ 标记完成，点 ☑ 取消；操作后立即刷新，任务随即在 sheet 间移动
- **状态按钮**：任务名后面跟着一个扁平按钮（显示 `Ongoing` 或 `To be start`），**点一下在两者间来回切换**，卡片分类与风险标注**当场重算**
  - **每条任务互相独立**：状态按任务逐条存（key = `原表文件名|任务|子任务|计划开始`），
    切换任务 A **不会**影响任务 B；同一任务可以无限次来回切
  - 取反基于「界面当前显示的状态」而非 Excel 原始值，所以连点两次一定切得动
- **「面板立即变、Excel 延后写」两层机制**：
  - **界面**：点完立刻按新状态显示、重新判定风险（存在 `data/status_overrides.json` 覆盖层）
  - **Excel**：改动只进队列，**不立即生成文件**；到点才按**最终结果**一次性写入
    ——**关闭小组件时**、或**每 `flush_interval_minutes` 分钟**（默认 180 = 3 小时）
  - 落盘时**每张表只生成 1 个新版本**（同一单元格反复切换只保留最终值），落到
    `data/updated_plans/<原表名>__vN_<时间戳>.xlsx`，原表文件**一个字节都不动**；
    写完后覆盖层自动撤销，此后界面改由 Excel 说了算
  - 窗口顶部黄色提示条显示「⏳ N 处改动待回写」，**点它可立即落盘**（不等 3 小时）
  - 意外被杀（关机/注销）也没关系：队列在磁盘上，下次启动自动补写
  - 想回到「只读原表」：删掉 `data/plan_versions.json` 里对应条目即可
- **标识口径**：key = `原表文件名|任务|子任务|计划开始`（用原表名，跨版本稳定）。改了任务名或计划开始日期，key 变化需重新勾选
- **锁定项**：Excel 状态列原本就是 Completed 且本地没有勾选记录的项显示 ☑ 但锁定
- **已完成清单上限**：默认最多渲染 200 项（`digest.py` 的 `COMPLETED_MAX`），超出的按计划结束日倒序截断

### 回写为什么不用 openpyxl 另存

真实计划表里 Task ID 是公式（`="AR-"&"0"&ROW()-22`），Excel 在文件里存了公式缓存值；
openpyxl 另存**不写缓存值**，新版本再解析会得到空 ID、整张表崩掉。所以 `plan_writer.py`
改成「zip 外科手术式」改写：复制原 zip 所有条目 → 只替换目标 sheet XML 里那一个单元格 →
其余字节原样保留（公式、缓存值、条件格式、图表、图片全部无损）。目标单元格写成 inlineStr。

### 窗口外观（内容区不透明 + 标题栏深灰半透明）

默认形态：**内容区保持原来的不透明深灰 `#17181c`**，只有顶部**自绘标题栏**是深灰半透明。

实现要点（见 `mica.py`）：

- 窗口用 `frameless=True` + `easy_drag=True`：去掉系统标题栏（那层是系统强调色蓝），
  由 pywebview 负责拖动；顶部 `#titlebar` 由 HTML 自绘，**浮在内容层之上**
  （`position:absolute`），背景 `rgba(23,24,28, titlebar_alpha)`，并带
  `backdrop-filter: blur(...) brightness(...)`；
- 内容层 `#scroll` **铺满整个窗口高度**（`position:absolute` 四边归零，顶部留
  `padding-top:34px`），所以**滚动时任务卡片会从标题栏下方经过**——标题栏透出的
  是卡片本身的红/绿，而不是桌面。这正是原先的观感。

> ⚠️ 为什么标题栏要自绘：Windows 原生标题栏不支持 alpha 透明，只能整体着色
> （`DWMWA_CAPTION_COLOR` 不透明）。想让标题栏"透"，只能去掉系统标题栏自己画。

> ⚠️ **标题栏背后是什么，决定了该用哪套参数**。早期版本内容层只占"标题栏以下"，
> 那时背后透出的是**桌面**（偏亮），所以曾引入压暗；现在内容层铺满整窗，背后是
> **深色内容**，压暗必须关掉（`titlebar_dim: 1.0`），否则红绿卡片会被压成黑。
> 只在 `transparent_window=false`、或你确实想透桌面时才调 dim。
> 调通透感请改 `titlebar_alpha`：越小越透，透出的红绿越明显。

在 `config.json` 里调整：

| 配置项 | 默认 | 说明 |
|---|---|---|
| `transparent_window` | `true` | 是否启用无边框+自绘标题栏；`false` 则回到系统原生窗口（蓝色标题栏） |
| `titlebar_alpha` | `0.55` | 标题栏深灰遮罩 alpha，**越小越透**（0=全透，1=全实）；透出的是滚动内容的红/绿 |
| `titlebar_blur_px` | `12` | 背后内容的模糊半径（px），越大越朦胧 |
| `titlebar_dim` | `1.0` | 背后内容的压暗系数（<1 变暗）。**当前背后是深色内容，保持 1.0 不要调**，调小会把红绿卡片压成黑 |
| `caption_color` | `#17181c` | 标题栏底色 / 系统标题色 |
| `window_backdrop` | `none` | 可选 `mica`/`mica_alt`/`acrylic` 铺系统磨砂材质（会跟随壁纸取色） |
| `extend_frame` | `false` | 边框延伸到客户区（配合 backdrop 用） |

**改完只重启组件即可（这几个都是 config 项，不用重新打包）**：
`Stop-ScheduledTask` 关掉再 `Start-ScheduledTask`，或直接双击 exe。

> 踩过的坑：① pywebview 是**异步建窗**的，`start()` 后句柄未必就绪，必须等
> `events.shown` 再调 DWM，否则静默失效 ② `background_color` **只接受 6/3 位
> HEX**，写 8 位会 `ValueError` 崩启动 ③ 窗口每次激活时 pywebview 会**重置
> DWM 属性**，`caption_guard` 后台线程每 5 秒重申一次兜住 ④ **改布局时别给
> `body` 留不透明背景**（会垫在标题栏背后让半透明和模糊全失效），不透明背景下沉
> 到 `#scroll`；也别用 `body{overflow:hidden}` 裁圆角（会连带禁掉页面滚动）。

## 自启动与每日推送

- **开机自启**：Windows 计划任务 `PM_Workbench_Widget`（登录触发）拉起 `Project Pacer.exe`。
  - 立即启动：`Start-ScheduledTask -TaskName PM_Workbench_Widget`
  - 卸载：`Unregister-ScheduledTask -TaskName PM_Workbench_Widget -Confirm:$false`
- **每日推送**（WorkBuddy 自动化「项目管理工作台·每日任务推送」，每天 08:30）：
  1. 运行 `run_daily.py` 刷新数据；
  2. `output/digest.md` 以文件卡片附在回复里（手机可直接打开）；
  3. `config.json` 的 `mail_to` 非空时，经 Agent Mail 发送到该邮箱。
  - 依赖 WorkBuddy 客户端保持运行（已配置开机自启）。

## 目录说明

```
pm_workbench/
  Project Pacer.exe     # 打包成品（双击打开）
  Project Pacer.spec    # 打包配置（重新打包用）
  config.json            # 唯一配置入口
  common.py parser.py    # 配置/解析（frozen 适配/表头自动探测/读最新版本）
  risk_engine.py db.py   # 风险规则/SQLite
  digest.py run_daily.py # 摘要生成/主入口
  done_store.py          # 勾选状态存储
  plan_writer.py         # 状态回写：zip 级改写生成新版本 Excel（原表不动）
  status_store.py        # 界面状态覆盖层（让面板点击后立即切换）
  mica.py                # Win11 磨砂玻璃背景（Mica / Acrylic）
  widget.pyw             # 桌面小组件源码
  real_plans/            # 真实计划表（AR882 / VASQ352，只读）
  sample_plans/          # 示例计划表（中文模板）
  data/                  # workbench.db、done_flags.json、plan_versions.json、
                         # pending_writes.json（待回写队列）、
                         # status_overrides.json（界面状态覆盖层）、widget.log
  data/updated_plans/    # 回写生成的新版本 Excel（v1/v2…，批量合并写入）
  output/                # digest.json / digest.md
 编译中间产物 build_exe/ build_exe_dbg/   # 可删
```

## 路线图

- **一期（已完成）**：解析 → 风险 → 桌面看板 → 每日推送 → exe 打包 → 真实计划表接入
- **二期**：结项项目 WBS 资源库、新项目相似匹配与子任务工期建议、CPM 关键路径
- **三期**：法规/新技术文章周频搜集入库（WorkBuddy 自动化）、月度复盘 HTML+CSV（读周报归档路径）

## 常见问题

- 小组件没弹出来 → 看 `data/widget.log`；手动跑 `run_daily.py` 看是否报错
- 视图是空的 → 若有红色警告横幅，说明解析到 0 条任务：检查 `config.json` 的 `plans_dirs`/`sheet_name`/列映射；**若刚改过 `.py` 代码，先重新打包 exe**（旧 exe 跑的是旧逻辑）。没有横幅才是「真的没任务」
- 换了计划表没变化 → 等 10 分钟（后台自动刷新），或 `Start-ScheduledTask -TaskName PM_Workbench_Widget` 立即重启
- exe 闪退无提示 → 换干净启动上下文验证（见「exe 打包与部署」的已知坑）
