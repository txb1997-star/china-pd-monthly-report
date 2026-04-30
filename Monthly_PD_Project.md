# Monthly PD Report Project

*最后更新：2026-04-30*
*负责人：Summer Tan (PMO)*
*状态：HTML 已上线（4-21 胡总确认），构建系统重构完成（template + build.py + translations.json），中英双版自动产出，4-30 加入产品渲染图自动抽取 + umbrella SKU 拆卡*

---

## 1. 项目概述

Summer 在做一个 Monthly PD Report（HTML），取代 Shine 现有的"长邮件正文 + 横向 PD Table Excel"月报格式。HTML 在 2026-04-21 给胡总过，胡总满意，确认后续所有 China PD Monthly Update 按此格式产出。

**主要受众：** US Sales。但做得好的话 PM 们、Engineering 也可能看，受众面比预想的广。设计时优先满足 Sales 的使用场景。

**Sales 使用 HTML 的三个场景：**
- **Pre-PO（找东西卖）** — actively 在开发但还没拿到 PO 的产品，Sales 看了去 pitch 客户
- **Post-PO（盯自己的货）** — 已下单进入 MP 的产品，Sales 追生产进度
- **开发中已被预订** — 客户提前预订、还在开发的产品，Sales 关心 development 走到哪一步

**语言：** 全英文（Ralph、Shine、US Sales / PM / Engineering）

---

## 2. 起因

Shine Hu 在 2026-04-20 转发两封邮件给 Summer：

**邮件 A — Merlin Yu Engineering Weekly Report（2026-04-17）**
Merlin 每周发 Weekly Projects Status 给美国 Engineering 团队。Ralph 回复说用 Claude AI 花 5 分钟做了一版 8 页结构化 PPT，建议团队 "use these tools to have cleaner and more focused reports"。

**邮件 B — China PD Monthly Update（2026-04-06）**
Shine 发给 Ralph + 全体 US 团队的月度 PD 汇总，按品类/PM 列出所有中国项目进展，附 SharePoint Excel：China PD updates Feb 2026.xlsx。

**Shine 的诉求：**
> "the current we have multiple reports, to find a way to better summary and highlight the issues to be focus is critical for our work… and save time for management."
> "Let's work on to have a new report format for this monthly PD update to US team and see if you can coordinate to get this report run for April."

**两件事不能混做：** PD Monthly Update 和 Engineering Weekly Report 是两条不同的线。Engineering 那边数据怎么拿、tracker 怎么建都没搞清楚——先做月报，月报跑通后再考虑 Engineering。

---

## 3. 胡总 4-21 会议确认

**形式确认：**
- 满意 HTML 形式，所有后续 PD Monthly Update 按此格式产出
- Home page 加 Filter，默认只显示 Project List 上的项目（给 Sales 重点看，可 toggle 看全部）
- 联系 IT 找 hosting 方案，让其他人通过 Link 访问（CC 胡总）

**CRD & Milestone 政策（胡总批准）：**
- PM 必须至少每月更新一版 CRD
- Kick-off 时的 CRD 是 Estimated CRD，约一个月延迟属正常，超出需说明原因
- Weekly Template 中 Milestone Change 列：PM 标注无法按时完成的节点
- Milestone Change 填写 Guidance：Summer 已发给 PM

---

## 4. 三数据源结构

HTML 通过 SKU join 三份数据源生成。**SKU = 唯一的 join key。**

| 数据源 | 文件位置 | 维护人 | 更新频率 | 数据性质 |
|--------|----------|--------|----------|----------|
| **Weekly Tracker** | `Weekly Tracker/China_PD_Weekly_Tracker_WK{周数}.xlsx`（最新一版） | Summer | 周 / 天（PM 周五交，Summer 整理） | 项目进度：阶段、风险、Action、CRD、Milestones、PO 状态 |
| **Summer's Monthly PD Table** | `Monthly PD Report/Summer_Monthly_PD_Table.xlsx` | Summer（基于 Shine 的 China PD Table 整理） | 低频（Shine 发新版时更新） | 商业静态：Brand、Description、Features、Cost、Factory、Port、Duty、MSRP、Sample ETA |
| **Project List** | `Monthly PD Report/Project List {日期}.xlsx` | China PM 们 | 低频（PM 们偶尔更新） | Sales 重点关注白名单 + Project Team 团队配置（Lab / ME / QE / Sourcing / Purchasing / CPM / UI/UX） |

**字段权威性规则（出现冲突时按此处理）：**
- **CRD / Milestone 类**（Kick off / FOT / EB / PP / MP / Inspection 等阶段日期、风险、Current Status、Action）：以 **Weekly Tracker** 为准
- Tracker 没有的字段才 fallback 到 PD Table（如 Est. 1st Inspection 这类商业预估）
- **商业字段**（Cost、Port、Duty、Features、MSRP 等）：以 **Summer's Monthly PD Table** 为准
- **Project Team 团队配置**（哪个 ME / QE / Sourcing 在跟某个项目）：以 **Project List** 为准

**Project List 使用方式（2026-04-28 确认）：**
- **当前：仅做 filter（白名单）。** HTML 默认只显示 Project List 上有的 SKU，可 toggle 看全部。不导入 Project List 的任何字段到 HTML。
- **可选升级（暂不做）：** Project List 中有 Team 配置字段（Lab / ME / QE / Sourcing / Purchasing / CPM / UI/UX）和 Product Release 状态（MOU / Creative briefs / PRD / PA / Certificate Ready / PT Ready / Life test result），这些是另外两个数据源没有的独有信息。如果未来 Sales 需要在 HTML 里查看"这个项目谁在跟"，可以把 Team 配置导入，建议放在详情弹窗的折叠区域。前提：需确认 Sales 是否真的会用、以及 PM 们能否保持 Project List 的 Team 信息及时更新。

---

## 5. 数据更新流程

**核心原则：** 我（Claude）不主动监听邮件、不主动找文件。所有数据更新由 Summer 主动触发。

### 5.1 周更 / 月更产出（最高频）

**触发：** Summer 让我"用最新的 WKxx Tracker 重新生成月报"

**完整 7 步：**

| Step | 谁做 | 做什么 |
|---|---|---|
| 1 | Summer | 把新文件扔进 `Monthly PD Report/`：<br>• 新 `Weekly Tracker/China_PD_Weekly_Tracker_WK{N}.xlsx`<br>• 新 `Summers_Monthly_PD_Table.xlsx`（如果走过 §5.2 SOP）<br>• 新 `China PD updates {Mon} {Year}.xlsx`（图片源，build.py 按 mtime 自动找最新） |
| 2 | Summer | 一句话告诉 Claude："用 WK{N} 重新生成月报" |
| 3 | Claude（沙箱） | 跑 `python3 build.py` → 产出 CN + EN（EN 此时会有部分中文未翻译，build 会 warn 列出来） |
| 4 | Summer | 检查 CN 版本：数据准确性、umbrella SKU 是否有新出现等 |
| 5 | Claude（对话） | 把 build.py warn 的"未翻译中文串"逐条翻译，追加到 `translations.json`（详见 §5.5） |
| 6 | Claude（沙箱） | 重跑 `python3 build.py` → EN 干净 0 warning |
| 7 | 双方 | CN + EN 两份 HTML 是当月 / 当周月报，发给受众 |

### 5.2 PD Table 更新（被动，Summer push 文件触发）

**PM 提交截止日（胡总确认，2026-04-30 邮件通知全体 PM + Shine）：** 每月 26 号前，PM 必须确认 PD Table 和 Project List 已 up to date，SKU 与周报 Tracker 保持一致。

**触发：** Shine 发新版 `China PD updates {月} {年}.xlsx` 给 Summer，**Summer 把文件给我**。

**详细流程文档：** → [`China_PD_Table_Update.md`](China_PD_Table_Update.md)（字段映射、SKU 匹配规则、openpyxl 注意事项、gap analysis、PM 沟通模板）

**SOP 摘要：**
1. 逐 Sheet 水平→纵向 transpose + 按 B 列 label 做字段映射（24 列）
2. 精确 SKU 匹配覆盖现有行数据；不一致的标黄列出给 Summer 找 PM 确认
3. Gap analysis：Tracker 有但 PD Table 无 → 列出让 PM 补商业信息
4. 输出 FINAL + diff 摘要 + PM 中文消息 draft
5. PM 回复后合并最终版，Summer 确认后转正

**关键规则：**
- **绝对不做 SKU 模糊匹配**（后缀有业务含义：SS/CA/CO/MX/D/M/AM 等）
- **数据精度优先级：** Weekly Tracker > PD Table > Draft
- **格式约束：** 24 列结构、Century Gothic 10pt、Short Date，只更新数据不动格式

### 5.3 Project List 更新（被动，Summer push 文件触发）

**触发：** PM 们偶尔更新 Project List，Summer 拿到新版给我。

**步骤：**
1. 新文件放 `Monthly PD Report/`，命名 `Project List {日期}.xlsx`
2. 旧版移 Archive
3. HTML build 时读最新一版

### 5.4 月报新文件命名

- 当月 HTML：`China_PD_Monthly_Report_{月}{年}.html`（如 `_Apr2026.html`）
- 月底切月：保留旧月文件 + 新建下月文件
- 周内迭代版本：`_v{YYYYMMDD}_{改动描述}.html` 备份

### 5.5 EN 版翻译步骤（§5.1 第 5 步详解）

build.py 跑完会 warn 列出未翻译的中文串（PM 写的 issue / next action 等自由文本）。这一步**由对话里的 Claude 完成**，不要试图给 build.py 加 Anthropic API 自动翻译——Cowork 沙箱代理会拦截带 `x-api-key` 的出站请求，技术上行不通。详见 memory `feedback_translation_via_chat`。

**手动翻译流程（每月一次，5-10 分钟）：**
1. 跑 build → warn 列表
2. Claude 逐条翻译（中→英），保留语义 + 缩写（CRD / FOT / PP / EB 不翻；📎 emoji 保留；日期格式 `5/8`/`5月初` 转 `May 8` / `early May`；保留 ✅ ⚠️ 等图标）
3. 全部追加到 `translations.json`（追加到 JSON 末尾，注意 trailing comma 规则；JSON 文件末必须保持 valid）
4. 重跑 build → 0 warning，EN 干净
5. CN + EN 两份 HTML 出版

**翻译风格指南（保持一致性）：**
- 中文里夹杂的英文（如 "Artwork 已交工厂"）→ 翻译时整句重写，不直接保留中文部分
- 工厂 / PM 名字（Cottee / Rowling / 田工 / 陈工等）→ 田工 / 陈工 译成 "Engineer Tian" / "Engineer Chen"，名字音译保留
- 内部缩写（DFM / OTP / FOT / EB / PP / MP / CRD / Oodle 等）→ 一律不翻
- 单位 / 数字 / 型号（cycles / 50pcs / RJ38-... / 6/12 等）→ 一律不翻
- 客户名（Andrew / Ryan / Todd / Tamer / Simon / Denisse / Sarah / Josh / Merlin / Jared）→ 不翻
- 客户公司（Kohl's / BJ's / Costco / KCL / FedEx / UPS）→ 不翻

### 5.6 跨月切换提醒

build.py 顶部 `MONTH_NAME = 'Apr'` 是硬编码的。5 月 Tracker 第一次产出时，Summer 提一句"切到 May"，Claude 改一行：
```python
MONTH_NAME = 'May'
YEAR = '2026'
```
注意 `MONTH_NAME` 用英文三字母简写（Jan/Feb/Mar/Apr/May/Jun/Jul/Aug/Sep/Oct/Nov/Dec），跟 HTML 文件名约定一致。

### 5.7 新 umbrella SKU 检测

如果 PM 在 PD updates 里又搞出新的合并写法（一列多 SKU），build.py 不会自动拆——Claude 跑完后会注意到这种新模式，**先问 Summer** 才能加进 `SPLIT_UMBRELLA_SKUS`。详见 memory `feedback_umbrella_sku_split`。

---

## 6. HTML 当前形态（快照，待迭代）

> 当前 HTML 形态记录如下，但 Summer 仍有不少改造想法。**本次更新只稳数据源，HTML 改造单独再聊**，详见 Todo_List。

**三页结构：**
- **Page 1: PD Table（Sales 选品目录）** — 卡片布局，按 Category 分组，点击展开详情弹窗。详情含 Description / Features / Key Dates / Commercial Details / Notes。Filter: **For Sales (active) / All toggle**（默认 ON，过滤 Project List 白名单）、Category dropdown、PO toggle、Search box
- **Page 2: Pipeline Timeline** — 横向 12 阶段（Kick off → Inspection），每阶段显示项目数 + 点击下钻。**默认进 tab 自动激活 Kick off**，active stage 有 inset shadow + scale(.94) "按下" 视觉。详情表含 SKU / Category / PM / Risk / **PO** / Next Action 列
- **Page 3: Weekly Tracker** — 项目进度详情，Filter: PM / Category / Risk / Location / **PO (Placed/No PO)** / **For (buyer dropdown)** / Search

**顶栏 Stats Bar（4-21 胡总确认）：**
- 5 个浮动可点击卡片：Total Projects / High Risk / Medium Risk / Tier 1 (CSM) / In MP
- 视觉：白色圆角卡片 + 顶部色条（Total/T1=深蓝渐变, High=红, Mid=橙, MP=绿），数字统一深蓝
- 交互：Total → 跳 Pipeline；High/Mid Risk → 展开详情面板（表格：SKU/PM/Status/Issues/NextAction/CRD/Location），**点同一 stat card 收起 / 点别处也收起**；T1/MP → 筛选 Page 1 卡片
- **数字规则（4-29 改）：** stats 必须等于 page1 "All" toggle 下能展开的卡片数。计算公式：从 page1Data 里筛 `category != ''` 的项算 risk/status/tier 各项分布。**不再用 Tracker 全量**——这样保证 stats 和点击后看到的卡片数严格一致，避免 Sales 看到数字与卡片对不上的困惑。

**视觉与交互偏好：**
- 顶栏卡片要浮动质感（圆角阴影），不要平铺表格
- 颜色与语义一致：High=红, Mid=橙, MP=绿, Total/T1=深蓝
- 数字不要彩色底色，统一深蓝
- Risk Detail Panel 不需要 Category 列（大家都知道）
- 数据命名：RJ38-G4 就叫 RJ38-G4，不要"玻璃碗"

**Excel 数据源格式约束：**
- 字体统一 Century Gothic
- 日期：具体日期 Short Date（不带时间），模糊日期如 "2026 April" 保持文字原样
- Summer's Monthly PD Table 与 Weekly Tracker 保持相同 PM 分组和 SKU 顺序

**Page1 可见性规则（4-29 确定）：**
- **PD Table 是 page1 卡片的权威数据源。** PD Table 没填 tier 或 category 的 SKU **不会**出现在卡片里。
- 这规则是 intentional 的：US-side tracking-only 项目（如 C55-11/C60/C45/C15）China PM 只在 Tracker 跟进度，不维护商业字段，自然不会被 Sales 看到。
- 后果：stats（也只数有 category 的）会比 Tracker 全量小。例如 Tracker 70 个，page1 67 个；Tracker T1=7，page1 T1=3。差额对应 US-side 项目，是预期行为不是 bug。

**Banner 规则（4-29 确定）：**
- 触发：某 PM 累计 ≥3 个 pending SKU（PD Table 待确认/Gap section）→ banner 显示该 PM 负责的 category 列表（英文）
- 文案模板：`X categor(y/ies) currently lack complete commercial data — pending updates from related PM.`
- 不点名 PM 个人（避免针对感）；自动从 build.py 内的 `PM_SECTION_TO_CATEGORIES` 字典查映射
- 阈值 ≥3 是为了避免单个 SKU 缺数据触发 category-级 warning（数据正常波动）

**双语输出（4-29 加）：**
- 每次跑 build 同时产出 CN（给 Shine + 国内）和 EN（给 US Sales）两份 HTML
- 翻译字典：`Monthly PD Report/translations.json`，约 140 条 PM 写的中文短语 → 英文映射
- 自动翻译字段：Page1 (currentStatus/category/crd) + Page3 (issue/nextAction/currentStatus/category/poRaw/crd) + Pipeline projects (action/category)
- 不翻译：SKU、PM 名字、buyer 名字、英文/数字、PD Table 预填的英文 description/features
- build.py 报漏译 → Summer 加进 translations.json 即可（不改 build.py）

---

## 7. 产品渲染图（2026-04-30 上线）

**结论：** PM 在 `China PD updates {Mon} {Year}.xlsx` 里内嵌的产品渲染图可以**自动抽出来嵌进 HTML**，不需要单独 folder 维护。

**机制：**
- xlsx 本质是 zip 包，图片放在 `xl/media/` 里。openpyxl 通过 `ws._images` 拿到每张图的锚点（`anchor._from.col / row / colOff / rowOff`）。
- 对每张图：anchor `_from.col + 1` 找它落在哪一列，去 Row 10 取那一列的 Model（SKU）当 key。
- 图片用 PIL 缩到 **300×300px + JPEG q78**，base64 内嵌到 HTML 里。一张图 ~8-15KB，46 张总共让 HTML 从 154KB 涨到 ~600KB（仍然单文件邮件可发）。
- 每月跑 `build.py` 时自动重新抽，不需要手动同步。

**文件命名约定：**
- 源文件：`Monthly PD Report/China PD updates {Mon} {Year}.xlsx`
- build.py 用 `glob('China PD updates *.xlsx')` + 按 mtime 取最新，下月 Shine 发新文件直接拖进同目录即可，不用改路径

**多 SKU / 多图同列处理（umbrella SKU 拆卡，2026-04-30 Summer 确认）：**

PM 在 PD updates 里偶尔会把多个变体放一列：

| 类型 | 例子 | 处理 |
|---|---|---|
| 多 SKU 共享 1 张图 | `RJ50-SFDAF-25D(SS) / RJ50-BFDAF-25D(BLK)`（颜色变体共用一张图） | 所有 SKU 注册同一张图；HTML 里拆 2 张卡 |
| 多 SKU 各有图（2×2 网格） | `RJ64-10-PTC / BTR / LVD / Aqu`（4 色冰淇淋，每色一张图） | 按 anchor (rowOff bucket → colOff) reading order 1:1 配对；HTML 里拆 4 张卡 |
| 单 SKU 单图 | 普通情况 | 直接映射 |

**Page 1 拆 / Page 2、3 合并的设计：**
- Page 1（Sales 选品目录）：每个变体显示成独立卡片，复制 umbrella 行的所有商业字段（cost / MSRP / description / features / factory / port / duty / 备注…），用变体自己的图（找不到则 fallback umbrella 共享图）
- Page 2（Pipeline）和 Page 3（Weekly Tracker）：保持 umbrella 一行不动，开发进度按 umbrella SKU 跟踪

**确认的 umbrella → variants 映射（写在 build.py 顶部 `SPLIT_UMBRELLA_SKUS` 字典）：**

```python
SPLIT_UMBRELLA_SKUS = {
    'RJ50-SFDAF-25D':     ['RJ50-SFDAF-25D(SS)', 'RJ50-BFDAF-25D(BLK)'],
    'RJ62-20A-Series':    ['RJ62-BLK', 'RJ62-WHT'],
    'RJ64-10-new colors': ['RJ64-10-PTC', 'RJ64-10-BTR', 'RJ64-10-LVD', 'RJ64-10-Aqu'],
}
```

新增映射**必须先问 Summer**，不要自动检测就拆。

**括号尾标剥离 alias（image-only）：** PD updates 里 `RJ50-SFDAF-25D(SS)` 注册时同时注册 bare 版 `RJ50-SFDAF-25D` 作 alias，让 PD Table 的 parent SKU 命中变体的图。仅图片字典使用，业务字段（PD Table 合并 / Tracker join 等）仍然精确匹配。

**当前覆盖率（4-30 跑出来）：** 72 个可见 page1 卡片，37 张有图（51%）。剩下 35 张是 placeholder，原因：
- Liz 在休假，水壶/微波炉系列没上传图（10+ 个）
- 部分 RJ34 / RJ07 / RJ15 系列 PM 当月没补图
- C 系列 / Aquart（US-side 项目，PM 不维护）
- 区域版本（MX / CA / EU 等通常不出独立图）

**没图的卡片** 在 template.html 里 fallback 到 `getCatIcon(category)` emoji 占位图标，不影响其他功能。

---

## 8. 构建系统（4-29 重构）

**核心文件（`Monthly PD Report/`）：**
- `template.html` — 完整单文件 HTML 模板，包含 CSS / DOM 结构 / 所有渲染 JS。数据位置用 5 个占位符：
  - `{{PAGE1_DATA}}`、`{{PIPELINE_DATA}}`、`{{PAGE3_DATA}}`、`{{SUMMARY_STATS}}`、`{{BANNER_BLOCK}}`
- `build.py` — 主构建脚本。读三个 xlsx + 抽 PD updates 图 → SKU 精确 join → 渲染 4 份 JSON + banner HTML → 灌进模板 → 输出
- `translations.json` — CN→EN 翻译字典（约 140 条）

**路径处理（4-30 改）：** `BASE = Path(__file__).resolve().parent.parent` 自动从脚本位置推导工作目录，不再硬编码 session id。`SCRATCH = Path(os.environ.get('CLAUDE_SCRATCH', '/tmp/pd_report_scratch'))`，由 Cowork 注入，跨 session 自动跟随。`PDUPDATES_PATH` 用 glob 找最新 `China PD updates *.xlsx`。

**数据流：**
```
3 xlsx 数据源 + China PD updates xlsx (图片源)
    ↓ load_tracker / load_pd_table / load_project_list / extract_sku_images
内存中的 dict / list + images dict {sku → base64 data URI}
    ↓ build_page1_data (umbrella SKU 拆卡 + 图片 join) / build_page3_data / build_pipeline_data
4 份 JSON
    ↓ build_summary_stats（从 page1 visible 算）
summaryStats
    ↓ build_banner_html（PM 阈值检测）
banner HTML
    ↓ render_template（占位符替换）
最终 HTML
    ↓ write_with_rotation（current → _prev → 写新版）
两份输出文件（CN + EN）
```

**图片抽取关键函数（`extract_sku_images`）：**
- 遍历 PD updates 每个 sheet 的 `ws._images`
- 按 anchor `_from.col + 1` 找列，去 Row 10 找 SKU
- 多图同列时按 `(rowOff // 300000, colOff)` reading order 排序，N 张图 N 个 SKU 1:1 配对（如 4 色冰淇淋）；不足时所有 SKU 共享每张图（如 SS/BLK 共享）
- PIL `thumbnail((300, 300), LANCZOS)` + `JPEG q=78 optimize=True` → base64 data URI
- 括号尾标 alias（`_sku_image_aliases`）：仅图片字典用，详见 §7

**输出文件命名（4-29 确定）：**
- CN：`China_PD_Monthly_Report_{Mon}{Year}.html`（如 `_Apr2026.html`）
- EN：`China_PD_Monthly_Report_{Mon}{Year}_EN.html`
- 上一版备份：`_prev.html` 后缀（每月份 family 各保留一份 prev）
- 月份按"数据所属月份"命名（April 报告 = `_Apr2026.html`），不按运行日期。月底切换到 May 时手动改 build.py 里的 `MONTH_NAME` 常量。

**Rotation 规则：**
- 同一月份命名 family：跑第 N 次 → 当前 → 改名 `_prev`，旧 `_prev` 删掉，新内容写当前
- 跨月不影响（4 月跑产出 `_Apr2026.html`，5 月跑产出 `_May2026.html`，互不干扰）
- 4-24 那批 baseline HTML（`China-PD-Monthly-Update.html` / `index.html` / `_v20260421_*`）一律不动

**写文件双跳避免 OneDrive 截断：**
- build.py 先写到 outputs scratch（`/sessions/.../mnt/outputs/`）
- 然后 `shutil.copyfile` 到 OneDrive 工作目录
- 如果遇到 PermissionError，回退到 `out_path.write_text` 直写

**已知技术坑：**
- **OneDrive Files On-Demand bug：** 文件 mtime/大小看着正常，但内容是全 0 字节。读到 BadZipFile / "File is not a zip file" 错误。修复：右键文件 → "Always keep on this device" 等同步成绿对勾，或在 Excel 里打开强制下载，或 Save As 到非 OneDrive 路径。
- **Excel 锁文件：** Excel 还开着时读 xlsx 会报 BadZipFile。先关 Excel 再让 build.py 读。
- **Edit / Write tool 截断 build.py：** 大文件被工具截断或注入 null bytes 是常见事故（4-30 当天截断了两次）。修复办法：head -n 到 last clean line + Python 脚本 append 剩余部分（heredoc 也行但小心 `!` 转义），或者 `data.replace(b'\x00', b'')` 清 null bytes 后重写。每次大改完用 `python3 -c "import ast; ast.parse(open(f).read())"` 验证语法。
- **bash heredoc 会转义 `!`：** 写 HTML / JS 用 Python 文件 IO，不用 heredoc。
- **WMF 图片格式：** openpyxl 读 PD updates 时如果遇到 WMF 格式的图会 warn 并丢弃。WMF 是少数情况（绝大多数 PM 用 PNG/JPEG），目前没专门处理。
- **PIL 处理透明度：** RGBA / LA / P 模式图先在白底上 paste 一次再保存为 JPEG（JPEG 不支持 alpha），避免黑底。

---

## 9. 当前数据源状态（2026-04-30）

- **Weekly Tracker WK17** — 74 行 SKU。Summer 持续维护。PO 列已完整填写。
- **Summers Monthly PD Table** — 82 个主表 SKU + 12 个 pending（待确认 / Gap）。其中 Liz 的 Kettle/Microwave section 有 10 个 pending（驱动 banner 触发）。
- **Project list.xlsx → China Projects sheet** — 41 个白名单 SKU（C4 列 Model）。命名是 `Project list.xlsx`（无日期后缀）。
- **China PD updates Apr 2026** — Shine 4-29 给的版本，已经走完一遍 §5.2 SOP，更新合并到 PD Table。**4-30 起也作为图片源**（46 张内嵌产品渲染图，build.py 自动抽出来）。

**HTML 输出：**
- `China_PD_Monthly_Report_Apr2026.html` — 中文版，Page 3 issue/action 保留中文
- `China_PD_Monthly_Report_Apr2026_EN.html` — 英文版，所有数据字段翻译
- 都是 4-21 胡总确认的三页结构 + Stats Bar + Risk Panel
- 4-29 新增功能：Project List filter toggle (For Sales / All)、Pipeline 默认 Kick off 激活、Pipeline / Tracker 加 PO 列、Tracker 加 PO/Buyer filter、Banner（PM 阈值触发）、双语自动产出
- 4-30 新增功能：产品渲染图自动抽取嵌入卡片、umbrella SKU 拆卡（`SPLIT_UMBRELLA_SKUS`）、PD updates 文件自动找最新

**Banner 当前触发：** Liz Liu — Kettle and Microwave categories（10 个 pending SKU）

**Stats Bar 当前数字：** Total=72, High=4, Med=12, MP=15, T1=3（从 page1 visible 算，4-30 因为拆卡 +5）

**图片覆盖：** 72 个可见卡片中 37 个有真实渲染图（51%），其他 fallback 到 emoji 占位图标。详见 §7 末尾的覆盖率分析。

---

## 10. 关键文件

| 类型 | 文件 |
|------|------|
| 项目文档（本文件） | `Monthly PD Report/Monthly_PD_Project.md` |
| PD Table 更新流程 | `Monthly PD Report/China_PD_Table_Update.md` |
| 任务追踪 | `Monthly PD Report/Todo_List.md` |
| HTML 输出（CN） | `Monthly PD Report/China_PD_Monthly_Report_{月}{年}.html` |
| HTML 输出（EN） | `Monthly PD Report/China_PD_Monthly_Report_{月}{年}_EN.html` |
| 数据源 1（Tracker） | `Weekly Tracker/China_PD_Weekly_Tracker_WK{周数}.xlsx` |
| 数据源 2（PD Table） | `Monthly PD Report/Summers_Monthly_PD_Table.xlsx` |
| 数据源 3（Project List） | `Monthly PD Report/Project list.xlsx`（China Projects sheet） |
| HTML 模板 | `Monthly PD Report/template.html`（5 个 `{{...}}` 占位符） |
| 构建脚本 | `Monthly PD Report/build.py` |
| 翻译字典 | `Monthly PD Report/translations.json` |

---

*参考：Email_Tracking_Rules.md | Company Org Chart April 2026.pdf*
