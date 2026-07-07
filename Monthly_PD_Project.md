# Monthly PD Report Project

*最后更新：2026-07-04（全文一致性清理：清 Cowork 时代残留 + 列数矛盾 + 过时快照，详见 §11 变更记录）*
*负责人：Summer Tan (PMO)*
*状态：HTML 已上线（4-21 胡总确认），构建系统重构完成（template + build.py + translations.json），中英双版自动产出。5-04 大改：纯镜像 PD Table 重建 + Stats Bar 重做 + Pipeline US/MX 拆分 + ASI/NPD filter + Placeholder 占位卡片 + Category 合并 + "Other" 收纳。5-07 大改：Tracker 25 列适配（5/6 P/V 列加进来）+ umbrella 拆卡机制彻底删除（PD updates 端 PM 拆开 + rebuild 端自动按 cell 拆行）+ MONTH_NAME 切到 May。5-19 大改：Tracker 26 列适配（5/19 加 col E "NPD/ASI"）+ ASI 来源从 config 切到 Tracker col E（pd_table_config.json `after_sales_improvement` 字段删除）+ build.py 新增 Excel 365 "image-in-cell" 解析支持（与传统 image-over-cell 并存）+ zero-area 幽灵图过滤（修 RJ44-CB 配错图 bug）。**5-26 大改：Page 1 加 Compare modal（per-category "⇄ Compare" 按钮 → 浮层 + 遮罩 + 5 列 Spec Cards 横向对比，每张卡直显 Cost / MSRP / Factory / Port / Duty / 40'HC / CRD / Sample ETA 八个商业字段，Other 不出按钮、placeholder 在 modal 内隐藏、Page 1 所有 filter 状态全部继承；触发自 Shine 5-26 批准 "as it is"）。* **7-1 大改：Weekly Tracker 删除 11 个阶段列（Kick off→MP，原 Q–AA），Tracker 回 16 列；build.py 清 `stage_label_map`/27列fallback/`stages` 死代码（HTML 从不渲染阶段列，Page2 按 Current Status 分桶，同源双跑验证删除零影响）；PA 列扩为 4 态（加 待批复 / Waiting for Signing，paNorm/paBadge/PA_LABELS/translations 已改）；4 位采购周报的 PA/6A/PO/CRD 一手数据并入 Tracker 四列（单表合并、dimtable 映射，详见 Weekly Tracker/PM_Weekly_Tracker.md）。***

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

## 4. 数据源结构（2026-07-02 起两个，Project List 已退役）

HTML 通过 SKU join 两份数据源生成。**SKU = 唯一的 join key。**

| 数据源 | 文件位置 | 维护人 | 更新频率 | 数据性质 |
|--------|----------|--------|----------|----------|
| **Weekly Tracker** | `Weekly Tracker/China_PD_Weekly_Tracker_WK{周数}.xlsx`（最新一版） | Summer | 周 / 天（PM 周五交，Summer 整理） | 项目进度：阶段、风险、Action、CRD、Milestones、PO 状态 |
| **Summer's Monthly PD Table** | `Monthly PD Report/Summers_Monthly_PD_Table.xlsx` | Summer（基于 Shine 的 China PD Table 整理） | 低频（Shine 发新版时更新） | 商业静态：Brand、Description、Features、Cost、Factory、Port、Duty、MSRP、Sample ETA |
| ~~Project List~~ | **已退役（2026-07-02）** | — | — | 原用途：Sales 白名单 filter。Summer 删除 HTML 的 For Sales/All toggle 后失去消费方，build.py 死代码同日清除，xlsx 删除 |

**字段权威性规则（出现冲突时按此处理）：**
- **CRD / Milestone 类**（Kick off / FOT / EB / PP / MP / Inspection 等阶段日期、风险、Current Status、Action）：以 **Weekly Tracker** 为准
- Tracker 没有的字段才 fallback 到 PD Table（如 Est. 1st Inspection 这类商业预估）
- **商业字段**（Cost、Port、Duty、Features、MSRP 等）：以 **Summer's Monthly PD Table** 为准
> **Project List 退役记录（2026-07-02）：** 原本仅做白名单 filter（2026-04-28 确认）。Summer 后来删除了 HTML 的 For Sales/All toggle，白名单没有了消费方；7-2 验证 template.html 中 `onProjectList` 零引用后，清除 build.py 全部相关死代码并删除 xlsx。历史上曾考虑的"导入 Team 配置到详情弹窗"升级一并作废。

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
| 3 | Claude（本机） | 跑 `python build.py` → 产出 CN + EN（EN 此时会有部分中文未翻译，build 会 warn 列出来） |
| 4 | Summer | 检查 CN 版本：数据准确性、umbrella SKU 是否有新出现等 |
| 5 | Claude（对话） | 把 build.py warn 的"未翻译中文串"逐条翻译，追加到 `translations.json`（详见 §5.5） |
| 6 | Claude（本机） | 重跑 `python build.py` → EN 干净 0 warning |
| 7 | 双方 | CN + EN 两份 HTML 是当月 / 当周月报，发给受众 |

### 5.2 PD Table 更新（被动，Summer push 文件触发）—— **2026-05-04 大改：纯镜像重建**

**PM 提交截止日（胡总确认，2026-04-30 邮件通知全体 PM + Shine）：** 每月 26 号前，PM 必须确认 PD Table 已 up to date，SKU 与周报 Tracker 保持一致。（原要求还包括 Project List，其已于 2026-07-02 退役。）

**触发：** Shine 发新版 `China PD updates {月} {年}.xlsx` 给 Summer，**Summer 把文件给我**。

**详细流程文档：** → [`China_PD_Table_Update.md`](China_PD_Table_Update.md)

**核心原则（2026-05-04 与上一版的关键差异）：**
- **每月从零重建**：新 PD Table = 当前 PD updates 文件的纯镜像。**不读上一版作 baseline、不 merge、不保留 PM 已删除的 SKU。**
- **保留隐藏列**：openpyxl 默认会读到 PM 隐藏的列，全部保留进 PD Table。
- **TBD/TBC 占位符**：跳过（不是真产品）。
- **多 SKU 单元格**：按换行/制表符拆分，每个 SKU 独立一行共享商业字段。
- **不再做 SKU 一致性比对**：不是 §5.2 的事；交给 §5.2 完成后 build 阶段 / Tracker 比对自动出 diff。

**新 SOP 摘要（一键脚本 `rebuild_pdtable.py`）：**
1. 读 `Monthly PD Report/pd_table_config.json`（ASI 列表、umbrella 拆卡、manual_additions 等）
2. 找最新 `China PD updates *.xlsx`（项目目录优先，否则 uploads；OneDrive Files On-Demand bug 自动 fallback）
3. 逐 Sheet 水平→纵向 transpose + 按 B 列 label 做 24 列映射，跳过 TBD/TBC
4. 追加 manual_additions 行（PM 还没填的 placeholder business info）
5. 输出新 `Summers_Monthly_PD_Table.xlsx`（直接覆盖）
6. **自动**对比 PD Table SKU 和最新 Weekly Tracker SKU，输出 3 段 diff：
   - **A 类**：Tracker 有 PD Table 没有（已过滤 ASI/MP）→ PM 需补商业信息
   - **B 类**：PD Table 有 Tracker 没有 → 项目状态待 PM 明确
   - **C 类**：Tracker 已 MP（Project Released）→ 不需要 PD Table info

**关键规则：**
- **绝对不做 SKU 模糊匹配**（后缀有业务含义：SS/CA/CO/MX/D/M/AM 等）
- **数据精度优先级：** Weekly Tracker > PD Table > Draft
- **格式约束：** 24 列结构、Century Gothic 10pt、Short Date

### 5.2.1 pd_table_config.json 配置（2026-05-04 新增，5-07 + 5-19 简化）

`Monthly PD Report/pd_table_config.json` 是 PD Table + HTML build 的统一配置入口。Summer 直接编辑这个文件就能调整以下行为，**不用改代码**：

| 字段 | 用途 |
|------|------|
| `sku_aliases` | Tracker SKU → PD Table canonical 名映射，用于 rebuild gap analysis + build.py tracker-row matching。如 `RJ38-G4-AS → RJ38-G4-LS`（Cottee typo）、`RJ38-9TW-V3 → RJ38-9TW-V2`、`RJ56-DIS-V2-CA-CO → RJ56-DIS-V2`。 |
| `mp_overrides` | PD Table SKU 强制当 MP（Released）。用于 PM 在 PD updates 把多色拆开但 Tracker 只有一行 MP 时。如 `RJ64-10-V2-{PTC,BTR,AQU,LVD}`。 |
| `manual_additions` | 注入 PD Table 的额外行。当 PM 还没把项目放进 PD updates 但 Summer 已经有 business info 时用这个。每条 entry 含 SKU 列表 + PM + 字段映射。 |

**已删字段：**
- `umbrella_to_variants`（5-07 删）：死代码（build.py 用自己硬编码的 SPLIT_UMBRELLA_SKUS，根本没读这个 config 字段）。新机制：PD updates 端 PM 把多色拆成独立列（每列一 SKU + 独立图），rebuild 自动每 cell 一行。共享一张图的多 SKU（如 RJ50 SS/BLK）仍可写一个 cell，rebuild `parse_sku_cell()` 按换行/制表符拆行，商业字段填重复。
- `after_sales_improvement`（5-19 删）：ASI 来源切到 Tracker col E "NPD/ASI"，PM/Summer 在 Tracker 维护，不再在 config 里写死。详见 §5.2.2。

**配置现状（会随周/月更漂移，以 json 实物为准；2026-07-04 实测）：**
- 3 个 sku_aliases
- 4 个 mp_overrides（RJ64-10-V2 四色 PTC/BTR/AQU/LVD）
- 3 个 manual_additions
- 5 个 crd_changes（6-30 加的字段，review-gated，每周重建）

### 5.2.2 ASI 数据源切到 Tracker col E（2026-05-19 新增）

**触发：** 5/19 WK21 Tracker 加了 NPD/ASI 列（col E），PM 可以直接在 Tracker 标 ASI。

**新规则：**
- **Tracker col E 是 ASI 的 single source of truth**：空白 = NPD（默认），写 `ASI` 就是 After Sales Improvement
- build.py `load_tracker` 读 col E 进 `npd_asi` 字段；main() 计算 `asi_set = {r['sku'] for r in tracker_rows if r['npd_asi'].upper() == 'ASI'}`
- rebuild_pdtable.py `load_tracker_skus` 把 npd_asi 放进 tuple 位置 3；`compare_pd_vs_tracker` 用 `asi = {sku for sku, info in tracker_skus.items() if info[3] == 'ASI'}`
- 删了 `pd_table_config.json` 的 `after_sales_improvement` 字段

**5-19 WK21 ASI 集（7 个）：** RJ38-10-RDO-V3, RJ54-G-SS, RJ54-G-SS-D-BLK, RJ54-SS-15-D-UK-EU, RJ62-20A-Series, RJ64-10-V2-WHT, RJ64-10-new colors

**改 ASI 的操作：** 改 Tracker col E（不是 config 了）→ 跑 build。

### 5.3 ~~Project List 更新~~（已作废，2026-07-02 Project List 退役）

### 5.4 月报新文件命名

- 当月 HTML：`China_PD_Monthly_Report_{月}{年}.html`（如 `_Apr2026.html`）
- 月底切月：保留旧月文件 + 新建下月文件
- 周内迭代版本：`_v{YYYYMMDD}_{改动描述}.html` 备份

### 5.5 EN 版翻译步骤（§5.1 第 5 步详解）

build.py 跑完会 warn 列出未翻译的中文串（PM 写的 issue / next action 等自由文本）。这一步**由对话里的 Claude 完成**，不给 build.py 加 Anthropic API 自动翻译。（原始原因是 Cowork 沙箱代理拦截 `x-api-key` 出站请求；2026-07-02 迁本地后该限制已消失，但流程维持对话翻译——每月一次 5-10 分钟，不值得为此接 API + 管理 key。）

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

### 5.6 跨月切换（2026-06-02 起已自动化，无需改代码）

build.py 按 **"每月 10 号切月"规则（Summer rule 2026-06-02）自动推导** `MONTH_NAME` / `YEAR` / `REPORT_PERIOD`：构建日 ≥ 当月 10 号 → 报当月；< 10 号 → 报上月（1 月自动跨年）。`DATA_ASOF` 同样自动取"构建日之前最近的周五"。**"Summer 提一句、Claude 改一行"的旧流程已废除**（旧文档描述硬编码 `MONTH_NAME = 'Apr'` 为 2026-06-02 之前的历史，2026-07-04 审计发现本节漏更、已改）。

### 5.7 PD updates 多 SKU cell 处理（5-07 改）

5-07 起 umbrella 字典彻底删掉。PM 在 PD updates 里如果一列写多 SKU（换行分隔，如 RJ50-SFDAF-25D(SS) / RJ50-BFDAF-25D(BLK) 共享一张图）：rebuild_pdtable.py 的 `parse_sku_cell()` 自动按换行/制表符拆成 N 行 PD Table，每行独立 SKU + 共享商业字段（cost / port / duty 等填重复）。Page 1 直接 1 PD Table 行 = 1 卡片。

PM 也可以选择把多色拆成独立列（每列一 SKU + 独立图，5 月 PD updates 已经这样写了 RJ64-10 4 色 + RJ15-7-LL 3 色）——也 OK，rebuild 行为一致。

不再需要任何字典或人工映射。

---

## 6. HTML 当前形态（快照，待迭代）

> 当前 HTML 形态记录如下，但 Summer 仍有不少改造想法。**本次更新只稳数据源，HTML 改造单独再聊**，详见 Todo_List。

**四页结构（2026-05-04 加 MX 拆分 + ASI filter + Placeholder 卡片）：**
- **Page 1: PD Table（Sales 选品目录）** — 卡片布局，按**合并后的 canonical category** 分组（详见下方分类规则），每段内 PD Table 真卡在前、placeholder 卡在后。**ASI 和 MP 项目不显示卡片**；**Tracker 有但 PD Table 没的 SKU 渲染成虚线占位卡片**（PENDING 标 + 黄底）。Filter: Category dropdown、PO toggle、Search box（原 For Sales/All toggle 已随 Project List 退役删除，2026-07-02）。**Compare modal（5-26 加）**：每个 category section 标题旁挂一个 `⇄ Compare` pill 按钮（Other 跳过、count < 2 跳过），点击弹出 modal + 半透明遮罩（`rgba(15,25,40,.58)` + 2px blur），modal 内是该品类的 Spec Cards 5 列横向排列，每张卡直显 **Cost / MSRP / Factory / Port / Duty / 40'HC / CRD / Sample ETA** 八个核心商业字段 + Top Feature + 图 + Status badge + Tier + Risk dot。点 Spec Card 仍打开标准详情 modal（叠在 Compare 之上，z-index 1000 vs 900）。关闭：X / 点遮罩 / Esc 三选一。Page 1 当前 filter（catF / poF / PLFilter / stat / search）全部继承到 Compare modal，placeholder 在 modal 内隐藏（8 个商业字段都是空，比较无意义）。
- **Page 2A: Pipeline US** — 横向 **11 阶段**（Kick off → MP，Inspection 合并进 MP）。**只显示非 -MX SKU**。每阶段显示项目数 + 点击下钻。**默认进 tab 自动激活 Kick off**。详情表含 SKU / Category / PM / Risk / **PO** / Next Action 列。顶栏右上角 **Type toggle**（NPD only / ASI only / All，默认 NPD only）— 切换时实时重算各阶段计数。
- **Page 2B: Pipeline MX** — 与 Pipeline US 同结构、独立的 ASI filter / 展开状态。**只显示 -MX 后缀 SKU**（严格 `endswith('-MX')`，不模糊匹配）。
- **Page 3: Weekly Tracker** — 项目进度详情，Filter: **PM / Location / PO / For (buyer) / Search**（5-04 精简，去掉 Category 和 Risk dropdown）+ 右上角 **Type toggle**（NPD only / ASI only / All，默认 NPD only）。**MP 项目仍在 Tracker 详情列表里**。

**Page 1 Category 合并规则（2026-05-04 加）：**
PD Table / Tracker 里 PM 写的 category 字段五花八门（"Microwave Oven" / "T1 Microwave" / "Water Kettle" / "Kettle (CA)"…），build.py 用 `CATEGORY_RULES` 关键字匹配统一映射到 canonical bucket。当前 19 条规则，按"具体优先"顺序匹配（如 "Air Fryer (Oven)" 必须排在 "Oven" 之前）：
- **Air Fryer (Oven)** vs **Air Fryers** 分开（Summer 2026-05-04 确认，烤箱-空气炸锅一体机和普通空气炸锅是不同品类）
- **Microwave** 收 "Microwave / Microwave Oven / Microwave (MX) / T1 Microwave"
- **Kettle** 收 "Kettle / Water Kettle / Kettle (CA) / Kettle (EU/UK)"
- **Iceman** 收 "ICEMAN Dispenser / Ice Maker / Icemaker+water dispenser / Slush Maker / T1 ICEMAN / T1 Slushy"
- **Mixer** 收 "Hand Mixer / Stand Mixer"
- **Blender** 收 "Hand Blender"
- **Pressure Cooker** 收 "MULTI-PRESSURE COOKER"
- 其它 Slow Cooker / Rice Cooker / Deep Fryer / Bread Maker / Roaster Oven / Oven / Coffee / Griddle / Vacuum / Grill / Ice Cream / Water Dispenser 按关键字归并

**< 3 个品的 category 合并为 "Other"**（Summer 2026-05-04 确认，避免太多碎片化分组）：
- 阈值 `SMALL_CAT_THRESHOLD = 3` 在 build.py main() 里，跑完 page1+placeholder 之后统计每 bucket card 数，<3 的全部 reassign 到 "Other"
- 模板里 "Other" 永远 pin 到所有 category 段的最后

**Placeholder 卡片（2026-05-04 加）：**
- Tracker 有但 PD Table 没的 SKU（且不是 ASI、不是 MP）→ 渲染虚线黄边占位卡，左上角橙色 "PENDING" 小标，卡片底部斜体 "Awaiting PM input"
- 点击 → 简化 modal，黄底 banner "Awaiting PM input. PM hasn't supplied commercial info yet" + 仅显示 Tracker 字段（Issues / Next Action / CRD）
- **计入所有 stats**（Total / High Risk / Mid Risk）；Stat number == Risk Detail Panel 行数（同一 filter）
- （历史）For Sales toggle ON 时 placeholder 仍然显示——该 toggle 已于 2026-07-02 随 Project List 退役删除，placeholder 现无条件显示

**顶栏 Stats Bar（2026-05-04 大改）：**
- 5 个浮动可点击卡片：**Total Projects / High Risk / Medium Risk / Tier 1 (CSM) / Project Released**（旧版 In MP 改名）
- 视觉：白色圆角卡片 + 顶部色条（Total/T1=深蓝渐变, High=红, Mid=橙, Released=绿），数字统一深蓝
- **新计数规则（Shine + Summer 2026-05-04 确认，5-04 晚再改一次让 stat == panel）：**
  - `Total Projects` = page1 visible cards (有 category) + ASI 非 MP（**不再等于 page1 卡片数**）。Placeholder 一旦获得 category（被合并到 "Other" 时）就计入 visible。
  - `High Risk / Medium Risk` = **Tracker 行计数**（risk 匹配 + 非 ASI + 非 MP）。**含 placeholder**（Tracker 有 PD Table 没的也算）。**与 Risk Detail Panel 同一 filter，stat 数 == panel 数**。
  - `Tier 1 (CSM)` = Tracker 全部 T1 项目（**包含 MP T1**，唯一例外）
  - `Project Released` = Tracker Current Status="MP" **或** "Inspection" 的全部 SKU 数（Pipeline 把 Inspection 合并进了 MP 阶段，所以 Project Released 也包含两种状态——保持 stat 数 = Pipeline MP 列 = Released 下拉行数）
- 交互：
  - Total → 跳 Pipeline US tab
  - High / Mid → 展开 Risk Detail Panel（SKU/PM/Status/Issues/NextAction/CRD/Location 列）
  - **Project Released** → 展开 Released Detail Panel（**SKU / PM / Category / PO Status / Buyer / CRD** 列；**不展示卡片，只看清单**——MP 不需要业务字段）
  - T1 → 筛选 Page 1 卡片
  - 同一 stat card 再点收起、点别处也收起

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

**Page1 可见性规则（2026-05-04 改，5-19 ASI 数据源变更）：**
- **PD Table 是 page1 卡片的主数据源**，但额外两个过滤：
  - **ASI 列表**（5-19 起来自 **Tracker col E "NPD/ASI"**；旧 `pd_table_config.json` `after_sales_improvement` 字段已删）→ 不显示卡片（仅 Page 2/3 显示）
  - **MP 状态**（来自 Tracker `Current Status == 'MP'`）→ 不显示卡片（计入 Project Released stat）
- PD Table 没填 tier/category 的 SKU 也不出卡片（沿用旧规则）

**PO 状态规则（2026-06-01 大改：数据源换 Tracker col N + 三态分级）：**
- **数据源换了**：page1 卡片的 PO 状态以前读 PD Table 的 "PO Placed?" 列，但那列大半空 / 偶有旧值（如 MX SKU PD Table 写 NO 但 Tracker 已下 PO），filter 不准。6-01 起 **page1 和 page3 都从 Weekly Tracker col N（PO/订单状态）读**，经 `parse_po()` 解析（与 page3 统一）。
- **`parse_po()` 返回 4 态**：`YES`（已下 PO）/ `INTENT`（意向·询单·即将下PO）/ `NO`（无PO·暂无订单·项目取消）/ `''`（col N 空白）。buyer（渠道名）独立提取。
  - 判定顺序：负向短语 `PO_NEGATIVE_PHRASES` → NO；否则提 buyer（`PO_BUYER_KEYWORDS` 或 "for X"）；含 `PO_INTENT_PHRASES=['意向','询单','即将']` → INTENT；其余实质文本（含"已下PO"/纯渠道名）→ YES。
  - **只写渠道名没写意向/询单的算 YES**（Summer 6-01 确认，如 Loblaws / Sam's / Canadian Tire / Menard's / for MX）。
- **Page 1 homepage = 2 态**：`build_page1_data` 把 INTENT 折叠进 No PO（只有确认 PO 才算 PO Placed——意向/询单不是 PO 已 place，给 Sales 看会误导）。filter toggle（PO Placed / No PO）、卡片 "PO ✓" 角标、详情面板 "PO Status (Tracker)" 全部读 `poStatus`。
- **Page 3 详细 Tracker = 3 态显示**：PO 列分 `无PO / XX意向 / XX已PO`（buyer + 状态词，`poBadge` 渲染：已PO 绿 / 意向 橙 / 无PO 灰）。PO 筛选下拉加了 **Intent** 选项（All / Placed / Intent / No PO）。Project Released 面板的 PO Status 列同样三态。
- **双语 label**：`render_template(..., po_labels=)` + template `{{PO_LABELS}}` 占位符。CN 版 `已PO / 意向 / 无PO`，EN 版 `PO Placed / Intent / No PO`；buyer 是英文渠道名两版通用。
- **6-01 May2026 结果**：page1 72 卡 → PO Placed 19 / No PO 44 / unknown 9；page3 84 行 → 已PO 31 / 意向 14 / 无PO 38 / unknown 1。

**Location 显示（2026-06-01）：** Page 3 / Risk Detail / Released 面板里的 Location "Both" 一律显示为 **"CN + US"**（实测大家看不懂 Both 是什么）。改的是 `locBadge` 文案 + Location 筛选下拉 option 文案，底层数据值仍是 'Both'（筛选不受影响）。

**Banner 规则（2026-05-04 改）：**
- 触发：某 PM 在 Tracker 上有 SKU 但 PD Table 没有（且不是 ASI、不是 MP）≥ 3 个 → banner 显示该 PM 负责的 category 列表（英文）
- 数据源变了（旧：PD Table 待确认/Gap section；新：Tracker vs PD Table diff）；其他规则不变
- 文案模板：`X categor(y/ies) currently lack complete commercial data — pending updates from related PM.`
- 不点名 PM 个人，自动从 `PM_SECTION_TO_CATEGORIES` 字典查映射
- 阈值 ≥3 是为了避免单个 SKU 缺数据触发 category-级 warning

**双语输出（4-29 加，5-04 翻译字典扩到 ~294 条）：**
- 每次跑 build 同时产出 CN（给 Shine + 国内）和 EN（给 US Sales）两份 HTML
- 翻译字典：`Monthly PD Report/translations.json`，**~294 条** PM 写的中文短语 → 英文映射
- 自动翻译字段：Page1 (currentStatus/category/crd) + Page3 (issue/nextAction/currentStatus/category/poRaw/crd) + Pipeline projects (action/category)
- 不翻译：SKU、PM 名字、buyer 名字、英文/数字、PD Table 预填的英文 description/features
- build.py 报漏译 → 对话里 Claude 翻译 → 加进 translations.json（不改 build.py，不调 Anthropic API；详见 §5.5 + memory `feedback_translation_via_chat`）

**~~OneDrive Files On-Demand bug 应对~~（历史，Cowork 时代；2026-07-02 迁本地后不再适用）：**
Cowork 沙箱读 OneDrive 云端占位文件会拿到全 0 字节（BadZipFile），当时靠 uploads fallback + "Always keep on this device" 等手段绕过。本机 Claude Code 读文件会自动触发 OneDrive hydrate，此坑消失；build.py / rebuild_pdtable.py 里的 uploads fallback 代码仍在（本机指向不存在路径，等于空操作），无害保留。

---

## 7. 产品渲染图（2026-04-30 上线，5-19 加 image-in-cell 支持）

**结论：** PM 在 `China PD updates {Mon} {Year}.xlsx` 里内嵌的产品渲染图可以**自动抽出来嵌进 HTML**，不需要单独 folder 维护。**PM 用 Excel 的"浮于单元格上"（image-over-cell）或新版"嵌入单元格"（image-in-cell）任何一种插入方式都支持**。

**机制：**

**Image-over-cell（传统）：**
- xlsx 本质是 zip 包，图片放在 `xl/media/` 里。openpyxl 通过 `ws._images` 拿到每张图的锚点（`anchor._from.col / row / colOff / rowOff`）。
- 对每张图：anchor `_from.col + 1` 找它落在哪一列，去 Row 10 取那一列的 Model（SKU）当 key。

**Image-in-cell（Excel 365 新功能，5-19 加）：**
- openpyxl `ws._images` **看不见这种图**，它走 Excel 的 rich-data 链路存储。
- build.py `_extract_image_in_cell_raw()` 自己解 xlsx zip：
  ```
  cell.vm (1-based) → metadata.xml/valueMetadata[vm-1].rc.@v
                    → rdrichvalue.xml/rv[v].<v>0</v> = LocalImageIdentifier L
                    → richValueRel.xml/rel[L].@r:id = rIdN
                    → _rels/richValueRel.xml.rels[rIdN] → media/image*.png
  ```
- 每个图都按 col 匹配 Row 10 SKU（跟 image-over-cell 同一套 col → SKU 映射）。
- 5-19 May PD updates 抽到 image-in-cell 15 张（Kettle 7 + Microwaves 1 + Grill&Hand Mixer 7 — Chris 那个新 sheet 全是 image-in-cell）。

**Zero-area 幽灵图过滤（5-19 加）：**
- 现象：PM 拖图错位 / 撤销 / 替换时，Excel 偶尔把图二进制留在 zip 里但把 anchor 矩形压缩到 0 宽 0 高（视觉上看不见）。
- bug 表现：RJ44-CB Coffee grinder 卡片显示成蓝色水桶——Coffee&Iceman col 5 有两张图，幽灵蓝水桶图在前抢占了真磨豆机图。
- 修复：build.py 在 `for img in ws._images` 循环里加判断，TwoCellAnchor 的 `_from == to`（同 col + 同 colOff，或同 row + 同 rowOff）就 skip。

**通用处理：**
- 图片用 PIL 缩到 **300×300px + JPEG q78**，base64 内嵌到 HTML 里。一张图 ~8-15KB，~60 张总共让 HTML 从 154KB 涨到 ~780KB（仍然单文件邮件可发）。
- 每月跑 `build.py` 时自动重新抽，不需要手动同步。

**文件命名约定：**
- 源文件：`Monthly PD Report/China PD updates {Mon} {Year}.xlsx`
- build.py 用 `glob('China PD updates *.xlsx')` + 按 mtime 取最新，下月 Shine 发新文件直接拖进同目录即可，不用改路径

**多 SKU / 多图同列处理（5-07 后简化，无 umbrella 字典）：**

PM 在 PD updates 里写多变体的两种方式都支持：

| 类型 | 例子 | 处理 |
|---|---|---|
| 共享 1 张图，多 SKU 写一个 cell | `RJ50-SFDAF-25D(SS) / RJ50-BFDAF-25D(BLK)` | rebuild 按换行拆 PD Table 2 行（共享图 + 共享商业字段）；图片提取时所有 SKU 共享该图 |
| 每色独立列 + 独立图 | RJ64-10 4 色（PTC/BTR/AQU/LVD 各占一列）/ RJ15-7-LL 3 色 / RJ38 系列 | 每列 PD Table 一行 + 各自独立图 |
| 单 SKU 单图 | 普通情况 | 直接映射 |

**5-07 起的简单规则：1 PD Table 行 = 1 Page 1 卡 = 1 张图（或 fallback 到 sibling 共享图）**。Page 2 / Page 3 也按 1 Tracker 行 = 1 项处理（含 P/V variant 子行）。

**括号尾标剥离 alias（image-only，仍保留）：** PD updates 里 `RJ50-SFDAF-25D(SS)` 注册时同时注册 bare 版 `RJ50-SFDAF-25D` 作 alias，让 Tracker 端 base SKU 也能命中图。仅图片字典使用，业务字段（PD Table 合并 / Tracker join 等）仍然精确匹配。

**当前覆盖率（4-30 跑出来）：** 72 个可见 page1 卡片，37 张有图（51%）。剩下 35 张是 placeholder，原因：
- Liz 在休假，水壶/微波炉系列没上传图（10+ 个）
- 部分 RJ34 / RJ07 / RJ15 系列 PM 当月没补图
- C 系列 / Aquart（US-side 项目，PM 不维护）
- 区域版本（MX / CA / EU 等通常不出独立图）

**没图的卡片** 在 template.html 里 fallback 到 `getCatIcon(category)` emoji 占位图标，不影响其他功能。

---

## 8. 构建系统（4-29 重构 + 5-04 大改 + 5-26 加 Compare modal）

**核心文件（`Monthly PD Report/`）：**
- `template.html` — 完整单文件 HTML 模板。占位符共 13 个（2026-07-04 实测核对）：
  - 数据 8 个：`{{PAGE1_DATA}}`、`{{PIPELINE_US_DATA}}`、`{{PIPELINE_MX_DATA}}`、`{{PAGE3_DATA}}`、`{{SUMMARY_STATS}}`、`{{RELEASED_DATA}}`、`{{BANNER_BLOCK}}`、`{{CRD_CHANGE_DATA}}`（6-30 加）
  - 双语 label 字典 3 个：`{{PO_LABELS}}`、`{{PA_LABELS}}`、`{{SIXA_LABELS}}`
  - 元信息 2 个：`{{REPORT_PERIOD}}`、`{{DATA_ASOF}}`
  - **5-26 加**：Compare modal CSS 块（`.compare-btn` / `.compare-overlay` / `.compare-modal*` / `.spec-grid` / `.spec-card*`）+ Compare modal HTML 容器（`#compareOverlay` / `#compareModalBody`）+ JS 三个新函数 `buildSpecCard(p,idx)` / `openCompareModal(category)` / `closeCompareModal()` + `closeModal()` 改成"只在 Compare modal 不活时才解锁 body scroll" + Escape handler 改成级联（详情 modal 优先关，再关 Compare modal）。`filterPage1()` 渲染 cat-header 那行加按钮注入（Other 跳过 / count < 2 跳过）。**不动 build.py 不动数据流**——所有字段已在 page1Data 里。
- `build.py` — HTML 构建脚本。读三个 xlsx + 抽 PD updates 图 → 应用 ASI/MP 过滤 → 渲染 5 份 JSON + banner HTML → 灌进模板 → 输出
- `rebuild_pdtable.py` — **新（5-04）** PD Table 重建脚本（详见 §5.2）。一键完成 transpose + manual_additions + 自动 Tracker 比对。
- `pd_table_config.json` — **新（5-04）** ASI / umbrella / manual_additions 配置（详见 §5.2.1）。**两个脚本都读这个文件**。
- `translations.json` — CN→EN 翻译字典（随月更持续追加；2026-07-02 时 863 条，正文他处出现的 220/294 等数字为历史时点值）

**路径处理（2026-07-02 本地化后）：**
- `BASE` 自动从脚本位置推导（`MONTHLY_DIR.parent`）
- `SCRATCH = CLAUDE_SCRATCH 环境变量，缺省用系统临时目录`（`tempfile.gettempdir()/pd_report_scratch`，原 hardcode `/tmp` 已改）
- stdout/stderr 强制 UTF-8（Windows 控制台 cp1252 遇中文会崩）
- `TRACKER_PATH` / `PDUPDATES_PATH`：glob 找项目目录内最新可读文件；Cowork 时代的 uploads fallback 代码保留但本机为空操作

**数据流：**
```
[Phase A] PD Table 重建（rebuild_pdtable.py，每月 PD updates 来时手动跑）
    pd_table_config.json + China PD updates xlsx
        ↓ load_pdupdates / apply_manual_additions
    新 Summers_Monthly_PD_Table.xlsx (覆盖旧版)
        ↓ compare_pd_vs_tracker (vs Weekly Tracker, 过滤 ASI / MP / umbrella)
    A/B/C 三段 diff（PM 邮件用）

[Phase B] HTML 构建（build.py）
    Tracker xlsx + PD Table xlsx + PD updates(图源) + config
        ↓ load_tracker / load_pd_table / extract_sku_images
    内存 dict + images
        ↓ compute_mp_set (Tracker Current Status='MP') + asi_set (Tracker col E, 5-19)
        ↓ build_page1_data (排除 ASI 和 MP) / build_page3_data (带 isASI 标签) / build_pipeline_data × 2 (US/MX 按 -MX 后缀拆，各自带 isASI 标签)
    JSON × 5（page1 + pipelineUS + pipelineMX + page3 + released）
        ↓ build_summary_stats(page1, tracker, asi_set, mp_set)
    summaryStats {total, high, mid, t1, released}
        ↓ build_released_data (MP SKU 列表，给 Project Released 下拉用)
    releasedData
        ↓ build_banner_html (Tracker - PD Table - ASI - MP, 按 PM 分组阈值 ≥3)
    banner HTML
        ↓ render_template（13 个占位符替换，清单见上方 template.html 条目）
    最终 HTML
        ↓ write_with_rotation
    CN + EN 两份输出
```

**Pipeline 11 阶段（5-04 改，Inspection 合并进 MP）：**
Kick off → Detail Design → Prototype → Tooling → FOT → EB → Culinary EB → Culinary Claims → PP → Culinary PP → MP

**Pipeline US/MX 拆分（5-04 加）：**
- `is_mx_sku(sku)` = `sku.upper().endswith('-MX')` —— 严格后缀匹配，不模糊匹配
- main() 把 tracker_rows 拆成 us_rows + mx_rows，各自跑一次 `build_pipeline_data`
- 模板里 `PipelineView(data, ids)` 是工厂函数，US 和 MX 各实例化一次，独立 ASI filter / 展开状态 / DOM
- 状态映射不到 Pipeline 阶段的 SKU（如 `RJ55-7-VN-MX / SMR-VN-MX` 状态是"—"）会被忽略，跟 US 行为一致

**ASI / NPD filter（5-04 加）：**
- `build_page3_data` 和 `build_pipeline_data` 都接 `asi_set`，每条数据加 `isASI` 字段
- Page 2 (Pipeline US + MX) 和 Page 3 (Weekly Tracker) 顶部右上角 Type toggle：`NPD only` / `ASI only` / `All`，默认 NPD only
- Pipeline 切 filter 时各阶段计数实时重算（不只是隐藏行），并自动 close 已展开的 stage 避免数据 stale
- ASI 行 / 项目带灰色 `ASI` 小标签

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
- 月份按"数据所属月份"命名（June 报告 = `_Jun2026.html`），不按运行日期。月份由 build.py 按"每月 10 号切月"规则自动推导（详见 §5.6），无需手动改常量。

**Rotation 规则：**
- 同一月份命名 family：跑第 N 次 → 当前 → 改名 `_prev`，旧 `_prev` 删掉，新内容写当前
- 跨月不影响（4 月跑产出 `_Apr2026.html`，5 月跑产出 `_May2026.html`，互不干扰）
- 4-24 那批 baseline HTML（`China-PD-Monthly-Update.html` / `index.html` / `_v20260421_*`）一律不动

**写文件路径（2026-07-02 本地化后）：**
- build.py 先写 SCRATCH（系统临时目录）再 copy 到工作目录；Cowork 时代"双跳避免 OneDrive 截断"的动机已消失，流程保留是因为无害且顺手。

**已知技术坑（2026-07-04 按本机环境重新标注）：**

*仍然有效（本机）：*
- **Excel 锁文件：** Excel 还开着时读 xlsx 会报 BadZipFile。先关 Excel 再让 build.py 读。
- **WMF 图片格式：** openpyxl 读 PD updates 时如果遇到 WMF 格式的图会 warn 并丢弃。WMF 是少数情况（绝大多数 PM 用 PNG/JPEG），目前没专门处理。
- **PIL 处理透明度：** RGBA / LA / P 模式图先在白底上 paste 一次再保存为 JPEG（JPEG 不支持 alpha），避免黑底。
- **大改 build.py 后验证语法：** `python -c "import ast; ast.parse(open(f, encoding='utf-8').read())"`，好习惯保留。

*历史（Cowork 沙箱专属，2026-07-02 迁本地后不再适用）：*
- ~~OneDrive Files On-Demand 全 0 字节 bug~~（本机自动 hydrate）
- ~~Edit / Write tool 截断大文件 / 注入 null bytes~~（Cowork 沙箱 + OneDrive 同步竞态所致）
- ~~bash heredoc 转义 `!`~~（当时被迫用 heredoc 写大文件的场景已不存在；本机直接用文件工具）

---

## 9. 当前数据源状态（2026-07-04 重写：本节只记"现在"，历史时点快照一律看 §11 变更记录）

> 旧版本节曾长期冻结在 5 月的快照（WK22 / 26 列 / May 数据），与 §1、§11 记录的 7-1 改动相互矛盾，2026-07-04 清理重写。**原则：本节不再堆每次跑完的 stats 数字**（那些数字只在当次有意义），要历史去 §11 找。

- **Weekly Tracker** — `Weekly Tracker/` 下 WK 号最大的一份。**16 列结构（2026-07-01 删 11 个阶段列后）**，列定义见 `Weekly Tracker/PM_Weekly_Tracker.md`。
- **Summers Monthly PD Table** — 每月 Shine 发新 PD updates 后用 rebuild_pdtable.py 重建（纯镜像）。
- **China PD updates {Mon} {Year}.xlsx** — 图源 + PD Table 重建源，build.py 按 mtime 自动找最新。
- **pd_table_config.json** — sku_aliases / mp_overrides / manual_additions / crd_changes（6-30 加，review-gated）。
- ~~Project list.xlsx~~ — 已退役并删除（2026-07-02）。
- **最近一次构建：** 2026-06-30，WK26 Tracker + China PD updates Jun 2026，产出 `China_PD_Monthly_Report_Jun2026(.html/_EN.html)`；7-2 迁本地后重跑 EN 0 warning（translations.json 863 条）。详见 §11。

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
| HTML 模板 | `Monthly PD Report/template.html`（13 个 `{{...}}` 占位符，清单见 §8） |
| 构建脚本 | `Monthly PD Report/build.py` |
| 翻译字典 | `Monthly PD Report/translations.json` |

---

*参考：Email_Tracking_Rules.md | Company Org Chart April 2026.pdf*


---

## 11. 变更记录

### 2026-06-30（WK26）
- **数据更新**：读 China PD updates Jun 2026 + WK26 Tracker，跑 rebuild_pdtable.py + build.py 重建中英文 HTML。A 段（Tracker 有、PD 缺、非 ASI/MP）5 个按 Summer 决定全出 PENDING 占位卡（未加 alias）。
- **风险展开页加列**：High/Medium Risk 点开的明细面板每行新增 **PO / PA / 6A** 三列（复用 poBadge/paBadge/sixABadge，中英文自动一致）。改的是 template.html 的 showRiskDetail。
- **英文翻译补全**：57 条 PM 状态文字加进 translations.json（802 → 859 条），EN 版报「OK all Chinese strings translated」。
- **新增 CRD Change 可点统计块**：顶栏第 6 个方块（深红 #8B0000，位置在 High Risk 前）。点开展开明细面板，列 SKU/PM/Status/Risk/CRD 变化(old→新)/**Delay Reason**/PO。
  - 数据存 `pd_table_config.json` 的 `crd_changes` 数组（含 reason/reasonEN 双语），**review-gated**：每周人工确认入选项才写进去。规则：只收「CRD 推迟 + 有 PO」和「高风险 + 有 PO」，提前的不收；基线用上一周 Tracker。
  - build.py 新增 `stats['crd']` + `{{CRD_CHANGE_DATA}}` 注入；EN 版把 reason 换成 reasonEN。template.html 新增 sc-crd 方块 + showCRDChangeDetail 面板。
  - WK25→WK26 首版 5 条：RJ11-18-PL-TI-V2、RJ38-3DBW-V2、RJ07-15-SS-V2-CA、RJ11-GN-BLK-V3、RJ11-18-SCTI-HP-V2。
- **踩坑**：OneDrive 把 template.html / pd_table_config.json / build.py 在编辑后同步成截断版，导致 build 出的 HTML 丢 `</script>`。修法：改文件一律走 bash python 写入 + fsync + round-trip 校验，不用 Edit 工具直接改这目录里 build 要读的文件。（此坑为 Cowork 沙箱专属，2026-07-02 迁本地后不再适用。）

### 2026-07-02（迁移到 Claude Code + Project List 退役）

- **全工作流从 Cowork 迁至本机 Claude Code**：build.py / rebuild_pdtable.py 本地化——scratch 目录改用系统临时目录（原 hardcode `/tmp`）、stdout/stderr 强制 UTF-8（Windows 控制台 cp1252 遇中文即崩，首跑就栽在这）；本机装 Pillow 12.3.0；Cowork 时代的 OneDrive Files On-Demand / 上传中转 / 双跳写文件等坑全部消失（本地读文件自动 hydrate）。
- **Project List 白名单退役**：Summer 此前已删 HTML 的 For Sales/All toggle，7-2 验证 `onProjectList` 在 template.html 零引用后，清除 build.py 死代码（`_find_latest_project_list` / `load_project_list` / main 调用），删除 `Project list_5-20-2026.xlsx`。**HTML 数据源三 → 二**（Weekly Tracker + Summers Monthly PD Table）。（2026-07-04 审计核准表述：placeholder 卡仍输出 `'onProjectList': False` 字段、`stage_label_map` 保留空壳 + 兼容循环——均为 template 零引用的无害残留，故意保留不算未清；"死代码清除"指有调用链的部分。）
- translations.json 860 → 863（补 3 条 WK26 采购/PM 状态），Jun2026 双语 HTML 重出、EN 0 warning。
- 同日外围变化（不影响本项目数据流）：Sales Tracker 重构为 `Sales FollowUp/` 活清单；周五定时扫描任务迁至 Code Scheduled Tasks。

### 2026-07-04（全文一致性清理，不改代码只改本文档）

7-2 迁移当天只做了重点章节的 spot-fix，本次把正文与 §1/§11 已记录的现实对齐：
- **Cowork 残留清理**：§5.1"Claude（沙箱）"→ 本机；§5.5 翻译不走 API 的理由更新（沙箱拦截已成历史，维持对话翻译是成本决定）；§6 OneDrive Files On-Demand 应对整段、§8 双跳写文件、§8 技术坑清单——Cowork 专属项全部标注历史，本机仍有效的（Excel 锁文件/WMF/PIL 透明度）单独留存。
- **§8 路径描述更新**为 7-2 本地化后的实际代码（系统临时目录 scratch、UTF-8 stdout、uploads fallback 本机空操作）。
- **§9 整节重写**：原节冻结在 5-26 快照（WK22 / "26 列稳定" / May 数据，与 7-1 的 16 列矛盾），改为"只记现在、历史看 §11"，不再堆单次 stats 数字。
- **占位符数目修正**：template.html 实测 13 个（数据 8 + label 3 + 元信息 2），§8/§10 原写 5/7 个均已改。translations.json 条数注明"随月更增长"，正文历史数字不再逐处追改。
