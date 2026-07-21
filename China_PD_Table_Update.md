# China PD Table 更新流程

*最后更新：2026-07-04（审计整改：§5.3 Cowork 双跳流程标历史、§7.2 提醒模板删 Project List；此前 2026-05-19 ASI 数据源切到 Tracker col E）*
*负责人：Summer Tan (PMO)*
*关联文档：Monthly_PD_Project.md §5.2 / §5.2.1 / §5.2.2*

---

## 1. 概况

Summer's Monthly PD Table 是给 US Sales 看的产品商业信息汇总表（Brand、Features、Cost、MSRP 等），也是 Monthly HTML Report 的 Page 1 数据源。

**数据来源：** PM 们各自填写的 `China PD updates {月} {年}.xlsx`（由 Shine 汇总发出）
**输出文件：** `Monthly PD Report/Summers_Monthly_PD_Table.xlsx`（**每月从零重建，覆盖**）
**触发条件：** Summer 把新版 PD updates 文件放进目录或 chat 上传，告诉 Claude 跑重建
**频率：** 月度。PM 提交截止日为每月 26 号（胡总确认）

**数据流（5-04 大改后）：**

```
China PD updates {月}.xlsx（PM 填，Shine 发）+ pd_table_config.json
        ↓ 水平→纵向 transpose + 字段映射 + manual_additions 注入
Summer's Monthly PD Table（24 列，完全替换旧版，不 merge）
        ↓ 自动对比最新 Weekly Tracker
   A/B/C 三段 diff（PM 邮件用）
        ↓
   Summer 发邮件 broadcast → PMs 本周内对齐
```

**关键转变（5-04）：** 这一步不再做"基于上版 merge"，也不在这一步做 SKU 一致性比对作为输出。比对结果作为 **副产物自动生成**，但不影响 PD Table 内容——PD Table 严格等于当前 PD updates 的镜像。

---

## 2. 源文件结构（China PD updates）

### 2.1 Sheet 列表（9 个品类 Sheet）

| Sheet 名 | PM | 典型 SKU 数 |
|-----------|-----|-------------|
| Kettle | Liz Liu | 14 |
| Air Fryers | Cottee Wei | 7 |
| Microwaves | Liz Liu | 3 |
| Coffee&Iceman | Serena Sun | 7 |
| Rice Cooker | Liz Liu / Serena / Rowling | 8 |
| Juicer | TBD | 1 |
| OVEN&Bread maker&Deep fryer&Ric | Rowling Luo | 6 |
| Roaster ovn&Waffle maker | Rowling Luo | 5 |
| Sourcing | Chris Zhou | 4 |

**注意：** Sheet 名和 SKU 数量会随 PM 更新变化，每次都要动态遍历所有 Sheet。

### 2.2 每个 Sheet 的布局

**水平布局：** B 列是字段名，C 列起每一列是一个产品。

**关键行位置（所有 Sheet 统一）：**

| 行号 | 内容 |
|------|------|
| Row 2 | Category |
| Row 3 | Project Manager |
| Row 4 | Tier |
| Row 5 | Initial Market |
| Row 6 | Factory |
| Row 7 | Sales Sample(s) ETA |
| Row 8 | Image（跳过，内嵌图无法读取） |
| Row 9 | Brand |
| **Row 10** | **Model（= SKU，join key）** |
| Row 11 | Description |
| Row 12 | MSRP 或 PO Placed?（视 Sheet 而定） |
| Row 13+ | 其余字段（Cost、Port、Duty、Features 等） |

**⚠️ Row 12 以下字段顺序因 Sheet 而异。** 有的 Sheet 在 Row 12 放 MSRP，有的放 PO Placed?，有的有 "Project stage" 行。**必须按 B 列的 label 文字做映射，不能硬编码行号。**

### 2.3 特殊 SKU 格式

有些单元格里包含多个 SKU，用换行符分隔：
- `RJ50-SFDAF-25D(SS)\nRJ50-BFDAF-25D(BLK)` — 同一产品的两个颜色变体，共享同一列的数据
- `RJ64-10-PTC    \tPistachio\nRJ64-10-BTR\t    Butter\n...` — SKU + tab + 颜色描述

处理方式：按 `\n` split，每行取 tab 前的部分作为 SKU，所有变体共享该列的商业数据。

---

## 3. 字段映射

### 3.1 映射表（PD updates → Summer's Monthly PD Table）

| PD updates 字段（B 列 label） | PD Table 列号 | PD Table 列名 | 备注 |
|-------------------------------|--------------|---------------|------|
| Model | 1 | SKU | join key |
| Category | 2 | Category | |
| Tier | 3 | Tier | |
| Brand | 4 | Brand | |
| Description | 5 | Description | |
| Top Feature | 6 | Top Feature | |
| Unique Feature（第 1 个） | 7 | Unique Feature 1 | 按出现顺序 |
| Unique Feature（第 2 个） | 8 | Unique Feature 2 | |
| Unique Feature（第 3 个） | 9 | Unique Feature 3 | |
| MSRP | 10 | MSRP | |
| Sales Sample(s) ETA | 11 | Sales Sample ETA | |
| PO Placed? | 12 | PO Placed? | 不是每个 Sheet 都有 |
| Estimated 1st Inspection | 13 | Est. 1st Inspection | |
| Factory | 14 | Factory | |
| Initial Market | 15 | Initial Market | |
| 1st Cost Estimate | 16 | 1st Cost Estimate | **加 $ 前缀** |
| Buffer Addt'l | 17 | Buffer Addt'l | |
| Port | 18 | Port | |
| Duty (into US) | 19 | Duty | |
| 40'HC Estimate | 20 | 40'HC | |
| Key Competitive Model | 21 | Key Competitive Model | |
| Key RJ Brands Difference | 22 | Key RJ Brands Difference | |
| Note (1) | 23 | Note 1 | |
| Note (2) | 24 | Note 2 | |
| Project Manager | — | 不进表 | 仅用于 PM 分组 |
| Image | — | 跳过 | 内嵌图无法读取 |

### 3.2 数据清洗规则

- **1st Cost Estimate：** 非空裸数字加 `$` 前缀（`12.50` → `$12.50`），已有 `$` 的不重复加
- **SKU：** 去掉尾部中文（`RJ38-G4 玻璃碗` → `RJ38-G4`），trim 空格
- **日期：** Short Date 格式（MM/DD/YYYY），模糊日期如 "2026 April" 保持原样

---

## 4. SKU 匹配规则

### 4.1 核心原则：绝对不做模糊匹配

SKU 后缀有业务含义，即使只差一两个字母也可能是完全不同的产品变体：

| 后缀 | 含义 |
|------|------|
| SS | Stainless Steel（不锈钢材质/颜色） |
| BLK / WHT | Black / White（颜色） |
| CA | Canada 市场 |
| CO | Costco 渠道 |
| MX | Mexico 市场 |
| EU / UK | 欧洲 / 英国 |
| AM | Amazon 渠道 |
| D / M | Digital / Mechanical（数字/机械控制） |
| V2 / V3 | 版本迭代 |
| HP | 升级版本（如壶嘴壶盖升级 SS） |
| PL | 塑料材质 |

**匹配流程（5-04 改）：**

PD Table 重建阶段**不再做 PD vs Tracker 交叉匹配**——这一步只是单纯把 PD updates 镜像成纵向 24 列。一致性比对作为副产物在重建脚本最后跑（详见 §6）。

阶段内的 SKU 处理只有两条规则：
1. **精确镜像**：PD updates 里有的 SKU（含隐藏列）直接转入 PD Table；TBD/TBC 占位符跳过；多 SKU 单元格按 \\n 拆分成多行
2. **manual_additions**：从 `pd_table_config.json` 注入额外行（PM 还没填、Summer 已经有 info 的项目）

### 4.2 已确认的 SKU 对应关系（仅供 reference，**不用于自动恢复**）

以下是历次 PM 确认过的 SKU 含义和命名映射，留作下次同名 SKU 出现时的参考。**不要用这些信息把上月已删除的 SKU 自动加回来**——PM 删了就是删了。

**Rowling Luo：**
- RJ50-SFDAF-25D(SS) / RJ50-BFDAF-25D(BLK) 是 RJ50-SFDAF-25D 的两个颜色变体，都是新行
- RJ34-10C-M-V3、RJ34-16C-M、RJ34-2C-M、RJ34-6C-M、RJ34-12C-M 是 Rice Cooker M（Mechanical）系列，与 D（Digital）系列并存
- RJ07-32-SS 按 Weekly Tracker 数据更新（Summer 确认）

**Serena Sun：**
- RJ62-BLK / RJ62-WHT 是 RJ62-20A-Series 的 gen1/gen2 颜色变体
- RJ64-20 — on hold（Serena 确认）
- RJ64-10-PTC / BTR / LVD / Aqua 是 RJ64-10-new colors 的具体颜色变体名
- **ICM1239X = RJ56-BUL-12-V2 同一产品**（Summer 2026-07-07 确认，Tracker 已删 ICM1239X 行只留 BUL）——今后任何数据源出现 ICM1239X 一律映射到 RJ56-BUL-12-V2 再问 Summer 确认

**Tammy / Chris Zhou：**
- RJ59-HNC-MX ≠ RJ59-HNC-V2-MX — Tammy 确认是两个不同产品，不能合并
- RJ40-8（Sourcing sheet）vs RJ40-8-MX（Draft）— 不同市场版本

**Cottee Wei：**
- RJ38-2D-AM 是 RJ38-2D-V2 的 Amazon 渠道版本，独立新行

**Liz Liu（写法颠倒惯犯，2026-07-07 两例）：**
- **RJ11-SS-12 / RJ11-SSD-12（PD updates 写法）= RJ11-12-SS / RJ11-12-SSD（Tracker 写法）**——Liz 曾说不是、Summer 判断是并按同一产品处理（sku_aliases 已配 `RJ11-12-SS→RJ11-SS-12`；SSD 走 `pd_exclude_skus` 排除了 PD 侧重复旧列、保留 Tracker 写法列）。若 Liz 反转结论，删 alias/exclusion 重跑即可
- Liz 在 PD updates 里可能同一产品填两列（一列新数据一列复制的旧壳）——重建后 diff 时看到她的"新 SKU"先怀疑是写法变体或重复列

**Liz Liu（2026-04 休假中，10 个 SKU 待确认）：**
- RJ11-12-SSTI-D、RJ11-15-SSD、RJ11-12-SCTI、RJ11-17-CTI-DG、RJ11-18-CTI-HP-V3、RJ11-GN-BLK-V2、RJ11-GN-BLK-AM、RJ11-12-SS-TI-MX、RJ55-7-VN-MX、RJ55-7-SMR-VN-MX
- 这些都是 PD Table 里有但与 Tracker SKU 写法不同的项目，等 Liz 回来确认

**已删除：**
- C56-Nugget (Welly) — 确认走 Aquart，Welly 版删除（2026-04-29）

---

## 5. openpyxl 注意事项

### 5.1 水平→纵向 Transpose 方法

```python
# 读取源文件时：
# 1. 遍历每个 Sheet
# 2. 扫 B 列建立 row_number → field_label 映射
# 3. 按 field_label（不是行号）做字段映射
# 4. 从 C 列起，每列提取一个产品的全部字段
# 5. Row 10 (Model) 是 SKU，作为 join key

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    row_labels = {}
    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 2).value
        if label:
            row_labels[r] = str(label).strip()
    
    # 找 Model row（不要硬编码 row 10）
    model_row = None
    for r, label in row_labels.items():
        if label.lower() in ('model', 'model no.', 'model no'):
            model_row = r
            break
```

### 5.2 合并单元格处理

**问题：** Draft 模板的 PM header 行（如 "Cottee Wei — 空气炸锅 + T1 项目"）可能是合并单元格，openpyxl 读取合并区域中非首格会返回 MergedCell 对象，直接写入会报 `'MergedCell' object attribute 'value' is read-only`。

**解决方案（与 PM_Weekly_Tracker.md 一致）：**
1. **方案 A（推荐）：** 创建全新 Workbook，从零写入所有数据和样式，不继承 Draft 的合并单元格
2. **方案 B：** 如果要基于 Draft 修改，先 `ws.unmerge_cells()` 解除所有合并，操作完再重新合并

### 5.3 文件损坏预防（2026-07-04 更新：Cowork 时代的坑已消失）

**历史坑（Cowork 沙箱专属，2026-07-02 迁本地后不再适用）：** 沙箱直接写 OneDrive 挂载路径可能截断文件（zip EOCD 丢失 → `BadZipFile`），当时靠"先写 `/sessions/.../mnt/outputs/` 再 shutil.copy 回工作目录"的双跳规避。

**现行做法（本机）：** 先写系统临时目录（`tempfile.gettempdir()` 下的 scratch）再 copy 到工作目录——rebuild_pdtable.py / build.py 已内置，无需手动操作。写完 reload 验证能正常打开即可。

### 5.4 格式保护

- **只写入 cell 值，不修改已有格式**（font, fill, alignment, column width）
- 如果是全新 Workbook，手动设置全表统一格式：
  - 字体：Century Gothic, 10pt
  - PM header 行：Century Gothic, 10pt, Bold, 白字蓝底 (FF4472C4)
  - 新增行：浅黄底 (FFFFF2CC)
  - 待确认行：黄底 (FFFFFF00)
  - Gap analysis header：白字红底 (FFC00000)
  - 对齐：wrap_text=True, vertical='top'
  - 日期列：Short Date

---

## 6. Tracker 对比（5-04 替代旧版 Gap Analysis）

**目的：** PD Table 重建完成后**自动**对比 PD Table SKU 集和最新 Weekly Tracker SKU 集，列出差异给 Summer 一次性 broadcast 给所有 PM。

**为什么改：** 旧版 Gap Analysis 把比对结果写到 PD Table 文件底部（红底 section），并在 §5.2 重建过程中混着做。新版剥开两件事——PD Table 严格等于 PD updates 镜像，比对作为副产物输出到对话/控制台不污染 xlsx，更清爽。

**自动逻辑（在 `rebuild_pdtable.py` 末尾跑）：**

```
PD Table SKUs ≡ Tracker SKUs，差异分三段：
  A) Tracker 有但 PD Table 没有：过滤 ASI 和 MP 后 → PM 需补 business info
  B) PD Table 有但 Tracker 没有：状态待 PM 明确（停了就删 PD updates，没停就加进 Tracker）
  C) Tracker 已 MP（Project Released）：合规情况，PD Table 不需要这些 SKU
```

**ASI 数据源（5-19 改）：** ASI 集从 Tracker col E "NPD/ASI" 实时计算，不再读 config。`compare_pd_vs_tracker` 里 `asi = {sku for sku, info in tracker_skus.items() if info[3] == 'ASI'}`。

**SKU rename / canonical 化处理：** 通过 `pd_table_config.json` 的 `sku_aliases` 字段做归一化，diff 不会因为命名差异而误报（如 `RJ38-G4-AS` → `RJ38-G4-LS`、`RJ38-9TW-V3` → `RJ38-9TW-V2`、`RJ56-DIS-V2-CA-CO` → `RJ56-DIS-V2`）。5-07 删了 `umbrella_to_variants` 字段（PD updates 端 PM 直接按列 / 按 cell 拆多色，rebuild 自动处理）。

**Banner 触发：** 上面 A 段如果某 PM ≥ 3 个 SKU 缺失，HTML build 时 banner 自动显示该 PM 负责的 category 列表（详见 `Monthly_PD_Project.md` §6 Banner 规则）。

**对账规则（6-15 Summer 定，A/B 两段的处置口径）：**

- **B 段（PD updates 有、Tracker 没有）= 可接受，不报警。** 直接在 card 页显示商业信息即可，不要求 PM 把它放进 Weekly Tracker，也不需要加 `sku_aliases` 强行映射。PD updates 端可以比 Tracker 多（如新一代 V2、color variant、PD 先行的新品）。
- **A 段（Tracker 有、且"未 MP"也"不是 ASI"、但 PD updates 没有）= 必须报警。** 这类才是真正的缺口（PM 在 Tracker 立了项却没在 PD updates 补商业信息）→ 渲染为 PENDING 占位卡片 + 触发 banner。
- 一句话：**方向性不对称** —— 多在 PD updates 这边无所谓，缺在 PD updates 这边（而 Tracker 有进行中的非 ASI 项目）才要喊。

---

## 7. PM 沟通模板

所有发给 PM 的消息**必须写中文**（PM 英文不好）。

### 7.1 月度对齐邮件（推荐：5-04 改后的标准模板）

5-04 起 Summer 直接用一封 broadcast 邮件给 5 位 PM + Shine CC，贴上每个 PM 的 A/B 两类 SKU。模板见 `China_PD_Table_Update_Email_20260504.md`（Summer 自己留底，Claude 不主动维护这个文件）。

核心结构：

```
主题：请本周内对齐 PD Table 与 Weekly Tracker —— Monthly PD Report 数据同步

各位 PM，

[原则陈述：Tracker 上的 SKU 必须在 PD Table 也有；反之亦然]

[每个 PM 的 A 类 / B 类列表]

A 类（Tracker 有 PD Table 没有）→ 下次 PD updates 补 business info
B 类（PD Table 有 Tracker 没有）→ 项目停了就删 PD updates，没停就加进 Tracker

deadline: 本周五下班前
```

### 7.2 月度提交提醒（每月 26 号前）

```
Hi 各位 PM，

提醒一下，每月 26 号前请确认 China PD updates 已更新到最新——
所有在研项目的商业信息（Description、Features、Cost 等）。

特别注意：PD updates 里的 SKU 请与 Weekly Tracker 保持一致。

谢谢配合！
```

> （2026-07-04 改：原模板第 2 条"Project List"已随其 2026-07-02 退役删除。）

---

## 8. 输出文件结构

### 8.1 PD Table 24 列

| 列号 | 列名 | 宽度 |
|------|------|------|
| 1 | SKU | 30 |
| 2 | Category | 20 |
| 3 | Tier | 6 |
| 4 | Brand | 12 |
| 5 | Description | 40 |
| 6 | Top Feature | 35 |
| 7 | Unique Feature 1 | 25 |
| 8 | Unique Feature 2 | 25 |
| 9 | Unique Feature 3 | 25 |
| 10 | MSRP | 10 |
| 11 | Sales Sample ETA | 18 |
| 12 | PO Placed? | 12 |
| 13 | Est. 1st Inspection | 18 |
| 14 | Factory | 15 |
| 15 | Initial Market | 15 |
| 16 | 1st Cost Estimate | 15 |
| 17 | Buffer Addt'l | 12 |
| 18 | Port | 10 |
| 19 | Duty | 10 |
| 20 | 40'HC | 10 |
| 21 | Key Competitive Model | 25 |
| 22 | Key RJ Brands Difference | 25 |
| 23 | Note 1 | 25 |
| 24 | Note 2 | 25 |

### 8.2 行结构

```
Row 1:  Header（Bold）
Row 2:  Cottee Wei — 空气炸锅 + T1 项目（PM header，蓝底白字 Bold）
Row 3+: Cottee 的 SKU 行（正常字体）
        ↳ 新增行标浅黄底
Row N:  Rowling Luo — 烤箱 / 面包机 / 饭煲 / 慢炖锅 / 油炸锅
Row N+: Rowling 的 SKU 行
...（Serena → Chris → Liz 同理）
Row X:  ▼ Gap Analysis（红底白字）
Row X+: 按 PM 分组的缺失 SKU
Row Y:  ⚠️ 待确认（橙底白字）
Row Y+: SKU 写法不一致需 PM 确认的行（黄底）
```

### 8.3 PM 分组顺序（固定）

1. Cottee Wei — 空气炸锅 + T1 项目
2. Rowling Luo — 烤箱 / 面包机 / 饭煲 / 慢炖锅 / 油炸锅
3. Serena Sun — ICEMAN / 咖啡 / 冰淇淋
4. Chris Zhou — 烤盘 / 搅拌类 + MX 项目
5. Liz Liu — 水壶 + 微波炉

---

## 9. 完整更新 Checklist（5-04 简化）

旧版 11 步 Checklist 已合并进 `rebuild_pdtable.py` 里，每次跑只需要：

1. ☐ Summer 把新版 `China PD updates {Mon} {Year}.xlsx` 拖进 `Monthly PD Report/`（或给出 SharePoint 路径由 Claude 拷入，旧版移入 Archive）
2. ☐ 跟 Claude 说"重建 PD Table"
3. ☐ Claude 跑 `python3 rebuild_pdtable.py`：
    - **脚本自动**把旧版 PD Table 备份到 `Archive/`（2026-07-07 起写进脚本；此前靠人记得，06-30 / 07-07 两次都漏，遂固化）
    - 读 config + PD updates → 输出新 PD Table（覆盖旧版）
    - 自动比对 Tracker，输出 A/B/C 三段 diff
4. ☐ **硬关卡 ✋（2026-07-07 Summer 重申，Cowork 时代原有、迁本地后一度被跳过）：A/B 两段 diff 按 PM 分组列给 Summer，她逐项 confirm（加 alias / mp_override / 删 manual_addition / 保持占位 / 不管）之前，不得跑 build.py 生成 HTML。**
5. ☐ Summer 顺带决定是否给 PM 发 broadcast 邮件（用 §7.1 模板）
6. ☐ Summer confirm 后 Claude 落配置、跑 build.py（进入 Monthly_PD_Project.md §5.1 第 5 步起的流程）

**完成判定：** 新 `Summers_Monthly_PD_Table.xlsx` 已写入 + diff 已展示 + **Summer 已逐项 confirm**。**不再需要 PM 单独确认才能转正**——PD Table 就是 PD updates 的直接镜像。

---

*本文件记录 China PD Table 的更新流程和实操细节，高层 SOP 见 Monthly_PD_Project.md §5.2 / §5.2.1*
