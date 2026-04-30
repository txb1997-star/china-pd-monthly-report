# China PD Table 更新流程

*最后更新：2026-04-29*
*负责人：Summer Tan (PMO)*
*关联文档：Monthly_PD_Project.md §5.2*

---

## 1. 概况

Summer's Monthly PD Table 是给 US Sales 看的产品商业信息汇总表（Brand、Features、Cost、MSRP 等），也是 Monthly HTML Report 的 Page 1 数据源。

**数据来源：** PM 们各自填写的 `China PD updates {月} {年}.xlsx`（由 Shine 汇总发出）
**输出文件：** `Monthly PD Report/Summers_Monthly_PD_Table.xlsx`
**触发条件：** Summer 把新版 PD updates 文件放进目录，告诉 Claude 路径
**频率：** 月度。PM 提交截止日为每月 26 号（胡总确认，2026-04-30 邮件通知全体 PM + Shine）

**数据流：**

```
China PD updates {月}.xlsx（PM 填，Shine 发）
        ↓ 水平→纵向 transpose + 字段映射
Summer's Monthly PD Table（24 列纵向格式）
        ↓ SKU 交叉验证
   与 Weekly Tracker 做 gap analysis
        ↓
   输出 FINAL + diff 摘要 + PM 消息 draft
```

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

**匹配流程：**
1. **精确匹配** → 安全，直接用 PD 数据覆盖 Draft 对应行
2. **SKU 不一致** → 标黄，列出所有不一致对按 PM 分组给 Summer，由 Summer 发消息给 PM 确认
3. **PD 有、Tracker 没有** → 新增行，标注 "PD 新增，Tracker 暂无"
4. **Tracker 有、PD 没有** → 进 gap analysis 列表（见 §6）

### 4.2 已确认的 SKU 对应关系（历史记录）

以下是 2026-04 首次跑 §5.2 时 PM 确认的结果，供后续更新参考：

**Rowling Luo：**
- RJ50-SFDAF-25D(SS) / RJ50-BFDAF-25D(BLK) 是 RJ50-SFDAF-25D 的两个颜色变体，都是新行
- RJ34-10C-M-V3、RJ34-16C-M、RJ34-2C-M、RJ34-6C-M、RJ34-12C-M 是 Rice Cooker M（Mechanical）系列，与 D（Digital）系列并存
- RJ07-32-SS 按 Weekly Tracker 数据更新（Summer 确认）

**Serena Sun：**
- RJ62-BLK / RJ62-WHT 是 RJ62-20A-Series 的 gen1/gen2 颜色变体
- RJ64-20 — on hold（Serena 确认）
- RJ64-10-PTC / BTR / LVD / Aqua 是 RJ64-10-new colors 的具体颜色变体名

**Tammy / Chris Zhou：**
- RJ59-HNC-MX ≠ RJ59-HNC-V2-MX — Tammy 确认是两个不同产品，不能合并
- RJ40-8（Sourcing sheet）vs RJ40-8-MX（Draft）— 不同市场版本

**Cottee Wei：**
- RJ38-2D-AM 是 RJ38-2D-V2 的 Amazon 渠道版本，独立新行

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

### 5.3 文件损坏预防

**已知坑：** 直接写入 OneDrive 挂载路径可能导致文件截断（zip EOCD 丢失），表现为 `BadZipFile: File is not a zip file`。

**正确流程：**
```python
# 1. 先存到 outputs 目录
wb.save("/sessions/.../mnt/outputs/FINAL_new.xlsx")

# 2. 用 shutil.copy 复制到工作目录
shutil.copy(
    "/sessions/.../mnt/outputs/FINAL_new.xlsx",
    "/sessions/.../mnt/PMO General Email Tracking/Monthly PD Report/Summers_Monthly_PD_Table_FINAL.xlsx"
)
```

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

## 6. Gap Analysis 流程

**目的：** 找出 Weekly Tracker 上有的活跃项目但 PD Table 里没有商业信息的 SKU，让 PM 补充。

**为什么需要：** Weekly Tracker 是项目进度的 ground truth，如果 Tracker 有某个 SKU 说明 PM 正在做这个项目。但 PM 可能没把它填进 PD Table，导致 Sales 在 Monthly Report 里看不到这个产品的商业信息。

**步骤：**
1. 收集当前 PD Table 所有 SKU（包括已更新 + 新增的）
2. 收集 Weekly Tracker 所有 SKU
3. 精确比对，找出 Tracker 有但 PD Table 没有的
4. 按 PM 分组，标注品类和风险等级
5. 写入 FINAL 文件底部的 Gap Analysis section（红色 header）
6. 生成 PM 消息 draft 给 Summer

**输出格式（在 FINAL 文件底部）：**
- 红底白字 header：`▼ Gap Analysis: Tracker 有但 PD Table 无商业信息（需 PM 补充）`
- 按 PM 分组子 header：浅蓝底 (FFD9E2F3)
- 每个缺失 SKU 一行：浅黄底，Note 2 标注 "来自 Weekly Tracker WK{xx}"

---

## 7. PM 沟通模板

所有发给 PM 的消息**必须写中文**（PM 英文不好）。

### 7.1 SKU 不一致确认

```
Hi {PM名}，

我在整理 Monthly PD Report 的时候发现以下 SKU 在你的 PD Table 和 Weekly Tracker 里写法不一致，
想跟你确认一下：

{按 SKU 列表}

- PD Table 里的 SKU: xxx
- Weekly Tracker 里的 SKU: yyy
→ 这两个是同一个产品吗？还是不同的产品？

麻烦确认一下，谢谢！
```

### 7.2 Gap Analysis — 商业信息缺失

```
Hi {PM名}，

以下项目在 Weekly Tracker 里有记录，但 PD Table（China PD updates）里还没有填写商业信息
（Description、Features、Cost、MSRP 等），Sales 那边需要这些信息做选品参考。

{SKU 列表 + 品类}

麻烦在下次更新 PD Table 时把这几个补上，谢谢！
```

### 7.3 月度提交提醒（每月 26 号前）

```
Hi 各位 PM，

提醒一下，每月 26 号前请确认以下两个文件已更新到最新：

1. China PD updates — 所有在研项目的商业信息（Description、Features、Cost 等）
2. Project List — China Projects sheet 的项目清单

特别注意：PD Table 和 Project List 里的 SKU 请与 Weekly Tracker 保持一致。

谢谢配合！
```

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

## 9. 完整更新 Checklist

每次跑 §5.2 时按此顺序执行：

1. ☐ 读取新版 China PD updates，确认 Sheet 数量和 SKU 总数
2. ☐ 读取当前 Draft / 上次 FINAL（作为格式和现有数据基底）
3. ☐ 读取最新 Weekly Tracker（用于交叉验证和 gap analysis）
4. ☐ 逐 Sheet transpose + 字段映射 + 数据清洗
5. ☐ 精确 SKU 匹配，覆盖已有行的数据
6. ☐ 不一致 SKU 标黄，列出待确认清单
7. ☐ PD 新增（Tracker 暂无）加到 PM section 尾部
8. ☐ Gap analysis：Tracker 有但 PD 无的 SKU 列出
9. ☐ 输出 FINAL + diff 摘要 + PM 消息 draft
10. ☐ Summer 审核 → PM 确认 → 合并最终版
11. ☐ 正式版替换 `Summers_Monthly_PD_Table_FINAL.xlsx`

---

*本文件记录 China PD Table 的更新流程和实操细节，高层 SOP 见 Monthly_PD_Project.md §5.2*
