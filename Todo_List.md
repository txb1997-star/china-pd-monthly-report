# Monthly PD Report — To Do List

*最后更新：2026-05-19*
*关联文档：Monthly_PD_Project.md*

---

## 当前阶段

HTML 已上线（4-21 胡总确认通过），数据更新 SOP 已跑通，PM 协同节奏稳定。5-19 这次月报更新覆盖了 Tracker 新加 NPD/ASI 列、Excel image-in-cell 解析、幽灵图过滤几项关键改造。

---

## 立即（本周）

- [ ] **EN 版 119 条新中文待翻译** — May 数据这次产 EN 时 build.py warn 出 119 条新中文卡点/Action。等 Summer 决定时机后对话里翻译追加 translations.json + 重跑 build
- [ ] **追踪 PM 反馈** — Teams broadcast 已发，等 A 类 11 个 + B 类 8 个 SKU 回填 / 确认（Cottee 2 / Liz 10 / Rowling 7；Rowling RJ15-7-LL 系列命名一致性也待 Rowling 确认）

## 紧接着

- [ ] **HTML 改造（Summer 仍有多处想改）** — 单独梳理后再做，先稳数据源

## 中期

- [ ] **联系 IT hosting 方案** — 让其他人通过 link 访问（CC 胡总）
- [ ] **补齐剩余 placeholder SKU 商业数据** — 11 个 placeholder 卡片（A 类 PM 回填后会变成真卡）

## 未来（月报跑通后再考虑）

- [ ] **Engineering 板块方向** — 合进 HTML 还是独立页面
- [ ] **Engineering Tracker 设计** — 字段定义、数据获取方式
- [ ] **和 Merlin 协作方式** — 周报数据怎么拿

---

## 已完成（参考）

### 上线 + 4 月基础
- [x] HTML 三页结构上线（PD Table / Pipeline / Weekly Tracker）
- [x] 顶栏 Stats Bar 5 卡片（Total / High Risk / Mid Risk / Tier 1 / Project Released）
- [x] Risk Detail Panel（点 High/Mid Risk 展开 Tracker 风格表格）
- [x] 4-21 胡总确认 HTML 形式
- [x] CRD & Milestone 政策制定
- [x] Milestone Change 填写 Guidance 发给 PM
- [x] 三数据源结构确定（Tracker / PD Table / Project List）
- [x] 数据更新流程 SOP（Monthly_PD_Project.md §5）
- [x] 4-29 双语自动产出（CN + EN）+ Project List filter + Banner
- [x] 4-30 产品渲染图自动抽取嵌入

### 5 月改造
- [x] 5-04 PD Table 纯镜像重建 + Stats Bar 重做 + Pipeline US/MX 拆 + Category 合并 + Placeholder 卡
- [x] 5-07 Tracker 25 列适配（P/V 列）+ umbrella 字典彻底删 + MONTH_NAME → May
- [x] 5-19 Tracker 26 列适配（NPD/ASI col E）
- [x] 5-19 ASI 数据源切到 Tracker col E + 删 `pd_table_config.json` `after_sales_improvement` 字段
- [x] 5-19 build.py 加 `_extract_image_in_cell_raw()` 支持 Excel 365 image-in-cell（rich-data 链路解析）
- [x] 5-19 加 zero-area 幽灵图过滤（修 RJ44-CB 配错图 bug）
- [x] 5-19 用 WK21 + May PD updates 跑出新 PD Table + HTML

---

*完成项打勾。新增任务直接加。本文件不存历史，状态推进时直接覆盖更新。*
