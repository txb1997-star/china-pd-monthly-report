# Monthly PD Report — To Do List

*最后更新：2026-04-28*
*关联文档：Monthly_PD_Project.md*

---

## 当前阶段

HTML 已上线（4-21 胡总确认通过），重点是**把数据更新跑成固定流程**，然后再回头改 HTML 呈现。

---

## 立即（本周）

- [ ] **跑通 PD Table 更新 SOP** — 用 `China PD updates Apr 2026.xlsx` 测一遍 Monthly_PD_Project.md §5.2 流程，验证转置 + 字段映射 + 清洗 + diff 摘要能否一次跑通
- [ ] **Summer's Monthly PD Table draft → 正式版** — 去 draft 后缀，重命名 `Summer_Monthly_PD_Table.xlsx`
- [ ] **删除冗余 MD 文件** — 经 Summer 确认后删 `Monthly_PD_Report.md` 和 `HTML_Update_Process.md`（内容已折叠进 `Monthly_PD_Project.md`）

## 紧接着

- [ ] **HTML build 脚本改造** — 从 3 个独立文件读（Tracker / Summer's PD Table / Project List），不再依赖测试版双 Sheet 结构
- [ ] **Project List 默认 Filter 上 HTML** — Home page 默认只显示 Project List 上的项目，可 toggle 看全部
- [ ] **HTML 改造（Summer 仍有多处想改）** — 单独梳理后再做，先稳数据源

## 中期

- [ ] **Image folder 建立** — 位置 + 命名约定 Summer 定，HTML 改造为按路径读图（替代当前占位图）
- [ ] **联系 IT hosting 方案** — 让其他人通过 link 访问（CC 胡总）
- [ ] **补齐 15 个未匹配 SKU 商业数据** — 主要是 T1 项目 + 部分 Chris/Liz 项目，逐步补

## 未来（月报跑通后再考虑）

- [ ] **Engineering 板块方向** — 合进 HTML 还是独立页面
- [ ] **Engineering Tracker 设计** — 字段定义、数据获取方式
- [ ] **和 Merlin 协作方式** — 周报数据怎么拿

---

## 已完成（参考）

- [x] HTML 三页结构上线（PD Table / Pipeline / Weekly Tracker）
- [x] 顶栏 Stats Bar 5 卡片（Total / High Risk / Mid Risk / Tier 1 / In MP）
- [x] Risk Detail Panel（点 High/Mid Risk 展开 Tracker 风格表格）
- [x] 4-21 胡总确认 HTML 形式
- [x] CRD & Milestone 政策制定
- [x] Milestone Change 填写 Guidance 发给 PM
- [x] Sheet 2 从测试版 Tracker 拆出 → Summer's Monthly PD Table draft
- [x] 三数据源结构确定（Tracker / PD Table / Project List）
- [x] 数据更新流程 SOP 起草（Monthly_PD_Project.md §5）

---

*完成项打勾。新增任务直接加。本文件不存历史，状态推进时直接覆盖更新。*
