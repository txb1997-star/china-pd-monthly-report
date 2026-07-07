# Monthly PD Report — To Do List

*最后更新：2026-07-02*
*关联文档：Monthly_PD_Project.md*

---

## 当前阶段

工作流已迁至本机 Claude Code（2026-07-02），构建链路本地验证通过（Jun2026 双语 HTML、EN 0 warning）。数据源精简为两个（Project List 退役）。月报节奏稳定：每月 26 号 PM 确认 PD Table → 月初跑 rebuild + build → 双语出版。

---

## 立即（本周）

- [ ] **追踪 PM 回填** — 6-30 Jun 报告中 A 段（Tracker 有、PD Table 缺商业信息）5 个 SKU 出 PENDING 占位卡，等 PM 补 PD updates 后消掉
- [ ] **git commit + push 迁移改动** — build.py/rebuild_pdtable.py 本地化 + 文档更新还没进 git；push 前核实 claude_api_key.txt / 推送 Token 从未进过历史

## 紧接着

- [ ] **HTML 改造（Summer 仍有多处想改）** — 单独梳理后再做
- [ ] **index.html 更新机制待定** — GitHub Pages 首页目前是人工拷贝的最新 EN 版（现为 Jun2026），build.py 不自动更新它，每月要手动同步一次；要么让 build 自动同步最新 EN 版，要么改成月份目录页（待 Summer 定）

## 中期

- [ ] **7 月月报周期** — 7/26 前 PM 确认 PD Table；7/10 后 build 自动切 Jul（MONTH 已自动化，无需改代码）

## 未来（月报跑通后再考虑）

- [ ] **Engineering 板块方向** — 合进 HTML 还是独立页面
- [ ] **Engineering Tracker 设计** — 字段定义、数据获取方式
- [ ] **和 Merlin 协作方式** — 周报数据怎么拿

---

## 已完成（近期，详细历史见 Monthly_PD_Project.md §11）

- [x] 7-02 迁移至 Claude Code：本地构建链路跑通、Pillow 安装、路径/编码修复
- [x] 7-02 Project List 白名单退役（数据源三→二）、build.py 死代码清除
- [x] 7-02 EN 版翻译补全（863 条，0 warning）—— 原"119 条待翻译"任务随各周构建陆续消化完毕
- [x] 6-30 CRD Change 可点统计块上线（第 6 个 stat 方块 + 明细面板，review-gated crd_changes 配置）
- [x] 6-30 风险明细面板加 PO/PA/6A 三列
- [x] ~~联系 IT hosting~~ — 已用 GitHub Pages（月报）+ Azure Static Web App（看板）解决，不再找 IT
- [x] 5-26 Page 1 Compare modal 上线
- [x] 5-19 image-in-cell 支持 + 幽灵图过滤 + ASI 切 Tracker col E

---

*完成项打勾。新增任务直接加。本文件不存历史，状态推进时直接覆盖更新。*
