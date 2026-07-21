# China PD Monthly Report

Chefman 中国 PD 团队月度产品开发进度报表的构建系统与最新成品。

**Live 地址:**
- 默认(English, 永远是最新一期): https://txb1997-star.github.io/china-pd-monthly-report/
- 当月单独文件: `.../China_PD_Monthly_Report_{Mon}{Year}.html`(中文) 和 `..._EN.html`(英文),如 `China_PD_Monthly_Report_Jun2026.html`。旧月份从 repo 移除,需要时在 git 历史里找。

---

## 这是什么

这个 repo 把每月一次的 China PD 进度报表从 Excel 数据源转成可交互、可分享的 HTML 网页。报表覆盖四页:

1. **Page 1 — SKU 卡片视图**:按 Category 分组的产品卡片(含渲染图),可筛可搜,每品类带 ⇄ Compare 横向对比浮层。顶部 Stats Bar 五个可点统计块(Total / CRD Change / **PA·6A 未完成** / High Risk / Medium Risk,2026-07-07 版;Tier 1 被 PA/6A 取代、Project Released 暂时下线可一行恢复),各有独立计数口径(不等于可见卡片数)。
2. **Page 2A — Pipeline US**:11 阶段横向管线(Kick off → MP),按 Current Status 分桶,点击下钻;NPD/ASI 切换。
3. **Page 2B — Pipeline MX**:同结构,只收 `-MX` 后缀 SKU。
4. **Page 3 — Weekly Tracker 明细表**:全项目行级视图,按 PM / Location / PO / Buyer 筛选。右上 **⇓ Export CSV** 按钮(2026-07-14 加,Moshi 需求)导出当前筛选后的行为 CSV(UTF-8 BOM,Excel 直接打开不乱码),文件名含日期和行数。

中英双语同源构建,改一处自动两边同步。月份自动按"每月 10 号切月"规则推导,无需改代码。

---

## 文件结构

```
.
├── build.py                              # 构建脚本(读 xlsx → 套 template → 写 HTML)
├── template.html                         # HTML 模板(布局、样式、JS 交互)
├── translations.json                     # 中英文术语映射表
├── index.html                            # GitHub Pages 默认页(最新 EN 版的副本)
├── China_PD_Monthly_Report_{Mon}{Year}.html     # 当月中文成品(如 _Jun2026.html)
├── China_PD_Monthly_Report_{Mon}{Year}_EN.html  # 当月英文成品
├── Monthly_PD_Project.md                 # 项目说明 / 决策记录
├── China_PD_Table_Update.md              # 报表更新注意事项
├── Todo_List.md                          # 进度清单
├── .gitignore                            # 排除数据源、API key 等
└── README.md                             # 本文件
```

---

## 怎么跑 build.py(本地刷新报表)

### 前置

- Python 3.9+
- 装依赖:`pip install openpyxl pillow`(pillow 用于产品图抽取;不需要 jinja2,模板是自研占位符替换)
- 数据源 xlsx(`China PD updates *.xlsx` 等)放在同目录下,**不会上 repo**

### 命令

```bash
cd "Monthly PD Report"
python build.py
```

跑完会生成:
- `China_PD_Monthly_Report_<month><year>.html`(中文)
- `China_PD_Monthly_Report_<month><year>_EN.html`(英文)

### 改翻译

编辑 `translations.json`,key 是中文,value 是英文。改完重跑 `build.py` 两个 HTML 都会更新。

---

## 怎么 push 最新版到 GitHub

> **分工（2026-07-07 定）：push 由 Summer 本人执行。** Claude 每次更新只负责到"本地文件就绪"（HTML + index.html 同步）为止，不主动 push。

每次出新月度报表后:

**手动方式(任意 PowerShell)**

```powershell
cd "C:\Users\xtan\OneDrive - Chefman\Desktop\Trial\PMO General Email Tracking\Monthly PD Report"
git add .
git commit -m "Update: <月份> 月度报表"
git push
```

**自动方式**

双击同目录下的 `Push_NOW.bat`,会自动 add + commit + push。

---

## 排除清单(永远不要 push)

- `claude_api_key.txt` — Anthropic API key
- 所有 `.xlsx`(内部 PD 数据)
- `*_prev.html`(上一版备份)
- `__pycache__/`

详见 `.gitignore`。

---

## 维护人

[@txb1997-star](https://github.com/txb1997-star) — Summer Tan, Chefman PMO

如有数据问题或新视图需求,提 Issue 或直接联系 Summer。
