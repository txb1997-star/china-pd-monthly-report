# China PD Monthly Report

Chefman 中国 PD 团队月度产品开发进度报表的构建系统与最新成品。

**Live 地址:**
- 默认(English): https://txb1997-star.github.io/china-pd-monthly-report/
- 中文版: https://txb1997-star.github.io/china-pd-monthly-report/China_PD_Monthly_Report_Apr2026.html
- English version: https://txb1997-star.github.io/china-pd-monthly-report/China_PD_Monthly_Report_Apr2026_EN.html

---

## 这是什么

这个 repo 把每月一次的 China PD 进度报表从 Excel 数据源转成可交互、可分享的 HTML 网页。报表覆盖三页:

1. **Page 1 — SKU 卡片视图**:按 Tier / Category 分组,可筛可搜。Stats Bar 顶部数字 = 当前可见卡片数。
2. **Page 2 — PD Tracker 完整表**:含所有项目,合并 Umbrella SKU。
3. **Page 3 — 关键节点 / Launch Date 视图**。

中英双语同源构建,改一处自动两边同步。

---

## 文件结构

```
.
├── build.py                              # 构建脚本(读 xlsx → 套 template → 写 HTML)
├── template.html                         # HTML 模板(布局、样式、JS 交互)
├── translations.json                     # 中英文术语映射表
├── index.html                            # GitHub Pages 默认页(EN 版的副本)
├── China_PD_Monthly_Report_Apr2026.html  # 当月中文成品
├── China_PD_Monthly_Report_Apr2026_EN.html  # 当月英文成品
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
- 装依赖:`pip install openpyxl jinja2`
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

每次出新月度报表后:

**手动方式(任意 PowerShell)**

```powershell
cd "C:\Users\xtan\OneDrive - Chefman\Desktop\Trial\PMO General Email Tracking\Monthly PD Report"
git add .
git commit -m "Update: <月份> 月度报表"
git push
```

**自动方式**

双击同目录下的 `push_to_github.bat`,会自动 add + commit + push。

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

如有数据问题或希望加入 Sales Tracker 等其它视图,提 Issue 或直接联系 Summer。
