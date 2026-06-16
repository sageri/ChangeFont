# AGENTS.md — Font Unifier

本文件为 AI 编码助手提供本项目的上下文与规范。全局行为准则见用户级 AGENTS.md（想清楚再写、简单优先、外科手术式改动、目标驱动执行）。

## 语言约定

- 对话输出默认简体中文。
- 代码注释、命令说明、日志说明默认使用日文。
- 命令、代码、日志本体保持原文。

## 项目概述

Font Unifier 是基于 Python + PyQt6 的 Windows 桌面工具，批量统一 Office 文件（.docx / .xlsx / .pptx）的字体。核心逻辑为三个纯函数 + 一个 GUI 类，处理在后台 QThread 执行。

## 技术栈

- Python 3.12+
- GUI：PyQt6
- Office 处理：python-docx（Word）、openpyxl（Excel）、python-pptx（PowerPoint）
- 测试：pytest + pytest-cov
- Lint：flake8（max-line-length=120）

## 项目结构

```
ChangeFont/
├── src/font_unifier.py   # 主程序：核心字体函数 + QThread worker + GUI
├── tests/                # pytest 单元测试（conftest.py 把 src/ 加入 sys.path）
├── docs/PRD.md           # 产品需求文档
├── config/               # 配置文件目录
├── run_font_unifier.bat  # Windows 启动脚本
└── requirements.txt
```

## 常用命令

所有命令在项目根目录执行，使用项目虚拟环境解释器：

```powershell
# 运行程序
.\.venv\Scripts\python.exe src\font_unifier.py

# 运行测试
.\.venv\Scripts\python.exe -m pytest tests\ -v

# 覆盖率
.\.venv\Scripts\python.exe -m pytest tests\ --cov=font_unifier --cov-report=term

# Lint（max-line-length=120）
.\.venv\Scripts\python.exe -m flake8 src\font_unifier.py tests\ --max-line-length=120
```

> 改动代码后必须运行测试与 flake8，两者均通过方可视为完成。

## 架构要点

- **核心函数**（纯函数，易测）：`change_word_font` / `change_excel_font` / `change_ppt_font`，输入 `(path, font_name)`，返回内存对象，不负责保存。
- **调度函数**：`process_office_file(path, font_name)` 用 `os.path.splitext` 切分，按扩展名（大小写不敏感）分发并保存为 `原名_modified.ext`。
- **字体属性**：Word 写入 `w:ascii/hAnsi/eastAsia/cs`，PowerPoint 写入 `a:latin/ea/cs`，确保中日文字符生效。
- **主题字体覆盖（重要）**：仅改字体名不够，必须清除主题引用，否则 Excel/Word 仍按主题字体渲染：
  - Excel：`_replace_all_fonts` 遍历 `workbook._fonts`，替换每个字体的 `name` 并把 `scheme` 置 `None`（`scheme` 与 `name` 共存时 Excel 会忽略 `name`、改用主题东亚字体，导致多 sheet 未改）。此法一次覆盖所有单元格、命名样式与默认（Normal）字体，且不再触发 `Font.copy()` 的 DeprecationWarning。
  - Word：`_set_docx_run_font` 设显式名后 `attrib.pop` 掉 `w:asciiTheme/hAnsiTheme/eastAsiaTheme/cstheme`。
  - PowerPoint：`_set_pptx_run_font` 直接 `set('typeface', ...)` 覆盖现有 `a:latin/ea/cs`，主题引用（`+mn-lt` 等）一并被替换，无需额外处理。
- **后台处理**：`FontProcessingWorker(QThread)` 通过信号 `finished`/`error` 回主线程更新 UI，避免大文件卡死。
- **图表兜底**：`_process_chart_fonts` 对单形状异常 try/except，单一图表异常不得拖垮整个文件。
- **GUI**：`FontUnifierApp(QMainWindow)`。全局 QSS（`APP_QSS` 常量）挂在 `QApplication` 上统一样式；卡片布局、靛蓝强调色、busy `QProgressBar`、`_set_status(text, kind)` 用动态属性 `kind` + `unpolish/polish` 切换状态色块。字体框为可编辑 `QComboBox`，配合 `QCompleter`（大小写不敏感前缀匹配）与 `eventFilter`（点击输入框即弹列表），行为对标 Excel。字体列表来自 `QFontDatabase.families()`。

## 代码规范

- 遵循 flake8，单行不超过 120 字符。
- 缩进 4 空格；模块级函数/类之间空 2 行。
- 不主动添加注释（除非逻辑不直观）；新增注释用日文。
- 不修改与本次需求无关的代码、格式或文件。

## 注意事项

- 输出文件命名为 `原名_modified.扩展名`，**绝不覆盖原文件**。
- python-pptx 没有 `chart.x_axis/y_axis`，应使用 `category_axis/value_axis`；`series.data_labels` 不可逐点迭代。
- 验证字体是否真生效，不能只看 openpyxl/python-docx 的再読込，必须直接检查 `xl/styles.xml`、`word/document.xml`、`xl/theme/theme1.xml`（Excel 的 `<scheme>` 与 Word 的 `asciiTheme` 等主题引用是常见陷阱）。
- .venv、.kilo/、规划草稿文件（task_plan.md/findings.md/progress.md）不应提交（见 .gitignore）。
