# Font Unifier

## 项目简介

Font Unifier 是一个基于 Python 的桌面应用程序，用于批量统一 Microsoft Office 文件（Word、Excel、PowerPoint）的字体。该工具通过直观的图形用户界面，让用户轻松选择文件并指定目标字体，实现字体统一化处理。

## 主要功能

- **支持多种文件格式**：.docx、.xlsx、.pptx
- **批量字体更改**：统一文档中所有文本的字体
- **主题字体覆盖**：正确处理 Excel/Word 的主题字体引用（`scheme`/`asciiTheme`），避免多 sheet 或 CJK 文本回退到旧字体
- **现代图形界面**：浅色卡片式布局、靛蓝强调色、加载动画、状态色块
- **智能字体选择**：下拉框枚举系统全部字体，支持输入并前缀自动匹配（行为对标 Excel）
- **后台处理**：大文件处理在后台线程执行，界面不卡顿
- **自动保存**：生成修改后的新文件，原文件保持不变

## 安装说明

### 系统要求
- Python 3.12 或更高版本
- Windows 操作系统
- Microsoft Office（用于查看处理后的文件）

### 依赖包安装
```bash
pip install PyQt6 python-docx openpyxl python-pptx
```

### 运行环境设置
1. 创建虚拟环境（推荐）：
```bash
python -m venv .venv
```

2. 激活虚拟环境：
```bash
.venv\Scripts\activate
```

3. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 启动应用程序
双击运行 `run_font_unifier.bat` 文件，或在命令行中执行：
```bash
python src/font_unifier.py
```

### 操作步骤
1. 点击 "Browse…" 按钮选择要处理的 Office 文件
2. 在 "目标字体" 框中选择目标字体（默认为 "Meiryo UI"）：可下拉选择，也可直接输入，输入时按前缀自动匹配系统已安装字体
3. 点击 "Start Processing" 按钮开始处理
4. 处理完成后，查看状态信息和保存路径

### 支持的文件类型
- **Word 文档** (.docx)：更改所有段落和表格文本的字体（同时清除主题字体引用）
- **Excel 工作簿** (.xlsx)：替换全部字体定义（含默认/Normal 字体）并清除 `scheme`，覆盖所有工作表
- **PowerPoint 演示文稿** (.pptx)：更改所有幻灯片文本（含表格、图表、嵌套组形状）的字体

## 输出说明

处理后的文件将以 `原文件名_modified.扩展名` 的格式保存到原文件所在目录。例如：
- `document.docx` → `document_modified.docx`
- `workbook.xlsx` → `workbook_modified.xlsx`

## 注意事项

- **原文件保护**：工具不会修改原文件，只生成新的修改版本
- **字体兼容性**：确保指定的目标字体在系统中已安装
- **文件大小**：大文件处理可能需要较长时间，请耐心等待
- **错误处理**：如果处理过程中出现错误，请检查文件是否损坏或字体名称是否正确

## 项目结构

```
ChangeFont/
├── src/
│   └── font_unifier.py      # 主程序文件
├── tests/                   # 单元测试（pytest）
├── docs/
│   └── PRD.md               # 产品需求文档
├── config/                  # 配置文件目录
├── run_font_unifier.bat     # Windows 启动脚本
├── README.md                # 项目说明文档
├── requirements.txt         # Python 依赖包列表
├── .gitignore               # Git 忽略文件配置
└── .venv/                   # 虚拟环境目录（运行时创建）
```

## 开发信息

- **开发语言**：Python 3.12+
- **GUI 框架**：PyQt6
- **Office 处理库**：
  - python-docx（Word 文件处理）
  - openpyxl（Excel 文件处理）
  - python-pptx（PowerPoint 文件处理）
- **字体选择**：下拉框枚举系统已安装的全部字体，支持输入与前缀自动匹配

## 许可证

本项目为个人实验项目，遵循 MIT 许可证。

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。

## 更新日志

### v1.0.0
- 初始版本发布
- 支持 Word、Excel、PowerPoint 文件的字体统一
- 提供图形用户界面
- 实现基本的错误处理和状态反馈

### v1.1.0
- 字体选择方式改为下拉框，提供16种预定义字体选项
- 默认字体设置为 Meiryo UI
- 改进用户体验，减少手动输入错误

### v1.2.0
- 修复中日文字体不生效问题（正确写入 eastAsia/ea 字体属性）
- 修复 Excel 处理覆盖原有字号、粗体、颜色的问题（改为仅替换字体名）
- 修复 PowerPoint 图表分支崩溃及轴 API 误用问题
- 支持文件扩展名大小写（如 `.PPTX`）
- 字体处理迁移到后台线程（QThread），解决大文件界面卡死
- 新增单元测试（pytest）

### v1.3.0
- 界面现代化：浅色卡片式布局、靛蓝强调色、加载动画、状态色块
- 字体选择增强：枚举系统全部字体，可下拉+输入，前缀自动匹配（对标 Excel）
- 修复 Excel 多 sheet 字体未生效：遍历全部字体定义并清除 `scheme`，覆盖默认/Normal 字体
- 修复 Word 主题字体引用残留：清除 `asciiTheme` 等属性
- 代码简化（`process_office_file` 路径处理、`process_shape_text`、GUI 小方法抽取）