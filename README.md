# Font Unifier

## 项目简介

Font Unifier 是一个基于 Python 的桌面应用程序，用于批量统一 Microsoft Office 文件（Word、Excel、PowerPoint）的字体。该工具通过直观的图形用户界面，让用户轻松选择文件并指定目标字体，实现字体统一化处理。

## 主要功能

- **支持多种文件格式**：.docx、.xlsx、.pptx
- **批量字体更改**：统一文档中所有文本的字体
- **图形用户界面**：简单易用的操作界面
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
python -m venv venv
```

2. 激活虚拟环境：
```bash
venv\Scripts\activate
```

3. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 启动应用程序
双击运行 `run_font_unifier.bat` 文件，或在命令行中执行：
```bash
python font_unifier.py
```

### 操作步骤
1. 点击 "Browse..." 按钮选择要处理的 Office 文件
2. 在 "Target Font" 输入框中输入目标字体名称（默认为 "Meiryo UI"）
3. 点击 "Start Processing" 按钮开始处理
4. 处理完成后，查看状态信息和保存路径

### 支持的文件类型
- **Word 文档** (.docx)：更改所有段落和表格文本的字体
- **Excel 工作簿** (.xlsx)：更改所有工作表单元格的字体
- **PowerPoint 演示文稿** (.pptx)：更改所有幻灯片文本的字体

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
├── font_unifier.py          # 主程序文件
├── run_font_unifier.bat     # Windows 启动脚本
├── PRD.md                   # 产品需求文档
├── README.md                # 项目说明文档
├── AGENTS.md                # 开发规范文档
├── requirements.txt         # Python 依赖包列表
├── .gitignore               # Git 忽略文件配置
└── venv/                    # 虚拟环境目录（运行时创建）
```

## 开发信息

- **开发语言**：Python 3.12+
- **GUI 框架**：PyQt6
- **Office 处理库**：
  - python-docx（Word 文件处理）
  - openpyxl（Excel 文件处理）
  - python-pptx（PowerPoint 文件处理）

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