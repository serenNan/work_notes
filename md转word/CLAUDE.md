# CLAUDE.md

## 项目概述

这是一个 Markdown 转 Word 文档的转换工具项目，使用 Pandoc 作为核心转换引擎，提供了自定义字体、表格边框、目录生成等高级功能。

## 项目特点

- **核心功能**：使用 Pandoc 将 Markdown 文件转换为 Word 文档
- **字体自定义**：正文使用宋体，代码使用 Times New Roman
- **目录生成**：自动生成目录（TOC），支持最多 3 级深度
- **表格格式**：自动添加完整的表格边框
- **数学支持**：支持 LaTeX 数学公式（行内和块级）
- **代码高亮**：支持多种编程语言的语法高亮
- **批量处理**：可批量转换当前目录的所有 Markdown 文件

## 关键配置

### Pandoc 路径

```
S:\Tools\Miniconda\envs\pandoc\Library\bin\pandoc.exe
```

如果 Pandoc 路径变更，需要修改 `convert.py` 第 20 行的 `PANDOC_PATH` 变量。

### 依赖库

- `python-docx`：用于 Word 文档的后处理
  - 字体设置
  - 表格边框添加
  - 样式调整

安装方法：
```bash
pip install python-docx
```

## 使用方法

### 基础用法

```bash
# 转换当前目录的所有 .md 文件
python convert.py

# 转换指定文件
python convert.py demo.md

# 指定输出文件名
python convert.py demo.md output.docx

# 列出当前目录的 .md 文件
python convert.py --list

# 显示帮助
python convert.py --help
```

### 拖放转换

将 .md 文件拖放到 `convert.py` 上可以快速转换。

## 主要函数说明

### `preprocess_markdown(content: str) -> str`

预处理 Markdown 内容，移除水平分割线（---）。这是因为在 Markdown 中 --- 可能被错误解析为 H2 标题。

### `create_reference_docx(output_path: str) -> bool`

创建自定义 Word 参考模板，设置：
- 正文样式（Normal）：宋体，12pt
- 标题样式（Heading 1-9）：宋体，加粗，黑色
- 目录样式（TOC 1-9）：Times New Roman，10.5pt，单倍行距
- 代码样式：Times New Roman，10pt

### `convert_md_to_docx(input_file: str, output_file: str, reference_doc: str = None) -> bool`

主转换函数，执行以下步骤：
1. 预处理 Markdown 文件
2. 使用 Pandoc 进行转换
3. 设置工作目录以正确处理相对路径中的图片

关键选项：
- `--from markdown+tex_math_dollars+raw_tex`：支持 LaTeX 数学公式
- `--toc`：生成目录
- `--toc-depth=3`：目录深度为 3 级
- `--highlight-style tango`：代码高亮风格

### `apply_chinese_fonts_to_docx(docx_path: str) -> bool`

对生成的 Word 文档进行后处理：
1. 应用字体设置
2. 添加表格边框
3. 设置行间距

### `add_table_borders(table) -> None`

为 Word 表格添加完整的边框（上、下、左、右、内部横线、内部竖线）。

## 工作流程

```
输入 Markdown 文件
    ↓
1. 预处理（移除分割线）
    ↓
2. 创建参考模板
    ↓
3. Pandoc 转换
    ↓
4. 后处理（字体、边框）
    ↓
输出 Word 文档
```

## 文件说明

- `convert.py`：主程序文件，包含所有转换逻辑
- `demo.md`：基础演示文档，展示常见格式
- `demo_advanced.md`：高级演示文档，展示复杂格式（多层级、嵌套列表、表格等）
- `test_tables.md`：表格测试文件（如果存在）
- `reference_template.docx`：临时参考模板（运行时生成，事后删除）

## 常见问题排查

### Pandoc 找不到

确保 Pandoc 已安装在 `S:\Tools\Miniconda\envs\pandoc\Library\bin\pandoc.exe`，或修改 `PANDOC_PATH` 变量。

### 字体显示不正确

确保安装了 `python-docx` 库，运行后处理步骤需要该库。

### 图片路径错误

转换时临时文件和原 Markdown 文件必须在同一目录，以便正确解析相对路径。脚本已自动处理此情况。

### 表格没有边框

确保 `python-docx` 已安装，后处理步骤会自动添加边框。

## 开发指南

### 修改字体和样式

在 `create_reference_docx()` 和 `apply_chinese_fonts_to_docx()` 中修改字体设置：

- 正文字体：修改 `'宋体'`
- 代码字体：修改 `'Times New Roman'`
- 字体大小：修改 `Pt()` 参数值

### 自定义表格边框

在 `add_table_borders()` 中修改 XML 配置：

```python
'<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
```

其中：
- `w:val`：边框样式（single、double 等）
- `w:sz`：边框宽度（八分之一磅）
- `w:color`：边框颜色（十六进制）

### 调整目录深度

在 `convert_md_to_docx()` 中修改：

```python
'--toc-depth=3',  # 改为需要的深度
```

## 注意事项

1. **Pandoc 版本**：脚本需要 Pandoc 2.x 或更高版本
2. **Python 版本**：建议使用 Python 3.7+
3. **文件编码**：所有 Markdown 文件应使用 UTF-8 编码
4. **临时文件**：`_temp_processed.md` 会自动清理，不需手动处理
5. **参考模板**：`reference_template.docx` 是临时文件，会自动删除

## 扩展建议

- 支持自定义 CSS 样式
- 支持页眉页脚设置
- 支持水印和背景图片
- 支持多列排版
- 支持交叉引用和书签
