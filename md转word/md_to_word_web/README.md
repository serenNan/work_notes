# Markdown 转 Word Web 应用

一个简单的 Web 应用，支持将 Markdown 文件转换为 Word 文档。

## 功能特点

- ✅ 拖拽上传 Markdown 文件
- ✅ 简洁的用户界面
- ✅ 转换选项配置（目录、代码高亮等）
- ✅ 实时转换进度反馈
- ✅ 直接下载转换后的 Word 文档
- ✅ 自动清理临时文件

## 系统要求

- Python 3.7+
- Pandoc 2.x+ （推荐：S:\Tools\Miniconda\envs\pandoc\Library\bin\pandoc.exe）
- python-docx

## 安装

### 1. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 2. 确保 Pandoc 已安装

应用依赖 Pandoc 进行 Markdown 到 Word 的转换。可以通过 Conda 安装：

```bash
conda install -c conda-forge pandoc
```

或访问 https://pandoc.org/installing.html

### 3. 验证 Pandoc 可用性

```bash
pandoc --version
```

## 运行应用

### 方式 1：直接运行

```bash
python app.py
```

然后打开浏览器访问 http://localhost:5000

### 方式 2：使用启动脚本

创建 `run.bat` 文件（Windows）：

```batch
@echo off
pip install -r requirements.txt
python app.py
```

或创建 `run.sh` 文件（Linux/Mac）：

```bash
#!/bin/bash
pip install -r requirements.txt
python app.py
```

## 使用方法

1. **打开应用** - 在浏览器中访问 http://localhost:5000
2. **上传文件** - 拖拽 .md 文件到上传区域，或点击选择文件
3. **配置选项** - 选择转换选项（生成目录、代码高亮风格等）
4. **开始转换** - 点击"开始转换"按钮
5. **下载文件** - 转换完成后，点击"下载文件"按钮下载 Word 文档

## 项目结构

```
md_to_word_web/
├── app.py                    # Flask 主应用
├── convert.py               # 转换核心逻辑（复制自主项目）
├── requirements.txt         # Python 依赖
├── templates/
│   └── index.html           # Web 界面
├── static/
│   └── style.css            # 样式表
│   └── app.js               # 前端脚本
└── uploads/                 # 临时文件存储
    └── temp/                # 临时文件目录
```

## 支持的格式

- **输入**：.md、.markdown
- **输出**：.docx（Word 2007+ 格式）
- **最大文件大小**：50MB

## 转换选项

### 生成目录
- 勾选此选项将在 Word 文档中生成目录
- 默认：勾选

### 代码高亮
- 选择代码块的高亮风格
- 支持的风格：Tango、Pygments、Espresso、Monochrome
- 默认：Tango

## 输出文档特点

- **正文字体**：宋体，12pt
- **标题字体**：宋体，加粗，黑色
- **代码字体**：Times New Roman，10pt
- **表格**：自动添加完整边框
- **目录**：支持最多 3 级深度

## 后续功能扩展建议

- [ ] 自定义字体选择
- [ ] 页边距调整
- [ ] 文件名自定义
- [ ] 支持多个文件批量转换
- [ ] 转换历史记录
- [ ] 预设配置保存

## 故障排除

### 问题：无法找到 Pandoc

**解决**：确保 Pandoc 已安装且在系统 PATH 中，或修改 convert.py 中的 PANDOC_PATH 变量。

### 问题：转换失败

**解决**：
- 检查 Markdown 文件格式是否正确
- 确保文件使用 UTF-8 编码
- 查看终端输出的错误信息

### 问题：文件下载后无法打开

**解决**：
- 确保安装了 python-docx：`pip install python-docx`
- 尝试用 Microsoft Word 或 LibreOffice 打开文件
- 检查文件是否完整下载

## 技术栈

- **后端框架**：Flask 2.3.0
- **文件处理**：Werkzeug
- **Word 处理**：python-docx 0.8.11
- **Markdown 转换**：Pandoc
- **前端**：HTML5 + CSS3 + Vanilla JavaScript

## 注意事项

- 上传的文件存储在 `uploads/temp/` 目录中
- 临时文件会在 24 小时后自动删除
- 每次刷新页面都需要重新上传文件
- 不支持在线协作编辑

## 许可证

个人使用

## 支持和反馈

如有问题或建议，请查看原始的 convert.py 项目文档。

## 版本历史

- **1.0** - 初始版本
  - 基本的拖拽上传功能
  - Web 界面转换
  - 配置选项支持
