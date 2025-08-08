# undoom-pdf-mcp

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-0.2.0-orange.svg)](https://github.com/kk520879/undoom_pdf_mcp)

一个功能强大的PDF转换工具MCP服务器，基于MCP (Model Context Protocol) 协议，集成了多种文件转换功能。

## ✨ 特性

- 🔄 **PDF转图片**: 支持单个/批量PDF转换为高质量图片
- 📄 **Office转PDF**: Word、Excel、PowerPoint文件转PDF
- 🔒 **PDF加密**: 为PDF文件添加密码保护
- 🖼️ **图片转PDF**: 单张或多张图片合并为PDF
- 📊 **PDF信息**: 获取PDF文件详细信息
- 🚀 **批量处理**: 支持批量文件转换
- 💾 **内存优化**: 自动内存管理，避免内存泄漏

## 功能特性

### PDF转图片功能
- **单个PDF转图片**: 将PDF文件转换为JPG图片
- **批量PDF转图片**: 批量处理多个PDF文件
- **页码选择**: 支持指定转换特定页面
- **质量控制**: 支持多种图片质量设置

### Office文件转PDF功能
- **Word转PDF**: 支持.doc和.docx格式
- **Excel转PDF**: 支持.xls和.xlsx格式
- **PowerPoint转PDF**: 支持.ppt和.pptx格式
- **批量转换**: 支持批量转换Office文件

### PDF安全功能
- **PDF加密**: 为PDF文件设置密码保护
- **权限控制**: 支持设置PDF文件的访问权限

### 图片转PDF功能
- **单张图片转PDF**: 将单张图片转换为PDF文件
- **多张图片合并PDF**: 将多张图片合并为一个PDF文件
- **页面大小设置**: 支持A4、A3、Letter等多种页面大小

### 其他功能
- **PDF信息查看**: 获取PDF文件的详细信息
- **内存优化**: 自动清理内存，避免内存泄漏

## 🚀 快速开始

### 环境要求

- Python 3.10+
- Windows系统（Office文件转换功能需要）
- 已安装Microsoft Office（Word、Excel、PowerPoint）

### 安装

#### 方法1: 使用uv（推荐）

```bash
# 克隆仓库
git clone https://github.com/kk520879/undoom_pdf_mcp.git
cd undoom_pdf_mcp

# 安装依赖
uv sync
```

#### 方法2: 使用pip

```bash
# 克隆仓库
git clone https://github.com/kk520879/undoom_pdf_mcp.git
cd undoom_pdf_mcp

# 创建虚拟环境
python -m venv venv
venv\Scripts\activate  # Windows

# 安装依赖
pip install -e .
```

### 启动服务

```bash
# 使用uv运行
uv run python undoom_pdf_mcp/main.py

# 或直接运行
python undoom_pdf_mcp/main.py
```

## 主要依赖

- `mcp[cli]>=1.12.4` - MCP协议支持
- `PyMuPDF>=1.23.0` - PDF处理库
- `Pillow>=10.0.0` - 图像处理库
- `pywin32>=306` - Windows COM接口（Office文件转换需要）
- `tkinterdnd2>=0.3.0` - GUI拖拽支持

## 使用方法

### 启动MCP服务器

```bash
python main.py
```

### 可用工具

#### 1. pdf_to_images
将PDF文件转换为图片

**参数:**
- `pdf_path` (必需): PDF文件的绝对路径
- `pages` (可选): 要转换的页码，格式如'1,2,3-5'，留空转换所有页
- `quality` (可选): 图片质量倍数，可选值：0.25, 0.5, 1.0, 2.0, 4.0，默认2.0
- `output_dir` (可选): 输出目录路径，留空使用PDF同目录

**示例:**
```json
{
  "pdf_path": "C:\\Documents\\example.pdf",
  "pages": "1,3-5",
  "quality": 2.0
}
```

#### 2. batch_convert_pdfs
批量转换PDF文件为图片

**参数:**
- `folder_path` (必需): 包含PDF文件的文件夹路径
- `page_settings` (必需): 文件名到页码设置的映射
- `quality` (可选): 图片质量倍数，默认2.0

**示例:**
```json
{
  "folder_path": "C:\\Documents\\PDFs",
  "page_settings": {
    "file1.pdf": "1,2,3-5",
    "file2.pdf": "1-10",
    "file3.pdf": ""
  },
  "quality": 2.0
}
```

#### 3. word_to_pdf
将Word文档转换为PDF

**参数:**
- `word_path` (必需): Word文件的绝对路径
- `output_path` (可选): 输出PDF文件路径，留空自动生成

#### 4. excel_to_pdf
将Excel文档转换为PDF

**参数:**
- `excel_path` (必需): Excel文件的绝对路径
- `output_path` (可选): 输出PDF文件路径，留空自动生成

#### 5. ppt_to_pdf
将PowerPoint文档转换为PDF

**参数:**
- `ppt_path` (必需): PowerPoint文件的绝对路径
- `output_path` (可选): 输出PDF文件路径，留空自动生成

#### 6. batch_office_to_pdf
批量转换Office文件为PDF

**参数:**
- `folder_path` (必需): 包含Office文件的文件夹路径
- `file_types` (可选): 要转换的文件类型列表，默认包含所有Office格式

#### 7. get_pdf_info
获取PDF文件信息

**参数:**
- `pdf_path` (必需): PDF文件的绝对路径

#### 8. encrypt_pdf
加密PDF文件

**参数:**
- `pdf_path` (必需): PDF文件的绝对路径
- `password` (必需): 加密密码
- `output_path` (可选): 输出PDF文件路径，留空自动生成

**示例:**
```json
{
  "pdf_path": "C:\\Documents\\example.pdf",
  "password": "mypassword123",
  "output_path": "C:\\Documents\\example_encrypted.pdf"
}
```

#### 9. images_to_pdf
将多张图片合并为PDF

**参数:**
- `image_paths` (必需): 图片文件路径列表
- `output_path` (必需): 输出PDF文件路径
- `page_size` (可选): 页面大小，如A4、A3、Letter等，默认A4

**示例:**
```json
{
  "image_paths": [
    "C:\\Images\\page1.jpg",
    "C:\\Images\\page2.png",
    "C:\\Images\\page3.jpg"
  ],
  "output_path": "C:\\Documents\\merged.pdf",
  "page_size": "A4"
}
```

#### 10. single_image_to_pdf
将单张图片转换为PDF

**参数:**
- `image_path` (必需): 图片文件的绝对路径
- `output_path` (可选): 输出PDF文件路径，留空自动生成
- `page_size` (可选): 页面大小，如A4、A3、Letter等，默认A4

**示例:**
```json
{
  "image_path": "C:\\Images\\document.jpg",
  "page_size": "A4"
}
```

## 页码格式说明

支持以下页码格式：
- `1` - 单页
- `1,2,3` - 多个单页
- `1-5` - 页码范围
- `1,3-5,7` - 混合格式
- 留空 - 转换所有页

## 图片质量说明

- `0.25` - 低质量（文件小）
- `0.5` - 中低质量
- `1.0` - 原始分辨率
- `2.0` - 高质量（默认）
- `4.0` - 超高质量（文件大）

## 注意事项

1. **Office文件转换**: 需要在Windows系统上安装相应的Office软件（Word、Excel、PowerPoint）
2. **文件路径**: 所有路径必须使用绝对路径
3. **权限**: 确保对输入文件有读取权限，对输出目录有写入权限
4. **内存管理**: 处理大文件时会自动进行内存清理

## 错误处理

服务器会捕获并返回详细的错误信息，包括：
- 文件不存在错误
- 权限错误
- 格式不支持错误
- Office应用程序错误

## 🛠️ 开发说明

本项目基于MCP协议开发，集成了多种PDF和Office文件处理功能：

1. **PDF转图片**: 基于PyMuPDF的高质量PDF渲染
2. **Office转PDF**: 利用Windows COM接口调用Office应用程序
3. **PDF加密**: 使用PyMuPDF的安全功能
4. **图片处理**: 基于Pillow的图像处理能力

所有功能都通过MCP协议暴露，可以被支持MCP的AI助手或应用程序调用。

### 项目结构

```
undoom_pdf_mcp/
├── undoom_pdf_mcp/
│   ├── __init__.py
│   └── main.py          # 主程序文件
├── pyproject.toml       # 项目配置
├── README.md           # 项目说明
├── LICENSE             # 许可证
└── test_converter.py   # 测试文件
```

## 🤝 贡献

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📝 更新日志

### v0.2.0 (2024-12-19)
- ✨ 新增PDF加密功能
- ✨ 新增图片转PDF功能
- 🐛 修复内存泄漏问题
- 📚 完善文档和示例

### v0.1.0 (2024-12-18)
- 🎉 初始版本发布
- ✨ PDF转图片功能
- ✨ Office文件转PDF功能
- ✨ 批量处理功能

## 📞 联系方式

- 作者: undoom
- 邮箱: kaikaihuhu666@163.com
- GitHub: [@kk520879](https://github.com/kk520879)

## 📄 许可证

本项目采用 [MIT License](LICENSE) 许可证。

## ⭐ 支持项目

如果这个项目对您有帮助，请给它一个星标 ⭐！

---

**注意**: 本项目主要在Windows系统上测试，Office文件转换功能需要安装相应的Microsoft Office软件。