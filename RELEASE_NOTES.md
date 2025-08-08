# 发布说明

## v0.2.1 (2024-12-19)

### 🔧 修复和改进
- 修复MCP服务器配置问题
- 添加`__main__.py`支持模块化运行
- 优化MCP客户端配置示例
- 改进项目构建配置
- 更新文档和配置说明

### 📦 技术改进
- 添加hatchling构建后端支持
- 优化包结构和模块导入
- 改进MCP服务器启动方式

## v0.2.0 (2024-12-19)

### 🎉 项目发布

这是 undoom-pdf-mcp 的首次正式发布！

### ✨ 主要功能

- **PDF转图片**: 支持单个和批量PDF文件转换为高质量图片
- **Office转PDF**: Word、Excel、PowerPoint文件转换为PDF
- **PDF加密**: 为PDF文件添加密码保护
- **图片转PDF**: 单张或多张图片合并为PDF文件
- **PDF信息**: 获取PDF文件的详细信息
- **批量处理**: 支持批量文件转换操作

### 🔧 技术特性

- 基于MCP (Model Context Protocol) 协议
- 使用PyMuPDF进行PDF处理
- 集成Windows COM接口支持Office文件转换
- 自动内存管理，避免内存泄漏
- 支持多种图片质量设置

### 📦 安装要求

- Python 3.10+
- Windows系统（Office转换功能）
- Microsoft Office软件（Word、Excel、PowerPoint）

### 🚀 快速开始

```bash
git clone https://github.com/kk520879/undoom_pdf_mcp.git
cd undoom_pdf_mcp
uv sync
uv run python undoom_pdf_mcp/main.py
```

### 📝 文档

完整的使用文档和API说明请参考 [README.md](README.md)

### 🤝 贡献

欢迎提交Issue和Pull Request！

### 📄 许可证

MIT License