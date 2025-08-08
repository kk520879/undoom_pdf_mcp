#!/usr/bin/env python3
"""
PDF转换工具MCP服务器
集成PDF转JPG、PDF批量转图片、Office文件转PDF等功能
"""

import asyncio
import json
import os
import tempfile
import gc
from typing import Any, Dict, List, Optional
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

from mcp.server.models import InitializationOptions
from mcp.server import NotificationOptions, Server
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
    LoggingLevel
)

server = Server("undoom-pdf-mcp")

class PDFConverter:
    """PDF转换工具类"""
    
    @staticmethod
    def pdf_to_images(pdf_path: str, pages: Optional[List[int]] = None, 
                     quality: float = 2.0, output_dir: Optional[str] = None) -> List[str]:
        """将PDF转换为图片
        
        Args:
            pdf_path: PDF文件路径
            pages: 要转换的页码列表，None表示转换所有页
            quality: 图片质量倍数 (0.25, 0.5, 1.0, 2.0, 4.0)
            output_dir: 输出目录，None表示使用PDF同目录
            
        Returns:
            生成的图片文件路径列表
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
            
        # 打开PDF文档
        doc = fitz.open(pdf_path)
        
        # 设置输出目录
        if output_dir is None:
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.join(os.path.dirname(pdf_path), f"{base_name}_images")
        
        os.makedirs(output_dir, exist_ok=True)
        
        # 确定要转换的页码
        total_pages = len(doc)
        if pages is None:
            pages = list(range(1, total_pages + 1))
        
        # 设置缩放矩阵
        matrix = fitz.Matrix(quality, quality)
        
        output_files = []
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        try:
            for page_num in pages:
                if 1 <= page_num <= total_pages:
                    # 加载页面
                    page = doc.load_page(page_num - 1)  # fitz使用0基索引
                    
                    # 渲染为图片
                    pix = page.get_pixmap(matrix=matrix)
                    
                    # 生成输出文件名
                    if len(pages) == 1 and len(pages) == total_pages:
                        output_file = os.path.join(output_dir, f"{base_name}.jpg")
                    else:
                        output_file = os.path.join(output_dir, f"{base_name}_page_{page_num}.jpg")
                    
                    # 保存图片
                    pix.save(output_file)
                    output_files.append(output_file)
                    
                    # 清理内存
                    pix = None
                    page = None
                    gc.collect()
                    
        finally:
            doc.close()
            
        return output_files
    
    @staticmethod
    def parse_page_numbers(page_string: str) -> List[int]:
        """解析页码字符串
        
        Args:
            page_string: 页码字符串，如 "1,2,3-5,7"
            
        Returns:
            页码列表
        """
        if not page_string.strip():
            return []
            
        pages = set()
        for part in page_string.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-', 1))
                pages.update(range(start, end + 1))
            else:
                pages.add(int(part))
        
        return sorted(list(pages))
    
    @staticmethod
    def batch_convert_pdfs(folder_path: str, page_settings: Dict[str, str], 
                          quality: float = 2.0) -> Dict[str, List[str]]:
        """批量转换PDF文件
        
        Args:
            folder_path: PDF文件夹路径
            page_settings: 文件名到页码设置的映射
            quality: 图片质量倍数
            
        Returns:
            文件名到输出图片列表的映射
        """
        results = {}
        
        for filename, page_string in page_settings.items():
            pdf_path = os.path.join(folder_path, filename)
            if os.path.exists(pdf_path) and filename.lower().endswith('.pdf'):
                try:
                    pages = PDFConverter.parse_page_numbers(page_string) if page_string else None
                    output_files = PDFConverter.pdf_to_images(pdf_path, pages, quality)
                    results[filename] = output_files
                except Exception as e:
                    results[filename] = [f"错误: {str(e)}"]
                    
        return results
    
    @staticmethod
    def encrypt_pdf(pdf_path: str, password: str, output_path: Optional[str] = None) -> str:
        """加密PDF文件
        
        Args:
            pdf_path: PDF文件路径
            password: 加密密码
            output_path: 输出文件路径，None表示覆盖原文件
            
        Returns:
            加密后的PDF文件路径
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
            
        if output_path is None:
            base_name = os.path.splitext(pdf_path)[0]
            output_path = f"{base_name}_encrypted.pdf"
            
        doc = fitz.open(pdf_path)
        try:
            # 设置加密参数
            encrypt_meth = fitz.PDF_ENCRYPT_AES_256  # 使用AES-256加密
            owner_pass = password  # 所有者密码
            user_pass = password   # 用户密码
            
            # 设置权限（允许所有操作）
            permissions = fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ANNOTATE
            
            # 保存加密的PDF
            doc.save(output_path, 
                    encryption=encrypt_meth,
                    owner_pw=owner_pass,
                    user_pw=user_pass,
                    permissions=permissions)
            
            return output_path
            
        finally:
            doc.close()
    
    @staticmethod
    def images_to_pdf(image_paths: List[str], output_path: str, 
                     page_size: str = "A4") -> str:
        """将多张图片合并为PDF
        
        Args:
            image_paths: 图片文件路径列表
            output_path: 输出PDF文件路径
            page_size: 页面大小，如"A4", "A3", "Letter"等
            
        Returns:
            生成的PDF文件路径
        """
        if not image_paths:
            raise ValueError("图片路径列表不能为空")
            
        # 检查所有图片文件是否存在
        for img_path in image_paths:
            if not os.path.exists(img_path):
                raise FileNotFoundError(f"图片文件不存在: {img_path}")
        
        # 创建新的PDF文档
        doc = fitz.open()
        
        # 定义页面大小
        page_sizes = {
            "A4": fitz.paper_rect("a4"),
            "A3": fitz.paper_rect("a3"),
            "A5": fitz.paper_rect("a5"),
            "Letter": fitz.paper_rect("letter"),
            "Legal": fitz.paper_rect("legal")
        }
        
        rect = page_sizes.get(page_size.upper(), fitz.paper_rect("a4"))
        
        try:
            for img_path in image_paths:
                # 打开图片
                img = Image.open(img_path)
                
                # 转换为RGB模式（如果需要）
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 保存为临时文件
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as temp_file:
                    img.save(temp_file.name, 'JPEG', quality=95)
                    temp_path = temp_file.name
                
                try:
                    # 创建新页面
                    page = doc.new_page(width=rect.width, height=rect.height)
                    
                    # 插入图片
                    img_rect = fitz.Rect(0, 0, rect.width, rect.height)
                    page.insert_image(img_rect, filename=temp_path, keep_proportion=True)
                    
                finally:
                    # 清理临时文件
                    if os.path.exists(temp_path):
                        os.unlink(temp_path)
                
                img.close()
            
            # 保存PDF
            doc.save(output_path)
            return output_path
            
        finally:
            doc.close()
    
    @staticmethod
    def single_image_to_pdf(image_path: str, output_path: Optional[str] = None,
                           page_size: str = "A4") -> str:
        """将单张图片转换为PDF
        
        Args:
            image_path: 图片文件路径
            output_path: 输出PDF文件路径，None表示自动生成
            page_size: 页面大小
            
        Returns:
            生成的PDF文件路径
        """
        if output_path is None:
            base_name = os.path.splitext(image_path)[0]
            output_path = f"{base_name}.pdf"
            
        return PDFConverter.images_to_pdf([image_path], output_path, page_size)

class OfficeConverter:
    """Office文件转换工具类"""
    
    @staticmethod
    def word_to_pdf(word_path: str, output_path: Optional[str] = None) -> str:
        """Word转PDF"""
        if not WIN32_AVAILABLE:
            raise RuntimeError("需要安装pywin32库才能转换Office文件")
            
        if not os.path.exists(word_path):
            raise FileNotFoundError(f"Word文件不存在: {word_path}")
            
        if output_path is None:
            output_path = os.path.splitext(word_path)[0] + '.pdf'
            
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            doc = word.Documents.Open(word_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF格式
            doc.Close()
            
            return output_path
            
        finally:
            if word:
                word.Quit()
                word = None
                gc.collect()
    
    @staticmethod
    def excel_to_pdf(excel_path: str, output_path: Optional[str] = None) -> str:
        """Excel转PDF"""
        if not WIN32_AVAILABLE:
            raise RuntimeError("需要安装pywin32库才能转换Office文件")
            
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
            
        if output_path is None:
            output_path = os.path.splitext(excel_path)[0] + '.pdf'
            
        excel = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            workbook = excel.Workbooks.Open(excel_path)
            workbook.ExportAsFixedFormat(0, output_path)  # 0 = PDF格式
            workbook.Close()
            
            return output_path
            
        finally:
            if excel:
                excel.Quit()
                excel = None
                gc.collect()
    
    @staticmethod
    def ppt_to_pdf(ppt_path: str, output_path: Optional[str] = None) -> str:
        """PowerPoint转PDF"""
        if not WIN32_AVAILABLE:
            raise RuntimeError("需要安装pywin32库才能转换Office文件")
            
        if not os.path.exists(ppt_path):
            raise FileNotFoundError(f"PowerPoint文件不存在: {ppt_path}")
            
        if output_path is None:
            output_path = os.path.splitext(ppt_path)[0] + '.pdf'
            
        ppt = None
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Visible = True
            
            presentation = ppt.Presentations.Open(ppt_path)
            presentation.SaveAs(output_path, 32)  # 32 = PDF格式
            presentation.Close()
            
            return output_path
            
        finally:
            if ppt:
                ppt.Quit()
                ppt = None
                gc.collect()
    
    @staticmethod
    def batch_office_to_pdf(folder_path: str, file_types: List[str] = None) -> Dict[str, str]:
        """批量转换Office文件为PDF
        
        Args:
            folder_path: 文件夹路径
            file_types: 要转换的文件类型列表，如['docx', 'xlsx', 'pptx']
            
        Returns:
            原文件名到PDF文件路径的映射
        """
        if file_types is None:
            file_types = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']
            
        results = {}
        
        for filename in os.listdir(folder_path):
            file_ext = filename.lower().split('.')[-1]
            if file_ext in file_types:
                file_path = os.path.join(folder_path, filename)
                try:
                    if file_ext in ['doc', 'docx']:
                        output_path = OfficeConverter.word_to_pdf(file_path)
                    elif file_ext in ['xls', 'xlsx']:
                        output_path = OfficeConverter.excel_to_pdf(file_path)
                    elif file_ext in ['ppt', 'pptx']:
                        output_path = OfficeConverter.ppt_to_pdf(file_path)
                    else:
                        continue
                        
                    results[filename] = output_path
                    
                except Exception as e:
                    results[filename] = f"错误: {str(e)}"
                    
        return results

@server.list_tools()
async def handle_list_tools() -> List[Tool]:
    """列出可用的工具"""
    return [
        Tool(
            name="pdf_to_images",
            description="将PDF文件转换为图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件的绝对路径"
                    },
                    "pages": {
                        "type": "string",
                        "description": "要转换的页码，格式如'1,2,3-5'，留空转换所有页",
                        "default": ""
                    },
                    "quality": {
                        "type": "number",
                        "description": "图片质量倍数，可选值：0.25, 0.5, 1.0, 2.0, 4.0",
                        "default": 2.0
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录路径，留空使用PDF同目录",
                        "default": ""
                    }
                },
                "required": ["pdf_path"]
            }
        ),
        Tool(
            name="batch_convert_pdfs",
            description="批量转换PDF文件为图片",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "包含PDF文件的文件夹路径"
                    },
                    "page_settings": {
                        "type": "object",
                        "description": "文件名到页码设置的映射，如{'file1.pdf': '1,2,3-5'}",
                        "additionalProperties": {"type": "string"}
                    },
                    "quality": {
                        "type": "number",
                        "description": "图片质量倍数",
                        "default": 2.0
                    }
                },
                "required": ["folder_path", "page_settings"]
            }
        ),
        Tool(
            name="word_to_pdf",
            description="将Word文档转换为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "word_path": {
                        "type": "string",
                        "description": "Word文件的绝对路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径，留空自动生成",
                        "default": ""
                    }
                },
                "required": ["word_path"]
            }
        ),
        Tool(
            name="excel_to_pdf",
            description="将Excel文档转换为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "excel_path": {
                        "type": "string",
                        "description": "Excel文件的绝对路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径，留空自动生成",
                        "default": ""
                    }
                },
                "required": ["excel_path"]
            }
        ),
        Tool(
            name="ppt_to_pdf",
            description="将PowerPoint文档转换为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "ppt_path": {
                        "type": "string",
                        "description": "PowerPoint文件的绝对路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径，留空自动生成",
                        "default": ""
                    }
                },
                "required": ["ppt_path"]
            }
        ),
        Tool(
            name="batch_office_to_pdf",
            description="批量转换Office文件为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder_path": {
                        "type": "string",
                        "description": "包含Office文件的文件夹路径"
                    },
                    "file_types": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "要转换的文件类型列表，如['docx', 'xlsx', 'pptx']",
                        "default": ["doc", "docx", "xls", "xlsx", "ppt", "pptx"]
                    }
                },
                "required": ["folder_path"]
            }
        ),
        Tool(
            name="get_pdf_info",
            description="获取PDF文件信息",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件的绝对路径"
                    }
                },
                "required": ["pdf_path"]
            }
        ),
        Tool(
            name="encrypt_pdf",
            description="加密PDF文件",
            inputSchema={
                "type": "object",
                "properties": {
                    "pdf_path": {
                        "type": "string",
                        "description": "PDF文件的绝对路径"
                    },
                    "password": {
                        "type": "string",
                        "description": "加密密码"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径，留空自动生成",
                        "default": ""
                    }
                },
                "required": ["pdf_path", "password"]
            }
        ),
        Tool(
            name="images_to_pdf",
            description="将多张图片合并为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "image_paths": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "图片文件路径列表"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径"
                    },
                    "page_size": {
                        "type": "string",
                        "description": "页面大小，如A4、A3、Letter等",
                        "default": "A4"
                    }
                },
                "required": ["image_paths", "output_path"]
            }
        ),
        Tool(
            name="single_image_to_pdf",
            description="将单张图片转换为PDF",
            inputSchema={
                "type": "object",
                "properties": {
                    "image_path": {
                        "type": "string",
                        "description": "图片文件的绝对路径"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出PDF文件路径，留空自动生成",
                        "default": ""
                    },
                    "page_size": {
                        "type": "string",
                        "description": "页面大小，如A4、A3、Letter等",
                        "default": "A4"
                    }
                },
                "required": ["image_path"]
            }
        )
    ]

@server.call_tool()
async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
    """处理工具调用"""
    try:
        if name == "pdf_to_images":
            pdf_path = arguments["pdf_path"]
            pages_str = arguments.get("pages", "")
            quality = arguments.get("quality", 2.0)
            output_dir = arguments.get("output_dir", "") or None
            
            pages = PDFConverter.parse_page_numbers(pages_str) if pages_str else None
            output_files = PDFConverter.pdf_to_images(pdf_path, pages, quality, output_dir)
            
            return [TextContent(
                type="text",
                text=f"成功转换PDF为图片！\n生成的图片文件：\n" + "\n".join(output_files)
            )]
            
        elif name == "batch_convert_pdfs":
            folder_path = arguments["folder_path"]
            page_settings = arguments["page_settings"]
            quality = arguments.get("quality", 2.0)
            
            results = PDFConverter.batch_convert_pdfs(folder_path, page_settings, quality)
            
            result_text = "批量转换结果：\n"
            for filename, output_files in results.items():
                result_text += f"\n{filename}:\n"
                if isinstance(output_files, list) and output_files:
                    if output_files[0].startswith("错误:"):
                        result_text += f"  {output_files[0]}\n"
                    else:
                        result_text += "\n".join(f"  - {f}" for f in output_files) + "\n"
                        
            return [TextContent(type="text", text=result_text)]
            
        elif name == "word_to_pdf":
            word_path = arguments["word_path"]
            output_path = arguments.get("output_path", "") or None
            
            result_path = OfficeConverter.word_to_pdf(word_path, output_path)
            
            return [TextContent(
                type="text",
                text=f"成功将Word文档转换为PDF！\n输出文件：{result_path}"
            )]
            
        elif name == "excel_to_pdf":
            excel_path = arguments["excel_path"]
            output_path = arguments.get("output_path", "") or None
            
            result_path = OfficeConverter.excel_to_pdf(excel_path, output_path)
            
            return [TextContent(
                type="text",
                text=f"成功将Excel文档转换为PDF！\n输出文件：{result_path}"
            )]
            
        elif name == "ppt_to_pdf":
            ppt_path = arguments["ppt_path"]
            output_path = arguments.get("output_path", "") or None
            
            result_path = OfficeConverter.ppt_to_pdf(ppt_path, output_path)
            
            return [TextContent(
                type="text",
                text=f"成功将PowerPoint文档转换为PDF！\n输出文件：{result_path}"
            )]
            
        elif name == "batch_office_to_pdf":
            folder_path = arguments["folder_path"]
            file_types = arguments.get("file_types", ["doc", "docx", "xls", "xlsx", "ppt", "pptx"])
            
            results = OfficeConverter.batch_office_to_pdf(folder_path, file_types)
            
            result_text = "批量转换Office文件结果：\n"
            for filename, output_path in results.items():
                if output_path.startswith("错误:"):
                    result_text += f"\n{filename}: {output_path}"
                else:
                    result_text += f"\n{filename} -> {output_path}"
                    
            return [TextContent(type="text", text=result_text)]
            
        elif name == "get_pdf_info":
            pdf_path = arguments["pdf_path"]
            
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
                
            doc = fitz.open(pdf_path)
            try:
                info = {
                    "文件路径": pdf_path,
                    "文件大小": f"{os.path.getsize(pdf_path) / 1024 / 1024:.2f} MB",
                    "页数": len(doc),
                    "标题": doc.metadata.get('title', '未知'),
                    "作者": doc.metadata.get('author', '未知'),
                    "创建时间": doc.metadata.get('creationDate', '未知'),
                    "修改时间": doc.metadata.get('modDate', '未知')
                }
                
                info_text = "PDF文件信息：\n"
                for key, value in info.items():
                    info_text += f"{key}: {value}\n"
                    
                return [TextContent(type="text", text=info_text)]
                
            finally:
                doc.close()
        
        elif name == "encrypt_pdf":
            pdf_path = arguments["pdf_path"]
            password = arguments["password"]
            output_path = arguments.get("output_path", "") or None
            
            result_path = PDFConverter.encrypt_pdf(pdf_path, password, output_path)
            
            return [TextContent(
                type="text",
                text=f"成功加密PDF文件！\n输出文件：{result_path}"
            )]
            
        elif name == "images_to_pdf":
            image_paths = arguments["image_paths"]
            output_path = arguments["output_path"]
            page_size = arguments.get("page_size", "A4")
            
            result_path = PDFConverter.images_to_pdf(image_paths, output_path, page_size)
            
            return [TextContent(
                type="text",
                text=f"成功将{len(image_paths)}张图片合并为PDF！\n输出文件：{result_path}"
            )]
            
        elif name == "single_image_to_pdf":
            image_path = arguments["image_path"]
            output_path = arguments.get("output_path", "") or None
            page_size = arguments.get("page_size", "A4")
            
            result_path = PDFConverter.single_image_to_pdf(image_path, output_path, page_size)
            
            return [TextContent(
                type="text",
                text=f"成功将图片转换为PDF！\n输出文件：{result_path}"
            )]
        
        else:
            return [TextContent(
                type="text",
                text=f"未知的工具: {name}"
            )]
            
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"执行工具 {name} 时发生错误: {str(e)}"
        )]

async def main():
    """主函数"""
    # 运行服务器
    from mcp.server.stdio import stdio_server
    
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="undoom-pdf-mcp",
                server_version="0.1.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    asyncio.run(main())
