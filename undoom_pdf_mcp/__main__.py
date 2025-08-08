#!/usr/bin/env python3
"""
__main__.py - 支持 python -m undoom_pdf_mcp 运行
"""

import asyncio
import sys

def main():
    """同步入口点函数"""
    # 导入异步main函数
    from .main import main as async_main
    
    # 运行异步main函数
    try:
        asyncio.run(async_main())
    except KeyboardInterrupt:
        print("\n服务器已停止")
        sys.exit(0)
    except Exception as e:
        print(f"服务器运行出错: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()