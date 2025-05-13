"""
mcp_server_okppt - 一个用于将SVG图像插入到PowerPoint演示文稿的MCP服务器包。
"""

import os
import sys

# 导出主要功能
from .svg_module import insert_svg_to_pptx

__all__ = ["insert_svg_to_pptx"]

# 添加命令行入口点
def run_server():
    """
    启动MCP服务器，使其可以通过命令行调用。
    这是package命令行入口点。
    """
    # 导入main模块，它定义了MCP服务器和工具
    from importlib.util import spec_from_file_location, module_from_spec
    
    # 获取main.py的路径
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    main_path = os.path.join(current_dir, "main.py")
    
    if not os.path.exists(main_path):
        print(f"错误: 找不到main.py文件: {main_path}")
        sys.exit(1)
    
    # 动态导入main模块
    spec = spec_from_file_location("mcp_main", main_path)
    main_module = module_from_spec(spec)
    spec.loader.exec_module(main_module)
    
    # 确保必要的目录存在
    if hasattr(main_module, 'get_tmp_dir'):
        main_module.get_tmp_dir()
    if hasattr(main_module, 'get_output_dir'):
        main_module.get_output_dir()
    
    # 运行MCP服务器
    if hasattr(main_module, 'mcp'):
        main_module.mcp.run()
    else:
        print("错误: main.py中未找到mcp对象")
        sys.exit(1)
    
    return 0 