from mcp.server.fastmcp import FastMCP, Context
from pptx.util import Inches, Pt, Cm, Emu
from typing import Optional, Union, List, Tuple
import os
import datetime
import traceback
from svg_module import insert_svg_to_pptx, to_emu, create_svg_file, get_pptx_slide_count

# 创建MCP服务器实例
mcp = FastMCP(name="main")

# 主要的SVG插入工具
@mcp.tool()
def insert_svg(
    pptx_path: str,
    svg_path: str,
    slide_number: int = 1,
    x_inches: Optional[float] = None,
    y_inches: Optional[float] = None,
    width_inches: Optional[float] = None,
    height_inches: Optional[float] = None,
    output_path: Optional[str] = None,
    create_if_not_exists: bool = True
) -> str:
    """
    将SVG图像插入到PPTX文件的指定位置。
    如果未提供PPTX路径，将自动创建一个临时文件。
    如果未提供输出路径，将覆盖原始文件。
    如果未提供坐标，默认对齐幻灯片左上角。
    如果未提供宽度和高度，默认覆盖整个幻灯片（16:9）。

    Args:
        pptx_path: PPTX文件路径，如果未提供则自动创建一个临时文件，最好使用英文路径
        svg_path: SVG文件路径，最好使用英文路径
        slide_number: 目标幻灯片编号（从1开始）
        x_inches: X坐标（英寸），如果未指定则默认为0
        y_inches: Y坐标（英寸），如果未指定则默认为0
        width_inches: 宽度（英寸），如果未指定则使用幻灯片宽度
        height_inches: 高度（英寸），如果未指定则根据宽度计算或使用幻灯片高度
        output_path: 输出文件路径，如果未指定则覆盖原始文件
        create_if_not_exists: 如果为True且PPTX文件不存在，将自动创建一个新文件
        
    Returns:
        操作结果消息，包含详细的错误信息（如果有）
    """
    # 收集错误信息
    error_messages = []

    if not os.path.isabs(pptx_path):
        pptx_path = os.path.abspath(pptx_path)
    
    # 确保PPTX文件的父目录存在
    pptx_dir = os.path.dirname(pptx_path)
    if not os.path.exists(pptx_dir):
        try:
            os.makedirs(pptx_dir, exist_ok=True)
            print(f"已创建PPTX目录: {pptx_dir}")
            error_messages.append(f"已创建PPTX目录: {pptx_dir}")
        except Exception as e:
            error_msg = f"创建PPTX目录 {pptx_dir} 时出错: {e}"
            error_messages.append(error_msg)
            return error_msg
    
    # 将英寸转换为Inches对象
    x = Inches(x_inches) if x_inches is not None else None
    y = Inches(y_inches) if y_inches is not None else None
    width = Inches(width_inches) if width_inches is not None else None
    height = Inches(height_inches) if height_inches is not None else None
    
    # 检查SVG文件是否存在，如果是相对路径则转换为绝对路径
    if not os.path.isabs(svg_path):
        svg_path = os.path.abspath(svg_path)
    
    # 确保SVG文件的父目录存在
    svg_dir = os.path.dirname(svg_path)
    if not os.path.exists(svg_dir):
        try:
            os.makedirs(svg_dir, exist_ok=True)
            print(f"已创建SVG目录: {svg_dir}")
            error_messages.append(f"已创建SVG目录: {svg_dir}")
        except Exception as e:
            error_msg = f"创建SVG目录 {svg_dir} 时出错: {e}"
            error_messages.append(error_msg)
            return error_msg
        
    # 如果SVG文件不存在且create_if_not_exists为True，则创建一个简单的SVG文件
    if not os.path.exists(svg_path) and create_if_not_exists:
        svg_created = create_svg_file(svg_path)
        if not svg_created:
            error_msg = f"错误：无法创建SVG文件 {svg_path}"
            error_messages.append(error_msg)
            return error_msg
    elif not os.path.exists(svg_path):
        error_msg = f"错误：SVG文件 {svg_path} 不存在"
        error_messages.append(error_msg)
        return error_msg
    
    # 如果提供了输出路径且是相对路径，转换为绝对路径
    if output_path and not os.path.isabs(output_path):
        output_path = os.path.abspath(output_path)
    
    # 如果提供了输出路径，确保其父目录存在
    if output_path:
        output_dir = os.path.dirname(output_path)
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                print(f"已创建输出目录: {output_dir}")
                error_messages.append(f"已创建输出目录: {output_dir}")
            except Exception as e:
                error_msg = f"创建输出目录 {output_dir} 时出错: {e}"
                error_messages.append(error_msg)
                return error_msg
    
    try:
        # 调用改进后的函数，它现在返回一个元组 (成功标志, 错误消息)
        result = insert_svg_to_pptx(
            pptx_path=pptx_path,
            svg_path=svg_path,
            slide_number=slide_number,
            x=x,
            y=y,
            width=width,
            height=height,
            output_path=output_path,
            create_if_not_exists=create_if_not_exists
        )
        
        # 检查返回值类型
        if isinstance(result, tuple) and len(result) == 2:
            success, error_details = result
        else:
            # 向后兼容
            success = result
            error_details = ""
        
        if success:
            result_path = output_path or pptx_path
            was_created = not os.path.exists(pptx_path) and create_if_not_exists
            creation_msg = "（已自动创建PPTX文件）" if was_created else ""
            return f"成功将SVG文件 {svg_path} 插入到 {result_path} 的第 {slide_number} 张幻灯片 {creation_msg}"
        else:
            # 返回详细的错误信息
            return f"插入SVG到PPTX文件失败，详细错误信息：\n{error_details}"
    except Exception as e:
        # 收集异常堆栈
        error_trace = traceback.format_exc()
        error_msg = f"插入SVG时发生错误: {str(e)}\n\n详细堆栈跟踪：\n{error_trace}"
        error_messages.append(error_msg)
        return error_msg

@mcp.tool()
def list_files(directory: str = ".", file_type: Optional[str] = None) -> str:
    """
    列出目录中的文件，可选按文件类型过滤。
    
    Args:
        directory: 要列出文件的目录路径
        file_type: 文件类型过滤，可以是 "svg" 或 "pptx"
        
    Returns:
        文件列表（每行一个文件）
    """
    import os
    
    if not os.path.exists(directory):
        return f"错误：目录 {directory} 不存在"
    
    files = os.listdir(directory)
    
    if file_type:
        file_type = file_type.lower()
        extensions = {
            "svg": [".svg"],
            "pptx": [".pptx", ".ppt"]
        }
        
        if file_type in extensions:
            filtered_files = []
            for file in files:
                if any(file.lower().endswith(ext) for ext in extensions[file_type]):
                    filtered_files.append(file)
            files = filtered_files
        else:
            files = [f for f in files if f.lower().endswith(f".{file_type}")]
    
    if not files:
        return f"未找到{'任何' if not file_type else f'{file_type}'} 文件"
    
    return "\n".join(files)

@mcp.tool()
def get_file_info(file_path: str) -> str:
    """
    获取文件信息，如存在状态、大小等。
    
    Args:
        file_path: 要查询的文件路径
        
    Returns:
        文件信息
    """
    import os
    
    if not os.path.exists(file_path):
        return f"文件 {file_path} 不存在"
    
    if os.path.isdir(file_path):
        return f"{file_path} 是一个目录"
    
    size_bytes = os.path.getsize(file_path)
    size_kb = size_bytes / 1024
    size_mb = size_kb / 1024
    
    if size_mb >= 1:
        size_str = f"{size_mb:.2f} MB"
    else:
        size_str = f"{size_kb:.2f} KB"
    
    modified_time = os.path.getmtime(file_path)
    from datetime import datetime
    modified_str = datetime.fromtimestamp(modified_time).strftime("%Y-%m-%d %H:%M:%S")
    
    # 获取文件扩展名
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    
    file_type = None
    if ext == ".svg":
        file_type = "SVG图像"
    elif ext in [".pptx", ".ppt"]:
        file_type = "PowerPoint演示文稿"
    else:
        file_type = f"{ext[1:]} 文件" if ext else "未知类型文件"
    
    return f"文件: {file_path}\n类型: {file_type}\n大小: {size_str}\n修改时间: {modified_str}"

# 添加一个将SVG转换为PNG的工具
@mcp.tool()
def convert_svg_to_png(
    svg_path: str,
    output_path: Optional[str] = None
) -> str:
    """
    将SVG文件转换为PNG图像。
    
    Args:
        svg_path: SVG文件路径
        output_path: 输出PNG文件路径，如果未指定则使用相同文件名但扩展名为.png
        
    Returns:
        操作结果消息
    """
    from reportlab.graphics import renderPM
    from svglib.svglib import svg2rlg
    import os
    
    if not os.path.exists(svg_path):
        return f"错误：SVG文件 {svg_path} 不存在"
    
    if not output_path:
        # 获取不带扩展名的文件名，然后添加.png扩展名
        base_name = os.path.splitext(svg_path)[0]
        output_path = f"{base_name}.png"
    
    try:
        drawing = svg2rlg(svg_path)
        if drawing is None:
            return f"错误：无法读取SVG文件 {svg_path}"
        
        renderPM.drawToFile(drawing, output_path, fmt="PNG")
        return f"成功将SVG文件 {svg_path} 转换为PNG文件 {output_path}\n宽度: {drawing.width}px\n高度: {drawing.height}px"
    except Exception as e:
        return f"转换SVG到PNG时发生错误: {str(e)}"

@mcp.tool()
def get_pptx_info(pptx_path: str) -> str:
    """
    获取PPTX文件的基本信息，包括幻灯片数量。
    
    Args:
        pptx_path: PPTX文件路径
        
    Returns:
        包含文件信息和幻灯片数量的字符串
    """
    import os
    
    # 确保路径存在
    if not os.path.isabs(pptx_path):
        pptx_path = os.path.abspath(pptx_path)
    
    # 先获取基本文件信息
    if not os.path.exists(pptx_path):
        return f"错误：文件 {pptx_path} 不存在"
    
    size_bytes = os.path.getsize(pptx_path)
    size_kb = size_bytes / 1024
    size_mb = size_kb / 1024
    
    if size_mb >= 1:
        size_str = f"{size_mb:.2f} MB"
    else:
        size_str = f"{size_kb:.2f} KB"
    
    modified_time = os.path.getmtime(pptx_path)
    from datetime import datetime
    modified_str = datetime.fromtimestamp(modified_time).strftime("%Y-%m-%d %H:%M:%S")
    
    # 获取幻灯片数量
    slide_count, error = get_pptx_slide_count(pptx_path)
    
    if error:
        slide_info = f"获取幻灯片数量失败：{error}"
    else:
        slide_info = f"幻灯片数量：{slide_count}张"
    
    return f"PPT文件: {pptx_path}\n大小: {size_str}\n修改时间: {modified_str}\n{slide_info}"

# 启动服务器
if __name__ == "__main__":
    mcp.run() 