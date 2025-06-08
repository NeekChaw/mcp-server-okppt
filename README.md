# MCP OKPPT Server

[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-blue)](https://github.com/anthropics/anthropic-tools)

一个基于Model Context Protocol (MCP)的服务器工具，专门用于将SVG图像插入到PowerPoint演示文稿中。它能够保留SVG的矢量特性，确保在PowerPoint中显示的图像保持高品质和可缩放性。

## 设计理念

此项目是让大型语言模型（如Claude、GPT等）能够自主设计PowerPoint演示文稿的"曲线救国"解决方案。通过让AI生成SVG图像，再借助本工具将其全屏插入PPT幻灯片，我们成功实现了AI完全控制PPT设计的能力，而无需直接操作复杂的PPT对象模型。

这种方法带来三大核心优势：
1. **AI完全控制**：充分发挥现代AI的图形设计能力，同时避开PPT编程的复杂性
2. **用户可编辑**：Office PowerPoint提供了强大的SVG编辑功能，插入后的SVG元素可以像原生PPT元素一样直接编辑、调整和重新着色，让用户能轻松地在AI生成基础上进行二次修改
3. **矢量级质量**：保持高品质可缩放的矢量特性，确保演示内容在任何尺寸下都清晰锐利

这一创新思路通过SVG作为AI与PPT之间的桥梁，既保证了设计的高度自由，又兼顾了最终成果的实用性和可维护性。

## 功能特点

- **矢量图保留**: 将SVG作为真实矢量图插入PPTX，保证高品质和可缩放性
- **批量批处理**: 支持一次操作多个SVG文件和幻灯片
- **全新演示文稿**: 直接从SVG文件创建完整的演示文稿
- **幻灯片复制与替换**: 智能复制SVG幻灯片并替换现有内容
- **SVG代码处理**: 支持直接从SVG代码创建文件
- **格式转换支持**: 内置SVG到PNG的转换功能

## PPT效果示例

以下是一些使用MCP OKPPT Server生成的PPT效果图：

![2008年金融危机](example/Financial_Crisis_2008.png)
*2008年金融危机分析PPT封面*

![小红书爆款指南](example/小红书如何写出爆款.png)
*小红书爆款内容分析报告PPT页面*

## 安装方法

### 方法一：从PyPI安装

```bash
# 使用pip安装
pip install mcp-server-okppt

# 或使用uv安装
uv pip install mcp-server-okppt
```

### 方法二：配置Claude Desktop

在Claude Desktop配置文件中添加服务器配置：

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

添加以下配置：

```json
{
  "mcpServers": {
    "okppt": {
      "command": "uvx",
      "args": [
        "mcp-server-okppt"
      ]
    }
  }
}
```

### 方法三：从源码安装并配置Cursor本地开发环境

在Cursor IDE中，可以通过本地配置文件来设置MCP服务器：

**Windows**: `C:\Users\用户名\.cursor\mcp.json`  
**macOS**: `~/.cursor/mcp.json`

添加以下配置：

```json
{
  "mcpServers": {
    "okppt": {
      "command": "uv",
      "args": [
        "--directory",
        "D:\\本地项目路径\\mcp-server-okppt\\src\\mcp_server_okppt",
        "run",
        "cli.py"
      ]
    }
  }
}
```

这种配置方式适合本地开发和测试使用，可以直接指向本地代码目录。

## 使用方法

### 使用Claude Desktop

1. 安装并配置Claude Desktop
2. 在配置文件中添加上述MCP服务器配置
3. 重启Claude Desktop
4. 在对话中使用PPTX相关工具

### 使用MCP CLI进行开发

```bash
# 运行测试
mcp test server.py
```

## 可用工具

### 1. 插入SVG图像 (insert_svg)

```python
def insert_svg(
    pptx_path: str,
    svg_path: List[str],
    slide_number: int = 1,
    x_inches: float = 0,
    y_inches: float = 0,
    width_inches: float = 16,
    height_inches: float = 9,
    output_path: str = "",
    create_if_not_exists: bool = True
) -> str
```

将SVG图像插入到PPTX文件的指定位置。

**参数**:
- `pptx_path`: PPTX文件路径
- `svg_path`: SVG文件路径或路径列表
- `slide_number`: 要插入的幻灯片编号（从1开始）
- `x_inches`: X坐标（英寸）
- `y_inches`: Y坐标（英寸）
- `width_inches`: 宽度（英寸）
- `height_inches`: 高度（英寸）
- `output_path`: 输出文件路径
- `create_if_not_exists`: 如果PPTX不存在是否创建

**返回**: 操作结果消息

### 2. 列出目录文件 (list_files)

```python
def list_files(
    directory: str = ".",
    file_type: Optional[str] = None
) -> str
```

列出目录中的文件。

**参数**:
- `directory`: 目录路径
- `file_type`: 文件类型过滤，可以是"svg"或"pptx"

**返回**: 文件列表

### 3. 获取文件信息 (get_file_info)

```python
def get_file_info(
    file_path: str
) -> str
```

获取文件信息。

**参数**:
- `file_path`: 文件路径

**返回**: 文件信息

### 4. 转换SVG为PNG (convert_svg_to_png)

```python
def convert_svg_to_png(
    svg_path: str,
    output_path: Optional[str] = None
) -> str
```

将SVG文件转换为PNG图像。

**参数**:
- `svg_path`: SVG文件路径
- `output_path`: 输出PNG文件路径

**返回**: 操作结果消息

### 5. 获取PPTX信息 (get_pptx_info)

```python
def get_pptx_info(
    pptx_path: str
) -> str
```

获取PPTX文件的基本信息。

**参数**:
- `pptx_path`: PPTX文件路径

**返回**: 包含文件信息和幻灯片数量的字符串

### 6. 保存SVG代码 (save_svg_code)

```python
def save_svg_code(
    svg_code: str
) -> str
```

将SVG代码保存为SVG文件并返回保存的绝对路径。

**参数**:
- `svg_code`: SVG代码内容

**返回**: 操作结果消息和保存的文件路径

### 7. 删除幻灯片 (delete_slide)

```python
def delete_slide(
    pptx_path: str,
    slide_number: int,
    output_path: Optional[str] = None
) -> str
```

从PPTX文件中删除指定编号的幻灯片。

**参数**:
- `pptx_path`: PPTX文件路径
- `slide_number`: 要删除的幻灯片编号
- `output_path`: 输出文件路径

**返回**: 操作结果消息

### 8. 插入空白幻灯片 (insert_blank_slide)

```python
def insert_blank_slide(
    pptx_path: str,
    slide_number: int,
    layout_index: int = 6,
    output_path: Optional[str] = None,
    create_if_not_exists: bool = True
) -> str
```

在PPTX文件的指定位置插入一个空白幻灯片。

**参数**:
- `pptx_path`: PPTX文件路径
- `slide_number`: 插入位置
- `layout_index`: 幻灯片布局索引，默认为6（空白布局）
- `output_path`: 输出文件路径
- `create_if_not_exists`: 如果PPTX不存在是否创建

**返回**: 操作结果消息

### 9. 复制SVG幻灯片 (copy_svg_slide)

```python
def copy_svg_slide(
    source_pptx_path: str,
    target_pptx_path: str = "",
    source_slide_number: int = 1,
    target_slide_number: Optional[int] = None,
    output_path: Optional[str] = None,
    create_if_not_exists: bool = True
) -> str
```

复制包含SVG图像的幻灯片。

**参数**:
- `source_pptx_path`: 源PPTX文件路径
- `target_pptx_path`: 目标PPTX文件路径
- `source_slide_number`: 要复制的源幻灯片编号
- `target_slide_number`: 要插入到目标文件的位置
- `output_path`: 输出文件路径
- `create_if_not_exists`: 如果目标PPTX不存在是否创建

**返回**: 操作结果消息

## 最佳实践

### 替换幻灯片内容的推荐方法

#### 方法一：完全替换法（最可靠）

```python
# 步骤1：删除要替换的幻灯片
delete_slide(
    pptx_path="演示文稿.pptx",
    slide_number=3,
    output_path="临时文件.pptx"
)

# 步骤2：在同一位置插入空白幻灯片
insert_blank_slide(
    pptx_path="临时文件.pptx",
    slide_number=3,
    output_path="临时文件2.pptx"
)

# 步骤3：将新SVG插入到空白幻灯片
insert_svg(
    pptx_path="临时文件2.pptx",
    svg_path=["新内容.svg"],
    slide_number=3,
    output_path="最终文件.pptx"
)
```

## 注意事项

1. **避免内容叠加**：直接对现有幻灯片插入SVG会导致新内容叠加在原内容上，而非替换
2. **批量处理**：批量插入SVG时，`svg_path`参数必须是数组形式，即使只有一个文件
3. **SVG代码转义**：在使用`save_svg_code`时，特殊字符（如"&"）需要正确转义为"&amp;"
4. **文件路径**：尽量使用英文路径，避免路径中出现特殊字符
5. **检查结果**：每次操作后应检查输出文件以确认修改是否成功

## 常见问题解答

### Q: SVG插入后变成了位图而非矢量图？
A: 请确保使用`copy_svg_slide`或`create_pptx_from_svg`函数，这些函数专门设计用于保留SVG的矢量特性。

### Q: 如何批量处理多个SVG文件？
A: 可以使用`insert_svg`函数并将多个SVG路径作为列表传入，或者使用`create_pptx_from_svg`一次性创建包含多个SVG的演示文稿。

### Q: 文件名变得很长且复杂？
A: 这是因为每次操作都会添加时间戳。建议使用"新文件法"一次性创建最终文件，或在最后一步操作中指定简洁的输出文件名。

## 版本信息

当前最新版本: v0.2.0

查看所有版本和更新信息: [GitHub Releases](https://github.com/NeekChaw/mcp-server-okppt/releases)

## 致谢

本项目在开发过程中受益于[Model Context Protocol(MCP) 编程极速入门](https://github.com/liaokongVFX/MCP-Chinese-Getting-Started-Guide)这一优质资源。该项目提供了全面而清晰的MCP开发指南，涵盖了从基础概念到实际部署的各个方面，极大地降低了开发者学习MCP协议的门槛。特别感谢其在服务配置、工具开发和部署流程等方面的详细示例和说明，为MCP生态的发展和普及做出了宝贵贡献。推荐所有对MCP开发感兴趣的开发者参考这份指南，它将帮助你快速掌握MCP服务器的开发与配置技能。

## 贡献指南

欢迎提交问题和拉取请求到[项目仓库](https://github.com/NeekChaw/mcp-server-okppt)！以下是一些潜在的改进方向：

- 添加更多幻灯片布局支持
- 增强SVG处理和兼容性
- 添加批量SVG处理的进度报告
- 改进错误处理和诊断功能
- 添加图表和表格的特殊处理功能

## 许可证

本项目采用MIT许可证。