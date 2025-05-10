# SVG-to-PPTX MCP服务器

一个基于Model Context Protocol (MCP)的服务器，允许通过Claude等大型语言模型将SVG图像插入到PowerPoint演示文稿中。

## 功能特点

- 将SVG文件插入到PPTX文件的指定位置
- 支持自定义图像大小和位置
- 自动转换SVG为PNG作为备用图像（用于不支持SVG的应用程序）
- 文件管理工具（列出文件、获取文件信息）
- 支持批量处理多个SVG文件
- 提供SVG到PNG的直接转换功能

## 技术架构

项目由以下主要组件构成：

1. **svg_module.py** - 核心功能模块
   - `to_emu()` - 单位转换函数
   - `insert_svg_to_pptx()` - 主要功能函数，处理SVG插入PPTX

2. **mcp_server.py** - MCP服务器和工具定义
   - 提供5个MCP工具：
     - `insert_svg` - 插入单个SVG到PPTX
     - `convert_svg_to_png` - 将SVG转换为PNG
     - `batch_insert_svgs` - 批量处理多个SVG文件
     - `list_files` - 列出目录中的文件
     - `get_file_info` - 获取文件信息

3. **requirements.txt** - 项目依赖

### 技术变更

- 使用reportlab和svglib替代cairosvg进行SVG到PNG转换
- 模块化设计，分离核心功能和接口层

## 安装步骤

1. 克隆此代码库
2. 安装依赖项：
   ```
   pip install -r requirements.txt
   ```
3. 运行MCP服务器：
   ```
   python mcp_server.py
   ```

## 配置Claude Desktop

要在Claude Desktop中使用此服务器：

1. 编辑Claude Desktop配置文件：
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - MacOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. 添加服务器配置：
   ```json
   {
     "mcpServers": {
       "svg-to-pptx": {
         "command": "python",
         "args": [
           "C:\\绝对路径\\到项目\\mcp_server.py"
         ]
       }
     }
   }
   ```

3. 重启Claude Desktop

## 使用方法

在Claude中，您可以使用以下命令：

### 列出文件

```
请列出当前目录中的所有SVG文件
```

### 插入SVG到PPTX

```
请将example.svg插入到presentation.pptx的第1张幻灯片中，位置在(1, 1)英寸处，宽度为4英寸
```

### 查看文件信息

```
请告诉我example.svg的信息
```

### 转换SVG到PNG

```
请将example.svg转换为PNG格式
```

### 批量处理SVG文件

```
请将svg_folder目录中的所有SVG文件插入到presentation.pptx中
```

## 开发者指南

### 添加新功能

要添加新功能，您可以：

1. 在`svg_module.py`中添加核心功能
2. 在`mcp_server.py`中使用`@mcp.tool()`装饰器定义新工具

### 运行测试

```
python -m unittest discover tests
```

## 依赖项

- MCP (Model Context Protocol) SDK >=1.8.0
- lxml >=4.9.2 - XML处理
- python-pptx >=0.6.21 - PowerPoint文件操作
- reportlab >=4.0.0 - 图形处理
- svglib >=1.5.1 - SVG解析和转换

## 故障排除

- **问题**: SVG文件无法转换为PNG
  **解决方案**: 确保SVG文件格式正确，且不含有svglib不支持的特性

- **问题**: 无法连接到MCP服务器
  **解决方案**: 检查Claude Desktop配置文件中的路径是否正确