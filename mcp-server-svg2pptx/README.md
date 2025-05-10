# SVG-to-PPTX MCP服务器

一个基于Model Context Protocol (MCP)的服务器，允许通过Claude等大型语言模型将SVG图像插入到PowerPoint演示文稿中。

## 功能特点

- 将SVG文件插入到PPTX文件的指定位置
- 支持自定义图像大小和位置
- 自动转换SVG为PNG作为备用图像（用于不支持SVG的应用程序）
- 自动创建足够的幻灯片数量（当指定幻灯片编号超出现有范围时）
- 查看PPTX文件的幻灯片数量和基本信息
- 文件管理工具（列出文件、获取文件信息）
- 支持批量处理多个SVG文件
- 提供SVG到PNG的直接转换功能

## 技术架构

项目由以下主要组件构成：

1. **svg_module.py** - 核心功能模块
   - `to_emu()` - 单位转换函数
   - `insert_svg_to_pptx()` - 主要功能函数，处理SVG插入PPTX
   - `get_pptx_slide_count()` - 获取PPTX文件的幻灯片数量

2. **main.py** - MCP服务器和工具定义
   - 提供6个MCP工具：
     - `insert_svg` - 插入单个SVG到PPTX
     - `convert_svg_to_png` - 将SVG转换为PNG
     - `list_files` - 列出目录中的文件
     - `get_file_info` - 获取文件信息
     - `get_pptx_info` - 获取PPTX文件的页数和信息

3. **requirements.txt** - 项目依赖

### 技术变更

- 使用reportlab和svglib替代cairosvg进行SVG到PNG转换
- 模块化设计，分离核心功能和接口层
- 增强错误处理，提供详细错误信息
- 自动创建路径不存在的目录
- 自动创建缺失的幻灯片（当slide_number大于现有幻灯片数量时）

## 安装步骤

1. 克隆此代码库
2. 安装依赖项：
   ```
   pip install -r requirements.txt
   ```
3. 运行MCP服务器：
   ```
   python main.py
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
           "C:\\绝对路径\\到项目\\main.py"
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

### 查看PPT页数和信息

```
请查看presentation.pptx的幻灯片数量和信息
```

### 转换SVG到PNG

```
请将example.svg转换为PNG格式
```

## 特殊功能说明

### 自动创建幻灯片

当您指定的幻灯片编号超出现有幻灯片数量时，系统会自动创建足够的幻灯片。例如：

```
请将example.svg插入到presentation.pptx的第5张幻灯片中
```

如果presentation.pptx只有2张幻灯片，系统会自动创建第3、4、5张幻灯片，然后将SVG插入到第5张。

### 获取PPTX信息

您可以使用`get_pptx_info`工具获取PPTX文件的详细信息：

```
请查看presentation.pptx的信息
```

输出示例：
```
PPT文件: D:\example\presentation.pptx
大小: 2.45 MB
修改时间: 2024-05-20 14:30:25
幻灯片数量：12张
```

## 开发者指南

### 添加新功能

要添加新功能，您可以：

1. 在`svg_module.py`中添加核心功能
2. 在`main.py`中使用`@mcp.tool()`装饰器定义新工具

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

- **问题**: "Package not found at 'path'"错误
  **解决方案**: 此问题已修复，系统现在会自动创建不存在的父目录