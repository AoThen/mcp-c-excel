# ExcelMcp - MCP Server for Microsoft Excel

[![VS Code Marketplace Installs](https://img.shields.io/visual-studio-marketplace/i/sbroenne.excel-mcp?label=VS%20Code%20Installs)](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)
[![Downloads](https://img.shields.io/github/downloads/sbroenne/mcp-server-excel/total?label=GitHub%20Downloads)](https://github.com/sbroenne/mcp-server-excel/releases)

[![Build MCP Server](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-mcp-server.yml)
[![Build CLI](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml/badge.svg)](https://github.com/sbroenne/mcp-server-excel/actions/workflows/build-cli.yml)
[![Release](https://img.shields.io/github/v/release/sbroenne/mcp-server-excel)](https://github.com/sbroenne/mcp-server-excel/releases/latest)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10-blue.svg)](https://dotnet.microsoft.com/download/dotnet/10.0)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://github.com/sbroenne/mcp-server-excel)
[![Built with Copilot](https://img.shields.io/badge/Built%20with-GitHub%20Copilot-0366d6.svg)](https://copilot.github.com/)

**Automate Excel with AI - A Model Context Protocol (MCP) server for comprehensive Excel automation through conversational AI.**

**MCP Server for Excel** enables AI assistants (GitHub Copilot, Claude, ChatGPT) to automate Excel through natural language commands. Automate Power Query, DAX measures, VBA macros, PivotTables, Charts, formatting, and data transformations (25 tools with 230 operations).

**🛡️ 100% Safe - Uses Excel's Native COM API** - Zero risk of file corruption. Unlike third-party libraries that manipulate `.xlsx` files directly, this project uses Excel's official API ensuring complete safety and compatibility.

**💡 Interactive Development** - See results instantly in Excel. Create a query, run it, inspect the output, refine and repeat. Excel becomes your AI-powered workspace for rapid development and testing.

**🧪 LLM-Tested Quality** - Tool behavior validated with real LLM workflows using [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering). We test that LLMs correctly understand and use our tools.

**Technical Requirements:**
- ⚠️ **Windows Only** - COM interop is Windows-specific
- ⚠️ **Excel Required** - Microsoft Excel 2016 or later must be installed
- ⚠️ **Desktop Environment** - Controls actual Excel process (not for server-side processing)

## 🎯 What You Can Do

**25 specialized tools with 230 operations:**

- 🔄 **Power Query** (1 tool, 12 ops) - Atomic workflows, M code management, load destinations
- 📊 **Data Model/DAX** (2 tools, 19 ops) - Measures with auto-formatted DAX, relationships, model structure
- 🎨 **Excel Tables** (2 tools, 27 ops) - Lifecycle, filtering, sorting, structured references
- 📈 **PivotTables** (3 tools, 30 ops) - Creation, fields, aggregations, calculated members/fields
- 📉 **Charts** (2 tools, 29 ops) - Create, configure, series, formatting, data labels, trendlines
- 📝 **VBA** (1 tool, 6 ops) - Modules, execution, version control
- 📋 **Ranges** (4 tools, 46 ops) - Values, formulas, formatting, validation, protection
- 📄 **Worksheets** (2 tools, 16 ops) - Lifecycle, colors, visibility, cross-workbook moves
- 🔌 **Connections** (1 tool, 9 ops) - OLEDB/ODBC management and refresh
- 🏷️ **Named Ranges** (1 tool, 6 ops) - Parameters and configuration
- 📁 **Files** (1 tool, 6 ops) - Session management, workbook creation, IRM/AIP-protected file support
- 🧮 **Calculation Mode** (1 tool, 3 ops) - Get/set calculation mode and trigger recalculation
- 🎚️ **Slicers** (1 tool, 8 ops) - Interactive filtering for PivotTables and Tables
- 🎨 **Conditional Formatting** (1 tool, 2 ops) - Rules and clearing
- 📸 **Screenshot** (1 tool, 2 ops) - Capture ranges/sheets as PNG for LLM visual verification
- 🪧 **Window Management** (1 tool, 9 ops) - Show/hide Excel, arrange, position, status bar feedback

📚 **[Complete Feature Reference →](FEATURES.md)** - Detailed documentation of all 230 operations


## 💬 Example Prompts

**Create & Populate Data:**
- *"Create a new Excel file called SalesTracker.xlsx with a table for Date, Product, Quantity, Unit Price, and Total with sample data"*
- *"Put this data in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*
- *"Add a formula column that calculates Quantity times Unit Price"*

**Analysis & Visualization:**
- *"Create a PivotTable from this data showing total sales by Product, then add a bar chart"*
- *"Use Power Query to import products.csv, load it to the Data Model, and create a measure for Total Revenue"*
- *"Create a slicer for the Region field so I can filter the PivotTable interactively"*
- *"Create a relationship between the Orders and Products tables using ProductID"*

**Formatting & Styling:**
- *"Format the Price column as currency and highlight values over $500 in green"*
- *"Convert this range to an Excel Table with a blue style and add a totals row"*
- *"Make the headers bold with a dark background and auto-fit column widths"*
- *"Apply the same section-header styling to A1:G1, A12:G12, and A24:G24 in one step"*

Formatting split: number display formats use the `range` tool, while visual styling and auto-fit use `range_format`.

**Automation:**
- *"Export all Power Query M code to files for version control"*
- *"Run the UpdatePrices macro"*
- *"Show me Excel while you work"* - watch changes in real-time

**🪟 Agent Mode — Watch AI Work in Excel:**
- *"Show me Excel side-by-side while you build this dashboard"* - real-time visibility
- *"Let me watch while you create the chart"* - AI asks your preference, then shows Excel
- Status bar shows live progress: *"ExcelMcp: Building PivotTable from Sales data..."*

## 👥 Who Should Use This?

**Perfect for:**
- ✅ **Data analysts** automating repetitive Excel workflows
- ✅ **Developers** building Excel-based data solutions
- ✅ **Business users** managing complex Excel workbooks
- ✅ **Teams** maintaining Power Query/VBA/DAX code in Git

**Not suitable for:**
- ❌ Server-side data processing (use libraries like ClosedXML, EPPlus instead)
- ❌ Linux/macOS users (Windows + Excel installation required)
- ❌ High-volume batch operations (consider Excel-free alternatives)


## 🚀 Quick Start

| Platform | Installation |
|----------|-------------|
| **VS Code** | [Install Extension](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp) (one-click, recommended) |
| **Claude Desktop** | Download `.mcpb` from [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) |
| **Any MCP Client** | Download `mcp-excel.exe` from [latest release](https://github.com/sbroenne/mcp-server-excel/releases/latest) and add to PATH |
| **Details** | 📖 [Installation Guide](docs/INSTALLATION.md) |

**⚠️ Important:** Close all Excel files before using. The server requires exclusive access to workbooks during automation.


## 🔧 CLI vs MCP Server

This package provides both **CLI** and **MCP Server** interfaces. Choose based on your use case:

| Interface | Best For | Why |
|-----------|----------|-----|
| **CLI** (`excelcli`) | Coding agents (Copilot, Cursor, Windsurf) | **64% fewer tokens** - single tool, no large schemas. Auto-generated from Core code, ensuring 1:1 feature parity. |
| **MCP Server** | Conversational AI (Claude Desktop, VS Code Chat) | Rich tool discovery, persistent connection. Better for interactive, exploratory workflows. |

**⚡ CLI Commands:** Generated automatically from Core service definitions using Roslyn source generators. All 22 command categories maintain exact 1:1 parity with MCP tools through shared code generation. See [code generation docs](docs/DEVELOPMENT.md#-cli-command-code-generation) for details.

<details>
<summary>📊 Benchmark Results (same task, same model)</summary>

| Metric | CLI | MCP Server | Winner |
|--------|-----|------------|--------|
| **Tokens** | ~59K | ~163K | 🏆 CLI (64% fewer) |

**Key insight:** MCP sends 23 tool schemas to the LLM on each request (~100K+ tokens).

</details>

**Manual Installation:**
```powershell
# Primary: Download standalone executables from latest release (no .NET runtime required)
# https://github.com/sbroenne/mcp-server-excel/releases/latest
# - ExcelMcp-MCP-Server-{version}-windows.zip → extract mcp-excel.exe
# - ExcelMcp-CLI-{version}-windows.zip → extract excelcli.exe (optional, for scripting)

# Secondary: Install via .NET tool (requires .NET 10 runtime)
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# After installing either way, auto-configure all your coding agents:
npx add-mcp "mcp-excel" --name excel-mcp
```

> ⚠️ **Step 2 requires [Node.js](https://nodejs.org/)** for `npx`. Install with `winget install OpenJS.NodeJS.LTS` if needed.

```powershell
# Optional: Install agent skills for better AI guidance
npx skills add sbroenne/mcp-server-excel --skill excel-cli   # Coding agents
npx skills add sbroenne/mcp-server-excel --skill excel-mcp   # Conversational AI
```

> 💡 **Skills provide AI guidance** - The CLI skill is highly recommended (agents don't work perfectly with CLI without it). The MCP skill is recommended - it adds workflow best practices and reduces token usage.


## ⚙️ How It Works - COM Automation & Unified Service Architecture

**ExcelMcp uses Windows COM automation to control the actual Excel application (not just .xlsx files).**

Both the **MCP Server** and **CLI** communicate with a shared **ExcelMCP Service** that manages Excel sessions. This unified architecture enables:

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server        │     │   CLI (excelcli)    │
│  (AI assistants)    │     │  (coding agents)    │
└─────────┬───────────┘     └─────────┬───────────┘
          │                           │
          └──────────┬────────────────┘
                     ▼
          ┌─────────────────────────┐
          │   ExcelMCP Service      │
          │  (shared session mgmt)  │
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   Excel COM API         │
          │  (Excel.Application)    │
          └─────────────────────────┘
```

**Key Benefits:**
- ✅ **Shared Sessions** - CLI and MCP Server can access the same open workbooks
- ✅ **Single Excel Instance** - No duplicate Excel processes or file locks
- ✅ **System Tray UI** - Monitor active sessions via the ExcelMCP tray icon

**💡 Tip: Watch Excel While AI Works**
By default, Excel runs hidden for faster automation. To see changes in real-time, just ask:
- *"Show me Excel while you work"*
- *"Let me watch what you're doing"*
- *"Open Excel so I can see the changes"*

The AI will display the Excel window so you can watch every operation happen live - great for learning or verifying changes!

## 📋 Additional Information

📚 **[CLI Guide →](src/ExcelMcp.CLI/README.md)** | **[CLI Skill for Agents →](skills/excel-cli/SKILL.md)** | **[MCP Server Guide →](src/ExcelMcp.McpServer/README.md)** | **[All Agent Skills →](skills/README.md)**

**License:** MIT License - see [LICENSE](LICENSE) file

**Privacy:** See [PRIVACY.md](PRIVACY.md) for our privacy policy

**Contributing:** See [CONTRIBUTING.md](docs/CONTRIBUTING.md) for guidelines

**Built With:** This entire project was developed using GitHub Copilot AI assistance - mainly with Claude but lately with Auto-mode.

**Acknowledgments:**
- Microsoft Excel Team - For comprehensive COM automation APIs
- Model Context Protocol community - For the AI integration standard
- Open Source Community - For inspiration and best practices

## Related Projects

Other projects by the author:

- [pytest-skill-engineering](https://github.com/sbroenne/pytest-skill-engineering) — LLM-powered testing framework for AI agents
- [Windows MCP Server](https://windowsmcpserver.dev/) — AI-powered Windows automation via MCP
- [OBS Studio MCP Server](https://github.com/sbroenne/mcp-server-obs) — AI-powered OBS Studio automation
- [HeyGen MCP Server](https://github.com/sbroenne/heygen-mcp) — MCP server for HeyGen AI video generation

---

# 🇨🇳 中文说明

## 项目简介

**ExcelMcp** 是一个基于模型上下文协议（Model Context Protocol，MCP）的 Excel 自动化服务器，让 AI 助手（如 GitHub Copilot、Claude、ChatGPT）能够通过自然语言指令自动操作 Excel。

**核心优势：**
- 🛡️ **100% 安全** — 使用 Excel 原生 COM API，零文件损坏风险
- 💡 **交互式开发** — 在 Excel 中实时查看结果，边操作边验证
- 🧪 **LLM 测试验证** — 使用真实 LLM 工作流验证工具行为

## 功能概览

**25 个专业工具，230+ 操作：**

| 类别 | 工具数 | 操作数 | 主要功能 |
|------|--------|--------|----------|
| 🔄 Power Query | 1 | 12 | M 代码管理、导入加载、原子工作流 |
| 📊 数据模型/DAX | 2 | 19 | 度量值、关系、模型结构、DAX 格式化 |
| 🎨 Excel 表格 | 2 | 27 | 生命周期、筛选、排序、结构化引用 |
| 📈 数据透视表 | 3 | 30 | 创建、字段、聚合、计算字段/成员 |
| 📉 图表 | 2 | 29 | 创建、配置、系列、格式化、趋势线 |
| 📝 VBA 宏 | 1 | 6 | 模块、执行、版本控制 |
| 📋 单元格范围 | 4 | 46 | 值、公式、格式、验证、保护 |
| 📄 工作表 | 2 | 16 | 生命周期、颜色、可见性、跨工作簿移动 |
| 🔌 数据连接 | 1 | 9 | OLEDB/ODBC 管理和刷新 |
| 🏷️ 命名区域 | 1 | 6 | 参数和配置管理 |
| 📁 文件操作 | 1 | 6 | 会话管理、工作簿创建 |
| 🧮 计算模式 | 1 | 3 | 获取/设置计算模式、触发重算 |
| 🎚️ 切片器 | 1 | 8 | 数据透视表和表格的交互式筛选 |
| 🎨 条件格式 | 1 | 2 | 规则创建和清除 |
| 📸 截图 | 1 | 2 | 捕获区域/工作表为 PNG 图像 |
| 🪧 窗口管理 | 1 | 9 | 显示/隐藏 Excel、排列、位置、状态栏 |

## 系统要求

- ⚠️ **仅限 Windows** — COM 互操作是 Windows 特有的
- ⚠️ **需要 Excel** — 必须安装 Microsoft Excel 2016 或更高版本
- ⚠️ **桌面环境** — 控制实际的 Excel 进程（不适用于服务器端处理）

## 快速开始

### 安装方式

| 平台 | 安装方式 |
|------|----------|
| **VS Code** | [安装扩展](https://marketplace.visualstudio.com/items?itemName=sbroenne.excel-mcp)（一键安装，推荐） |
| **Claude Desktop** | 从[最新发布版](https://github.com/sbroenne/mcp-server-excel/releases/latest)下载 `.mcpb` 文件 |
| **任意 MCP 客户端** | 从[最新发布版](https://github.com/sbroenne/mcp-server-excel/releases/latest)下载 `mcp-excel.exe` 并添加到 PATH |

### 手动安装

```powershell
# 方式一：从最新发布版下载独立可执行文件（无需 .NET 运行时）
# https://github.com/sbroenne/mcp-server-excel/releases/latest
# - ExcelMcp-MCP-Server-{version}-windows.zip → 解压得到 mcp-excel.exe
# - ExcelMcp-CLI-{version}-windows.zip → 解压得到 excelcli.exe（可选，用于脚本）

# 方式二：通过 .NET 工具安装（需要 .NET 10 运行时）
dotnet tool install --global Sbroenne.ExcelMcp.McpServer
dotnet tool install --global Sbroenne.ExcelMcp.CLI

# 安装后，自动配置所有编程代理：
npx add-mcp "mcp-excel" --name excel-mcp
```

## 使用示例

### 创建和填充数据

```
"创建一个名为 SalesTracker.xlsx 的新 Excel 文件，包含日期、产品、数量、单价和总计的表格，并填充示例数据"
"将以下数据放入 A1:C4 - 姓名、年龄、城市 / 张三, 30, 北京 / 李四, 25, 上海"
"添加一个公式列，计算数量乘以单价"
```

### 分析和可视化

```
"从此数据创建数据透视表，按产品显示总销售额，然后添加柱状图"
"使用 Power Query 导入 products.csv，加载到数据模型，并创建总收入度量值"
"为地区字段创建切片器，以便交互式筛选数据透视表"
"使用产品ID在订单和产品表之间创建关系"
```

### 格式化和样式

```
"将价格列格式化为货币，并将超过500的值以绿色高亮显示"
"将此范围转换为蓝色样式的 Excel 表格，并添加汇总行"
"将标题设为粗体，深色背景，并自动调整列宽"
```

### 自动化

```
"导出所有 Power Query M 代码到文件进行版本控制"
"运行 UpdatePrices 宏"
"让我看看你在做什么" - 实时查看 Excel 变化
```

## CLI vs MCP 服务器

| 接口 | 适用场景 | 原因 |
|------|----------|------|
| **CLI** (`excelcli`) | 编程代理（Copilot、Cursor、Windsurf） | **节省 64% token** — 单一工具，无大型模式 |
| **MCP 服务器** | 对话式 AI（Claude Desktop、VS Code Chat） | 丰富的工具发现，持久连接，适合交互式探索工作流 |

## 架构说明

```
┌─────────────────────┐     ┌─────────────────────┐
│   MCP Server        │     │   CLI (excelcli)    │
│  (AI 助手)          │     │  (编程代理)         │
└─────────┬───────────┘     └─────────┬───────────┘
          │                           │
          └──────────┬────────────────┘
                     ▼
          ┌─────────────────────────┐
          │   ExcelMCP Service      │
          │  (共享会话管理)         │
          └─────────┬───────────────┘
                    ▼
          ┌─────────────────────────┐
          │   Excel COM API         │
          │  (Excel.Application)    │
          └─────────────────────────┘
```

**主要优势：**
- ✅ **共享会话** — CLI 和 MCP 服务器可以访问相同的打开工作簿
- ✅ **单一 Excel 实例** — 无重复的 Excel 进程或文件锁定
- ✅ **系统托盘 UI** — 通过 ExcelMCP 托盘图标监控活动会话

## 适用人群

**适合：**
- ✅ 自动化重复 Excel 工作流的数据分析师
- ✅ 构建 Excel 数据解决方案的开发者
- ✅ 管理复杂 Excel 工作簿的业务用户
- ✅ 在 Git 中维护 Power Query/VBA/DAX 代码的团队

**不适合：**
- ❌ 服务器端数据处理（请使用 ClosedXML、EPPlus 等库）
- ❌ Linux/macOS 用户（需要 Windows + Excel 安装）
- ❌ 高容量批量操作（请考虑无 Excel 的替代方案）

## 文档资源

- **[安装指南](docs/INSTALLATION.md)** — 所有 AI 助手的设置
- **[MCP 服务器指南](src/ExcelMcp.McpServer/README.md)** — 工具文档和示例
- **[CLI 指南](src/ExcelMcp.CLI/README.md)** — 命令行参考
- **[代理技能](skills/README.md)** — 跨平台 AI 助手指导
- **[贡献指南](docs/CONTRIBUTING.md)** — 开发指南

## 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件
