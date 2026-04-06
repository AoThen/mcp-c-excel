# ExcelMcp - Complete Feature Reference

**25 specialized tools with 230 operations for comprehensive Excel automation**

---

## 📁 File Operations (6 operations)

- **List Sessions:** View all active Excel sessions
- **Open:** Open workbook and create session (returns session ID for all subsequent operations). IRM/AIP-protected files are automatically detected and opened read-only with Excel visible for credential authentication — no extra parameters needed.
- **Close:** Close session with optional save
- **Close Workbook:** Close workbook without closing Excel
- **Create Empty:** Create new .xlsx or .xlsm workbook
- **Test:** Verify workbook can be opened and is accessible. Returns `isIrmProtected` flag for IRM/AIP-protected files.

---

## 🧮 Calculation Mode (3 operations)

- **Get Mode:** Query current calculation mode and calculation state
- **Set Mode:** Switch between automatic, manual, and semi-automatic modes
- **Calculate:** Explicitly recalculate workbook, sheet, or range

---

## 🔄 Power Query & M Code (12 operations)

**Atomic Operations** - Single-call workflows:
- **List:** List all Power Query queries in workbook
- **View:** View the M code of a Power Query
- **Create:** Import + load in one operation (atomic workflow) with automatic formatting
- **Update:** Update M code with automatic formatting and auto-refresh
- **Rename:** Rename a Power Query (trim + case-insensitive uniqueness check)
- **Refresh:** Refresh a Power Query with timeout detection
- **Refresh All:** Batch refresh all queries in workbook
- **Load To:** Configure load destination and refresh (atomic)
- **Get Load Config:** Get current load configuration
- **Unload:** Remove data from all destinations (keeps query definition)
- **Delete:** Remove Power Query from workbook
- **Evaluate:** Execute M code directly and return results (without creating a permanent query)

**Automatic M-Code Formatting:** M code is automatically formatted on write operations (Create, Update) using the powerqueryformatter.com API (by mogularGmbH, MIT License). Read operations return M code as stored in Excel. Formatting adds ~100-500ms network latency but dramatically improves readability with proper indentation, spacing, and line breaks. Graceful fallback returns original M code if formatting fails.

---

## 📊 Data Model & DAX (Power Pivot) (19 operations)

- **List Tables:** Discover all tables in the Data Model
- **Read Table:** Get specific table information
- **Rename Table:** Rename a Data Model table (best-effort via Power Query; returns clear error if not supported)
- **List Columns:** List columns for a table
- **List Measures:** List all DAX measures with formula previews
- **Read Info:** Get comprehensive model information
- **Create Measure:** Create new DAX measure with automatic formatting (format types: Currency, Percentage, Decimal, General)
- **Update Measure:** Modify existing measure with automatic formatting
- **Delete Measure:** Remove measure from model
- **Delete Table:** Remove table from Data Model
- **List Relationships:** View all table relationships
- **Read Relationship:** Get specific relationship info
- **Create Relationship:** Create relationship between tables
- **Update Relationship:** Modify relationship (toggle active/inactive)
- **Delete Relationship:** Remove relationship
- **Refresh:** Refresh entire Data Model
- **List Workbook Connections:** List Power Query sources available for integration
- **Evaluate:** Execute DAX EVALUATE queries and return tabular results (for ad-hoc analysis)
- **Execute DMV:** Execute SQL-like DMV (Dynamic Management View) queries for metadata discovery

**Automatic DAX Formatting:** DAX formulas are automatically formatted on write operations (CreateMeasure, UpdateMeasure) using the official Dax.Formatter library (SQLBI). Read operations return raw DAX as stored in Excel. Formatting adds ~100-500ms network latency but dramatically improves readability. Graceful fallback returns original DAX if formatting fails.

**Note:** DAX calculated columns not supported - use Excel UI for calculated columns

---

## 🎨 Excel Tables (ListObjects) (27 operations)

**Lifecycle:**
- List, read, create, rename, resize, delete tables

**Styling & Formatting:**
- Apply table styles
- Toggle totals row
- Set column totals

**Data Operations:**
- Append rows
- Get table data (with optional visible-only filtering)
- Add to Data Model

**DAX-Backed Tables:**
- Create from DAX (create Excel Table populated by DAX EVALUATE query)
- Update DAX (change the DAX query of an existing DAX-backed table)
- Get DAX (retrieve DAX query info from a table)

**Filter Operations:**
- Apply filter (criteria)
- Apply filter (values)
- Clear filters
- Get filter state

**Column Management:**
- Add, remove, rename columns

**Structured References:**
- Get structured reference (formula syntax for table columns/ranges)

**Sorting:**
- Single-column sort
- Multi-column sort (up to 3 levels)

**Number Formatting:**
- Get column number formats
- Set column number formats

---

## 📈 PivotTables (30 operations)

**Creation:**
- Create from range
- Create from Excel Table
- Create from Data Model

**Field Management:**
- List all fields (row, column, value, filter areas)
- Add row field, column field, value field, filter field
- Remove field

**Field Configuration:**
- Set field aggregation function (Sum, Average, Count, Min, Max, etc.)
- Set custom field name
- Set field number format
- Set field filter criteria
- Sort field (ascending/descending)

**Calculated Fields (Regular PivotTables):**
- List calculated fields
- Create calculated field
- Delete calculated field

**Calculated Members (OLAP/Data Model PivotTables):**
- List calculated members
- Create calculated member
- Delete calculated member

**Layout & Formatting:**
- Set layout (table or outline)
- Set subtotals display
- Set grand totals display

**Data Operations:**
- Get PivotTable data as 2D array
- Refresh PivotTable

**Lifecycle:**
- List PivotTables
- Read PivotTable info
- Delete PivotTable

---

## 📉 Charts (28 operations)

**Creation:**
- Create from range
- Create from PivotTable

**Series Management:**
- Add series
- Remove series
- Update series data

**Configuration:**
- Set data source range
- Set chart type
- Show/hide legend
- Set style

**Formatting:**
- Set chart title
- Set axis title
- Set axis number format
- Get axis number format

**Data Labels:**
- Configure data labels (show values, percentages, category names, etc.)
- Set label position (Center, InsideEnd, OutsideEnd, etc.)
- Apply to all series or specific series

**Axis Scale:**
- Get axis scale settings
- Set minimum/maximum scale
- Set major/minor units

**Gridlines:**
- Get gridlines configuration
- Set major/minor gridlines visibility

**Series Formatting:**
- Set marker style (Circle, Square, Diamond, Triangle, etc.)
- Set marker size
- Set marker colors

**Trendlines:**
- Add trendline (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)
- List trendlines on series
- Delete trendline
- Configure trendline (forecast forward/backward, display equation, display R²)

**Placement & Positioning:**
- Set chart placement (move/size with cells options)
- Fit to range (position and size to match a range)

**Lifecycle:**
- List charts
- Read chart info
- Move chart (to different worksheet or new sheet)
- Delete chart

---

## 📋 Ranges (46 operations)

Formatting split: use `range` for number display formats such as dates, currency, percentages, and text display. Use `range_format` for visual styling, validation, auto-fit, and size/layout changes.

**Data Operations:**
- Get values
- Set values
- Get formulas
- Set formulas
- Clear all
- Clear contents
- Clear formats
- Copy
- Copy values
- Copy formulas
- Insert cells
- Delete cells
- Insert rows
- Delete rows
- Insert columns
- Delete columns
- Find
- Replace
- Sort

**Discovery & Utilities:**
- Get used range
- Get current region
- Get range info (address, dimensions)

**Hyperlinks:**
- Add hyperlink
- Remove hyperlink
- List hyperlinks
- Get specific hyperlink

**Number Formatting (`range`):**
- Get number formats (as 2D array)
- Set number format (uniform)
- Set number formats (individual)

**Visual Formatting (`range_format`):**
- Get style
- Set style (built-in Excel styles)
- Format range (font, color, borders, alignment, orientation)
- Format multiple ranges with one shared formatting payload

**Data Validation (`range_format`):**
- Add validation rules (dropdowns, number/date/text rules)
- Get validation info
- Remove validation

**Merge Operations (`range_format`):**
- Merge cells
- Unmerge cells
- Get merge info

**Cell Protection:**
- Set cell lock status
- Get cell lock status

**Auto-Sizing (`range_format`):**
- Auto-fit columns
- Auto-fit rows

---

## 📄 Worksheets (16 operations)

**Lifecycle:**
- List worksheets
- Create worksheet
- Rename worksheet
- Copy worksheet
- Move worksheet
- Delete worksheet

**Cross-Workbook Operations:**
- Copy worksheet to file (atomic)
- Move worksheet to file (atomic)

**Tab Colors:**
- Set tab color (RGB)
- Get tab color
- Clear tab color

**Visibility:**
- Show worksheet
- Hide worksheet
- Very hide worksheet (hidden from UI)
- Get visibility status
- Set visibility status

---

## 🔌 Data Connections (9 operations)

- **List:** View all data connections
- **View:** Get connection details
- **Create:** Create OLEDB/ODBC connections (requires provider installed)
- **Test:** Verify connection validity
- **Refresh:** Refresh connection data
- **Delete:** Remove connection
- **Load To:** Load connection data to worksheet (when supported)
- **Get Properties:** Get connection string and metadata
- **Set Properties:** Update connection string, command text, and settings

**Supported Types:**
- OLEDB (requires Microsoft.ACE.OLEDB.16.0 or similar)
- ODBC (requires ODBC driver installed)
- Power Query connections (atomic redirect to powerquery)

**Automatic Fallback:**
- TEXT/WEB connections automatically redirect to powerquery for reliable imports

---

## 🏷️ Named Ranges (Parameters) (6 operations)

- **List:** List all named ranges with references
- **Read:** Get value of a named range
- **Write:** Set value of a named range (ideal for parameter automation)
- **Create:** Create new named range
- **Update:** Modify existing named range
- **Delete:** Remove named range

**Use Cases:**
- Workbook parameter management without touching worksheets
- Ideal for automation: update parameter → Power Query refreshes automatically

---

## 📝 VBA Macros (6 operations)

- **List:** List VBA components and discovered procedures
- **View:** Display component code without exporting
- **Import:** Create a new standard module from code or file input
- **Update:** Replace code in an existing VBA component
- **Delete:** Remove a VBA component by name
- **Run:** Execute a procedure with optional string parameters

**Features:**
- Procedural/module-focused VBA support for `.xlsm` workbooks
- Manual VBA trust prerequisite in Excel (no trust-configuration command)
- Import creates standard modules; list/view also cover class, form, and document components

---

## �️ Slicers (8 operations)

**PivotTable Slicers:**
- **Create Slicer:** Add slicer for PivotTable field with optional position
- **List Slicers:** List all PivotTable slicers in workbook
- **Set Selection:** Filter PivotTable by slicer selection (single or multi-select)
- **Delete Slicer:** Remove PivotTable slicer

**Table Slicers:**
- **Create Table Slicer:** Add slicer for Excel Table column
- **List Table Slicers:** List all Table slicers in workbook
- **Set Table Selection:** Filter Table by slicer selection
- **Delete Table Slicer:** Remove Table slicer

**Use Cases:**
- Interactive data filtering without modifying PivotTable/Table structure
- Dashboard creation with visual filter controls
- Multi-slicer filtering for complex data analysis

---

## �🎨 Conditional Formatting (2 operations)

- **Add Rule:** Create conditional formatting rules
  - Cell value comparisons (>, <, =, etc.)
  - Expression-based formulas (custom DAX/Excel formulas)
  - Color scales, data bars, icons
- **Clear Rules:** Remove formatting from ranges

---

## 📸 Screenshot (2 operations)

- **Capture Range:** Capture a specific range as a PNG image
- **Capture Sheet:** Capture the entire used area of a worksheet as a PNG image
  - Uses Excel's built-in rendering (CopyPicture) — captures formatting, charts, conditional formatting
  - MCP: Returns image directly as ImageContent (base64 PNG)
  - CLI: Returns JSON with base64-encoded image data

---

## 🪧 Window Management (9 operations)

- **Show:** Makes Excel visible and brings it to the foreground
- **Hide:** Hides the Excel window
- **Bring to Front:** Brings Excel to the foreground without changing visibility
- **Get Info:** Gets current window state (visibility, position, size, foreground status)
- **Set State:** Sets window state to normal, minimized, or maximized
- **Set Position:** Sets window position and size in points (left, top, width, height)
- **Arrange:** Arranges Excel window using preset layouts
- **Set Status Bar:** Displays custom text in Excel's status bar for real-time feedback
- **Clear Status Bar:** Restores the default status bar text

**Arrange Presets:**
- `left-half` / `right-half` — Side-by-side with other applications
- `top-half` / `bottom-half` — Stacked view
- `center` — Centered window (60% of screen)
- `full-screen` — Maximized

**Use Cases:**
- Interactive "agent mode" where users watch Excel respond to AI commands in real-time
- Side-by-side: Excel on one half, AI assistant on the other
- Visibility changes are reflected in session metadata (session list shows updated state)

---

## 📊 Total Operations Summary

| Category | Operations |
|----------|-----------|
| File Operations | 6 |
| Power Query | 12 |
| Data Model/DAX | 19 |
| Excel Tables | 27 |
| PivotTables | 30 |
| Charts | 29 |
| Ranges | 46 |
| Worksheets | 16 |
| Connections | 9 |
| Named Ranges | 6 |
| VBA Macros | 6 |
| Slicers | 8 |
| Conditional Formatting | 2 |
| Screenshot | 2 |
| Calculation Mode | 3 |
| Window Management | 9 |
| **Total** | **230** |

---

## 🚀 Key Capabilities

**Data Transformation:**
- Comprehensive Power Query M code management
- Atomic import + load workflows
- Calculated fields and members for analysis

**Data Model:**
- Full DAX measure lifecycle
- Relationship management
- Multi-table integration

**Analysis & Visualization:**
- PivotTable creation and configuration
- Chart automation
- Custom calculations

**Automation:**
- VBA macro execution and management
- Named range parameter automation
- Conditional formatting rules

**Data Loading:**
- Multiple connection type support
- OLEDB/ODBC management
- Power Query atomic workflows

---

## 🔧 Tool Selection Quick Reference

| Task | Tool |
|------|------|
| Import data | `powerquery` or `connection` |
| Create analysis | `pivottable` (data model-based for OLAP) |
| Visualize data | `chart` |
| Update parameters | `namedrange` (write operation) |
| Manage formulas | `range` (set-formulas) |
| Format data | `range` / `range_format` (`format-range`, `format-ranges`, `validate-range`) |
| Script automation | `vba` (run macro) |

---

## 📚 Documentation

- **[Installation Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/INSTALLATION.md)** - Setup for all AI assistants
- **[MCP Server Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.McpServer/README.md)** - Tool documentation and examples
- **[CLI Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/src/ExcelMcp.CLI/README.md)** - Command-line reference
- **[Agent Skills](https://github.com/sbroenne/mcp-server-excel/blob/main/skills/excel-mcp/SKILL.md)** - Cross-platform AI assistant guidance (agentskills.io)
- **[Contributing](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/CONTRIBUTING.md)** - Development guidelines
- **[Releases](https://github.com/sbroenne/mcp-server-excel/releases)** - Latest updates and features

---

# 🇨🇳 中文功能参考

**25 个专业工具，230+ 操作，全面的 Excel 自动化**

---

## 📁 文件操作（6 个操作）

- **列出会话：** 查看所有活动的 Excel 会话
- **打开：** 打开工作簿并创建会话（返回会话 ID 用于后续所有操作）。IRM/AIP 保护的文件会自动检测并以只读方式打开，Excel 可见以便进行凭据认证——无需额外参数。
- **关闭：** 关闭会话，可选择保存
- **关闭工作簿：** 关闭工作簿但不关闭 Excel
- **创建空白：** 创建新的 .xlsx 或 .xlsm 工作簿
- **测试：** 验证工作簿可以打开且可访问。返回 `isIrmProtected` 标志用于 IRM/AIP 保护文件。

---

## 🧮 计算模式（3 个操作）

- **获取模式：** 查询当前计算模式和计算状态
- **设置模式：** 在自动、手动和半自动模式之间切换
- **计算：** 显式重新计算工作簿、工作表或区域

---

## 🔄 Power Query & M 代码（12 个操作）

**原子操作** - 单次调用工作流：
- **列出：** 列出工作簿中的所有 Power Query 查询
- **查看：** 查看 Power Query 的 M 代码
- **创建：** 导入 + 加载一步完成（原子工作流），自动格式化
- **更新：** 更新 M 代码，自动格式化和自动刷新
- **重命名：** 重命名 Power Query（修剪 + 不区分大小写唯一性检查）
- **刷新：** 刷新 Power Query，带超时检测
- **全部刷新：** 批量刷新工作簿中的所有查询
- **加载到：** 配置加载目标并刷新（原子操作）
- **获取加载配置：** 获取当前加载配置
- **卸载：** 从所有目标移除数据（保留查询定义）
- **删除：** 从工作簿中删除 Power Query
- **执行：** 直接执行 M 代码并返回结果（不创建永久查询）

**自动 M 代码格式化：** 写操作（创建、更新）时自动格式化 M 代码，使用 powerqueryformatter.com API（mogularGmbH，MIT 许可证）。读操作返回 Excel 中存储的原始 M 代码。格式化增加约 100-500ms 网络延迟，但通过正确的缩进、间距和换行大幅提高可读性。格式化失败时优雅降级，返回原始 M 代码。

---

## 📊 数据模型 & DAX（Power Pivot）（19 个操作）

- **列出表格：** 发现数据模型中的所有表格
- **读取表格：** 获取特定表格信息
- **重命名表格：** 重命名数据模型表格（通过 Power Query 尽力支持；不支持时返回明确错误）
- **列出列：** 列出表格的列
- **列出度量值：** 列出所有 DAX 度量值及公式预览
- **读取信息：** 获取全面的模型信息
- **创建度量值：** 创建新的 DAX 度量值，自动格式化（格式类型：货币、百分比、小数、常规）
- **更新度量值：** 修改现有度量值，自动格式化
- **删除度量值：** 从模型中删除度量值
- **删除表格：** 从数据模型中删除表格
- **列出关系：** 查看所有表格关系
- **读取关系：** 获取特定关系信息
- **创建关系：** 在表格之间创建关系
- **更新关系：** 修改关系（切换活动/非活动）
- **删除关系：** 删除关系
- **刷新：** 刷新整个数据模型
- **列出工作簿连接：** 列出可用于集成的 Power Query 源
- **执行：** 执行 DAX EVALUATE 查询并返回表格结果（用于临时分析）
- **执行 DMV：** 执行类似 SQL 的 DMV（动态管理视图）查询用于元数据发现

**自动 DAX 格式化：** 写操作（创建度量值、更新度量值）时自动格式化 DAX 公式，使用官方 Dax.Formatter 库（SQLBI）。读操作返回 Excel 中存储的原始 DAX。格式化增加约 100-500ms 网络延迟，但大幅提高可读性。格式化失败时优雅降级，返回原始 DAX。

**注意：** 不支持 DAX 计算列 - 请使用 Excel UI 创建计算列

---

## 🎨 Excel 表格（ListObjects）（27 个操作）

**生命周期：**
- 列出、读取、创建、重命名、调整大小、删除表格

**样式和格式化：**
- 应用表格样式
- 切换汇总行
- 设置列汇总

**数据操作：**
- 追加行
- 获取表格数据（可选仅可见筛选）
- 添加到数据模型

**DAX 支持表格：**
- 从 DAX 创建（创建由 DAX EVALUATE 查询填充的 Excel 表格）
- 更新 DAX（更改现有 DAX 支持表格的 DAX 查询）
- 获取 DAX（从表格检索 DAX 查询信息）

**筛选操作：**
- 应用筛选（条件）
- 应用筛选（值）
- 清除筛选
- 获取筛选状态

**列管理：**
- 添加、删除、重命名列

**结构化引用：**
- 获取结构化引用（表格列/范围的公式语法）

**排序：**
- 单列排序
- 多列排序（最多 3 级）

**数字格式化：**
- 获取列数字格式
- 设置列数字格式

---

## 📈 数据透视表（30 个操作）

**创建：**
- 从范围创建
- 从 Excel 表格创建
- 从数据模型创建

**字段管理：**
- 列出所有字段（行、列、值、筛选区域）
- 添加行字段、列字段、值字段、筛选字段
- 删除字段

**字段配置：**
- 设置字段聚合函数（求和、平均值、计数、最小值、最大值等）
- 设置自定义字段名称
- 设置字段数字格式
- 设置字段筛选条件
- 排序字段（升序/降序）

**计算字段（常规数据透视表）：**
- 列出计算字段
- 创建计算字段
- 删除计算字段

**计算成员（OLAP/数据模型数据透视表）：**
- 列出计算成员
- 创建计算成员
- 删除计算成员

**布局和格式化：**
- 设置布局（表格或大纲）
- 设置小计显示
- 设置总计显示

**数据操作：**
- 获取数据透视表数据为二维数组
- 刷新数据透视表

**生命周期：**
- 列出数据透视表
- 读取数据透视表信息
- 删除数据透视表

---

## 📉 图表（28 个操作）

**创建：**
- 从范围创建
- 从数据透视表创建

**系列管理：**
- 添加系列
- 删除系列
- 更新系列数据

**配置：**
- 设置数据源范围
- 设置图表类型
- 显示/隐藏图例
- 设置样式

**格式化：**
- 设置图表标题
- 设置轴标题
- 设置轴数字格式
- 获取轴数字格式

**数据标签：**
- 配置数据标签（显示值、百分比、类别名称等）
- 设置标签位置（居中、内侧端部、外侧端部等）
- 应用于所有系列或特定系列

**坐标轴刻度：**
- 获取坐标轴刻度设置
- 设置最小/最大刻度
- 设置主要/次要单位

**网格线：**
- 获取网格线配置
- 设置主要/次要网格线可见性

**系列格式化：**
- 设置标记样式（圆形、方形、菱形、三角形等）
- 设置标记大小
- 设置标记颜色

**趋势线：**
- 添加趋势线（线性、指数、对数、多项式、幂、移动平均）
- 列出系列上的趋势线
- 删除趋势线
- 配置趋势线（前向/后向预测、显示方程、显示 R²）

**位置和定位：**
- 设置图表位置（随单元格移动/调整大小选项）
- 适应范围（位置和大小匹配范围）

**生命周期：**
- 列出图表
- 读取图表信息
- 移动图表（到不同工作表或新工作表）
- 删除图表

---

## 📋 单元格范围（46 个操作）

格式化分工：使用 `range` 处理数字显示格式（如日期、货币、百分比和文本显示）。使用 `range_format` 处理视觉样式、验证、自动调整和大小/布局更改。

**数据操作：**
- 获取值
- 设置值
- 获取公式
- 设置公式
- 全部清除
- 清除内容
- 清除格式
- 复制
- 复制值
- 复制公式
- 插入单元格
- 删除单元格
- 插入行
- 删除行
- 插入列
- 删除列
- 查找
- 替换
- 排序

**发现和工具：**
- 获取已用范围
- 获取当前区域
- 获取范围信息（地址、维度）

**超链接：**
- 添加超链接
- 删除超链接
- 列出超链接
- 获取特定超链接

**数字格式化（`range`）：**
- 获取数字格式（为二维数组）
- 设置数字格式（统一）
- 设置数字格式（单独）

**视觉格式化（`range_format`）：**
- 获取样式
- 设置样式（内置 Excel 样式）
- 格式化范围（字体、颜色、边框、对齐、方向）
- 使用一个共享格式化负载格式化多个范围

**数据验证（`range_format`）：**
- 添加验证规则（下拉列表、数字/日期/文本规则）
- 获取验证信息
- 删除验证

**合并操作（`range_format`）：**
- 合并单元格
- 取消合并单元格
- 获取合并信息

**单元格保护：**
- 设置单元格锁定状态
- 获取单元格锁定状态

**自动调整（`range_format`）：**
- 自动调整列宽
- 自动调整行高

---

## 📄 工作表（16 个操作）

**生命周期：**
- 列出工作表
- 创建工作表
- 重命名工作表
- 复制工作表
- 移动工作表
- 删除工作表

**跨工作簿操作：**
- 复制工作表到文件（原子操作）
- 移动工作表到文件（原子操作）

**标签颜色：**
- 设置标签颜色（RGB）
- 获取标签颜色
- 清除标签颜色

**可见性：**
- 显示工作表
- 隐藏工作表
- 深度隐藏工作表（从 UI 隐藏）
- 获取可见性状态
- 设置可见性状态

---

## 🔌 数据连接（9 个操作）

- **列出：** 查看所有数据连接
- **查看：** 获取连接详情
- **创建：** 创建 OLEDB/ODBC 连接（需要已安装提供程序）
- **测试：** 验证连接有效性
- **刷新：** 刷新连接数据
- **删除：** 删除连接
- **加载到：** 将连接数据加载到工作表（支持时）
- **获取属性：** 获取连接字符串和元数据
- **设置属性：** 更新连接字符串、命令文本和设置

**支持类型：**
- OLEDB（需要 Microsoft.ACE.OLEDB.16.0 或类似）
- ODBC（需要已安装 ODBC 驱动）
- Power Query 连接（原子重定向到 powerquery）

**自动回退：**
- TEXT/WEB 连接自动重定向到 powerquery 以获得可靠导入

---

## 🏷️ 命名区域（参数）（6 个操作）

- **列出：** 列出所有命名区域及引用
- **读取：** 获取命名区域的值
- **写入：** 设置命名区域的值（非常适合参数自动化）
- **创建：** 创建新的命名区域
- **更新：** 修改现有命名区域
- **删除：** 删除命名区域

**使用场景：**
- 无需触碰工作表即可管理工作簿参数
- 适合自动化：更新参数 → Power Query 自动刷新

---

## 📝 VBA 宏（6 个操作）

- **列出：** 列出 VBA 组件和发现的程序
- **查看：** 显示组件代码而不导出
- **导入：** 从代码或文件输入创建新的标准模块
- **更新：** 替换现有 VBA 组件中的代码
- **删除：** 按名称删除 VBA 组件
- **运行：** 执行程序，可选字符串参数

**特性：**
- 面向程序/模块的 VBA 支持，适用于 `.xlsm` 工作簿
- 需要在 Excel 中手动信任 VBA（无信任配置命令）
- 导入创建标准模块；列出/查看也覆盖类、窗体和文档组件

---

## 🎚️ 切片器（8 个操作）

**数据透视表切片器：**
- **创建切片器：** 为数据透视表字段添加切片器，可选位置
- **列出切片器：** 列出工作簿中的所有数据透视表切片器
- **设置选择：** 通过切片器选择筛选数据透视表（单选或多选）
- **删除切片器：** 删除数据透视表切片器

**表格切片器：**
- **创建表格切片器：** 为 Excel 表格列添加切片器
- **列出表格切片器：** 列出工作簿中的所有表格切片器
- **设置表格选择：** 通过切片器选择筛选表格
- **删除表格切片器：** 删除表格切片器

**使用场景：**
- 无需修改数据透视表/表格结构即可进行交互式数据筛选
- 创建带有可视化筛选控件的仪表板
- 多切片器筛选用于复杂数据分析

---

## 🎨 条件格式（2 个操作）

- **添加规则：** 创建条件格式规则
  - 单元格值比较（>、<、= 等）
  - 基于表达式的公式（自定义 DAX/Excel 公式）
  - 色阶、数据条、图标
- **清除规则：** 从范围中删除格式化

---

## 📸 截图（2 个操作）

- **捕获范围：** 将特定范围捕获为 PNG 图像
- **捕获工作表：** 将工作表的整个已用区域捕获为 PNG 图像
  - 使用 Excel 内置渲染（CopyPicture）— 捕获格式化、图表、条件格式
  - MCP：直接返回图像为 ImageContent（base64 PNG）
  - CLI：返回包含 base64 编码图像数据的 JSON

---

## 🪧 窗口管理（9 个操作）

- **显示：** 使 Excel 可见并将其置于前台
- **隐藏：** 隐藏 Excel 窗口
- **置于前台：** 将 Excel 置于前台而不改变可见性
- **获取信息：** 获取当前窗口状态（可见性、位置、大小、前台状态）
- **设置状态：** 将窗口状态设置为正常、最小化或最大化
- **设置位置：** 以点为单位设置窗口位置和大小（左、上、宽、高）
- **排列：** 使用预设布局排列 Excel 窗口
- **设置状态栏：** 在 Excel 状态栏中显示自定义文本以提供实时反馈
- **清除状态栏：** 恢复默认状态栏文本

**排列预设：**
- `left-half` / `right-half` — 与其他应用程序并排
- `top-half` / `bottom-half` — 堆叠视图
- `center` — 居中窗口（屏幕的 60%）
- `full-screen` — 最大化

**使用场景：**
- 交互式"代理模式"，用户实时观看 Excel 响应 AI 命令
- 并排显示：Excel 在一侧，AI 助手在另一侧
- 可见性更改反映在会话元数据中（会话列表显示更新状态）

---

## 📊 操作总数汇总

| 类别 | 操作数 |
|------|--------|
| 文件操作 | 6 |
| Power Query | 12 |
| 数据模型/DAX | 19 |
| Excel 表格 | 27 |
| 数据透视表 | 30 |
| 图表 | 29 |
| 单元格范围 | 46 |
| 工作表 | 16 |
| 连接 | 9 |
| 命名区域 | 6 |
| VBA 宏 | 6 |
| 切片器 | 8 |
| 条件格式 | 2 |
| 截图 | 2 |
| 计算模式 | 3 |
| 窗口管理 | 9 |
| **总计** | **230** |

---

## 🔧 工具选择快速参考

| 任务 | 工具 |
|------|------|
| 导入数据 | `powerquery` 或 `connection` |
| 创建分析 | `pivottable`（基于数据模型用于 OLAP） |
| 可视化数据 | `chart` |
| 更新参数 | `namedrange`（写入操作） |
| 管理公式 | `range`（set-formulas） |
| 格式化数据 | `range` / `range_format`（`format-range`、`format-ranges`、`validate-range`） |
| 脚本自动化 | `vba`（运行宏） |

---

## 📚 文档资源

- **[安装指南](docs/INSTALLATION.md)** — 所有 AI 助手的设置
- **[MCP 服务器指南](src/ExcelMcp.McpServer/README.md)** — 工具文档和示例
- **[CLI 指南](src/ExcelMcp.CLI/README.md)** — 命令行参考
- **[代理技能](skills/README.md)** — 跨平台 AI 助手指导
- **[贡献指南](docs/CONTRIBUTING.md)** — 开发指南
- **[发布版本](https://github.com/sbroenne/mcp-server-excel/releases)** — 最新更新和功能
