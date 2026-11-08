# 文档与表格处理程序

一个功能强大的Excel表格处理工具，提供直观的用户界面和便捷的命令操作。

## 功能特性

- **直观的用户界面**：类似电子表格的图形界面，支持直接编辑单元格
- **工作表管理**：创建、删除、切换工作表
- **格式模板系统**：内置多种预设格式模板（表头、数据、数字、标题、高亮等）
- **命令解析引擎**：通过简单命令快速执行复杂操作
- **便捷的格式设置**：填充颜色、设置对齐、添加边框等
- **自动调整列宽**：一键优化表格布局

## 安装依赖

```bash
pip install -r requirements.txt
```

## 运行程序

```bash
python main.py
```

## 使用说明

### 界面布局

1. **左侧面板**：工作表列表和格式模板列表
2. **中间区域**：表格编辑区域（50行×26列）
3. **右侧面板**：命令控制台

### 常用命令

| 命令 | 说明 | 示例 |
|------|------|------|
| `help` | 显示命令帮助 | `help` |
| `template` | 查看所有可用模板 | `template` |
| `table <范围> <模板名>` | 应用表格模板 | `table A1:E10 header` |
| `border <范围>` | 添加边框 | `border A1:E10` |
| `fill <范围> <颜色>` | 填充背景色 | `fill A1:A5 FFFF00` |
| `align <范围> <left/center/right>` | 设置对齐方式 | `align A1:E10 center` |
| `autofit` | 自动调整列宽 | `autofit` |
| `select <工作表名>` | 选择工作表 | `select Sheet1` |
| `create <工作表名>` | 创建新工作表 | `create 数据` |
| `set <单元格> <值>` | 设置单元格值 | `set A1 姓名` |

### 内置模板

- **header**：表头样式（加粗、灰色背景、居中、边框）
- **data**：数据单元格样式（左对齐、边框）
- **number**：数字单元格样式（右对齐、边框）
- **title**：标题样式（大字体、加粗、居中）
- **highlight**：高亮样式（黄色背景、加粗）

### 快捷键

- `Ctrl+N`：新建文件
- `Ctrl+O`：打开文件
- `Ctrl+S`：保存文件

## 项目结构

```
ExcelToolkit/
├── main.py              # 主程序入口和GUI界面
├── spreadsheet_editor.py # 表格编辑核心模块
├── format_templates.py  # 格式模板管理系统
├── command_parser.py    # 命令解析引擎
├── config.py            # 配置文件
├── card_view.py         # 卡片视图入口
├── integrated_app.py    # 集成版入口
├── requirements.txt     # 依赖包列表
├── run_main.bat         # 启动主程序
├── run_card_view.bat    # 启动卡片视图
├── run_integrated.bat   # 启动集成版
├── test_demo.py         # 测试演示脚本
└── README.md            # 说明文档
```

## 技术栈

- **Python 3.x**
- **openpyxl**：Excel文件处理
- **tkinter**：图形用户界面

## 启动方式

### 方式一：批处理脚本（推荐）
双击项目根目录下对应的 `.bat` 文件：
- `run_main.bat` — 启动主程序
- `run_card_view.bat` — 启动卡片视图
- `run_integrated.bat` — 启动集成版

### 方式二：命令行
```bash
python main.py
python card_view.py
python integrated_app.py
```

## 示例使用流程

1. 启动程序：双击 `run_main.bat` 或执行 `python main.py`
2. 在A1-E1单元格输入表头（姓名、年龄、性别、职业、城市）
3. 在命令控制台输入：`table A1:E1 header`
4. 在A2-E10输入数据
5. 输入：`table A2:E10 data`
6. 输入：`autofit`
7. 保存文件

## 许可证

MIT License
