# 分箱单生成工具

基于总单和附件备件清单，按分箱单模板批量生成装箱明细单。

## 环境要求

- Python 3.7+
- 依赖：`openpyxl`、`xlrd`

```bash
pip3 install openpyxl xlrd
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `template.xlsx` | 分箱单模板（格式基准） |
| `总单.xlsx` | 箱单总表（GPL sheet 为数据源） |
| `附件备件清单.xls` | 备件明细（多 sheet，按合同号匹配） |
| `generate_packing_list.py` | 核心生成逻辑 |
| `web_gui.py` | Web 图形界面 |
| `build_windows.bat` | Windows 可执行文件构建脚本 |
| `分箱单_生成结果.xlsx` | 输出文件 |

## 运行

### Web 图形界面（推荐）

```bash
python3 web_gui.py
```

浏览器自动打开 `http://localhost:8090`。模板文件内置，总单和备件清单通过页面上传。

### 命令行

```bash
python3 generate_packing_list.py
```

## 打包为可执行文件

### macOS

```bash
pip3 install pyinstaller
pyinstaller --onedir --name "分箱单生成工具" --add-data "template.xlsx:." --hidden-import email.policy --hidden-import xlrd --hidden-import openpyxl web_gui.py
```

### Windows

双击运行 `build_windows.bat`，或手动执行：

```cmd
pip install pyinstaller openpyxl xlrd
pyinstaller --onedir --name "分箱单生成工具" --add-data "template.xlsx;." --hidden-import email.policy --hidden-import xlrd --hidden-import openpyxl web_gui.py
```

## 输出

- 按 Case No. 排列在同一个 Sheet 中，每页末有强制分页符
- 格式与原模板一致（字体、边框、对齐、合并单元格）
- Accessories 页自动从备件清单匹配对应 sheet 并合并同编号数据

## 数据逻辑

1. 读取 `总单.xlsx` → GPL sheet，以 "MNS Low Voltage Switchgear" 行为界，下方为数据区
2. 逐行生成一页分箱单，字段映射：

| 模板位置 | 含义 | 数据来源 |
|----------|------|----------|
| C3 | 客户订单号 | GPL C17 |
| F3 | 装箱编号 | GPL B 列 |
| C4 | 尺寸 | GPL E 列 |
| C5 | 净重 | GPL F 列 |
| C6 | 毛重 | GPL G 列 |
| F4 | 项目号 | GPL I 列 |
| C10 | 品名规格 | GPL C 列 |
| D10 | 数量 | GPL D 列 |
| F10 | 合同号/系统号 | GPL H 列 |

3. 当 C10 为 "Accessories" 时，根据 F4 项目号匹配 `附件备件清单.xls` 中的 sheet，将备件明细填入 Row 11 起
