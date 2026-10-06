# Check_People_For_Students

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![Release](https://img.shields.io/github/v/release/Miyuki5887/Check_People_For_Students)

一个基于 Python + Tkinter 的桌面小工具，用于帮助班级同学**快速统计已到人员**并生成可粘贴的名单文本。

界面左侧按名单生成一排「按钮」，每点一次就在**出席（绿色）**和**缺席（红色）**之间切换；右侧分上下两栏，实时显示**已到**与**未到**名单，点一下「一键导出」按钮即可把整份名单复制到剪贴板，直接粘贴到 QQ / 微信群。

---

## 功能特性

- 从 Excel（`Name.xlsx`）自动读取班级名单，无需手写
- 一键切换出席 / 缺席状态，按钮颜色即时反馈
- 鼠标悬停变色，界面简洁直观
- 右侧**上下双栏**实时显示「已到名单」与「未到名单」，每 3 人一行排版
- **一键导出**已到 / 未到名单到剪贴板
- 支持自定义窗口图标（`icon.ico`）
- 按**程序所在目录**查找数据文件，从任何位置启动都能正常运行
- 按钮样式集中管理，便于统一换肤
- **自动适配高 DPI 显示屏**：125% / 150% 缩放下窗口按比例放大，文字与按钮保持原生像素清晰（v0.1.2 起）

---

## 运行环境

| 项目 | 要求 |
| --- | --- |
| Python | 3.x（建议 3.8 及以上） |
| 图形库 | `tkinter`（Python 官方自带，一般无需安装） |
| 第三方库 | `openpyxl` |
| 操作系统 | Windows（中文字体按 `Microsoft YaHei UI` 设置，其他系统需自行调整字体） |
| 运行方式 | 双击 `CheckPeople.pyw`（Windows 关联到 `pyw.exe`，无控制台窗口），或命令行 `pythonw CheckPeople.pyw` |

### 安装依赖

```bash
pip install openpyxl
```

> 若提示 `No module named tkinter`（常见于 Linux），需另行安装系统包：
> Ubuntu/Debian 下执行 `sudo apt install python3-tk`。

---

## 目录结构

```
CheckTool/
├── CheckPeople.pyw    # 主程序：窗口、名单读取、按钮逻辑、结果输出（双击即可运行）
├── ButtonChange.py    # 自定义按钮组件：配色方案与悬停/切换效果
├── icon.ico           # 窗口图标（由 CheckPeople.pyw 自动加载）
├── Name.xlsx          # 班级名单表格（需自行准备，见下节；已被 .gitignore 排除，不会随仓库上传）
├── LICENSE            # MIT 开源许可证
└── README.md          # 本说明文件
```

---

## 数据准备（重要）

程序启动时会读取**自身所在目录**下的 `Name.xlsx`，请按下表格式准备：

| 单元格 | 内容 |
| --- | --- |
| 工作表名 | `Sheet1`（必须一致） |
| A1、A2、A3 … | 依次填写学生姓名 |

示例：

| A |
| --- |
| 张三 |
| 李四 |
| 王五 |

⚠️ 注意事项：

1. 文件名必须是 `Name.xlsx`，且与 `CheckPeople.pyw` 放在**同一目录**。程序以**自身所在目录**为基准查找，因此在哪个目录执行命令、从哪个快捷方式启动都不影响。
2. 名单从 **A1 开始连续向下**读取，**遇到第一个空白单元格即停止**，因此中间不能留空行。
3. 若文件不存在，程序会弹出一个「启动失败」提示框，告诉你把 `Name.xlsx` 放到程序所在的文件夹，然后直接退出。

---

## 使用方法

1. 准备好 `Name.xlsx`，放到项目目录下。
2. 安装依赖：`pip install openpyxl`
3. 运行主程序：**双击 `CheckPeople.pyw`** 即可。`.pyw` 在 Windows 上关联到 `pyw.exe`，所以不会弹出黑色控制台窗口。

   也可以在命令行里运行：

   ```bash
   pythonw CheckPeople.pyw
   ```

4. 窗口打开后，默认所有同学都算「已到」（按钮为绿色），「未到名单」栏为空。
5. 点击某位同学的按钮：
   - 变**红色** → 标记为缺席，其姓名从「已到」栏移到「未到」栏；
   - 再次点击 → 恢复**绿色**，姓名回到「已到」栏。
6. 右侧上栏显示**已到名单**、下栏显示**未到名单**，均每行 3 人。
7. 需要复制时，点击「**一键导出已到名单**」或「**一键导出未到名单**」按钮，整份名单会写入剪贴板，直接粘贴发送即可。

> 自 v0.1.1 起，点击同学按钮**只切换状态并刷新两栏**，不再自动写入剪贴板。

---

## 界面说明

- 窗口标题：**学生查人软件**，窗口图标取自 `icon.ico`，默认尺寸 `1000 × 918`（100% 缩放下的基准值；125% 缩放时会按比例放大为 `1250 × 1148`）
- 顶部：蓝色标题栏，附「绿色=已到 红色=未到 右侧可以一键导出结果」提示
- 左侧：学生按钮区，**每列 10 个按钮**，超出后自动换到下一列
- 右侧上栏：**已到名单**（只读文本，固定 `300 × 414` px），底部为「一键导出已到名单」按钮
- 右侧下栏：**未到名单**（只读文本，固定 `300 × 404` px），底部为「一键导出未到名单」按钮

---

## 代码结构说明

### `CheckPeople.pyw`

模块级辅助函数：

| 函数 | 作用 |
| --- | --- |
| `EnableDpiAwareness()` | 把进程声明为 DPI 感知（依次尝试 per-monitor-v2 → per-monitor → system），避免高缩放屏下整窗被系统拉伸糊边 |
| `SystemScale()` | 读取系统 DPI（`GetDpiForSystem()`）并换算成缩放倍数（96 DPI = 1.0，120 DPI = 1.25） |
| `BaseDir()` | 返回程序所在目录（源码运行时为脚本所在目录，PyInstaller 打包后为 exe 所在目录） |
| `DataPath(FileName)` | 把数据文件名拼成 `BaseDir()` 下的绝对路径，用于定位 `Name.xlsx` 与 `icon.ico` |

`App` 类为主控制器：

| 方法 | 作用 |
| --- | --- |
| `ErrorTk(ErrorTitle, ErrorText)` | 弹出一个「启动失败」提示框，然后以退出码 1 结束程序 |
| `SysInitailize()` | 启动自检：检查 `openpyxl` 是否安装、`Name.xlsx` 是否存在（缺失时弹窗后退出） |
| `Px(Number)` | 把逻辑像素换算成当前 DPI 下的实际像素（`round(Number × self.Scale)`），所有尺寸参数都经过它 |
| `TkInit()` | 搭建界面：标题栏、左右分栏、两个只读文本框与两个导出按钮 |
| `ListInit()` | 打开 `Name.xlsx` 并选中 `Sheet1` |
| `ExcelGet()` | 读取 `Sheet1` 的 A 列姓名，并复制一份到 `ListName02` 作为原始备份 |
| `WindowPreSitting(Size)` | 设置窗口图标、尺寸与网格布局权重 |
| `MainButtonBuild()` | 遍历名单批量创建按钮，每满 10 个换列 |
| `OneButtonBuild(...)` | 创建单个按钮并绑定点击事件 |
| `ClickButton(...)` | 处理点击：切换出席 / 缺席状态 → 刷新右侧两栏 |
| `StyleChange(OrderStyle, OrderButton)` | 调用按钮组件的 `StyleChangeWhenRunning()` 动态换色 |
| `ListReduce(idx)` / `ListAdd(idx)` | 把对应位置置空 / 从备份恢复 |
| `ListOutput()` | 生成「已到」「未到」两份文本，每 3 人一行，末尾去掉多余逗号 |
| `Copying(Copytext)` | 把指定文本写入系统剪贴板 |
| `CopyButton(...)` | 创建「一键导出」按钮并绑定对应命令 |
| `WindowsBuild()` | 进入主事件循环 |

### `ButtonChange.py`

- `InitSTYLE` 定义三套配色：`primary`（蓝，`#3498db`）、`Attend`（绿，`#2ecc71`）、`Absent`（红，`#e74c3c`），每套含 `bg` / `fg` / `hover` / `active`。
- `ChangedButton` 继承自 `tk.Button`，新增：
  - `Entering` / `Leaving`：鼠标进入、离开时的悬停变色（禁用状态不变色）；
  - `StyleChangeWhenRunning(style)`：运行中动态切换配色；
  - `StyleGet()`：返回当前样式名，供主程序判断出席 / 缺席。

---

## 已知限制

- **状态不持久化**：关闭程序后所有标记丢失，每次运行都从「全部已到」开始。
- **名单长度上限**：按钮按每列 10 个排列，但窗口只为两列预留了空间，超过 **20 人**时按钮可能超出可视区域。
- **窗口尺寸随 DPI 放大，但不可拖拽重排**：`1000 × 918` 是 100% 缩放下的基准值，高 DPI 屏会等比放大（v0.1.2 起），窗口本身没有做自适应重排。
- **首次点击前导出「未到名单」会报错（已知问题）**：启动后若不点任何同学、直接点「一键导出未到名单」，会抛出 `AttributeError: 'App' object has no attribute 'AbsentCopy'`。先随便点一位同学即可正常工作。
- **无编辑功能**：姓名只能通过修改 `Name.xlsx` 变更。
- 输出使用中文全角逗号「，」作为分隔符，若需其他格式请修改 `ListOutput()`。

---

## 常见问题

**Q：启动时报 `FileNotFoundError: [Errno 2] No such file or directory: 'Name.xlsx'`？**
A：自 v0.1.1 起程序按**自身所在目录**查找名单，正常情况下不会再出现这个错误。若弹出了「启动失败」提示框，把 `Name.xlsx` 放到 **`CheckPeople.pyw` 所在的同一个文件夹**里即可。

**Q：没有安装 `openpyxl` 会怎样？**
A：程序会弹出「启动失败」并提示你下载，然后退出。在 PowerShell 里执行 `pip install openpyxl`，再重新打开程序即可。

**Q：点了同学按钮，剪贴板里什么都没有？**
A：点击按钮只负责切换状态并刷新右侧两栏；复制请使用「一键导出已到名单」/「一键导出未到名单」按钮。

**Q：界面中文显示成方块？**
A：系统缺少 `Microsoft YaHei UI` 字体（非 Windows 系统常见）。请把 `font=("Microsoft YaHei UI", ...)` 改为本机已有字体，例如 `("SimHei", 12)` 或 `("Arial", 11)`。

**Q：只用了一部分同学的名字，后面的没加载出来？**
A：检查 A 列中间是否有空单元格——读取遇到空格就会停止。

**Q：粘贴出来的内容和文本框显示的不一样？**
A：文本框每 3 人换行，而剪贴板内容是**连续一行**（不带换行），两者内容相同、排版不同。

---

## 后续可扩展方向

- 增加「全选 / 全不选」按钮
- 支持把考勤结果直接写回 Excel
- 名单过长时自动分页或加滚动条
- 打包为免安装的 `.exe`：`pyinstaller --onefile --noconsole --icon icon.ico --add-data "icon.ico:." CheckPeople.pyw`（`icon.ico` 是运行时读取的，必须用 `--add-data` 一起打进去；`--onefile` 在部分机器上可能解不出 `VCRUNTIME140.dll` 而启动失败，遇到时改用 `--onedir`）

---

## 更新日志

- **v0.1.2**：适配高 DPI 显示屏（125% / 150% 缩放下窗口不再被系统拉伸模糊）；主程序改名为 `CheckPeople.pyw`，双击运行不再弹出黑色控制台窗口；标题栏文字改为白色、提示字号调大
- **v0.1.1**：新增窗口图标、已到 / 未到双栏面板与一键导出按钮；修复从任意目录启动时找不到 `Name.xlsx` / `icon.ico` 的问题
- **v0.1.0**：首个可用版本

完整记录见 [Releases](https://github.com/Miyuki5887/Check_People_For_Students/releases)。

---

## 许可证

本项目基于 [MIT License](LICENSE) 开源。

```
MIT License

Copyright (c) 2026 Check_People_For_Students contributors
```

任何人都可以自由使用、复制、修改、合并、发布、分发、再授权或销售本软件，唯一的要求是保留上述版权声明和许可声明。本软件按「原样」提供，不附带任何形式的担保。
