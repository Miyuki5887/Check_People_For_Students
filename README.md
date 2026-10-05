# Check_People_For_Students

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python](https://img.shields.io/badge/Python-3.x-blue.svg)

一个基于 Python + Tkinter 的桌面小工具，用于帮助班级同学**快速统计已到人员**并生成可粘贴的名单文本。

界面左侧按名单生成一排「按钮」，每点一次就在**出席（绿色）**和**缺席（红色）**之间切换；右侧实时显示当前已到人员的名单，并且**每次点击都会自动把结果复制到剪贴板**，直接粘贴到 QQ / 微信群即可。

---

## 功能特性

- 从 Excel（`Name.xlsx`）自动读取班级名单，无需手写
- 一键切换出席 / 缺席状态，按钮颜色即时反馈
- 鼠标悬停变色，界面简洁直观
- 实时统计已到名单，每 3 人一行排版
- 自动复制到剪贴板，省去手动框选复制
- 按钮样式集中管理，便于统一换肤

---

## 运行环境

| 项目 | 要求 |
| --- | --- |
| Python | 3.x（建议 3.8 及以上） |
| 图形库 | `tkinter`（Python 官方自带，一般无需安装） |
| 第三方库 | `openpyxl` |
| 操作系统 | Windows（中文字体按 `Microsoft YaHei UI` 设置，其他系统需自行调整字体） |

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
├── CheckPeople.py     # 主程序：窗口、名单读取、按钮逻辑、结果输出
├── ButtonChange.py    # 自定义按钮组件：配色方案与悬停/切换效果
├── Name.xlsx          # 班级名单表格（需自行准备，见下节）
├── LICENSE            # MIT 开源许可证
└── README.md          # 本说明文件
```

---

## 数据准备（重要）

程序启动时会读取**同目录下**的 `Name.xlsx`，请按下表格式准备：

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

1. 文件名必须是 `Name.xlsx`，且与 `CheckPeople.py` 放在**同一目录**。
2. 名单从 **A1 开始连续向下**读取，**遇到第一个空白单元格即停止**，因此中间不能留空行。
3. 如果文件不存在，程序会直接抛出 `FileNotFoundError`。

---

## 使用方法

1. 准备好 `Name.xlsx`，放到项目目录下。
2. 安装依赖：`pip install openpyxl`
3. 运行主程序：

   ```bash
   python CheckPeople.py
   ```

4. 窗口打开后，默认所有同学都算「已到」（按钮为绿色）。
5. 点击某位同学的按钮：
   - 变**红色** → 标记为缺席，其姓名从右侧名单中移除；
   - 再次点击 → 恢复**绿色**，姓名重新加入名单。
6. 右侧文本框显示当前已到人员（每行 3 人），同时内容已自动复制到剪贴板，直接粘贴发送即可。

---

## 界面说明

- 窗口标题：**学生查人软件**，默认尺寸 `1000 × 918`
- 顶部：蓝色标题栏
- 左侧：学生按钮区，**每列 10 个按钮**，超出后自动换到下一列
- 右侧：结果展示区（固定宽度 300 px），只读文本，无法手动编辑

---

## 代码结构说明

### `CheckPeople.py`

`App` 类为主控制器：

| 方法 | 作用 |
| --- | --- |
| `ExcelGet()` | 读取 `Name.xlsx` 的 A 列姓名，并复制一份到 `ListName02` 作为原始备份 |
| `WindowPreSitting(Size)` | 设置窗口尺寸与网格布局权重 |
| `MainButtonBuild()` | 遍历名单批量创建按钮，每满 10 个换列 |
| `OneButtonBuild(...)` | 创建单个按钮并绑定点击事件 |
| `ClickButton(...)` | 处理点击：切换状态 → 刷新输出 → 写入剪贴板 |
| `ListReduce(idx)` / `ListAdd(idx)` | 把对应位置置空 / 从备份恢复 |
| `ListOutput()` | 生成右侧文本，每 3 人一行，末尾去除多余逗号 |
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
- **窗口尺寸固定**：`1000 × 918` 为写死的初值，未做自适应缩放与 DPI 适配。
- **无编辑功能**：姓名只能通过修改 `Name.xlsx` 变更。
- 输出使用中文全角逗号「，」作为分隔符，若需其他格式请修改 `ListOutput()`。

---

## 常见问题

**Q：点运行就报 `FileNotFoundError: [Errno 2] No such file or directory: 'Name.xlsx'`？**
A：说明当前工作目录下没有 `Name.xlsx`。请把它放到与 `CheckPeople.py` 同一目录，或在正确的目录下执行 `python CheckPeople.py`。

**Q：界面中文显示成方块？**
A：系统缺少 `Microsoft YaHei UI` 字体（非 Windows 系统常见）。请把两处 `font=("Microsoft YaHei UI", ...)` 改为本机已有字体，例如 `("SimHei", 12)` 或 `("Arial", 11)`。

**Q：只用了一部分同学的名字，后面的没加载出来？**
A：检查 A 列中间是否有空单元格——读取遇到空格就会停止。

**Q：粘贴出来的内容和文本框显示的不一样？**
A：文本框每 3 人换行，而剪贴板内容是**连续一行**（不带换行），两者内容相同、排版不同。

---

## 后续可扩展方向

- 增加「全选 / 全不选」与「保存结果」按钮
- 支持从 Excel 直接写回考勤结果
- 名单过长时自动分页或加滚动条
- 打包为免安装的 `.exe`（可用 `pyinstaller -F -w CheckPeople.py`）

---

## 许可证

本项目基于 [MIT License](LICENSE) 开源。

```
MIT License

Copyright (c) 2026 Check_People_For_Students contributors
```

任何人都可以自由使用、复制、修改、合并、发布、分发、再授权或销售本软件，唯一的要求是保留上述版权声明和许可声明。本软件按「原样」提供，不附带任何形式的担保。

