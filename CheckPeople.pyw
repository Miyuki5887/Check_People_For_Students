import ctypes
import os
import shutil
import subprocess
import sys
import tkinter as tk
import tkinter.messagebox as mb
from tkinter import filedialog

import ButtonChange as btnC

(
    GRAY,
    WHITE,
    BLUE,
    BLACK,
) = (
    "#e0e0e0",
    "#f2f3f5",
    "#1677FF",
    "#000000",
)


def EnableDpiAwareness():
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return "per-monitor-v2"
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return "per-monitor"
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        return "system"
    except Exception:
        return "none"


def SystemScale():
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem()
    except Exception:
        dpi = 96
    return (dpi or 96) / 96.0


def BaseDir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def DataPath(FileName):
    return os.path.join(BaseDir(), FileName)


class App:
    def __init__(self):
        self.Scale = SystemScale()
        self.SysInitailize()
        self.WindowsMain = tk.Tk()
        self.WindowsMain.title("学生查人软件")
        self.TkInit()
        self.ListInit()

    def All(self, state: bool):
        for Btn in self.ButtonArea.winfo_children():
            if Btn.On == state:
                self.ClickButton(Btn, Btn.Position)

    def AllButton(self):
        AllAttendButton = btnC.ChangedButton(
            self.AttendedArea,
            "       一键全到       ",
            command=lambda s=False: self.All(s),
            padx=self.Px(16),
            pady=self.Px(7),
        )
        AllAbsentButton = btnC.ChangedButton(
            self.AbsentArea,
            "     一键全未到     ",
            command=lambda s=True: self.All(s),
            padx=self.Px(16),
            pady=self.Px(7),
        )
        AllAttendButton.pack()
        AllAbsentButton.pack()

    def RebuildButton(self):
        for Btn in self.ButtonArea.winfo_children():
            Btn.destroy()
        self.AttendCopy = ""
        self.AbsentCopy = ""
        self.ListInit()
        self.ExcelGet()
        self.MainButtonBuild()

    def ImportFunc(self):
        # 只负责把选中的名单拷成程序目录里的 Name.xlsx，成功返回 True、取消返回 False
        FilePath = filedialog.askopenfilename(
            title="请选择文件", filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not FilePath:
            return False
        Target = DataPath("Name.xlsx")
        if os.path.normcase(os.path.abspath(FilePath)) == os.path.normcase(
            os.path.abspath(Target)
        ):
            mb.showerror("警告", "请勿选取已有的Name.xlsx文件！")
            return self.ImportFunc()
        try:
            shutil.copy(FilePath, Target)
        except shutil.SameFileError:
            mb.showerror("警告", "请勿选取已有的Name.xlsx文件！")
            return self.ImportFunc()
        except PermissionError:
            mb.showerror("警告", "Name.xlsx 正被 Excel 打开，请关闭后再试。")
            return False
        return True

    def ImportButton(self):
        # 给界面右上角那个按钮用：拷好名单后原地重建按钮
        if self.ImportFunc():
            self.RebuildButton()

    def Px(self, Number):
        return int(round(Number * self.Scale))

    def ErrorTk(self, ErrorTitle, ErrorText):
        ErrorWindows = tk.Tk()
        ErrorWindows.withdraw()
        mb.showerror(ErrorTitle, ErrorText)
        ErrorWindows.destroy()
        raise SystemExit(1)

    def SysInitailize(self):
        # 检查库是否安装
        try:
            import openpyxl

            self.ExcelImportMoudel = openpyxl
        except ModuleNotFoundError:
            self.ErrorTk(
                "启动失败",
                "你需要下载openpyxl，下载命令已存入剪切板，请前往PowerShell执行",
            )
            subprocess.run(
                "clip",
                input="pip install openpyxl",
                text=True,
                encoding="utf-8",
                check=True,
            )
            return
        # 检查xlsx文件是否导入
        try:
            TestList = openpyxl.load_workbook(DataPath("Name.xlsx"))
            del TestList
        except FileNotFoundError:
            if not self.ImportFunc():
                self.ErrorTk(
                    "启动失败",
                    "您需要在当前文件所在的文件夹内放入Name.xlsx文件，或者把您的名单改为Name.xlsx",
                )
            return

    def TkInit(self):
        self.header = tk.Frame(self.WindowsMain, bg=BLUE)
        self.header.grid(row=0, column=0, columnspan=2, sticky="nsew")
        tk.Label(
            self.header,
            text="学生查人软件",
            bg=BLUE,
            fg="#ffffff",
            font=("Microsoft YaHei UI", 18),
        ).grid(row=0, column=0, padx=self.Px(24), pady=self.Px(20), sticky="w")
        tk.Label(
            self.header,
            text="绿色=已到 红色=未到 右侧可以一键导出结果",
            bg=BLUE,
            fg="#ffffff",
            font=("Microsoft YaHei UI", 10),
        ).grid(row=1, column=0, padx=self.Px(24), pady=self.Px(6), sticky="w")
        btnC.ChangedButton(
            self.header,
            text="点击这里导入文件",
            command=lambda: self.ImportButton(),
        ).grid(
            row=0,
            column=1,
            rowspan=2,
            padx=self.Px(10),
            pady=self.Px(10),
            ipadx=self.Px(8),
            ipady=self.Px(7),
            sticky="e",
        )
        self.ButtonArea = tk.Frame(self.WindowsMain, bg=GRAY)
        self.ButtonArea.grid(row=1, column=0, rowspan=2, sticky="nsew")
        self.AttendedArea = tk.Frame(
            self.WindowsMain,
            bg=WHITE,
            width=self.Px(300),
            height=self.Px(414),
            bd=0,
            highlightthickness=0,
        )
        self.AttendedArea.grid(
            row=1,
            column=1,
            sticky="nsew",
        )
        self.AttendedArea.grid_propagate(False)
        self.AbsentArea = tk.Frame(
            self.WindowsMain,
            bg=WHITE,
            width=self.Px(300),
            height=self.Px(404),
            bd=0,
            highlightthickness=0,
        )
        self.AbsentArea.grid(row=2, column=1, sticky="nsew")
        self.AbsentArea.grid_propagate(False)
        self.AttendOutputText = tk.Text(
            self.AttendedArea,
            width=1,
            height=1,
            fg=BLACK,
            font=("Microsoft YaHei UI", 12),
            relief="flat",
            bd=0,
            highlightthickness=0,
            wrap="char",
        )
        self.AttendOutputText.pack(
            side="top",
            padx=self.Px(24),
            fill="both",
            expand=True,
        )
        self.AttendOutputText.config(state="disabled")
        self.AbsentOutputText = tk.Text(
            self.AbsentArea,
            width=1,
            height=1,
            fg=BLACK,
            font=("Microsoft YaHei UI", 12),
            relief="flat",
            bd=0,
            highlightthickness=0,
            wrap="char",
        )
        self.AbsentOutputText.pack(
            side="top",
            padx=self.Px(24),
            fill="both",
            expand=True,
        )
        self.AbsentOutputText.config(state="disabled")
        self.AttendCopy = ""
        self.CopyButton(
            self.AttendedArea,
            text="一键导出已到名单",
            command=lambda: self.Copying(self.AttendCopy),
        )
        self.CopyButton(
            self.AbsentArea,
            text="一键导出未到名单",
            command=lambda: self.Copying(self.AbsentCopy),
        )
        self.AllButton()

    def Copying(self, Copytext):
        self.WindowsMain.clipboard_clear()
        self.WindowsMain.clipboard_append(Copytext)
        self.WindowsMain.update()

    def ListInit(self):
        self.ListName01 = []
        self.ListName02 = []
        self.ListNumber = 0
        self.Excel01 = self.ExcelImportMoudel.load_workbook(DataPath("Name.xlsx"))
        self.Sheet01 = self.Excel01["Sheet1"]

    def CopyButton(self, master, text, command):
        btn = btnC.ChangedButton(
            master,
            text=text,
            style="primary",
            padx=self.Px(16),
            pady=self.Px(7),
        )
        btn.pack(side="bottom")
        btn.config(command=command)
        return btn

    def StyleChange(self, OrderStyle, OrderButton):
        OrderButton.StyleChangeWhenRunning(OrderStyle)

    def ListReduce(self, ReduceNumber):
        self.ListName01[ReduceNumber] = ""

    def ListAdd(self, AddNumber):
        self.ListName01[AddNumber] = self.ListName02[AddNumber]

    def WindowPreSitting(self, Size):
        Length = Size[0]
        Weigth = Size[1]
        self.WindowsMain.iconbitmap(DataPath("icon.ico"))
        self.WindowsMain.geometry(str(self.Px(Length)) + "x" + str(self.Px(Weigth)))
        self.WindowsMain.configure(bg=GRAY)
        self.WindowsMain.grid_columnconfigure(0, weight=1)
        self.WindowsMain.grid_columnconfigure(1, weight=0, minsize=self.Px(300))
        self.WindowsMain.grid_rowconfigure(0, weight=0, minsize=self.Px(100))
        self.WindowsMain.grid_rowconfigure(1, weight=1, uniform="right")
        self.WindowsMain.grid_rowconfigure(2, weight=1, uniform="right")
        self.header.grid_columnconfigure(0, weight=1)

    def WindowsBuild(self):
        self.WindowsMain.mainloop()

    def OneButtonBuild(self, Buttonname, Line, Group, ListPosition):
        btn = btnC.ChangedButton(self.ButtonArea, text=Buttonname, style="Attend")
        btn.grid(
            row=Line,
            column=Group,
            padx=self.Px(10),
            pady=self.Px(10),
            ipadx=self.Px(8),
            ipady=self.Px(7),
            sticky="w",
        )
        btn.config(command=lambda b=btn, idx=ListPosition: self.ClickButton(b, idx))
        btn.Position = ListPosition
        btn.On = True

    def MainButtonBuild(self):
        self.ListOutput()
        i = 0
        Line = 0
        Group = 0
        for i in range(self.ListNumber - 1):
            self.OneButtonBuild(self.ListName01[i], Line, Group, i)
            Line = Line + 1
            if Line == 10:
                Group = Group + 1
                Line = 0

    def ExcelGet(self):
        i = 1
        while True:
            ListPosition = "A" + str(i)
            result = self.Sheet01[ListPosition].value
            if result != None:
                self.ListName01.append(result)
            else:
                self.ListNumber = i
                break
            i = i + 1
        self.ListName02 = self.ListName01.copy()

    def ClickButton(self, OrderButton: btnC.ChangedButton, ListPosition):
        self.AttendCopy = ""
        self.AbsentCopy = ""
        NowStyle = OrderButton.StyleGet()
        if NowStyle == "Absent":
            self.ListAdd(ListPosition)
            self.StyleChange("Attend", OrderButton)
        else:
            self.ListReduce(ListPosition)
            self.StyleChange("Absent", OrderButton)
        self.ListOutput()
        OrderButton.On = not OrderButton.On

    def ListOutput(self):
        self.AttendOutputText.config(state="normal")
        self.AbsentOutputText.config(state="normal")
        AttendOutput = ""
        AbsentOutput = ""
        ListNum = 0
        i = 0
        j = 0
        for name in self.ListName01:
            if name != "":
                AttendOutput = AttendOutput + name + "，"
                i = i + 1
                self.AttendCopy = self.AttendCopy + name + "，"
            else:
                AbsentOutput = AbsentOutput + self.ListName02[ListNum] + "，"
                j = j + 1
                self.AbsentCopy = self.AbsentCopy + self.ListName02[ListNum] + "，"
            if i == 3:
                AttendOutput = AttendOutput + "\n"
                i = 0
            if j == 3:
                AbsentOutput = AbsentOutput + "\n"
                j = 0
            ListNum = ListNum + 1
        AttendOutput = AttendOutput[:-1]
        self.AttendCopy = self.AttendCopy[:-1]
        if AbsentOutput != "":
            AbsentOutput = AbsentOutput[:-1]
            self.AbsentCopy = self.AbsentCopy[:-1]
        self.AttendOutputText.delete("1.0", "end")
        self.AbsentOutputText.delete("1.0", "end")
        self.AttendOutputText.insert("1.0", AttendOutput)
        self.AbsentOutputText.insert("1.0", AbsentOutput)
        self.AttendOutputText.config(state="disabled")
        self.AbsentOutputText.config(state="disabled")

    def main(self):
        self.ExcelGet()
        self.WindowPreSitting([1000, 918])
        self.MainButtonBuild()
        self.WindowsBuild()


if __name__ == "__main__":
    EnableDpiAwareness()
    App().main()
