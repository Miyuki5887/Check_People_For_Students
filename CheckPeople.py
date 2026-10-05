import tkinter as tk

import ButtonChange as btnC
import openpyxl

GRAY, WHITE, BLUE = "#e0e0e0", "#ffffff", "#4169e1"


class App:
    def __init__(self):
        self.WindowsMain = tk.Tk()
        self.WindowsMain.title("学生查人软件")
        self.ListName01 = []
        self.ListName02 = []
        self.ListNumber = 0
        self.Excel01 = openpyxl.load_workbook("Name.xlsx")
        self.Sheet01 = self.Excel01["Sheet1"]
        self.header = tk.Frame(self.WindowsMain, bg=BLUE)
        self.header.grid(row=0, column=0, columnspan=2, sticky="nsew")
        tk.Label(
            self.header,
            text="学生查人软件",
            bg=BLUE,
            fg="#000000",
            font=("Microsoft YaHei UI", 18),
        ).pack(side="left", padx=24)
        self.ButtonArea = tk.Frame(self.WindowsMain, bg=GRAY)
        self.ButtonArea.grid(row=1, column=0, sticky="nsew")
        self.TextArea = tk.Frame(
            self.WindowsMain,
            bg=WHITE,
            width=300,
            height=518,
            bd=0,
            highlightthickness=0,
        )
        self.TextArea.grid(row=1, column=1, sticky="nsew")
        self.TextArea.grid_propagate(False)
        self.OutputText = tk.Text(
            self.TextArea,
            width=1,
            height=1,
            fg="#000000",
            font=("Microsoft YaHei UI", 12),
            relief="flat",
            bd=0,
            highlightthickness=0,
            wrap="char",
        )
        self.OutputText.pack(side="left", padx=24, fill="both", expand=True)
        self.OutputText.config(state="disabled")
        self.OutputForCopy = ""

    def StyleChange(self, OrderStyle, OrderButton):
        OrderButton.StyleChangeWhenRunning(OrderStyle)

    def ListReduce(self, ReduceNumber):
        self.ListName01[ReduceNumber] = ""

    def ListAdd(self, AddNumber):
        self.ListName01[AddNumber] = self.ListName02[AddNumber]

    def WindowPreSitting(self, Size):
        Length = Size[0]
        Weigth = Size[1]
        self.WindowsMain.geometry(str(Length) + "x" + str(Weigth))
        self.WindowsMain.configure(bg=GRAY)
        self.WindowsMain.grid_columnconfigure(0, weight=1)
        self.WindowsMain.grid_columnconfigure(1, weight=0, minsize=300)
        self.WindowsMain.grid_rowconfigure(0, weight=0, minsize=100)
        self.WindowsMain.grid_rowconfigure(1, weight=1)

    def WindowsBuild(self):
        self.WindowsMain.mainloop()

    def OneButtonBuild(self, Buttonname, Line, Group, ListPosition):
        btn = btnC.ChangedButton(self.ButtonArea, text=Buttonname, style="Attend")
        btn.grid(row=Line, column=Group, padx=10, pady=10, ipadx=8, ipady=7, sticky="w")
        btn.config(command=lambda b=btn, idx=ListPosition: self.ClickButton(b, idx))

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

    def ClickButton(self, OrderButton, ListPosition):
        self.OutputForCopy = ""
        NowStyle = OrderButton.StyleGet()
        if NowStyle == "Absent":
            self.ListAdd(ListPosition)
            self.StyleChange("Attend", OrderButton)
        else:
            self.ListReduce(ListPosition)
            self.StyleChange("Absent", OrderButton)
        self.ListOutput()
        self.WindowsMain.clipboard_clear()
        self.WindowsMain.clipboard_append(self.OutputForCopy)
        self.WindowsMain.update()

    def ListOutput(self):
        self.OutputText.config(state="normal")
        Output = ""
        i = 0
        for name in self.ListName01:
            if name != "":
                Output = Output + name + "，"
                i = i + 1
                self.OutputForCopy = self.OutputForCopy + name + "，"
            else:
                pass
            if i == 3:
                Output = Output + "\n"
                i = 0
        Output = Output[:-1]
        self.OutputForCopy = self.OutputForCopy[:-1]
        self.OutputText.delete("1.0", "end")
        self.OutputText.insert("1.0", Output)
        self.OutputText.config(state="disabled")

    def main(self):
        self.ExcelGet()
        self.WindowPreSitting([1000, 918])
        self.MainButtonBuild()
        self.WindowsBuild()
        print(self.ListName01)
        print(self.ListName02)


if __name__ == "__main__":
    App().main()
