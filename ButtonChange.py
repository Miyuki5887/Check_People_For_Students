import tkinter as tk

InitSTYLE = {
    "primary": {
        "bg": "#3498db",
        "fg": "#ffffff",
        "hover": "#5dade2",
        "active": "#2471a3",
    },
    "Attend": {
        "bg": "#2ecc71",
        "fg": "#ffffff",
        "hover": "#58d68d",
        "active": "#239b56",
    },
    "Absent": {
        "bg": "#e74c3c",
        "fg": "#ffffff",
        "hover": "#ec7063",
        "active": "#b03a2e",
    },
}


class ChangedButton(tk.Button):
    def __init__(
        self,
        master,
        text="",
        style="primary",
        command=None,
        font=None,
        padx=16,
        pady=7,
        **kwargs,
    ):
        conf = dict(InitSTYLE.get(style, InitSTYLE["primary"]))
        for key in ("bg", "background"):
            if key in kwargs:
                conf["bg"] = kwargs.pop(key)
        for key in ("fg", "foreground"):
            if key in kwargs:
                conf["fg"] = kwargs.pop(key)
        self._conf = conf
        self._font = font or ("Microsoft YaHei UI", 11)
        self._style = style
        super().__init__(
            master,
            text=text,
            command=command,
            font=self._font,
            bg=conf["bg"],
            fg=conf["fg"],
            activebackground=conf["active"],
            activeforeground=conf["fg"],
            relief="flat",
            bd=2,
            highlightthickness=0,
            cursor="hand2",
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self.bind("<Enter>", self.Entering)
        self.bind("<Leave>", self.Leaving)

    def Entering(self, _event=None):
        if str(self["state"]) != "disabled":
            self.config(bg=self._conf["hover"])

    def Leaving(self, _event=None):
        if str(self["state"]) != "disabled":
            self.config(bg=self._conf["bg"])

    def StyleChangeWhenRunning(self, style):
        self._conf = dict(InitSTYLE.get(style, InitSTYLE["primary"]))
        self._style = style
        self.config(
            bg=self._conf["bg"],
            fg=self._conf["fg"],
            activebackground=self._conf["active"],
        )

    def StyleGet(self):
        return self._style
