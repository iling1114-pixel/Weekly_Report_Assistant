"""週報小幫手視窗：貼上工作紀錄 → 勾選已完成待辦 → 產生 Word → 用 Outlook 或 Gmail 寄出。"""
import copy
import datetime as dt
import json
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import report as R

try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # 工作列顯示自己的圖示，不跟 pythonw 混在一起
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("WeeklyReportHelper")
except Exception:
    pass

ICON = str(R.RES / "Asset" / "app.ico")

# ---------- Notion 風格配色（淺色 / 深色） ----------
# 同一組裡每個色碼都要不同：切換主題時是靠色碼對照，把畫面上的舊色換成新色
PALETTES = {
    "light": dict(BG="#FFFFFF", SIDE="#F7F7F5", HOVER="#EFEFED", LINE="#E9E9E7", TEXT="#37352F",
                  MUTED="#787774", HINT_FG="#B4B4B0", BLUE="#2383E2", BLUE_HOVER="#0B6BCB",
                  GREEN="#448361", RED="#D44C47",
                  SCROLL_ACTIVE="#D3D3D0", SELECT_BG="#CCE3F9", CHECK_LINE="#A5A4A0"),
    "dark": dict(BG="#191919", SIDE="#202020", HOVER="#2F2F2F", LINE="#373737", TEXT="#E6E6E4",
                 MUTED="#9B9A97", HINT_FG="#5F5E5B", BLUE="#2383E2", BLUE_HOVER="#3A92E8",
                 GREEN="#52A37F", RED="#E5675F",
                 SCROLL_ACTIVE="#4A4A48", SELECT_BG="#1F4A73", CHECK_LINE="#6B6A67"),
}
THEME = "light"
BG = SIDE = HOVER = LINE = TEXT = MUTED = HINT_FG = BLUE = BLUE_HOVER = GREEN = RED = ""
SCROLL_ACTIVE = SELECT_BG = CHECK_LINE = ""  # 捲軸按住、選取反白、未勾選框線


def set_palette(name):
    """換掉模組裡的色票（之後新建的元件就用新色），回傳 {舊色碼: 新色碼}。"""
    global THEME
    name = name if name in PALETTES else "light"
    old, new = PALETTES[THEME], PALETTES[name]
    THEME = name
    globals().update(new)
    if "Button" in globals():
        Button.restyle()
    return {old[k].upper(): new[k] for k in new}


set_palette("light")

FAMILY = "Microsoft JhengHei UI"
FONT = (FAMILY, 11)
FONT_SM = (FAMILY, 9)
FONT_B = (FAMILY, 11, "bold")
FONT_H = (FAMILY, 13, "bold")
FONT_TITLE = (FAMILY, 22, "bold")
FONT_DATE = (FAMILY, 16, "bold")
FONT_DONE = (FAMILY, 11, "overstrike")
FIELD_HINT = "點這裡開始打，會自動編號・Tab 變子項目"
NEW_PROJ_HINT = "請輸入新專案名稱"
NEW_HINT = "點這裡開始打，會自動編號\nEnter 下一項・Tab 變子項目・Shift+Tab 退回"
ITEM_PREFIX = "{}. "    # 新增待辦框的主項編號
CHILD_PREFIX = "    - "  # 小項（縮排只是好看，解析時會去掉）
NUM = re.compile(r"^(\s*)(\d+)")  # 主項行開頭的號碼
WIN_FILE = R.BASE / "window.json"  # 視窗大小 / 位置 / 左右欄比例
SPLIT_DEFAULT, SPLIT_MIN, SPLIT_MAX = 0.6, 0.35, 0.75  # 左欄寬度比例：預設 3:2，拖拉上下限
SPIN = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"  # 處理中的轉圈動畫


class Button(tk.Label):
    """扁平按鈕：ghost（灰底 hover）、primary（藍底）、outline（細框）。"""

    STYLES = {}

    @classmethod
    def restyle(cls):
        cls.STYLES = {
            "ghost": dict(bg=BG, fg=MUTED, hover=HOVER, border=BG),
            "outline": dict(bg=BG, fg=TEXT, hover=HOVER, border=LINE),
            "primary": dict(bg=BLUE, fg="white", hover=BLUE_HOVER, border=BLUE),
        }

    def __init__(self, parent, text, command, kind="ghost", font=FONT, padx=10, pady=4):
        s = self.STYLES[kind]
        bg = parent.cget("bg") if kind == "ghost" else s["bg"]
        super().__init__(parent, text=text, font=font, bg=bg, fg=s["fg"], padx=padx, pady=pady,
                         cursor="hand2", highlightthickness=1,
                         highlightbackground=bg if kind == "ghost" else s["border"])
        self._bg, self._hover, self._cmd = bg, s["hover"], command
        self._fg, self._border, self._enabled = s["fg"], self.cget("highlightbackground"), True
        self.bind("<Enter>", lambda e: self._enabled and self.configure(bg=self._hover))
        self.bind("<Leave>", lambda e: self._enabled and self.configure(bg=self._bg))
        self.bind("<Button-1>", lambda e: self._enabled and self._cmd())

    def set_enabled(self, on):
        """停用時變灰、點了沒反應。"""
        if on == self._enabled:
            return
        self._enabled = on
        if on:
            self.configure(bg=self._bg, fg=self._fg, highlightbackground=self._border, cursor="hand2")
        else:
            self.configure(bg=LINE, fg=HINT_FG, highlightbackground=LINE, cursor="arrow")

    def remap(self, cmap):
        """切換主題：記住的顏色也換成新色。"""
        self._bg, self._hover, self._fg, self._border = (
            cmap.get(c.upper(), c) for c in (self._bg, self._hover, self._fg, self._border))


Button.restyle()


class Segmented(tk.Frame):
    """分段切換：灰底框裡並排選項，選中的白底粗字（取代 Radiobutton）。options: [(值, 文字)]。"""

    def __init__(self, parent, variable, options, on_change=None):
        super().__init__(parent, bg=SIDE, padx=3, pady=3, highlightthickness=1, highlightbackground=LINE)
        self.var, self._on_change, self._items = variable, on_change, {}
        for value, text in options:
            lb = tk.Label(self, text=text, font=FONT, padx=16, pady=3, cursor="hand2",
                          highlightthickness=1)
            lb.pack(side="left")
            lb.bind("<Button-1>", lambda e, v=value: self.var.set(v))
            lb.bind("<Enter>", lambda e, v=value: self._hover(v, True))
            lb.bind("<Leave>", lambda e, v=value: self._hover(v, False))
            self._items[value] = lb
        self.var.trace_add("write", lambda *_: self._paint())
        self._paint()

    def _paint(self):
        cur = self.var.get()
        for value, lb in self._items.items():
            on = value == cur
            lb.configure(bg=BG if on else SIDE, fg=TEXT if on else MUTED, font=FONT_B if on else FONT,
                         highlightbackground=LINE if on else SIDE)
        if self._on_change:
            self._on_change(cur)

    def _hover(self, value, inside):
        if value != self.var.get():
            self._items[value].configure(bg=HOVER if inside else SIDE)


COLOR_OPTS = ("bg", "fg", "highlightbackground", "highlightcolor", "insertbackground",
              "activebackground", "activeforeground", "selectbackground", "selectforeground",
              "selectcolor", "disabledforeground", "readonlybackground")


def recolor(widget, cmap):
    """整棵元件樹把舊色換新色：元件本身、Text 的 tag、Canvas 上畫的圖形、Button 記住的顏色。"""
    def swap(get, put, opts):
        for o in opts:
            try:
                v = str(get(o))
            except tk.TclError:
                continue
            if v.upper() in cmap:
                put(o, cmap[v.upper()])

    swap(widget.cget, lambda o, v: widget.configure(**{o: v}), COLOR_OPTS)
    if isinstance(widget, tk.Text):
        for tag in widget.tag_names():
            swap(lambda o: widget.tag_cget(tag, o), lambda o, v: widget.tag_configure(tag, **{o: v}),
                 ("foreground", "background", "selectbackground", "selectforeground"))
    elif isinstance(widget, tk.Canvas):
        for item in widget.find_all():
            swap(lambda o: widget.itemcget(item, o), lambda o, v: widget.itemconfigure(item, **{o: v}),
                 ("fill", "outline"))
    if isinstance(widget, ttk.Combobox):  # 下拉清單是 Tcl 自己建的，winfo_children 抓不到
        try:
            pd = widget.tk.call("ttk::combobox::PopdownWindow", widget)
            widget.tk.call(f"{pd}.f.l", "configure", "-background", BG, "-foreground", TEXT,
                           "-selectbackground", SELECT_BG, "-selectforeground", TEXT)
        except tk.TclError:
            pass
    if isinstance(widget, Button):
        widget.remap(cmap)
    elif isinstance(widget, Segmented):
        widget._paint()
    for child in widget.winfo_children():
        recolor(child, cmap)


def dark_titlebar(win):
    """Windows 10/11 標題列跟著主題變深色/淺色（不支援的版本就算了）。"""
    try:
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        on = ctypes.c_int(1 if THEME == "dark" else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE（新版 20、舊版 19）
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(on), 4) == 0:
                break
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"週報小幫手 v{R.VERSION}")
        try:
            self.iconbitmap(default=ICON)   # default= 讓對話框也套用
        except tk.TclError:
            pass
        self.geometry("1180x800")
        self.minsize(900, 600)
        self._split = SPLIT_DEFAULT  # 左欄佔中間區域的比例（拖中間分隔線調整）
        self._load_window()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.configure(bg=BG)
        if R.load_settings() is None:  # 第一次開：先填個人設定
            self.withdraw()
            if not SettingsDialog(self, first=True).saved:
                self.destroy()
                raise SystemExit
            self.deiconify()
        set_palette(R.THEME)  # 建畫面前先套主題
        self.configure(bg=BG)
        dark_titlebar(self)
        self.state_ = R.load_state()
        self.vars = {}          # key -> BooleanVar
        self.manual = set()     # 使用者手動改過的 key，自動判斷不覆蓋
        self._after = None
        self._saved = ""        # 最後一次儲存的工作紀錄，用來判斷「未儲存」
        self._wrap_labels = []  # (label, indent)：清單寬度變動時重算換行
        self._rows, self._heads = [], {}
        self._drag = self._press_at = self._click_job = None  # 拖拉 / 單擊延遲
        self._ignore_release = False
        self._edit_close = None  # 正在編輯的待辦輸入框 (entry, finish)
        self._inserted = set()  # 勾待辦時自動寫進工作紀錄的 (專案, 文字)，取消勾才能刪回來
        self._toast_job = None
        self._busy = False      # 背景工作進行中（_run_busy）
        self._style()
        self._build_ui()
        self.refresh_todos()
        self._update_preview()
        self.bind_all("<Control-s>", lambda e: self.save_work())
        self.bind_all("<Button-1>", self._click_away, add="+")
        self.after(1500, self._check_update)

    # ---------- 記住視窗大小 / 位置 ----------
    def _load_window(self):
        try:
            w = json.loads(WIN_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        try:
            self._split = min(max(float(w.get("split", SPLIT_DEFAULT)), SPLIT_MIN), SPLIT_MAX)
        except (TypeError, ValueError):
            pass
        m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", w.get("geometry", ""))
        if m:
            wd, ht, x, y = map(int, m.groups())
            # 位置在螢幕外（例如外接螢幕拔掉了）就只套大小
            self.geometry(f"{wd}x{ht}+{x}+{y}" if self._on_screen(x, y) else f"{wd}x{ht}")
        if w.get("zoomed"):
            self.state("zoomed")

    def _on_screen(self, x, y):
        try:  # 虛擬桌面（含所有螢幕）的範圍
            u = ctypes.windll.user32
            vx, vy, vw, vh = (u.GetSystemMetrics(i) for i in (76, 77, 78, 79))
        except Exception:
            vx, vy, vw, vh = 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()
        return vx - 50 <= x < vx + vw - 100 and vy - 50 <= y < vy + vh - 100

    def _on_close(self):
        if self._after:  # 剛打完字、還沒自動存的先存
            self.after_cancel(self._after)
            self._after_typing()
        if self._toast_job:
            self.after_cancel(self._toast_job)
        try:
            zoomed = self.state() == "zoomed"
            old = {}
            if WIN_FILE.exists():
                old = json.loads(WIN_FILE.read_text(encoding="utf-8"))
            # 最大化時記原本的大小，下次還原才不會變成全螢幕大小
            geo = old.get("geometry", "") if zoomed else self.geometry()
            WIN_FILE.write_text(json.dumps({"geometry": geo, "zoomed": zoomed, "split": round(self._split, 3)}),
                                encoding="utf-8")
        except Exception:
            pass
        self.destroy()

    def _style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        # 唯讀下拉選單：白底、不反白選取文字，看起來跟一般欄位一樣
        st.configure("Pick.TCombobox", bordercolor=LINE, lightcolor=BG, darkcolor=BG, arrowcolor=MUTED,
                     padding=(6, 3))
        st.map("Pick.TCombobox",
               fieldbackground=[("readonly", BG)], background=[("readonly", BG), ("active", HOVER)],
               selectbackground=[("readonly", BG)], selectforeground=[("readonly", TEXT)],
               foreground=[("readonly", TEXT)], bordercolor=[("focus", BLUE)])
        self.option_add("*TCombobox*Listbox.font", FONT_SM)
        for opt, color in (("background", BG), ("foreground", TEXT),
                           ("selectBackground", SELECT_BG), ("selectForeground", TEXT)):
            self.option_add(f"*TCombobox*Listbox.{opt}", color)
        st.configure("Thin.Vertical.TScrollbar", background=LINE, troughcolor=BG, bordercolor=BG,
                     lightcolor=LINE, darkcolor=LINE, arrowsize=0, gripcount=0, width=8)
        st.map("Thin.Vertical.TScrollbar", background=[("active", SCROLL_ACTIVE)])
        st.layout("Thin.Vertical.TScrollbar",
                  [("Vertical.Scrollbar.trough", {"children": [("Vertical.Scrollbar.thumb",
                                                                {"expand": "1", "sticky": "nswe"})],
                                                  "sticky": "ns"})])

    # ---------- 版面 ----------
    def _text(self, parent, height=None):
        t = tk.Text(parent, font=FONT, wrap="word", undo=True, relief="flat", bg=BG, fg=TEXT,
                    padx=14, pady=12, spacing1=2, spacing3=2, insertbackground=TEXT,
                    selectbackground=SELECT_BG, selectforeground=TEXT,
                    highlightthickness=1, highlightbackground=LINE, highlightcolor=BLUE)
        if height:
            t.configure(height=height)
        return t

    # ---------- 左右欄分隔線拖拉 ----------
    def _apply_split(self):
        # uniform 欄位會照 weight 比例分寬度；中間分隔線欄固定寬、不分配
        w = round(self._split * 1000)
        self._body.columnconfigure(0, weight=w, uniform="c")
        self._body.columnconfigure(1, weight=0)
        self._body.columnconfigure(2, weight=1000 - w, uniform="c")

    def _set_split(self, ratio):
        ratio = min(max(ratio, SPLIT_MIN), SPLIT_MAX)
        if abs(ratio - self._split) > 0.001:
            self._split = ratio
            self._apply_split()

    def _split_press(self, e):
        x, y, w, h = self._body.grid_bbox(0, 0)
        self._drag_split = (e.x_root, w)  # 按下時的滑鼠位置、左欄寬度

    def _split_move(self, e):
        if not self._drag_split:
            return
        x0, left_w = self._drag_split
        avail = self._body.grid_bbox(0, 0, 2, 0)[2] - self._body.grid_bbox(1, 0)[2]  # 左右兩欄總寬
        if avail > 0:
            self._set_split((left_w + e.x_root - x0) / avail)

    def _heading(self, parent, title, sub=""):
        f = tk.Frame(parent, bg=BG)
        tk.Label(f, text=title, font=FONT_H, bg=BG, fg=TEXT).pack(side="left")
        if sub:
            tk.Label(f, text=sub, font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(10, 0), pady=(3, 0))
        return f

    def _build_ui(self):
        # 頁首：週報日期列
        head = tk.Frame(self, bg=BG, padx=40, pady=24)
        head.pack(fill="x")
        Button(head, "⚙ 設定", self.open_settings, kind="ghost", font=FONT_SM).place(relx=1, y=0, anchor="ne")
        prop = tk.Frame(head, bg=BG)
        prop.pack(anchor="w")
        tk.Label(prop, text="週報日期", font=FONT_DATE, bg=BG, fg=MUTED, anchor="w").pack(side="left", padx=(0, 12))
        self.date = tk.StringVar(value=f"{R.this_friday():%Y.%m.%d}")
        ent = tk.Entry(prop, textvariable=self.date, width=11, font=FONT_DATE, relief="flat", bg=BG, fg=TEXT,
                       insertbackground=TEXT, highlightthickness=1, highlightbackground=BG,
                       highlightcolor=BLUE)
        ent.pack(side="left", ipady=3, ipadx=4)
        ent.bind("<Enter>", lambda e: ent.configure(bg=HOVER))
        ent.bind("<Leave>", lambda e: ent.configure(bg=BG))
        self.date_ent = ent
        self.date_tag = tk.Label(prop, text="", font=FONT_SM, bg=SIDE, fg=MUTED, padx=6, pady=1)
        self.date_tag.pack(side="left", padx=(6, 0))
        self.date_tag.bind("<Button-1>", lambda e: self._reset_date())
        self.date.trace_add("write", self._check_date)
        self._check_date()
        tk.Frame(self, bg=LINE, height=1).pack(fill="x", padx=40)

        # 底部列（先 pack，視窗縮小時才不會被擠掉）
        self._build_footer()
        tk.Frame(self, bg=LINE, height=1).pack(side="bottom", fill="x")

        body = tk.Frame(self, bg=BG, padx=40, pady=20)
        body.pack(fill="both", expand=True)
        # 處理中的遮罩：蓋住左右兩欄，避免看到待辦重畫、消失以為資料被刪（平常不顯示）
        self.mask = tk.Frame(body, bg=SIDE)
        center = tk.Frame(self.mask, bg=SIDE)
        center.place(relx=0.5, rely=0.45, anchor="center")
        self.mask_spin = tk.Label(center, text="", font=(FAMILY, 28), bg=SIDE, fg=BLUE)
        self.mask_spin.pack()
        self.mask_text = tk.Label(center, text="", font=FONT_H, bg=SIDE, fg=TEXT)
        self.mask_text.pack(pady=(8, 4))
        self.mask_sub = tk.Label(center, text="", font=FONT_SM, bg=SIDE, fg=MUTED, justify="center")
        self.mask_sub.pack()
        self._body = body
        self._apply_split()
        body.rowconfigure(0, weight=1)

        # 中：可左右拖拉的分隔線（滑過變藍、雙擊回預設比例）
        sash = tk.Frame(body, bg=BG, width=28, cursor="sb_h_double_arrow")
        sash.grid(row=0, column=1, sticky="ns")
        sash_line = tk.Frame(sash, bg=LINE, width=1, cursor="sb_h_double_arrow")
        sash_line.place(relx=0.5, y=0, relheight=1, anchor="n")
        for w in (sash, sash_line):
            w.bind("<Enter>", lambda e: sash_line.configure(bg=BLUE, width=3))
            w.bind("<Leave>", lambda e: self._drag_split is None and sash_line.configure(bg=LINE, width=1))
            w.bind("<ButtonPress-1>", self._split_press)
            w.bind("<B1-Motion>", self._split_move)
            w.bind("<ButtonRelease-1>", lambda e: (setattr(self, "_drag_split", None),
                                                   sash_line.configure(bg=LINE, width=1)))
            w.bind("<Double-Button-1>", lambda e: self._set_split(SPLIT_DEFAULT))
        self._drag_split = None

        # 左：工作紀錄
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, sticky="nsew")
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        bar = tk.Frame(left, bg=BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._heading(bar, "本週工作紀錄").pack(side="left")
        Button(bar, "全部清空", self.clear_work, kind="outline", font=FONT_SM).pack(side="right")
        self.save_state_lbl = tk.Label(bar, text="", font=FONT_SM, bg=BG, fg=MUTED)
        self.save_state_lbl.pack(side="right", padx=8)
        box = tk.Frame(left, bg=BG, highlightthickness=1, highlightbackground=LINE)
        box.grid(row=1, column=0, sticky="nsew")
        self.wcanvas = tk.Canvas(box, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.wcanvas.yview, style="Thin.Vertical.TScrollbar")
        self.fields_frame = tk.Frame(self.wcanvas, bg=BG)
        self.fields_frame.bind("<Configure>",
                               lambda e: self.wcanvas.configure(scrollregion=self.wcanvas.bbox("all")))
        self._wwin = self.wcanvas.create_window((0, 0), window=self.fields_frame, anchor="nw")
        self.wcanvas.bind("<Configure>", lambda e: self.wcanvas.itemconfigure(self._wwin, width=e.width))
        self.wcanvas.configure(yscrollcommand=sb.set)
        self.wcanvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self.fields = {}  # 專案 -> Text
        self._build_fields(R.load_draft())
        self._saved = self._work_text()
        self._update_saved_lbl()
        self.preview = tk.Label(left, text="", font=FONT_SM, bg=BG, fg=MUTED, anchor="w", justify="left")
        self.preview.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        self.preview.bind("<Configure>", lambda e: self.preview.configure(wraplength=e.width))

        # 右：待辦
        right = tk.Frame(body, bg=BG)
        right.grid(row=0, column=2, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)
        bar = tk.Frame(right, bg=BG)
        bar.grid(row=0, column=0, sticky="ew")
        self._heading(bar, "待辦事項").pack(side="left")
        self._todo_bar = bar
        self.import_btn = Button(bar, "從週報匯入", self.upload, kind="outline", font=FONT_SM)
        self.import_btn.pack(side="right")
        self.count_lbl = tk.Label(right, text="", font=FONT_SM, bg=BG, fg=MUTED, anchor="w")
        self.count_lbl.grid(row=1, column=0, sticky="ew", pady=(4, 8))

        box = tk.Frame(right, bg=BG, highlightthickness=1, highlightbackground=LINE)
        box.grid(row=2, column=0, sticky="nsew")
        self.canvas = tk.Canvas(box, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.canvas.yview,
                           style="Thin.Vertical.TScrollbar")
        self.list = tk.Frame(self.canvas, bg=BG)
        self.list.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.list, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self.canvas.bind_all("<MouseWheel>", self._wheel)

        add = tk.Frame(right, bg=BG)
        add.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        add.columnconfigure(0, weight=1)
        pick = tk.Frame(add, bg=BG)
        pick.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        tk.Label(pick, text="加到專案", font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=(0, 8))
        # 唯讀：整個框點下去都會展開清單（新專案從左邊「＋ 新增專案」加）
        self.add_proj = ttk.Combobox(pick, font=FONT_SM, values=self.state_["projects"], state="readonly",
                                     style="Pick.TCombobox", cursor="hand2")
        self.add_proj.pack(side="left", fill="x", expand=True)
        if self.state_["projects"]:
            self.add_proj.set(self.state_["projects"][0])
        self.new = self._text(add, height=3)
        self.new.grid(row=1, column=0, sticky="ew")
        self.new.insert("1.0", NEW_HINT)
        self.new.configure(fg=HINT_FG)
        self._setup_list_box(self.new, NEW_HINT, reset=True)
        self.new.bind("<Control-Return>", lambda e: (self.add(), "break")[1])
        Button(add, "＋ 新增待辦", self.add, kind="ghost", font=FONT_SM).grid(row=2, column=0, sticky="w", pady=(4, 0))
        tk.Label(add, text="Ctrl+Enter", font=FONT_SM, bg=BG, fg=HINT_FG).grid(row=2, column=0, sticky="e")

    def _build_footer(self):
        foot = tk.Frame(self, bg=SIDE, padx=40, pady=14)
        foot.pack(side="bottom", fill="x")
        self.status = tk.Label(foot, text="", font=FONT, bg=SIDE, fg=GREEN, anchor="w")
        self.status.pack(side="left")
        self.undo_btn = Button(foot, "復原", lambda: None, kind="outline", font=FONT_SM, padx=8, pady=2)

        # 由右往左：產生並寄送 → 分隔 → 重寄這份舊檔
        self.gen_btn = Button(foot, f"產生並開 {self._mail_name()}", self.generate, kind="primary",
                              font=FONT_B, padx=18, pady=7)
        self.gen_btn.pack(side="right")
        self.gen_step = tk.Label(foot, text="", font=FONT_SM, bg=SIDE, fg=MUTED)
        self.gen_step.pack(side="right", padx=(0, 6))
        tk.Frame(foot, bg=LINE, width=1).pack(side="right", fill="y", padx=18)
        self.resend_btn = resend = Button(foot, "重寄這份", self.send, kind="outline", font=FONT_SM, padx=10, pady=5)
        resend.pack(side="right")
        self._tooltip(resend, "不重新產生週報，直接用左邊這份 Word 開郵件", above=True)
        f = R.output_path(R.this_friday())
        self.mail_file = tk.StringVar(value=str(f) if f.exists() else "")
        self.file_chip = Button(foot, "", self.pick_mail_file, kind="ghost", font=FONT_SM, padx=8, pady=7)
        self.file_chip.pack(side="right", padx=(0, 6))
        self._tooltip(self.file_chip, "要寄出的週報 Word，點一下可以換成別份", above=True)
        self.mail_file.trace_add("write", lambda *_: self._update_chip())
        self._update_chip()

    def _update_chip(self):
        p = self.mail_file.get()
        self.file_chip.configure(text=f"附件　{os.path.basename(p)}" if p else "＋ 選擇要寄的檔案",
                                 fg=TEXT if p else MUTED)

    # ---------- 清單輸入框（新增待辦、本週工作紀錄）：自動編號、Tab 變子項目 ----------
    def _setup_list_box(self, t, hint, reset):
        """t 已放好內容或提示字（灰字）。reset：遇到專案名就從 1 重新編號（只有新增待辦框需要）。"""
        def focus_in(_):  # 點進來：拿掉提示字，空的就從「1. 」開始
            if t.cget("fg") == HINT_FG:
                t.delete("1.0", "end")
                t.configure(fg=TEXT)
            if not t.get("1.0", "end-1c").strip():
                t.delete("1.0", "end")
                t.insert("1.0", ITEM_PREFIX.format(1))

        def focus_out(_):  # 離開時只剩「1.」就放回提示字
            if t.cget("fg") != HINT_FG and t.get("1.0", "end-1c").strip() in ("", "1."):
                t.delete("1.0", "end")
                t.configure(fg=HINT_FG)
                t.insert("1.0", hint)
        t.bind("<FocusIn>", focus_in, add="+")
        t.bind("<FocusOut>", focus_out, add="+")
        t.bind("<Return>", lambda e: self._list_enter(t, reset))
        t.bind("<KP_Enter>", lambda e: self._list_enter(t, reset))
        t.bind("<Tab>", lambda e: self._list_indent(t, True, reset))
        t.bind("<Shift-Tab>", lambda e: self._list_indent(t, False, reset))
        t.bind("<<Paste>>", lambda e: self._list_paste(t, reset))

    def _list_paste(self, t, reset):
        """貼上多行文字時自動整理成清單：沒編號的行補「1. 」，「-」「•」開頭的變小項，最後重新編號。
        不整理的話，沒編號的行會被當成模組名，週報裡變 0 筆。單行文字照一般貼上。"""
        try:
            clip = self.clipboard_get()
        except tk.TclError:
            return "break"
        if "\n" not in clip.strip():
            return None
        names = set(self.state_["projects"]) | set(R.ALIASES)
        out = []
        for raw in clip.splitlines():
            s = raw.strip()
            if not s:
                continue
            if reset and s in names:  # 新增待辦框：專案名那行保留，照舊格式分專案
                out.append(s)
            elif R.LV1.match(s) and out and out[-1] not in names:
                out.append(CHILD_PREFIX + R.LV1.sub("", s))
            else:
                out.append(ITEM_PREFIX.format(1) + R.LV0.sub("", R.LV1.sub("", s)))
        if not out:
            return "break"
        try:
            t.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        if t.cget("fg") == HINT_FG:
            t.delete("1.0", "end")
            t.configure(fg=TEXT)
        line = t.get("insert linestart", "insert lineend")
        m = R.LV0.match(line) or R.LV1.match(line)
        if not line.strip() or (m and not line[m.end():].strip()):  # 空行或空的「1. 」→ 直接用這行
            t.delete("insert linestart", "insert lineend")
            lead = ""
        else:  # 游標在有字的行 → 接在這一項（含底下小項）後面，不切斷原本的字、不拆散小項
            row = int(t.index("insert").split(".")[0])
            last = int(t.index("end-1c").split(".")[0])
            while row < last and R.LV1.match(t.get(f"{row + 1}.0", f"{row + 1}.end")):
                row += 1
            t.mark_set("insert", f"{row}.end")
            lead = "\n"
        t.insert("insert", lead + "\n".join(out))
        self._renumber(t, reset)
        t.see("insert")
        return "break"

    def _list_enter(self, t, reset):
        """換行時自動帶編號：主項 → 下一個號碼，小項 → 繼續小項；空的那行再按 Enter 就退一層。"""
        line = t.get("insert linestart", "insert lineend")
        m0, m1 = R.LV0.match(line), R.LV1.match(line)
        if m1 and not line[m1.end():].strip():  # 空的小項 → 退回主項
            t.delete("insert linestart", "insert lineend")
            t.insert("insert", ITEM_PREFIX.format(1))
        elif m0 and not line[m0.end():].strip():  # 空的主項 → 拿掉編號（可接著打專案名或模組名）
            t.delete("insert linestart", "insert lineend")
        elif not line.strip():  # 空行就單純換行
            t.insert("insert", "\n")
        else:
            t.insert("insert", "\n" + (CHILD_PREFIX if m1 else ITEM_PREFIX.format(1)))
        self._renumber(t, reset)
        t.see("insert")
        return "break"

    def _list_indent(self, t, deeper, reset):
        """Tab：主項 → 上一筆的小項；Shift+Tab：小項 → 主項。"""
        if t.cget("fg") == HINT_FG:
            return "break"
        row = int(t.index("insert").split(".")[0])
        line = t.get(f"{row}.0", f"{row}.end")
        prev = t.get(f"{row - 1}.0", f"{row - 1}.end") if row > 1 else ""
        m = (R.LV0 if deeper else R.LV1).match(line)
        if m is None or (deeper and not (R.LV0.match(prev) or R.LV1.match(prev))):
            return "break"  # 第一筆前面沒有主項，不能變小項
        t.delete(f"{row}.0", f"{row}.{m.end()}")
        t.insert(f"{row}.0", CHILD_PREFIX if deeper else ITEM_PREFIX.format(1))
        self._renumber(t, reset)
        return "break"

    def _renumber(self, t, reset):
        """主項依序重編 1、2、3（只換數字，「1.」「1. 」原本的寫法不動）；小項不佔號。"""
        n = 0
        for i, line in enumerate(t.get("1.0", "end-1c").split("\n"), 1):
            if R.LV0.match(line):
                n += 1
                d = NUM.match(line)
                if d.group(2) != str(n):
                    t.delete(f"{i}.{d.start(2)}", f"{i}.{d.end(2)}")
                    t.insert(f"{i}.{d.start(2)}", str(n))
            elif reset and line.strip() and not R.LV1.match(line):
                n = 0

    def _wheel(self, e):
        x, y = self.winfo_pointerxy()
        w = str(self.winfo_containing(x, y))
        for c in (self.canvas, self.wcanvas):
            if w.startswith(str(c)):
                c.yview_scroll(int(-e.delta / 120), "units")

    def _on_canvas_resize(self, e):
        self.canvas.itemconfigure(self._win, width=e.width)
        for lbl, indent in self._wrap_labels:
            lbl.configure(wraplength=max(e.width - indent - 100, 100))  # 留位置給「自動」小標和 ✕

    # ---------- 工作紀錄：每個專案一格 ----------
    def _build_fields(self, text):
        """依專案清單重畫每一格；紀錄裡出現清單沒有的專案會自動加進清單。"""
        parts = R.split_log(text, self.state_["projects"])
        new = [p for p in parts if p not in self.state_["projects"]]
        for p in new:
            R.resolve(self.state_, p)
        if new:
            R.save_state(self.state_)
        for w in self.fields_frame.winfo_children():
            w.destroy()
        self.fields = {}

        # 新增專案放最上面；框內灰字提示，點進去就清掉
        add = tk.Frame(self.fields_frame, bg=BG)
        add.pack(fill="x", padx=14, pady=(14, 4))
        # Entry 沒有內距，包一層框：框負責邊線，Entry 左右留 10px，跟下面各專案輸入框一樣
        wrap = tk.Frame(add, bg=BG, highlightthickness=1, highlightbackground=LINE, cursor="xterm")
        wrap.pack(side="left", fill="x", expand=True)
        e = self.new_proj = tk.Entry(wrap, font=FONT, relief="flat", bg=BG, fg=HINT_FG, insertbackground=TEXT,
                                     highlightthickness=0, borderwidth=0)
        e.insert(0, NEW_PROJ_HINT)
        e.pack(fill="x", padx=10, pady=(4, 3))
        wrap.bind("<Button-1>", lambda _: e.focus_set())

        def focus_in(_):
            wrap.configure(highlightbackground=BLUE)
            if e.cget("fg") == HINT_FG:
                e.delete(0, "end")
                e.configure(fg=TEXT)

        def focus_out(_):
            wrap.configure(highlightbackground=LINE)
            if not e.get().strip():
                e.delete(0, "end")
                e.configure(fg=HINT_FG)
                e.insert(0, NEW_PROJ_HINT)
        e.bind("<FocusIn>", focus_in)
        e.bind("<FocusOut>", focus_out)
        e.bind("<Return>", lambda _: self.add_project())
        Button(add, "＋ 新增專案", self.add_project, kind="ghost", font=FONT_SM).pack(side="left", padx=(6, 0))

        for proj in self.state_["projects"]:
            self._field(proj, parts.get(proj, ""))
        tk.Frame(self.fields_frame, bg=BG, height=14).pack()  # 最後一格下方留白

    def _field(self, proj, content):
        h = tk.Frame(self.fields_frame, bg=BG)
        h.pack(fill="x", padx=14, pady=(14, 4))
        tk.Label(h, text=proj, font=FONT_B, bg=BG, fg=TEXT).pack(side="left")
        del_btn = Button(h, "✕ 刪除專案", lambda: self.delete_project(proj), kind="ghost", font=FONT_SM,
                         padx=6, pady=0)
        t = self._text(self.fields_frame, height=2)

        def show_del(_=None):  # 有內容 → 清空內容；空的 → 刪除專案
            if self._field_text(t).strip():
                del_btn.configure(text="✕ 清空內容")
                del_btn._cmd = lambda: (self.clear_field(proj), del_btn.pack_forget())
            else:
                del_btn.configure(text="✕ 刪除專案")
                del_btn._cmd = lambda: self.delete_project(proj)
            del_btn.pack(side="right")

        def hide_del(_=None):  # 滑鼠還在這個專案的標題或輸入框上就不收
            def check():
                if not del_btn.winfo_exists():  # 欄位已重畫
                    return
                w = self.winfo_containing(*self.winfo_pointerxy())
                if w is None or not (w in (h, t) or str(w).startswith(str(h) + ".")):
                    del_btn.pack_forget()
            self.after(30, check)
        for w in (h, t, del_btn):
            w.bind("<Enter>", show_del, add="+")
            w.bind("<Leave>", hide_del, add="+")
        t.configure(padx=10, pady=6)
        t.pack(fill="x", padx=14)
        if content:
            t.insert("1.0", content)
        else:
            t.configure(fg=HINT_FG)
            t.insert("1.0", FIELD_HINT)

        self._setup_list_box(t, FIELD_HINT, reset=False)
        t.edit_modified(False)
        t.bind("<<Modified>>", lambda e: self._on_work_change(t))
        t.bind("<Configure>", lambda e: self._fit(t), add="+")
        self.fields[proj] = t

    def _fit(self, t):
        """欄位高度跟著內容長（至少 2 行）。"""
        n = t.count("1.0", "end", "displaylines")
        n = (n[0] if isinstance(n, tuple) else n) or 1
        if int(t.cget("height")) != max(n, 2):
            t.configure(height=max(n, 2))

    def _field_text(self, t):
        return "" if t.cget("fg") == HINT_FG else t.get("1.0", "end")

    def _work_text(self):
        return R.join_log([(p, self._field_text(t)) for p, t in self.fields.items()])

    def _set_work(self, text):
        """把整段工作紀錄分回各專案欄位；有沒見過的專案就重畫欄位。"""
        parts = R.split_log(text, self.state_["projects"])
        if any(p not in self.fields for p in parts):
            self._build_fields(text)
            self._after_typing()
            return
        for proj, t in self.fields.items():
            content = parts.get(proj, "")
            if self._field_text(t).strip() == content.strip():
                continue
            t.delete("1.0", "end")
            if content:
                t.configure(fg=TEXT)
                t.insert("1.0", content)
            else:
                t.configure(fg=HINT_FG)
                t.insert("1.0", FIELD_HINT)

    def add_project(self):
        name = "" if self.new_proj.cget("fg") == HINT_FG else self.new_proj.get().strip()
        if not name:
            return
        known = self.state_["projects"]
        k = name if name in known else (R.ALIASES.get(name) if R.ALIASES.get(name) in known else None)
        if k:
            self.fields[k].focus_set()
            self._toast(f"已經有「{k}」", MUTED)
            return
        text = self._work_text()
        R.resolve(self.state_, name)
        R.save_state(self.state_)
        self._build_fields(text)
        self.refresh_todos(auto=False)
        self.fields[name].focus_set()
        self._toast(f"✓ 已新增專案「{name}」")

    def clear_field(self, proj):
        """只清掉這個專案的本週工作紀錄，專案和待辦都留著。"""
        t = self.fields[proj]
        if not messagebox.askyesno("週報小幫手", f"清空「{proj}」的本週工作紀錄？\n（專案和待辦都會留著）"):
            return
        t.delete("1.0", "end")
        t.configure(fg=HINT_FG)
        t.insert("1.0", FIELD_HINT)
        self._inserted = {k for k in self._inserted if k[0] != proj}
        self._after_typing()
        self._toast(f"✓ 已清空「{proj}」的紀錄")

    def delete_project(self, proj):
        n = sum(1 + len(it["children"]) for it in self.state_["todos"].get(proj, []))
        log = bool(self._field_text(self.fields[proj]).strip())
        what = "、".join(x for x in ["本週紀錄" if log else "", f"{n} 筆待辦" if n else ""] if x)
        if not messagebox.askyesno("週報小幫手", f"刪除專案「{proj}」？\n"
                                   + (f"{what}會一起刪掉，" if what else "")
                                   + "之後的週報 Word 也不會再有這一列。"):
            return
        known = self.state_["projects"]
        rest = [(p, c) for p, c in R.split_log(self._work_text(), known).items() if p != proj]
        known.remove(proj)
        self.state_["todos"].pop(proj, None)
        R.save_state(self.state_)
        self.manual = {k for k in self.manual if k[0] != proj}
        self._inserted = {k for k in self._inserted if k[0] != proj}
        self._build_fields(R.join_log(rest))
        self.refresh_todos(auto=False)
        self._after_typing()
        self._toast(f"✓ 已刪除專案「{proj}」")

    def _check_update(self):
        """背景查 GitHub 有沒有新版；有就在底部顯示提醒和「下載」按鈕。"""
        def run():
            r = R.check_update()
            if r:
                self.after(0, lambda: self._toast(f"有新版本 v{r[0]}（目前 v{R.VERSION}）", BLUE,
                                                  undo=lambda: __import__("webbrowser").open(r[1]),
                                                  label="下載"))
        threading.Thread(target=run, daemon=True).start()

    def _toast(self, text, color=None, undo=None, label="復原"):
        """成功訊息（綠）6 秒後清掉；進行中、提醒（灰）留著。undo：訊息旁顯示 label 按鈕，按了執行。"""
        if self._toast_job:
            self.after_cancel(self._toast_job)
            self._toast_job = None
        self.status.configure(text=text, fg=color or GREEN)  # 預設綠色在執行時才取，切換主題後才會跟著變
        self.undo_btn.pack_forget()
        self.undo_btn.configure(text=label)
        if undo:
            def do():
                self._clear_toast()
                undo()
            self.undo_btn._cmd = do
            self.undo_btn.pack(side="left", padx=(10, 0), after=self.status)
        if color == GREEN:
            self._toast_job = self.after(6000, self._clear_toast)

    def _clear_toast(self):
        self._toast_job = None
        self.status.configure(text="")
        self.undo_btn.pack_forget()

    def _run_busy(self, text, work, done, fail, prefix="", mask=None):
        """耗時工作丟背景執行緒，畫面不凍結；期間底部列轉圈、相關按鈕停用。
        做完回主執行緒呼叫 done(結果)，出錯呼叫 fail(例外)（tkinter 只能在主執行緒動畫面）。
        mask：給說明文字就用遮罩蓋住左右兩欄，中間顯示轉圈 + text + 說明。
        done 裡若接著再 _run_busy(mask=...)，同一輪事件內重新蓋上，不會閃。"""
        res = {}

        def run():
            try:
                res["ok"] = work()
            except Exception as e:
                res["err"] = e

        threading.Thread(target=run, daemon=True).start()
        self._clear_toast()
        btns = (self.resend_btn, self.import_btn)
        self._busy = True
        self.gen_btn.set_enabled(False)
        for b in btns:
            b.set_enabled(False)
        self.configure(cursor="watch")
        if mask is not None:
            self.mask_text.configure(text=text)
            self.mask_sub.configure(text=mask)
            self.mask.place(x=0, y=0, relwidth=1, relheight=1)
            self.mask.lift()

        def tick(i=0):
            if res:
                self._busy = False
                for b in btns:
                    b.set_enabled(True)
                self._update_preview()  # 產生鈕照「有沒有寫紀錄」恢復
                self.configure(cursor="")
                self.status.configure(text="")
                self.mask.place_forget()
                done(res["ok"]) if "ok" in res else fail(res["err"])
                return
            frame = SPIN[i % len(SPIN)]
            self.status.configure(text=f"{prefix}{frame} {text}", fg=MUTED)
            if mask is not None:
                self.mask_spin.configure(text=frame)
            self.after(80, tick, i + 1)
        tick()

    # ---------- 待辦清單 ----------
    def refresh_todos(self, auto=True):
        for w in self.list.winfo_children():
            w.destroy()
        self.vars = {}
        self._wrap_labels = []
        self._rows = []    # 畫面順序的每一列：key（佔位列為 None）、proj、frame、lbl
        self._heads = {}   # 專案 -> 標題 frame
        projects = self.state_["projects"]
        self.add_proj.configure(values=projects)
        if self.add_proj.get() not in projects:  # 選的專案被刪了 → 回第一個
            self.add_proj.set(projects[0] if projects else "")
        for proj in self.state_["projects"]:
            items = self.state_["todos"].setdefault(proj, [])
            h = tk.Frame(self.list, bg=BG)
            h.pack(fill="x", padx=14, pady=(14, 4))
            self._heads[proj] = h
            name = tk.Label(h, text=proj, font=FONT_B, bg=BG, fg=TEXT, cursor="hand2")
            name.pack(side="left")
            for w in (h, name):  # 點專案標題 → 下面「加到專案」切到這個專案
                w.bind("<Button-1>", lambda e, p=proj: self.add_proj.set(p))
            n = len(items) + sum(len(it["children"]) for it in items)
            if n:
                tk.Label(h, text=str(n), font=FONT_SM, bg=BG, fg=MUTED).pack(side="left", padx=6)
            if not items:
                self._placeholder(proj)
            for i, it in enumerate(items):
                self._check((proj, i), f"{i + 1}. {it['text']}", 10)
                for j, c in enumerate(it["children"]):
                    self._check((proj, i, j), f"- {c}", 36)
        tk.Frame(self.list, bg=BG, height=10).pack()
        self._drop_line = tk.Frame(self.list, bg=BLUE, height=2)
        self._on_canvas_resize(type("E", (), {"width": self.canvas.winfo_width()}))
        if auto:
            self.auto_check()
        self._update_count()

    def _placeholder(self, proj):
        row = tk.Frame(self.list, bg=BG)
        row.pack(fill="x", padx=(4, 4))
        lbl = tk.Label(row, text="雙擊新增，或把待辦拖到這裡", font=FONT_SM, bg=BG, fg=HINT_FG, anchor="w")
        lbl.pack(side="left", fill="x", expand=True, padx=(36, 0), pady=5)
        r = dict(key=None, proj=proj, frame=row, lbl=lbl)
        self._rows.append(r)

        def new(_):
            def done(t):
                if t.strip():
                    self._commit({proj: [(0, t.strip(), None)]})
            self._inline_entry(r, "", done, padx=(36, 10))
        for w in (row, lbl):
            w.bind("<Double-Button-1>", new)

    def _check(self, key, text, indent):
        v = tk.BooleanVar(value=False)
        self.vars[key] = v
        row = tk.Frame(self.list, bg=BG, cursor="hand2")
        row.pack(fill="x", padx=(4, 4))
        box = tk.Canvas(row, width=18, height=18, bg=BG, highlightthickness=0)
        box.pack(side="left", anchor="n", padx=(indent, 8), pady=6)
        lbl = tk.Label(row, text=text, font=FONT, bg=BG, fg=TEXT, anchor="w", justify="left")
        lbl.pack(side="left", fill="x", expand=True, pady=4)
        # ✕ 刪除鈕平常跟底色同色（看不到），滑到這列才顯示；位置固定，換行才不會跳
        x = tk.Label(row, text="✕", font=FONT_SM, bg=BG, fg=BG, cursor="hand2", padx=4)
        x.pack(side="right", anchor="n", padx=(0, 6), pady=5, before=lbl)
        tag = tk.Label(row, text="自動", font=FONT_SM, bg=BG, fg=BLUE)  # 依紀錄自動勾的才顯示
        self._wrap_labels.append((lbl, indent))
        self._rows.append(dict(key=key, proj=key[0], frame=row, lbl=lbl))

        def draw(*_):
            box.delete("all")
            if v.get():
                box.create_rectangle(1, 1, 17, 17, fill=BLUE, outline=BLUE)
                box.create_line(4.5, 9, 7.5, 12.5, 13.5, 5.5, fill="white", width=2)
                lbl.configure(fg=HINT_FG, font=FONT_DONE)
            else:
                box.create_rectangle(1, 1, 17, 17, fill=BG, outline=CHECK_LINE, width=1.5)
                lbl.configure(fg=TEXT, font=FONT)
            if v.get() and key not in self.manual:
                tag.pack(side="right", anchor="n", padx=(4, 10), pady=6, before=lbl)
            else:
                tag.pack_forget()
            self._update_count()
        v.trace_add("write", draw)
        draw()

        def hover(on):
            if self._drag:
                return
            c = HOVER if on else BG
            for w in (row, box, lbl, tag, x):
                w.configure(bg=c)
            x.configure(fg=MUTED if on else c)
        x.bind("<Enter>", lambda e: (hover(True), x.configure(fg=RED)))
        x.bind("<Leave>", lambda e: hover(False))
        x.bind("<Button-1>", lambda e: self.after_idle(lambda: self._delete_item(key)))
        for w in (row, box, lbl, tag):
            w.bind("<ButtonPress-1>", lambda e: self._press(e, key))
            w.bind("<B1-Motion>", self._motion)
            w.bind("<ButtonRelease-1>", lambda e, w=w: self._release(key, now=w is box))
            w.bind("<Enter>", lambda e: hover(True), add="+")
            w.bind("<Leave>", lambda e: hover(False), add="+")
        for w in (row, lbl):  # 勾選框不吃雙擊，連點就是勾兩次
            w.bind("<Double-Button-1>", lambda e: self._dbl(key))
        self._tooltip(tag, "依本週紀錄判斷已完成\n產生週報時會刪除，點一下可取消")

    def _tooltip(self, w, text, above=False):
        """滑鼠停在 w 上 0.4 秒後，旁邊浮出說明小框；above：浮在上方（視窗底部的按鈕用）。"""
        tip = {"win": None, "job": None}

        def show():
            tip["job"] = None
            t = tk.Toplevel(self)
            t.wm_overrideredirect(True)
            t.attributes("-topmost", True)
            tk.Label(t, text=text, font=FONT_SM, bg=TEXT, fg=BG, padx=8, pady=4,
                     justify="left").pack()
            t.update_idletasks()
            y = w.winfo_rooty() - t.winfo_reqheight() - 4 if above else w.winfo_rooty() + w.winfo_height() + 4
            t.geometry(f"+{w.winfo_rootx()}+{y}")
            tip["win"] = t

        def hide(_=None):
            if tip["job"]:
                self.after_cancel(tip["job"])
                tip["job"] = None
            if tip["win"]:
                tip["win"].destroy()
                tip["win"] = None
        w.bind("<Enter>", lambda e: tip.update(job=self.after(400, show)), add="+")
        w.bind("<Leave>", hide, add="+")
        w.bind("<ButtonPress-1>", hide, add="+")
        w.bind("<Destroy>", hide, add="+")  # 清單重畫時別留下浮框

    # ---------- 單擊勾選 / 雙擊改字 ----------
    def _toggle(self, key):
        self._click_job = None
        v = self.vars.get(key)
        if v is not None:
            self.manual.add(key)  # 先記成手動，重畫時才不會掛「自動」
            v.set(not v.get())
            self._toggled(key)

    def _dbl(self, key):
        if self._click_job:
            self.after_cancel(self._click_job)
            self._click_job = None
        self._ignore_release = True
        self._edit_item(key)

    def _inline_entry(self, r, text, on_done, padx=(0, 10)):
        """把該列文字換成輸入框；Enter / 點別處 = 存，Esc = 取消。"""
        lbl = r["lbl"]
        info = {("in_" if k == "in" else k): v for k, v in lbl.pack_info().items()}
        ent = tk.Entry(r["frame"], font=FONT, relief="flat", bg=BG, fg=TEXT, insertbackground=TEXT,
                       highlightthickness=1, highlightbackground=BLUE, highlightcolor=BLUE)
        ent.insert(0, text)
        lbl.pack_forget()
        ent.pack(side="left", fill="x", expand=True, padx=padx, pady=2, ipady=3)
        ent.focus_set()
        ent.select_range(0, "end")
        ent.icursor("end")
        closed = [False]

        def finish(save):
            if closed[0]:
                return "break"
            closed[0] = True
            self._edit_close = None
            value = ent.get()
            ent.destroy()
            lbl.pack(**info)
            self.list.focus_set()
            if save:
                self.after_idle(lambda: on_done(value))
            return "break"
        ent.bind("<Return>", lambda e: finish(True))
        ent.bind("<KP_Enter>", lambda e: finish(True))
        ent.bind("<Escape>", lambda e: finish(False))
        ent.bind("<FocusOut>", lambda e: finish(True))
        # 點空白處不會讓輸入框失焦，改由 _click_away 收掉；等這次雙擊處理完才生效，不然第二下就把它關了
        self.after_idle(lambda: closed[0] or setattr(self, "_edit_close", (ent, finish)))

    def _click_away(self, e):
        """編輯待辦中，點輸入框以外的任何地方 = 結束編輯（保留打好的字）。"""
        c = self._edit_close
        if c and e.widget is not c[0]:
            c[1](True)

    def _edit_item(self, key):
        r = next((r for r in self._rows if r["key"] == key), None)
        if r is None:
            return
        proj = key[0]
        it = self.state_["todos"][proj][key[1]]
        old = it["text"] if len(key) == 2 else it["children"][key[2]]

        def done(new):
            new = new.strip()
            if new == old:
                return
            if not new:
                self._delete_item(key)
                return
            flat = self._flat(proj)
            idx = next(i for i, f in enumerate(flat) if f[2] == key)
            flat[idx] = (flat[idx][0], new, key)
            self._commit({proj: flat})
        self._inline_entry(r, old, done)

    def _delete_item(self, key):
        """刪一筆待辦；底下有小項才先確認，單筆直接刪（todos.json.bak 留有上一版）。"""
        proj = key[0]
        it = self.state_["todos"][proj][key[1]]
        old = self._todo_text(key)
        n = len(it["children"]) if len(key) == 2 else 0
        if n and not messagebox.askyesno("週報小幫手", f"刪除「{old}」和底下 {n} 個小項？"):
            return
        # 刪之前的樣子，6 秒內按「復原」整個專案的待辦和勾選狀態放回去
        snap = copy.deepcopy(self.state_["todos"][proj])
        manual, checked = set(self.manual), {k for k, v in self.vars.items() if v.get()}

        def undo():
            self.state_["todos"][proj] = snap
            R.save_state(self.state_)
            self.manual = manual
            self.refresh_todos(auto=False)
            for k in checked:
                if k in self.vars:
                    self.vars[k].set(True)
            self.auto_check()
            self._toast(f"✓ 已復原「{old}」")

        flat = self._flat(proj)
        idx = next(i for i, f in enumerate(flat) if f[2] == key)
        del flat[idx:idx + 1 + n]
        self._commit({proj: flat})
        self._toast(f"✓ 已刪除「{old}」", undo=undo)

    # ---------- 拖拉排序 ----------
    def _flat(self, proj):
        """專案待辦攤平成 [(層級, 文字, 舊 key)]。"""
        out = []
        for i, it in enumerate(self.state_["todos"].get(proj, [])):
            out.append((0, it["text"], (proj, i)))
            out += [(1, c, (proj, i, j)) for j, c in enumerate(it["children"])]
        return out

    def _commit(self, flats):
        """把攤平的清單組回待辦、存檔、重畫，並把勾選狀態搬到新位置。"""
        checked = {k for k, v in self.vars.items() if v.get()}
        moved = {}
        for proj, flat in flats.items():
            items = []
            for lvl, text, old in flat:
                if lvl == 1 and items:
                    items[-1]["children"].append(text)
                    new = (proj, len(items) - 1, len(items[-1]["children"]) - 1)
                else:
                    items.append({"text": text, "children": []})
                    new = (proj, len(items) - 1)
                if old:
                    moved[old] = new
            self.state_["todos"][proj] = items
        # 沒動到的專案 key 不變
        for k in checked | self.manual:
            if k[0] not in flats:
                moved[k] = k
        R.save_state(self.state_)
        self.manual = {moved[k] for k in self.manual if k in moved}
        self.refresh_todos(auto=False)
        for k in checked:
            if moved.get(k) in self.vars:
                self.vars[moved[k]].set(True)
        self.auto_check()

    def _press(self, e, key):
        self._press_at = (e.x_root, e.y_root, key)
        self._drag = None

    def _motion(self, e):
        if not self._press_at:
            return
        x0, y0, key = self._press_at
        if self._drag is None:
            if abs(e.x_root - x0) < 6 and abs(e.y_root - y0) < 6:
                return
            if self._click_job:
                self.after_cancel(self._click_job)
                self._click_job = None
            block = {key} if len(key) == 3 else {k for k in self.vars if k[:2] == key}
            self._drag = dict(key=key, block=block, level=len(key) - 2, target=None)
            for r in self._rows:
                if r["key"] in block:
                    r["lbl"].configure(fg=HINT_FG)
        self._update_drop(e.x_root, e.y_root)

    def _update_drop(self, xr, yr):
        d = self._drag
        cy = yr - self.canvas.winfo_rooty()
        if cy < 24:
            self.canvas.yview_scroll(-1, "units")
        elif cy > self.canvas.winfo_height() - 24:
            self.canvas.yview_scroll(1, "units")
        rows = [r for r in self._rows if r["key"] not in d["block"]]
        projects = self.state_["projects"]
        target = next((r for r in rows if yr < r["frame"].winfo_rooty() + r["frame"].winfo_height() / 2), None)
        if target is None:
            proj, before = projects[-1], None
        else:
            proj, before = target["proj"], target
            first = next(r for r in rows if r["proj"] == proj) is target
            head = self._heads[proj]
            pi = projects.index(proj)
            if first and pi > 0 and yr < head.winfo_rooty() + head.winfo_height():
                proj, before = projects[pi - 1], None  # 游標在標題上 → 放到上一個專案最後
        mine = [r for r in rows if r["proj"] == proj and r["key"]]
        idx = mine.index(before) if before in mine else len(mine)
        lvl = d["level"]
        dx = xr - self._press_at[0]
        if dx > 30:
            lvl = 1
        elif dx < -30:
            lvl = 0
        if idx == 0:
            lvl = 0
        d["target"] = (proj, idx, lvl)

        # 藍色插入線
        if before is not None:
            y = before["frame"].winfo_y()
        elif mine:
            f = mine[-1]["frame"]
            y = f.winfo_y() + f.winfo_height()
        else:
            h = self._heads[proj]
            y = h.winfo_y() + h.winfo_height()
        x = 4 + (10 if lvl == 0 else 36)
        self._drop_line.place(x=x, y=y - 1, width=max(self.list.winfo_width() - x - 14, 20), height=2)
        self._drop_line.lift()

    def _release(self, key, now=False):
        d, self._drag, self._press_at = self._drag, None, None
        if d:
            self._drop_line.place_forget()
            if d["target"]:
                self._move(d["key"], *d["target"])
            return
        if self._ignore_release:  # 雙擊的第二下放開
            self._ignore_release = False
            return
        if now:  # 點勾選框：立刻勾
            self._toggle(key)
        else:  # 點文字：等一下確定不是雙擊才勾
            self._click_job = self.after(250, lambda: self._toggle(key))

    def _move(self, key, proj, idx, lvl):
        src = key[0]
        flat_src = self._flat(src)
        start = next(i for i, f in enumerate(flat_src) if f[2] == key)
        n = 1 if len(key) == 3 else 1 + len(self.state_["todos"][src][key[1]]["children"])
        block = flat_src[start:start + n]
        del flat_src[start:start + n]
        if lvl == 1:  # 降成小項：底下的小項一起併過去
            block = [(1, t, k) for _, t, k in block]
        else:
            block[0] = (0, block[0][1], block[0][2])
        flat_dst = flat_src if proj == src else self._flat(proj)
        flat_dst[idx:idx] = block
        self._commit({src: flat_src, proj: flat_dst})

    def _update_count(self):
        if not hasattr(self, "count_lbl"):
            return
        n = sum(v.get() for v in self.vars.values())
        if n:  # 有勾就整行收起來，下面清單往上補，間距跟左欄一樣
            self.count_lbl.grid_remove()
            self._todo_bar.grid_configure(pady=(0, 8))
        else:
            self.count_lbl.configure(text="打勾 = 已完成，產生週報時刪除（會依紀錄自動勾）")
            self.count_lbl.grid()
            self._todo_bar.grid_configure(pady=0)

    def _toggled(self, key):
        self.manual.add(key)  # 勾大項不連動小項，小項要各自勾
        self._sync_log(key)

    def _sync_log(self, key):
        """手動勾 → 待辦寫進工作紀錄（已經有就不重複加）；取消勾 → 從紀錄找到那行刪掉（重開程式後也有效）。"""
        proj = key[0]
        it = self.state_["todos"][proj][key[1]]
        text = it["text"] if len(key) == 2 else it["children"][key[2]]
        children = [c for j, c in enumerate(it["children"])  # 大項只帶已勾的小項
                    if len(key) == 2 and self.vars.get((proj, key[1], j)) and self.vars[(proj, key[1], j)].get()]
        known = self.state_["projects"]
        work = self._work_text()
        if len(key) == 3:  # 小項：寫在紀錄裡該大項底下的「-」小項
            if self.vars[key].get():
                new = R.insert_child_log_line(work, proj, it["text"], text, known)
            else:
                pv = self.vars.get(key[:2])
                new = R.remove_child_log_line(work, proj, it["text"], text, known,
                                              drop_parent=not (pv and pv.get()))
            if new != work:
                self._set_work(new)
            return
        if self.vars[key].get():
            if R.find_log_line(work, proj, text, known) is None:
                self._set_work(R.insert_log_line(work, proj, text, known, children))
            self._inserted.add((proj, text))
        else:
            self._inserted.discard((proj, text))
            new = R.remove_log_line(work, proj, text, known)
            if new != work:
                self._set_work(new)

    def _on_work_change(self, t):
        t.edit_modified(False)
        self._fit(t)
        if self.focus_get() is t:  # 自己在打字（不是勾待辦自動寫入）才捲
            self.after(20, lambda: self._scroll_to_cursor(t))
        if self._after:
            self.after_cancel(self._after)
        self._after = self.after(400, self._after_typing)

    def _scroll_to_cursor(self, t):
        """欄位變高後，外層清單捲到游標那行看得到的位置。"""
        if not t.winfo_exists():
            return
        bb = t.bbox("insert")
        total = self.fields_frame.winfo_height()
        if not bb or total <= 1:
            return
        y = t.winfo_rooty() - self.fields_frame.winfo_rooty() + bb[1]  # 游標在整份清單中的 y
        top, vis = self.wcanvas.canvasy(0), self.wcanvas.winfo_height()
        if y < top:
            self.wcanvas.yview_moveto((y - 20) / total)
        elif y + bb[3] + 20 > top + vis:
            self.wcanvas.yview_moveto((y + bb[3] + 20 - vis) / total)

    def _after_typing(self):
        text = self._work_text()
        if text.rstrip() != self._saved.rstrip():  # 自動存草稿，關掉重開還在
            R.save_draft(text)
            self._saved = text
        self.auto_check()
        self._update_preview()
        self._update_saved_lbl()

    def _update_preview(self):
        """工作紀錄下方：會寫進哪些專案列、各幾項；對不到的專案用紅字提醒。"""
        known = self.state_["projects"]
        counts, new = {}, []
        for name, items in R.parse(self._work_text(), known):
            k = R.match_project(name, known)
            if k is None and name not in new:
                new.append(name)
            k = k or name
            counts[k] = counts.get(k, 0) + len(items) + sum(len(i["children"]) for i in items)
        text = "→ 會寫進：" + "・".join(f"{k} {n}" for k, n in counts.items()) if counts else ""
        if new:
            text += f"　⚠ 新專案會多一列：{'、'.join(new)}"
        self.preview.configure(text=text, fg=RED if new else MUTED)
        # 還沒寫紀錄就不能產生
        has = sum(counts.values()) > 0
        self.gen_btn.set_enabled(has and not self._busy)  # 處理中也不能按
        self.gen_step.configure(text="" if has else "先在左邊寫本週紀錄", fg=MUTED if has else RED)

    def _update_saved_lbl(self):
        """打字 0.4 秒後就自動存了，這裡只顯示草稿最後存檔時間。"""
        m = R.draft_mtime()
        if not m:
            self.save_state_lbl.configure(text="")
            return
        t = dt.datetime.fromtimestamp(m)
        fmt = "%H:%M" if t.date() == dt.date.today() else "%m/%d %H:%M"
        self.save_state_lbl.configure(text=f"已自動儲存 {t:{fmt}}")

    def auto_check(self):
        done = R.suggest_done(self.state_, R.parse(self._work_text(), self.state_["projects"]))
        for k, v in self.vars.items():
            if k not in self.manual:
                v.set(k in done)  # 小項要自己的文字對得到才勾，不跟著大項

    def add(self):
        text = "" if self.new.cget("fg") == HINT_FG else self.new.get("1.0", "end")
        known = list(self.state_["projects"])
        if R.has_project_line(text, known):  # 照舊格式貼「專案名＋1.」的整段，可一次加多個專案
            n = R.add_todos(self.state_, text)
        else:
            proj = self.add_proj.get().strip()
            if not proj:
                messagebox.showinfo("週報小幫手", "請先在「加到專案」選一個專案。")
                return
            n = R.add_items(self.state_, proj, text) if R.parse_items(text) else 0
        if not n:
            messagebox.showinfo("週報小幫手", "沒有讀到新的待辦，請在框裡每行打一個項目。\n（跟現有待辦同名的不會重複加）")
            return
        new = [p for p in self.state_["projects"] if p not in known]
        R.save_state(self.state_)
        self.new.delete("1.0", "end")
        self.new.insert("1.0", ITEM_PREFIX.format(1))  # 可以直接接著打下一批
        if list(self.fields) != self.state_["projects"]:  # 待辦帶進新專案 → 左邊也多一格
            self._build_fields(self._work_text())
        self.refresh_todos()
        self._toast(f"✓ 已加入 {n} 筆待辦" + (f"，並新增專案「{'、'.join(new)}」" if new else ""))

    def save_work(self):
        text = self._work_text()
        R.save_draft(text)
        self._saved = text
        self._update_saved_lbl()
        self._toast(f"✓ 已儲存工作紀錄（{dt.datetime.now():%H:%M}）")

    def clear_work(self):
        if not self._work_text().strip():
            return
        if not messagebox.askyesno("週報小幫手", "清空所有專案的本週工作紀錄？\n（已產生的週報 Word 不受影響）"):
            return
        self._set_work("")
        R.clear_draft()
        self._saved = ""
        self._inserted.clear()
        self.manual.clear()  # 新的一週重新依紀錄自動勾
        self.auto_check()
        self._update_preview()
        self._update_saved_lbl()
        self._toast("✓ 已清空本週工作紀錄")

    def _import_preview(self, path, projects, todos, n):
        """匯入前的確認文字：每個專案目前幾筆 → 匯入後幾筆，標出新增和會消失的專案。"""
        count = lambda items: sum(1 + len(it["children"]) for it in items)
        cur = self.state_["todos"]
        lines = [f"檔案：{os.path.basename(path)}", f"讀到 {len(projects)} 個專案、{n} 筆待辦：", ""]
        for p in projects:
            new = count(todos[p])
            if p in cur:
                lines.append(f"・{p}：{count(cur[p])} → {new} 筆")
            else:
                lines.append(f"・{p}：{new} 筆（新專案）")
        gone = [p for p in self.state_["projects"] if p not in todos]
        if gone:
            lines += ["", "⚠ 這份週報沒有、匯入後待辦清單會移除的專案："]
            lines += [f"・{p}（目前 {count(cur.get(p, []))} 筆）" for p in gone]
            lines.append("（左邊有本週紀錄的專案會保留欄位）")
        lines += ["", "要用這份週報的待辦取代目前清單嗎？", "（舊清單會備份成 todos.json.bak）"]
        return "\n".join(lines)

    def upload(self):
        path = filedialog.askopenfilename(
            title="選擇最新的週報 Word", initialdir=R.OUT_DIR if R.OUT_DIR.exists() else R.BASE,
            filetypes=[("Word 週報", "*.docx")])
        if not path:
            return

        def loaded(res):
            projects, todos = res
            n = sum(1 + len(it["children"]) for items in todos.values() for it in items)
            if not messagebox.askyesno("週報小幫手", self._import_preview(path, projects, todos, n)):
                return
            text = self._work_text()
            self.state_ = {"projects": projects, "todos": todos}
            R.save_state(self.state_)
            self.manual.clear()
            self._build_fields(text)
            self.refresh_todos()
            self._toast(f"✓ 已從週報更新待辦：{len(projects)} 個專案、{n} 筆")

        def failed(e):
            if isinstance(e, PermissionError):
                messagebox.showwarning("週報小幫手", "這份 Word 正開著，請先關掉再上傳。")
            else:
                messagebox.showwarning("週報小幫手", f"讀不到這份週報：{e}")

        self._run_busy("正在讀取週報…", lambda: R.import_report(path), loaded, failed)

    # ---------- 產生 ----------
    def _parse_date(self):
        try:
            return dt.datetime.strptime(self.date.get().strip().replace("-", ".").replace("/", "."), "%Y.%m.%d").date()
        except ValueError:
            return None

    def _date(self):
        d = self._parse_date()
        if d is None:
            messagebox.showwarning("週報小幫手", "日期格式請寫成 2026.10.02")
        return d

    def _check_date(self, *_):
        """日期欄旁的小標：本週五 / 星期幾；格式錯或不是週五就紅字，不是本週五可點回本週五。"""
        d, friday = self._parse_date(), R.this_friday()
        back = "・回本週五"
        if d is None:
            text, bad = "格式像 2026.10.02" + back, True
        elif d == friday:
            text, bad = "本週五", False
        elif d.weekday() == 4:
            text, bad = "週五" + back, False
        else:
            text, bad = f"週{'一二三四五六日'[d.weekday()]} ⚠ 不是週五" + back, True
        self.date_tag.configure(text=text, fg=RED if bad else MUTED,
                                cursor="hand2" if d != friday else "")
        self.date_ent.configure(highlightbackground=RED if d is None else BG,
                                highlightcolor=RED if d is None else BLUE)

    def _reset_date(self):
        if self._parse_date() != R.this_friday():
            self.date.set(f"{R.this_friday():%Y.%m.%d}")

    def pick_mail_file(self):
        path = filedialog.askopenfilename(
            title="選擇要寄送的週報", initialdir=R.OUT_DIR if R.OUT_DIR.exists() else R.BASE,
            filetypes=[("Word 週報", "*.docx"), ("所有檔案", "*.*")])
        if path:
            self.mail_file.set(os.path.normpath(path))

    def send(self):
        path = self.mail_file.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("週報小幫手", "請先點「＋ 選擇要寄的檔案」挑要寄的週報。")
            return
        date = self._date()
        if date:
            self._open_mail(path, date)

    @staticmethod
    def _mail_name():
        return R.MAIL_NAMES[R.MAIL_VIA]

    def _open_mail(self, path, date, prefix="", mask=None):
        name = self._mail_name()
        app = "Outlook（classic）" if R.MAIL_VIA == "outlook" else "Gmail"
        done = ("Gmail 草稿已開好：把檔案總管選好的週報拖進信裡，確認後按傳送" if R.MAIL_VIA == "gmail"
                else f"Outlook 郵件已開好：{R.mail_subject(date)}，確認後按傳送")

        def failed(e):
            self._toast(prefix.rstrip("・"))
            messagebox.showwarning("週報小幫手", f"開不了 {app}：\n{e}\n\n週報 Word 在：{path}")

        self._run_busy(f"正在開啟 {name}…", lambda: R.send_mail(path, date),
                       lambda _: self._toast(prefix + done), failed, prefix, mask)

    def _todo_text(self, key):
        it = self.state_["todos"][key[0]][key[1]]
        return it["text"] if len(key) == 2 else it["children"][key[2]]

    def _confirm_generate(self, done, date):
        """產生前列出會刪掉的待辦、提醒會覆蓋同名週報。"""
        keys = [r["key"] for r in self._rows if r["key"] in done]  # 照畫面順序
        if keys:
            lines = [f"會刪除 {len(keys)} 筆已完成待辦："]
            lines += [f"・{self._todo_text(k)}" for k in keys[:5]]
            if len(keys) > 5:
                lines.append(f"…另外 {len(keys) - 5} 筆")
        else:
            lines = ["沒有勾選的待辦，待辦清單維持不變。"]
        path = R.output_path(date)
        if path.exists():
            lines.append(f"\n⚠ 已經有「{path.name}」，會被覆蓋。")
        return messagebox.askyesno("週報小幫手", "\n".join(lines) + "\n\n要產生週報嗎？")

    def generate(self):
        date = self._date()
        if not date:
            return
        work = R.parse(self._work_text(), self.state_["projects"])
        if not work:
            messagebox.showwarning("週報小幫手", "請先在左邊貼上本週工作紀錄。")
            return
        done = {k for k, v in self.vars.items() if v.get()}
        state = copy.deepcopy(self.state_)
        R.remove_todos(state, done)
        skip = ()
        empty = R.empty_projects(work, state)
        if empty:
            skip = EmptyProjectsDialog(self, empty).result
            if skip is None:  # 取消 = 先不產生
                return
        if not self._confirm_generate(done, date):
            return
        # 遮罩說明：讓使用者知道待辦是「整理進週報」而不是資料不見
        note = (f"已完成的 {len(done)} 筆待辦會寫進週報，並從待辦清單移除\n工作紀錄都會保留，請稍候"
                if done else "工作紀錄與待辦都會保留，請稍候")

        def built(path):
            self.state_ = state
            R.save_state(state)
            self.refresh_todos(auto=False)
            # 工作紀錄保留（按「清空」才清）；剩下的待辦是使用者決定保留的，之後改紀錄也不要再自動勾
            self.manual = set(self.vars)
            self._inserted.clear()  # 那些待辦已刪除，寫進紀錄的行就是正式工作項目了
            self.mail_file.set(str(path))
            self._open_mail(path, date, f"✓ {path.name}（刪除 {len(done)} 筆已完成待辦）・", note)

        def failed(e):
            if isinstance(e, PermissionError):
                messagebox.showwarning("週報小幫手", "同名的週報 Word 正開著，請先關掉再按一次。")
            else:
                messagebox.showwarning("週報小幫手", f"產生週報失敗：{e}")

        self._run_busy("正在產生週報…", lambda: R.build(work, state, date, skip=skip), built, failed, mask=note)

    # ---------- 個人設定 ----------
    def open_settings(self):
        if not SettingsDialog(self).saved:
            return
        self.gen_btn.configure(text=f"產生並開 {self._mail_name()}")
        if R.THEME != THEME:  # 換主題：不用重開，整個畫面直接換色
            recolor(self, set_palette(R.THEME))
            self._style()
            dark_titlebar(self)
        # 設定裡新列的專案補進清單（已有的專案、待辦不動）
        new = [p for p in R.FIXED_PROJECTS if p not in self.state_["projects"]]
        if new:
            self.state_["projects"] += new
            self.state_ = R.normalize(self.state_)
            R.save_state(self.state_)
            self.refresh_todos(auto=False)
        self._toast("✓ 設定已儲存")


class EmptyProjectsDialog(tk.Toplevel):
    """產生週報前：列出沒內容的專案，逐一勾要不要顯示。result = 不顯示的專案 tuple；按取消是 None。"""

    def __init__(self, parent, projects):
        super().__init__(parent)
        self.result = None
        self.title("週報小幫手")
        self.configure(bg=BG, padx=28, pady=20)
        self.resizable(False, False)
        tk.Label(self, text="這些專案本週工作紀錄和待辦事項都沒有內容", font=FONT_B, bg=BG, fg=TEXT,
                 anchor="w").pack(fill="x")
        tk.Label(self, text="勾選的會顯示在週報（空白列），沒勾的不顯示", font=FONT_SM, bg=BG, fg=MUTED,
                 anchor="w").pack(fill="x", pady=(2, 10))
        self.vars = {}
        for p in projects:
            self.vars[p] = self._check(p)
        foot = tk.Frame(self, bg=BG)
        foot.pack(anchor="e", pady=(16, 0))
        Button(foot, "取消", self.destroy, kind="outline", font=FONT_SM, padx=12, pady=5).pack(side="left", padx=(0, 8))
        Button(foot, "確定", self._ok, kind="primary", font=FONT_B, padx=18, pady=5).pack(side="left")
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.grab_set()
        self.focus_set()
        self.wait_window()

    def _check(self, text):
        """跟待辦清單同款的勾選框，整列點了都會切換。預設勾（照舊顯示）。"""
        v = tk.BooleanVar(value=True)
        row = tk.Frame(self, bg=BG, cursor="hand2")
        row.pack(fill="x")
        box = tk.Canvas(row, width=18, height=18, bg=BG, highlightthickness=0)
        box.pack(side="left", padx=(0, 8), pady=5)
        lbl = tk.Label(row, text=text, font=FONT, bg=BG, fg=TEXT, anchor="w")
        lbl.pack(side="left", fill="x", expand=True)

        def draw(*_):
            box.delete("all")
            if v.get():
                box.create_rectangle(1, 1, 17, 17, fill=BLUE, outline=BLUE)
                box.create_line(4.5, 9, 7.5, 12.5, 13.5, 5.5, fill="white", width=2)
                lbl.configure(fg=TEXT)
            else:
                box.create_rectangle(1, 1, 17, 17, fill=BG, outline=CHECK_LINE, width=1.5)
                lbl.configure(fg=HINT_FG)
        v.trace_add("write", draw)
        draw()
        for w in (row, box, lbl):
            w.bind("<Button-1>", lambda e: v.set(not v.get()))
            w.bind("<Enter>", lambda e: [x.configure(bg=HOVER) for x in (row, box, lbl)])
            w.bind("<Leave>", lambda e: [x.configure(bg=BG) for x in (row, box, lbl)])
        return v

    def _ok(self):
        self.result = tuple(p for p, v in self.vars.items() if not v.get())
        self.destroy()


class SettingsDialog(tk.Toplevel):
    """個人設定：週報表頭、收件人、郵件內文（簽名）、專案清單、簡稱。存成 settings.json。"""

    def __init__(self, parent, first=False):
        super().__init__(parent)
        self.saved = None
        self.title("週報小幫手｜個人設定")
        self.configure(bg=BG, padx=28, pady=20)
        self.resizable(False, False)
        s = R.load_settings() or dict(R.DEFAULT_SETTINGS)
        if first:
            tk.Label(self, text="第一次使用，請先填好你的資料（之後可按右上角「⚙ 設定」修改）",
                     font=FONT_B, bg=BG, fg=TEXT).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        self.fields = {}
        rows = [("header", "週報表頭", "例：王小明 UI/UX設計師", 1),
                ("mail_to", "收件人", "多位用 ; 隔開", 1),
                ("mail_cc", "副本", "多位用 ; 隔開", 1),
                ("mail_body", "郵件內文", "含簽名檔，請改成你的姓名與 Email", 10),
                ("projects", "專案", "每行一個，依週報順序", 5),
                ("aliases", "簡稱", "每行一個「簡稱=全名」，例：北專=台北新案專案", 3)]
        value = {"projects": "\n".join(s["projects"]),
                 "aliases": "\n".join(f"{k}={v}" for k, v in s["aliases"].items())}
        self.mail_via = tk.StringVar(value=s.get("mail_via", "outlook"))
        r = 1  # row 0 是第一次使用的說明
        for key, label, hint, h in rows:
            if key == "mail_to":  # 寄送方式放在收件人上面
                tk.Label(self, text="寄送方式", font=FONT, bg=BG, fg=MUTED, anchor="nw", width=8).grid(
                    row=r, column=0, sticky="nw", pady=(6, 0))
                tips = {"outlook": "用 Outlook（傳統版）開郵件，週報 Word 自動附上",
                        "gmail": "用瀏覽器開 Gmail 草稿，週報 Word 要自己拖進信裡"}
                tip = tk.Label(self, font=FONT_SM, bg=BG, fg=BLUE)
                tip.grid(row=r + 1, column=1, sticky="w")
                Segmented(self, self.mail_via, [("outlook", "Outlook"), ("gmail", "Gmail")],
                          on_change=lambda v: tip.configure(text=tips.get(v, ""))).grid(
                    row=r, column=1, sticky="w", pady=(6, 0))
                r += 2
            tk.Label(self, text=label, font=FONT, bg=BG, fg=MUTED, anchor="nw", width=8).grid(
                row=r, column=0, sticky="nw", pady=(6, 0))
            w = tk.Text(self, width=56, height=h, font=FONT, relief="flat", bg=SIDE, fg=TEXT,
                        wrap="none" if h == 1 else "word",
                        insertbackground=TEXT, highlightthickness=1, highlightbackground=LINE,
                        highlightcolor=BLUE, padx=6, pady=4)
            w.insert("1.0", value.get(key, s.get(key, "")).rstrip("\n"))
            w.grid(row=r, column=1, sticky="w", pady=(6, 0))
            if h == 1:
                w.bind("<Return>", lambda e: "break")  # 單行欄位不換行
            tk.Label(self, text=hint, font=FONT_SM, bg=BG, fg=HINT_FG).grid(row=r + 1, column=1, sticky="w")
            self.fields[key] = w
            r += 2
        tk.Label(self, text="外觀", font=FONT, bg=BG, fg=MUTED, anchor="nw", width=8).grid(
            row=r, column=0, sticky="nw", pady=(10, 0))
        self.theme = tk.StringVar(value=s.get("theme", "light"))
        Segmented(self, self.theme, [("light", "淺色"), ("dark", "深色")]).grid(
            row=r, column=1, sticky="w", pady=(10, 0))
        foot = tk.Frame(self, bg=BG)
        foot.grid(row=99, column=0, columnspan=2, sticky="e", pady=(16, 0))
        Button(foot, "取消", self.destroy, kind="outline", font=FONT_SM, padx=12, pady=5).pack(side="left", padx=(0, 8))
        Button(foot, "儲存", self._save, kind="primary", font=FONT_B, padx=18, pady=5).pack(side="left")
        self.transient(parent if parent.winfo_viewable() else None)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")
        dark_titlebar(self)
        self.grab_set()
        self.fields["header"].focus_set()
        self.wait_window()

    def _get(self, key):
        return self.fields[key].get("1.0", "end").strip()

    def _save(self):
        if not self._get("header") or not self._get("mail_to"):
            messagebox.showwarning("週報小幫手", "「週報表頭」和「收件人」一定要填。", parent=self)
            return
        if "（中文姓名）" in self._get("mail_body") or "（你的 Email）" in self._get("mail_body"):
            if not messagebox.askyesno("週報小幫手", "郵件簽名還是範例文字，確定要這樣存嗎？", parent=self):
                return
        aliases = {}
        for line in self._get("aliases").splitlines():
            if "=" in line:
                k, v = (x.strip() for x in line.split("=", 1))
                if k and v:
                    aliases[k] = v
        self.saved = R.save_settings({
            "header": self._get("header"),
            "mail_via": self.mail_via.get(),
            "mail_to": self._get("mail_to"),
            "mail_cc": self._get("mail_cc"),
            "mail_body": self._get("mail_body") + "\n",
            "projects": [p.strip() for p in self._get("projects").splitlines() if p.strip()],
            "aliases": aliases,
            "theme": self.theme.get(),
        })
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
