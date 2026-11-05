"""
Tabela Paster - Windows program za lepljenje celija iz Excel tabele.

Hotkey-evi (rade samo dok je program upaljen, imaju prednost nad Windowsom):
    Ctrl+C            kopira CELIJU u kojoj se nalazis (bez lepljenja i pomeranja)
    Ctrl+V            nalepi CELIJU, predje na sledecu (desno) i odmah je kopira
    Ctrl+strelice     kretanje po tabeli (strelice bez Ctrl rade normalno)
    Ctrl+N            novi red, od prve kolone
    Ctrl+Shift+P      pauza / nastavak (privremeno vraca normalne Windows tastere)

Koristi samo Windows funkcije (RegisterHotKey / SendInput), bez dodatnih
biblioteka za tastaturu. Potrebno: openpyxl, pyperclip.
"""
import csv
import ctypes
import datetime
import queue
import threading
import time
import tkinter as tk
from ctypes import wintypes
from tkinter import ttk, filedialog, messagebox

import pyperclip
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import fills, fonts, borders, alignment, protection


# Zaobilazi gresku u nekim xlsx fajlovima:
# "__init__() got an unexpected keyword argument 'extLst'"
def _ignore_extlst(cls):
    orig = cls.__init__

    def init(self, *args, **kwargs):
        kwargs.pop("extLst", None)
        orig(self, *args, **kwargs)

    cls.__init__ = init


for _cls in (fills.PatternFill, fills.GradientFill, fonts.Font, borders.Border,
             borders.Side, alignment.Alignment, protection.Protection):
    _ignore_extlst(_cls)


# ======================= Windows: hotkey-evi i slanje tastera =======================
user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

ULONG_PTR = ctypes.c_size_t

WM_HOTKEY = 0x0312
WM_APP_CMD = 0x8000 + 1
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_NOREPEAT = 0x4000
VK_CONTROL = 0x11
VK_C = 0x43
VK_V = 0x56
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ULONG_PTR)]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ULONG_PTR)]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD)]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("mi", MOUSEINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
user32.RegisterHotKey.restype = wintypes.BOOL
user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
user32.UnregisterHotKey.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.PeekMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND,
                                wintypes.UINT, wintypes.UINT, wintypes.UINT]
user32.PeekMessageW.restype = wintypes.BOOL
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostThreadMessageW.restype = wintypes.BOOL
user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT
user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.MapVirtualKeyW.argtypes = [wintypes.UINT, wintypes.UINT]
user32.MapVirtualKeyW.restype = wintypes.UINT
user32.PostQuitMessage.argtypes = [ctypes.c_int]
kernel32.GetCurrentThreadId.restype = wintypes.DWORD

# id: (naziv, modifikatori, virtuelni taster)
HK = {
    1: ("ctrl+left", MOD_CONTROL, 0x25),
    2: ("ctrl+right", MOD_CONTROL, 0x27),
    3: ("ctrl+up", MOD_CONTROL, 0x26),
    4: ("ctrl+down", MOD_CONTROL, 0x28),
    5: ("ctrl+n", MOD_CONTROL | MOD_NOREPEAT, 0x4E),
    6: ("ctrl+v", MOD_CONTROL | MOD_NOREPEAT, VK_V),
    7: ("ctrl+c", MOD_CONTROL | MOD_NOREPEAT, VK_C),
    8: ("ctrl+shift+p", MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT, 0x50),
}
ID_PASTE, ID_COPY, ID_PAUSE = 6, 7, 8

CMD_REG, CMD_UNREG, CMD_INIT, CMD_REG_ALL, CMD_UNREG_ALL, CMD_QUIT = range(1, 7)


def send_keys(events):
    """events: lista (vk, up). Salje sve odjednom, bez ikakvog umetanja drugih tastera."""
    n = len(events)
    arr = (INPUT * n)()
    for i, (vk, up) in enumerate(events):
        arr[i].type = INPUT_KEYBOARD
        arr[i].u.ki = KEYBDINPUT(vk, user32.MapVirtualKeyW(vk, 0),
                                 KEYEVENTF_KEYUP if up else 0, 0, 0)
    user32.SendInput(n, arr, ctypes.sizeof(INPUT))


def ctrl_is_down():
    return bool(user32.GetAsyncKeyState(VK_CONTROL) & 0x8000)


class HotkeyThread(threading.Thread):
    """Poseban thread koji prima Windows hotkey poruke (WM_HOTKEY)."""

    def __init__(self, callback):
        super().__init__(daemon=True)
        self.callback = callback
        self.tid = None
        self.ready = threading.Event()
        self.done = threading.Event()
        self.active = set()
        self.failed = []

    def _reg(self, hid):
        if hid in self.active:
            return
        name, mods, vk = HK[hid]
        if user32.RegisterHotKey(None, hid, mods, vk):
            self.active.add(hid)
        elif name not in self.failed:
            self.failed.append(name)

    def _unreg(self, hid):
        if hid in self.active:
            user32.UnregisterHotKey(None, hid)
            self.active.discard(hid)

    def run(self):
        self.tid = kernel32.GetCurrentThreadId()
        msg = wintypes.MSG()
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)   # napravi red poruka
        self.ready.set()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == WM_HOTKEY:
                self.callback(int(msg.wParam))
            elif msg.message == WM_APP_CMD:
                cmd, hid = int(msg.wParam), int(msg.lParam)
                if cmd == CMD_REG:
                    self._reg(hid)
                elif cmd == CMD_UNREG:
                    self._unreg(hid)
                elif cmd == CMD_INIT:
                    for h in HK:
                        self._reg(h)
                elif cmd == CMD_REG_ALL:
                    for h in HK:
                        if h != ID_PAUSE:
                            self._reg(h)
                elif cmd == CMD_UNREG_ALL:
                    for h in HK:
                        if h != ID_PAUSE:
                            self._unreg(h)
                elif cmd == CMD_QUIT:
                    for h in list(self.active):
                        self._unreg(h)
                    user32.PostQuitMessage(0)
                self.done.set()
        for h in list(self.active):
            self._unreg(h)

    def command(self, cmd, hid=0):
        self.ready.wait()
        self.done.clear()
        user32.PostThreadMessageW(self.tid, WM_APP_CMD, cmd, hid)
        self.done.wait(2)


# ======================= Program =======================
ROW_H = 22
HDR_H = 22
RN_W = 44
CHAR_PX = 7
PASTE_WAIT = 0.15   # koliko se ceka posle lepljenja (povecaj ako program kasni sa citanjem clipboarda)

COL_CURSOR = "#bcdcff"      # blago plava - trenutna celija
COL_CURSOR_BORDER = "#2b7de9"
COL_HDR = "#eeeeee"
COL_HDR_ACTIVE = "#cfe3ff"
COL_PINNED = "#fff3c4"      # pinovan prvi red


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%d.%m.%Y")
        return v.strftime("%d.%m.%Y %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


class App:
    def __init__(self, root):
        self.root = root
        root.title("Tabela Paster")
        root.geometry("540x300+20+20")
        root.minsize(280, 140)
        root.attributes("-topmost", True)

        self.q = queue.Queue()
        self.data = []
        self.widths = []
        self.nrows = 0
        self.ncols = 0
        self.r = 0
        self.c = 0
        self.sr = 0
        self.sc = 0
        self.path = None
        self.paused = False
        self.colpos = []

        top = ttk.Frame(root)
        top.pack(fill="x", padx=2, pady=2)
        ttk.Button(top, text="Učitaj tabelu", command=self.load).pack(side="left")
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(top, textvariable=self.sheet_var,
                                     state="readonly", width=10)
        self.sheet_cb.pack(side="left", padx=4)
        self.sheet_cb.bind("<<ComboboxSelected>>",
                           lambda e: self.load_sheet(self.sheet_var.get()))
        self.pin_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Pinuj prvi red", variable=self.pin_var,
                        command=self.refresh).pack(side="left", padx=4)
        self.pause_btn = ttk.Button(top, text="Pauza", command=self.toggle_pause)
        self.pause_btn.pack(side="right")

        self.status = tk.Label(root, anchor="w", font=("Segoe UI", 9, "bold"),
                               text="Nema učitane tabele")
        self.status.pack(fill="x", padx=4)

        self.canvas = tk.Canvas(root, bg="white", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.refresh())
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<MouseWheel>", self.on_wheel)

        self.hk = HotkeyThread(self.on_hotkey)
        self.hk.start()
        self.hk.command(CMD_INIT)
        if self.hk.failed:
            root.after(400, lambda: messagebox.showwarning(
                "Zauzeti tasteri",
                "Ove prečice je već zauzeo neki drugi program i neće raditi:\n\n"
                + "\n".join(self.hk.failed)))

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(5, self.poll)

    # ---------- hotkey-evi ----------
    def on_hotkey(self, hid):
        """Poziva se iz hotkey thread-a; samo stavlja komandu u red."""
        q = self.q
        if hid == 1:
            q.put(("move", 0, -1))
        elif hid == 2:
            q.put(("move", 0, 1))
        elif hid == 3:
            q.put(("move", -1, 0))
        elif hid == 4:
            q.put(("move", 1, 0))
        elif hid == 5:
            q.put(("newrow",))
        elif hid == ID_PASTE:
            q.put(("paste",))
        elif hid == ID_COPY:
            q.put(("copy",))
        elif hid == ID_PAUSE:
            q.put(("pause",))

    def toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.hk.command(CMD_UNREG_ALL)
            self.pause_btn.config(text="Nastavi")
        else:
            self.hk.command(CMD_REG_ALL)
            self.pause_btn.config(text="Pauza")
        self.update_status()

    # ---------- kopiranje / lepljenje ----------
    def set_clipboard(self, val):
        """Postavi clipboard i proveri da je zaista tu. Vraca True samo ako jeste."""
        want = val.replace("\r\n", "\n")
        for _ in range(15):
            try:
                pyperclip.copy(val)
                if pyperclip.paste().replace("\r\n", "\n") == want:
                    return True
            except Exception:
                pass
            time.sleep(0.02)
        return False

    def inject_combo(self, hid, vk):
        """Salje pravi Ctrl+<taster> programu u fokusu. Nas hotkey se na trenutak
        skida da ga vestacki taster ne bi progutao."""
        self.hk.command(CMD_UNREG, hid)
        try:
            if ctrl_is_down():
                send_keys([(vk, False), (vk, True)])          # Ctrl je vec drzan
            else:
                send_keys([(VK_CONTROL, False), (vk, False),
                           (vk, True), (VK_CONTROL, True)])    # Ctrl je vec pusten
            time.sleep(PASTE_WAIT)      # daj programu vremena da procita clipboard
        finally:
            if not self.paused:
                self.hk.command(CMD_REG, hid)

    def copy_cell(self):
        """Uvek upisuje trenutnu celiju u clipboard i proverava da je tu."""
        return self.set_clipboard(self.data[self.r][self.c])

    def do_copy(self):
        # Bez ucitane tabele - obican Ctrl+C
        if not self.data:
            self.inject_combo(ID_COPY, VK_C)
            return
        if self.copy_cell():
            self.update_status(prefix="✓ Kopirano   ")
        else:
            self.status.config(text="Clipboard je zauzet - pokušaj ponovo", fg="#b00020")

    def do_paste(self):
        # Bez ucitane tabele ne diramo clipboard - obican Ctrl+V
        if not self.data:
            self.inject_combo(ID_PASTE, VK_V)
            return
        # 1) provera: u kojoj je celiji indikator (cita se tek sada, na pritisak)
        val = self.data[self.r][self.c]
        if val != "":
            # 2) kopiraj bas tu celiju (uvek upisuje, ne oslanja se na stari clipboard)
            if not self.copy_cell():
                self.status.config(text="Clipboard je zauzet - pokušaj ponovo", fg="#b00020")
                return                      # ne lepimo nista pogresno
            # 3) nalepi
            self.inject_combo(ID_PASTE, VK_V)
        # 4) tek onda na sledecu celiju (desno); na kraju reda ostaje
        if self.c < self.ncols - 1:
            self.c += 1
        self.refresh()
        # unapred kopiraj novu celiju, da je sledeci Ctrl+V vec spreman
        if self.data[self.r][self.c] != "":
            self.copy_cell()

    # ---------- učitavanje ----------
    def load(self):
        path = filedialog.askopenfilename(
            title="Izaberi tabelu",
            filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.csv"), ("Svi fajlovi", "*.*")])
        if not path:
            return
        self.path = path
        try:
            if path.lower().endswith(".csv"):
                self.sheet_cb["values"] = []
                self.sheet_var.set("")
                self.load_csv(path)
            else:
                wb = load_workbook(path, read_only=True, data_only=True)
                names = wb.sheetnames
                wb.close()
                self.sheet_cb["values"] = names
                self.sheet_var.set(names[0])
                self.load_sheet(names[0])
        except Exception as e:
            messagebox.showerror("Greška", f"Ne mogu da učitam fajl:\n{e}")

    def load_csv(self, path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            rows = [list(r) for r in csv.reader(f, dialect)]
        self.set_data(rows)

    def load_sheet(self, name):
        try:
            wb = load_workbook(self.path, read_only=True, data_only=True)
            ws = wb[name]
            rows = [[fmt(v) for v in row] for row in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception as e:
            messagebox.showerror("Greška", f"Ne mogu da učitam list:\n{e}")
            return
        self.set_data(rows)

    def set_data(self, rows):
        while rows and not any(str(x).strip() for x in rows[-1]):
            rows.pop()
        ncols = max((len(r) for r in rows), default=0)
        rows = [[str(x) for x in r] + [""] * (ncols - len(r)) for r in rows]
        self.data = rows
        self.nrows = len(rows)
        self.ncols = ncols
        self.widths = []
        for j in range(ncols):
            longest = len(get_column_letter(j + 1))
            for i in range(min(self.nrows, 300)):
                longest = max(longest, len(rows[i][j]))
            self.widths.append(max(4, min(longest, 30)) * CHAR_PX + 12)
        self.r = self.c = self.sr = self.sc = 0
        self.refresh()

    # ---------- kretanje ----------
    def poll(self):
        # komande se obradjuju striktno redom, tako da Ctrl+V uvek
        # vidi tacnu, azurnu poziciju
        try:
            while True:
                cmd = self.q.get_nowait()
                if cmd[0] == "move":
                    self.move(cmd[1], cmd[2])
                elif cmd[0] == "newrow":
                    self.new_row()
                elif cmd[0] == "paste":
                    self.do_paste()
                elif cmd[0] == "copy":
                    self.do_copy()
                elif cmd[0] == "pause":
                    self.toggle_pause()
        except queue.Empty:
            pass
        self.root.after(5, self.poll)

    def move(self, dr, dc):
        if not self.data:
            return
        self.r = max(0, min(self.nrows - 1, self.r + dr))
        self.c = max(0, min(self.ncols - 1, self.c + dc))
        self.refresh()

    def new_row(self):
        if not self.data:
            return
        self.c = 0
        self.r = min(self.nrows - 1, self.r + 1)
        self.refresh()

    def min_row(self):
        return 1 if self.pin_var.get() else 0

    def top_h(self):
        return HDR_H + (ROW_H if self.pin_var.get() else 0)

    def on_wheel(self, e):
        if not self.data:
            return
        step = -3 if e.delta > 0 else 3
        self.sr = max(self.min_row(), min(self.nrows - 1, self.sr + step))
        self.draw()

    def on_click(self, e):
        if not self.data or e.y < HDR_H:
            return
        if self.pin_var.get() and e.y < HDR_H + ROW_H:
            i = 0
        else:
            i = max(self.sr, self.min_row()) + (e.y - self.top_h()) // ROW_H
        if i >= self.nrows:
            return
        for j, x, w in self.colpos:
            if x <= e.x < x + w:
                self.r, self.c = i, j
                self.refresh()
                return

    # ---------- prikaz ----------
    def ensure_visible(self):
        h = self.canvas.winfo_height()
        w = self.canvas.winfo_width()
        minr = self.min_row()
        vis = max(1, (h - self.top_h()) // ROW_H)
        if self.sr < minr:
            self.sr = minr
        if self.r >= minr:
            if self.r < self.sr:
                self.sr = self.r
            elif self.r >= self.sr + vis:
                self.sr = self.r - vis + 1
        if self.c < self.sc:
            self.sc = self.c
        else:
            while self.sc < self.c and sum(self.widths[self.sc:self.c + 1]) > w - RN_W:
                self.sc += 1

    def update_status(self, prefix=""):
        if self.paused:
            self.status.config(text="PAUZIRANO - normalni Windows tasteri", fg="#b00020")
            return
        if not self.data:
            self.status.config(text="Nema učitane tabele", fg="black")
            return
        val = self.data[self.r][self.c].replace("\n", " ")
        if len(val) > 40:
            val = val[:40] + "…"
        name = ""
        if self.pin_var.get() and self.r > 0:
            hn = self.data[0][self.c].strip()
            if hn:
                name = f" ({hn[:25]})"
        self.status.config(
            text=f"{prefix}Kolona {get_column_letter(self.c + 1)}{name}, red {self.r + 1}   |   {val}",
            fg="black")

    def refresh(self):
        if self.data:
            self.ensure_visible()
        self.update_status()
        self.draw()

    def draw_row(self, i, y, pinned_style=False):
        cv = self.canvas
        bg = COL_HDR_ACTIVE if i == self.r else COL_HDR
        cv.create_rectangle(0, y, RN_W, y + ROW_H, fill=bg, outline="#bbb")
        cv.create_text(RN_W / 2, y + ROW_H / 2, text=str(i + 1),
                       font=("Segoe UI", 9, "bold"))
        for j, x, cw in self.colpos:
            cur = (i == self.r and j == self.c)
            if cur:
                fill = COL_CURSOR
            elif pinned_style:
                fill = COL_PINNED
            else:
                fill = "white"
            cv.create_rectangle(x, y, x + cw, y + ROW_H, fill=fill,
                                outline=COL_CURSOR_BORDER if cur else "#ddd",
                                width=2 if cur else 1)
            txt = self.data[i][j].replace("\n", " ")
            n = max(1, (cw - 8) // CHAR_PX)
            if len(txt) > n:
                txt = txt[:n - 1] + "…"
            cv.create_text(x + 4, y + ROW_H / 2, text=txt, anchor="w",
                           font=("Segoe UI", 9, "bold" if pinned_style else "normal"))

    def draw(self):
        cv = self.canvas
        cv.delete("all")
        w, h = cv.winfo_width(), cv.winfo_height()
        if not self.data:
            cv.create_text(w / 2, h / 2, text="Klikni „Učitaj tabelu“", fill="#888")
            return

        self.colpos = []
        x = RN_W
        j = self.sc
        while j < self.ncols and x < w:
            self.colpos.append((j, x, self.widths[j]))
            x += self.widths[j]
            j += 1

        # zaglavlje kolona (slova)
        cv.create_rectangle(0, 0, RN_W, HDR_H, fill=COL_HDR, outline="#bbb")
        for j, x, cw in self.colpos:
            bg = COL_HDR_ACTIVE if j == self.c else COL_HDR
            cv.create_rectangle(x, 0, x + cw, HDR_H, fill=bg, outline="#bbb")
            cv.create_text(x + cw / 2, HDR_H / 2, text=get_column_letter(j + 1),
                           font=("Segoe UI", 9, "bold"))

        y = HDR_H
        pinned = self.pin_var.get()
        if pinned:
            self.draw_row(0, y, pinned_style=True)
            y += ROW_H
        i = max(self.sr, self.min_row())
        while i < self.nrows and y < h:
            self.draw_row(i, y)
            y += ROW_H
            i += 1

    def on_close(self):
        self.hk.command(CMD_QUIT)   # oslobadja sve prečice, Windows se vraca na normalno
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
