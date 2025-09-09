import json, time, random, threading, ctypes
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import List, Literal, Tuple, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pynput import keyboard, mouse
import ctypes.wintypes

SW_RESTORE = 9
record_hwnd: Optional[int] = None


user32 = ctypes.windll.user32
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
def get_cursor_pos() -> Tuple[int,int]:
    pt = POINT(); user32.GetCursorPos(ctypes.byref(pt)); return (pt.x, pt.y)
def set_cursor_pos(x:int,y:int): user32.SetCursorPos(int(x),int(y))

@dataclass
class MacroEvent:
    kind: Literal["key","mouse"]
    action: str
    data: dict
    delay_ms: int

@dataclass
@dataclass
class Macro:
    events: List[MacroEvent]
    target_hwnd: Optional[int] = None  # window handle recorded in

    def save(self, path: Path):
        data = {"events": [asdict(e) for e in self.events], "target_hwnd": self.target_hwnd}
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    @staticmethod
    def load(path: Path) -> "Macro":
        raw = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            events = [MacroEvent(**e) for e in raw["events"]]
            return Macro(events, raw.get("target_hwnd"))
        else:  # backward-compat for old files that were just a list
            return Macro([MacroEvent(**e) for e in raw], None)


kb = keyboard.Controller()
ms = mouse.Controller()

APP_DIR = Path.home() / "AppData" / "Roaming" / "MacroClicker"
APP_DIR.mkdir(parents=True, exist_ok=True)
CFG_PATH = APP_DIR / "settings.json"
DEFAULTS = {
    "hk_rec_start": "<f9>",
    "hk_rec_stop":  "<f10>",
    "hk_play":      "<f8>",
    "hk_click_toggle": "<f6>",
    "click_mode": "cps",
    "cps": 10.0,
    "interval_ms": 100,
    "button": "left",
    "mode": "current",
    "fixed_xy": (0,0),
    "jitter": 0.0
}
def load_cfg():
    try:
        if CFG_PATH.exists():
            data = json.loads(CFG_PATH.read_text(encoding="utf-8"))
            return {**DEFAULTS, **data}
    except: pass
    return DEFAULTS.copy()
def save_cfg(d):
    try: CFG_PATH.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except: pass

settings = load_cfg()

def key_to_str(k)->str:
    try:
        if isinstance(k, keyboard.KeyCode):
            return k.char if k.char else str(k.vk)
        return str(k).split(".")[-1]
    except: return str(k)
def button_to_str(b)->str: return str(b).split(".")[-1]
def str_to_button(s:str): return getattr(mouse.Button, s)
def extract_simple_key_names(hk: str)->List[str]:
    names=[]; tmp=hk.lower().replace(" ","")
    for part in tmp.split("+"):
        p=part.strip("<>"); 
        if p: names.append(p)
    return names

CONTROL_KEYS = set(sum([extract_simple_key_names(settings[k])
                        for k in ("hk_rec_start","hk_rec_stop","hk_play","hk_click_toggle")], []))

recording = False
playing = False
autoclicking = False
_rec_prev_t: float = 0.0
_rec_buf: List[MacroEvent] = []
_rec_lock = threading.Lock()
player_thread: Optional[threading.Thread] = None
clicker_thread: Optional[threading.Thread] = None
gh_listener: Optional[keyboard.GlobalHotKeys] = None
gh_lock = threading.Lock()

def start_record():
    global recording, _rec_prev_t, _rec_buf, record_hwnd
    if recording or playing or autoclicking:
        return
    record_hwnd = int(user32.GetForegroundWindow())
    with _rec_lock:
        _rec_buf = []
    _rec_prev_t = time.time()
    recording = True


def stop_record() -> Macro:
    global recording
    recording = False
    with _rec_lock:
        return Macro(list(_rec_buf), target_hwnd=record_hwnd)


def on_kb_event(key, pressed: bool):
    if not recording or playing or autoclicking: return
    name = key_to_str(key).lower()
    if name in CONTROL_KEYS or name in ("ctrl","alt","shift","cmd","windows","alt_gr"):
        return
    now = time.time()
    global _rec_prev_t
    delay_ms = int((now - _rec_prev_t)*1000)
    _rec_prev_t = now
    ev = MacroEvent(kind="key", action="press" if pressed else "release",
                    data={"key": name}, delay_ms=delay_ms)
    with _rec_lock: _rec_buf.append(ev)

def on_mouse_move(x, y): pass

def on_mouse_click(x, y, button, pressed):
    if not recording or playing or autoclicking: return
    now = time.time(); global _rec_prev_t
    delay_ms = int((now - _rec_prev_t)*1000); _rec_prev_t = now
    with _rec_lock:
        _rec_buf.append(MacroEvent("mouse","click",
                                   {"x":int(x),"y":int(y),
                                    "button": button_to_str(button),
                                    "pressed": bool(pressed)}, delay_ms))

def playback_macro(m: Macro, speed: float = 1.0, loop: int = 1):
    global playing
    if playing or not m.events:
        return

    # focus the recorded window if known
    try:
        if m.target_hwnd:
            user32.ShowWindow(ctypes.wintypes.HWND(m.target_hwnd), SW_RESTORE)
            user32.SetForegroundWindow(ctypes.wintypes.HWND(m.target_hwnd))
            time.sleep(0.15)
    except Exception:
        pass

    playing = True
    try:
        for _ in range(max(1, loop)):
            for e in m.events:
                if not playing:
                    return
                wait = max(0, (e.delay_ms / 1000.0) / max(0.01, speed))
                if wait:
                    time.sleep(wait)
                if e.kind == "key":
                    k = e.data["key"]
                    key_obj = getattr(keyboard.Key, k, None)
                    if key_obj is None:
                        key_obj = keyboard.KeyCode.from_char(k[:1])
                    (kb.press if e.action == "press" else kb.release)(key_obj)
                else:
                    if e.action == "move":
                        ms.position = (e.data["x"], e.data["y"])
                    elif e.action == "click":
                        btn = getattr(mouse.Button, e.data["button"])
                        (ms.press if e.data["pressed"] else ms.release)(btn)
    finally:
        playing = False


def click_loop(get_cfg):
    global autoclicking
    while autoclicking:
        cfg = get_cfg()
        if cfg["click_mode"] == "ms":
            delay = max(1, int(cfg["interval_ms"])) / 1000.0
        else:
            delay = 1.0 / max(0.1, float(cfg["cps"]))
        btn = str_to_button(cfg["button"])
        mode = cfg["mode"]; fx, fy = cfg["fixed_xy"]; jitter = float(cfg["jitter"])
        if mode == "fixed":
            tx, ty = fx, fy
        else:
            tx, ty = get_cursor_pos()
        if jitter > 0:
            tx += int(random.uniform(-jitter, jitter))
            ty += int(random.uniform(-jitter, jitter))
        if mode == "fixed":
            set_cursor_pos(tx, ty)
        ms.press(btn); ms.release(btn)
        time.sleep(delay)

def rebuild_hotkeys(app):
    global gh_listener
    with gh_lock:
        if gh_listener:
            try: gh_listener.stop()
            except: pass
            gh_listener = None
        def hk_rec_start():
            start_record(); app.set_status("Recording…" if recording else "Busy")
        def hk_rec_stop():
            if recording:
                m = stop_record(); app.load_macro_into_table(m); app.set_status(f"Recorded {len(m.events)} events.")
        def hk_play():
            if app.current_macro and not playing: app.play_macro()
        def hk_click_toggle():
            app.toggle_clicker()
        mapping = {
            settings["hk_rec_start"]: hk_rec_start,
            settings["hk_rec_stop"]:  hk_rec_stop,
            settings["hk_play"]:      hk_play,
            settings["hk_click_toggle"]: hk_click_toggle
        }
        gh_listener = keyboard.GlobalHotKeys(mapping)
        threading.Thread(target=gh_listener.run, daemon=True).start()
        

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Macro + Auto Clicker")
        ico = Path("Untitled-1.ico")
        if ico.exists():
            try: self.iconbitmap(ico.resolve())
            except: pass
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_macro = ttk.Frame(nb); nb.add(self.tab_macro, text="Macro")
        self.tab_click = ttk.Frame(nb); nb.add(self.tab_click, text="Auto Clicker")

        self.current_macro: Optional[Macro] = None

        f_hk = ttk.LabelFrame(self.tab_macro, text="Global Hotkeys")
        f_hk.pack(fill="x", padx=6, pady=6)
        self.var_hk_start = tk.StringVar(value=settings["hk_rec_start"])
        self.var_hk_stop  = tk.StringVar(value=settings["hk_rec_stop"])
        self.var_hk_play  = tk.StringVar(value=settings["hk_play"])
        r=0
        ttk.Label(f_hk, text="Start Record:").grid(row=r, column=0, sticky="w"); 
        ttk.Entry(f_hk, textvariable=self.var_hk_start, width=18).grid(row=r, column=1, padx=6); r+=1
        ttk.Label(f_hk, text="Stop Record:").grid(row=r, column=0, sticky="w"); 
        ttk.Entry(f_hk, textvariable=self.var_hk_stop,  width=18).grid(row=r, column=1, padx=6); r+=1
        ttk.Label(f_hk, text="Play Macro:").grid(row=r, column=0, sticky="w"); 
        ttk.Entry(f_hk, textvariable=self.var_hk_play,  width=18).grid(row=r, column=1, padx=6); r+=1
        ttk.Button(f_hk, text="Apply Hotkeys", command=self.apply_macro_hotkeys).grid(row=r, column=0, columnspan=2, pady=4)

        f_tbl = ttk.LabelFrame(self.tab_macro, text="Recorded Events")
        f_tbl.pack(fill="both", expand=True, padx=6, pady=6)
        cols = ("#", "type", "action", "detail", "delay_ms")
        self.tree = ttk.Treeview(f_tbl, columns=cols, show="headings", height=12)
        for c,w in zip(cols,(50,70,80,360,90)):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(f_tbl, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set); sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._edit_delay_cell)

        hint = ttk.Label(self.tab_macro, text="Tip: double-click the delay_ms cell to edit.", foreground="#6a6a6a")
        hint.pack(anchor="e", padx=8, pady=(0, 4))

        f_ed = ttk.Frame(self.tab_macro); f_ed.pack(fill="x", padx=6, pady=(0,6))
        self.sel_idx = None
        ttk.Label(f_ed, text="Selected row:").grid(row=0, column=0, sticky="w")
        self.lbl_sel = ttk.Label(f_ed, text="-"); self.lbl_sel.grid(row=0, column=1, sticky="w", padx=(4,12))
        ttk.Label(f_ed, text="Kind:").grid(row=0, column=2, sticky="e")
        self.var_kind = tk.StringVar(value="key")
        ttk.OptionMenu(f_ed, self.var_kind, "key", "key", "mouse").grid(row=0, column=3, sticky="w")
        ttk.Label(f_ed, text="Action:").grid(row=0, column=4, sticky="e")
        self.var_action = tk.StringVar(value="press")
        ttk.OptionMenu(f_ed, self.var_action, "press", "press", "release", "move", "click").grid(row=0, column=5, sticky="w")
        ttk.Label(f_ed, text="Key/Button:").grid(row=1, column=0, sticky="e")
        self.var_kb = tk.StringVar(value="a")
        ttk.Entry(f_ed, textvariable=self.var_kb, width=10).grid(row=1, column=1, sticky="w", padx=(4,12))
        ttk.Label(f_ed, text="X:").grid(row=1, column=2, sticky="e")
        self.var_x = tk.IntVar(value=0)
        ttk.Entry(f_ed, textvariable=self.var_x, width=8).grid(row=1, column=3, sticky="w")
        ttk.Label(f_ed, text="Y:").grid(row=1, column=4, sticky="e")
        self.var_y = tk.IntVar(value=0)
        ttk.Entry(f_ed, textvariable=self.var_y, width=8).grid(row=1, column=5, sticky="w")
        ttk.Label(f_ed, text="Pressed (mouse click):").grid(row=2, column=0, sticky="e")
        self.var_pressed = tk.BooleanVar(value=True)
        ttk.Checkbutton(f_ed, variable=self.var_pressed).grid(row=2, column=1, sticky="w", padx=(4,12))
        ttk.Label(f_ed, text="Delay before (ms):").grid(row=2, column=2, sticky="e")
        self.var_delay = tk.IntVar(value=0)
        ttk.Entry(f_ed, textvariable=self.var_delay, width=10).grid(row=2, column=3, sticky="w")
        f_ed_btn = ttk.Frame(self.tab_macro); f_ed_btn.pack(fill="x", padx=6, pady=(0,6))
        ttk.Button(f_ed_btn, text="Load Selected", command=self.load_selected).pack(side="left", padx=4)
        ttk.Button(f_ed_btn, text="Apply to Selected", command=self.apply_to_selected).pack(side="left", padx=4)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_selected_label())

        row_btn = ttk.Frame(self.tab_macro); row_btn.pack(fill="x", padx=6, pady=6)
        ttk.Button(row_btn, text="Start Record", command=start_record).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Stop Record",  command=self.stop_record_btn).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Play", command=self.play_macro).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Delete Selected", command=self.delete_selected).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Move Up", command=lambda: self.move_selected(-1)).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Move Down", command=lambda: self.move_selected(1)).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Save…", command=self.save_macro).pack(side="left", padx=4)
        ttk.Button(row_btn, text="Load…", command=self.load_macro).pack(side="left", padx=4)

        f_rate = ttk.LabelFrame(self.tab_click, text="Click Rate")
        f_rate.pack(fill="x", padx=6, pady=6)
        self.var_mode_rate = tk.StringVar(value=settings["click_mode"])
        self.var_cps = tk.DoubleVar(value=float(settings["cps"]))
        self.var_ms  = tk.IntVar(value=int(settings["interval_ms"]))
        ttk.Radiobutton(f_rate, text="CPS", variable=self.var_mode_rate, value="cps").grid(row=0, column=0, sticky="w")
        ttk.Entry(f_rate, textvariable=self.var_cps, width=10).grid(row=0, column=1, padx=6)
        ttk.Radiobutton(f_rate, text="Interval (ms)", variable=self.var_mode_rate, value="ms").grid(row=1, column=0, sticky="w")
        ttk.Entry(f_rate, textvariable=self.var_ms, width=10).grid(row=1, column=1, padx=6)

        f_click = ttk.LabelFrame(self.tab_click, text="Click Settings")
        f_click.pack(fill="x", padx=6, pady=6)
        self.var_button = tk.StringVar(value=settings["button"])
        self.var_where  = tk.StringVar(value=settings["mode"])
        self.var_fx = tk.IntVar(value=int(settings["fixed_xy"][0]))
        self.var_fy = tk.IntVar(value=int(settings["fixed_xy"][1]))
        self.var_jitter = tk.DoubleVar(value=float(settings["jitter"]))
        ttk.Label(f_click, text="Button:").grid(row=0, column=0, sticky="w")
        ttk.OptionMenu(f_click, self.var_button, self.var_button.get(), "left","right","middle").grid(row=0, column=1, sticky="w")
        ttk.Label(f_click, text="Mode:").grid(row=1, column=0, sticky="w")
        ttk.OptionMenu(f_click, self.var_where, self.var_where.get(), "current","fixed").grid(row=1, column=1, sticky="w")
        ttk.Label(f_click, text="Fixed X:").grid(row=2, column=0, sticky="w"); ttk.Entry(f_click, textvariable=self.var_fx, width=10).grid(row=2, column=1)
        ttk.Label(f_click, text="Fixed Y:").grid(row=3, column=0, sticky="w"); ttk.Entry(f_click, textvariable=self.var_fy, width=10).grid(row=3, column=1)
        ttk.Label(f_click, text="Jitter (px):").grid(row=4, column=0, sticky="w"); ttk.Entry(f_click, textvariable=self.var_jitter, width=10).grid(row=4, column=1)

        f_hkc = ttk.LabelFrame(self.tab_click, text="Global Toggle")
        f_hkc.pack(fill="x", padx=6, pady=6)
        self.var_hk_click = tk.StringVar(value=settings["hk_click_toggle"])
        ttk.Label(f_hkc, text="Hotkey:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f_hkc, textvariable=self.var_hk_click, width=18).grid(row=0, column=1, padx=6)
        ttk.Button(f_hkc, text="Apply", command=self.apply_click_hotkey).grid(row=0, column=2, padx=6)
        rowc = ttk.Frame(self.tab_click); rowc.pack(fill="x", padx=6, pady=6)
        ttk.Button(rowc, text="Start Clicker", command=lambda: self.toggle_clicker(True)).pack(side="left", padx=4)
        ttk.Button(rowc, text="Stop Clicker", command=lambda: self.toggle_clicker(False)).pack(side="left", padx=4)

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status).pack(anchor="w", padx=8, pady=(0,8))

        rebuild_hotkeys(self)
        self.k_listener = keyboard.Listener(on_press=lambda k: on_kb_event(k, True),
                                            on_release=lambda k: on_kb_event(k, False))
        self.k_listener.start()
        self.m_listener = mouse.Listener(on_move=on_mouse_move, on_click=on_mouse_click)
        self.m_listener.start()

    def set_status(self, s:str): self.status.set(s)

    def _macro_to_rows(self, m: Macro):
        rows=[]
        for i,e in enumerate(m.events, start=1):
            if e.kind=="key":
                detail=f"key={e.data.get('key','')}"
            else:
                if e.action=="click":
                    detail=f"{e.data.get('button','')} @({e.data.get('x','')},{e.data.get('y','')}) pressed={e.data.get('pressed')}"
                else:
                    detail=f"move @({e.data.get('x','')},{e.data.get('y','')})"
            rows.append((str(i), e.kind, e.action, detail, str(e.delay_ms)))
        return rows

    def load_macro_into_table(self, m: Macro):
        self.current_macro = m
        for i in self.tree.get_children(): self.tree.delete(i)
        for row in self._macro_to_rows(m): self.tree.insert("", "end", values=row)

    def _table_to_macro(self) -> Macro:
        evs=[]
        for item in self.tree.get_children():
            _, kind, action, detail, dly = self.tree.item(item, "values")
            delay_ms = int(float(dly)) if dly!="" else 0
            if kind=="key":
                key = detail.split("key=",1)[1] if "key=" in detail else detail
                evs.append(MacroEvent("key", action, {"key": key}, delay_ms))
            else:
                btn="left"; x=0; y=0; pressed=True
                if "@(" in detail:
                    try:
                        pre, rest = detail.split("@(",1)
                        if "pressed=" in rest:
                            coord, pressed_part = rest.split(") pressed=",1)
                            pressed = (pressed_part.strip().lower()=="true")
                        else:
                            coord = rest.split(")")[0]
                        x,y = map(int, coord.split(","))
                        if any(b in pre for b in ("left","right","middle")):
                            btn = pre.strip().split()[0]
                    except: pass
                evs.append(MacroEvent("mouse", action, {"button":btn,"x":x,"y":y,"pressed":pressed}, delay_ms))
        return Macro(evs)

    def delete_selected(self):
        sel=self.tree.selection()
        if not sel: return
        self.tree.delete(sel[0])
        for i,it in enumerate(self.tree.get_children(), start=1):
            vals=list(self.tree.item(it,"values")); vals[0]=str(i); self.tree.item(it, values=vals)

    def move_selected(self, delta:int):
        sel=self.tree.selection()
        if not sel: return
        it=sel[0]; idx=self.tree.index(it)
        new=max(0, min(idx+delta, len(self.tree.get_children())-1))
        if new==idx: return
        vals=self.tree.item(it,"values"); self.tree.delete(it)
        new_it = self.tree.insert("", new, values=vals)
        for i,it2 in enumerate(self.tree.get_children(), start=1):
            vals2=list(self.tree.item(it2,"values")); vals2[0]=str(i); self.tree.item(it2, values=vals2)
        self.tree.selection_set(new_it)

    def stop_record_btn(self):
        if recording:
            m = stop_record()
            self.load_macro_into_table(m)
            self.set_status(f"Recorded {len(m.events)} events.")

    def play_macro(self):
        global player_thread
        m = self._table_to_macro()
        self.current_macro = m
        if player_thread and player_thread.is_alive():
            self.set_status("Already playing."); return
        self.set_status("Playing…")
        def run():
            try: playback_macro(m, speed=1.0, loop=1)
            finally: self.set_status("Playback finished.")
        player_thread = threading.Thread(target=run, daemon=True); player_thread.start()

    def save_macro(self):
        m = self._table_to_macro()
        p = filedialog.asksaveasfilename(defaultextension=".json",
                                         filetypes=[("Macro JSON","*.json")],
                                         initialdir=str(APP_DIR), initialfile="macro.json")
        if p: m.save(Path(p)); self.set_status(f"Saved: {p}")

    def load_macro(self):
        p = filedialog.askopenfilename(filetypes=[("Macro JSON","*.json")], initialdir=str(APP_DIR))
        if p:
            try: self.load_macro_into_table(Macro.load(Path(p))); self.set_status(f"Loaded: {p}")
            except Exception as e: messagebox.showerror("Load Macro", f"Failed: {e}")

    def apply_macro_hotkeys(self):
        settings["hk_rec_start"]=self.var_hk_start.get().strip() or DEFAULTS["hk_rec_start"]
        settings["hk_rec_stop"] =self.var_hk_stop.get().strip()  or DEFAULTS["hk_rec_stop"]
        settings["hk_play"]     =self.var_hk_play.get().strip()  or DEFAULTS["hk_play"]
        save_cfg(settings)
        CONTROL_KEYS.clear()
        CONTROL_KEYS.update(sum([extract_simple_key_names(settings[k]) 
                                 for k in ("hk_rec_start","hk_rec_stop","hk_play","hk_click_toggle")], []))
        rebuild_hotkeys(self)
        self.set_status("Hotkeys applied.")

    def apply_click_hotkey(self):
        settings["hk_click_toggle"]=self.var_hk_click.get().strip() or DEFAULTS["hk_click_toggle"]
        save_cfg(settings); rebuild_hotkeys(self); self.set_status("Clicker hotkey applied.")

    def current_click_cfg(self):
        settings["click_mode"]=self.var_mode_rate.get()
        settings["cps"]=float(self.var_cps.get())
        settings["interval_ms"]=int(self.var_ms.get())
        settings["button"]=self.var_button.get()
        settings["mode"]=self.var_where.get()
        settings["fixed_xy"]=(int(self.var_fx.get()), int(self.var_fy.get()))
        settings["jitter"]=float(self.var_jitter.get())
        save_cfg(settings); return settings

    def toggle_clicker(self, want: Optional[bool]=None):
        global autoclicking, clicker_thread
        target = not autoclicking if want is None else want
        if target and not autoclicking:
            autoclicking=True; self.set_status("Auto clicker: ON")
            clicker_thread=threading.Thread(target=click_loop, args=(self.current_click_cfg,), daemon=True)
            clicker_thread.start()
        elif not target and autoclicking:
            autoclicking=False; self.set_status("Auto clicker: OFF")

    def _edit_delay_cell(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        colid = self.tree.identify_column(event.x)
        if colid != "#5": return
        rowid = self.tree.identify_row(event.y)
        if not rowid: return
        x, y, w, h = self.tree.bbox(rowid, colid)
        vals = list(self.tree.item(rowid, "values")); current = vals[4]
        e = tk.Entry(self.tree); e.insert(0, current); e.select_range(0, tk.END); e.focus()
        e.place(x=x, y=y, width=w, height=h)
        def commit(*_):
            try: new_ms = str(int(float(e.get())))
            except: new_ms = current
            vals[4] = new_ms; self.tree.item(rowid, values=vals); e.destroy()
        e.bind("<Return>", commit); e.bind("<FocusOut>", commit)

    def _update_selected_label(self):
        sel = self.tree.selection()
        self.sel_idx = self.tree.index(sel[0]) if sel else None
        self.lbl_sel.config(text=str(self.sel_idx+1) if self.sel_idx is not None else "-")

    def load_selected(self):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0]); self.sel_idx = idx
        self.lbl_sel.config(text=str(idx+1))
        _, kind, action, detail, dly = self.tree.item(sel[0], "values")
        self.var_kind.set(kind); self.var_action.set(action); self.var_delay.set(int(float(dly or 0)))
        if kind == "key":
            k = detail.split("key=",1)[1] if "key=" in detail else detail
            self.var_kb.set(k)
            self.var_x.set(0); self.var_y.set(0); self.var_pressed.set(True)
        else:
            btn="left"; x=0; y=0; pressed=True
            if "@(" in detail:
                try:
                    pre, rest = detail.split("@(",1)
                    if "pressed=" in rest:
                        coord, p2 = rest.split(") pressed=",1); pressed = (p2.strip().lower()=="true")
                    else:
                        coord = rest.split(")")[0]
                    x,y = map(int, coord.split(","))
                    if any(b in pre for b in ("left","right","middle")): btn = pre.strip().split()[0]
                except: pass
            self.var_kb.set(btn); self.var_x.set(x); self.var_y.set(y); self.var_pressed.set(pressed)

    def apply_to_selected(self):
        if self.sel_idx is None: return
        items = self.tree.get_children()
        it = items[self.sel_idx]
        kind = self.var_kind.get()
        action = self.var_action.get()
        dly = str(int(self.var_delay.get()))
        if kind == "key":
            detail = f"key={self.var_kb.get().strip() or 'a'}"
        else:
            detail = f"{self.var_kb.get().strip() or 'left'} @({int(self.var_x.get())},{int(self.var_y.get())}) pressed={bool(self.var_pressed.get())}"
        self.tree.item(it, values=(str(self.sel_idx+1), kind, action, detail, dly))

    def on_close(self):
        try: self.toggle_clicker(False)
        except: pass
        global recording, playing
        playing=False
        if recording:
            try: stop_record()
            except: pass
        try:
            if hasattr(self,"k_listener"): self.k_listener.stop()
            if hasattr(self,"m_listener"): self.m_listener.stop()
        except: pass
        try: rebuild_hotkeys(self)
        except: pass
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.geometry("900x720")
    app.mainloop()
