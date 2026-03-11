"""
Twitch Auto-Mod Tool v2.1
新機能:
  - 段階的ペナルティ（全違反共通カウント・段階数/秒数/BAN手動設定）
  - 連投速度検知（N秒以内にN回）
  - NGワードも段階的ペナルティに統合
  - 違反カウントのクールダウンリセット（発言禁止後N日）
  - 自動再接続（30秒ごと最大3回）
  - チャットリアルタイム表示
  - 履歴保持件数UI設定
  - 各機能ON/OFF
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading, socket, ssl, re, csv, os, json, time, subprocess
from datetime import datetime
from difflib import SequenceMatcher
from collections import deque
try:
    from ai_engine import get_engine
    HAS_AI = True
except ImportError:
    HAS_AI = False

# ─────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────
AEGISMOD_DIR = os.path.join(os.path.expanduser("~"), ".aegismod")
os.makedirs(AEGISMOD_DIR, exist_ok=True)
CONFIG_FILE  = os.path.join(AEGISMOD_DIR, "config.json")

DEFAULT_STEPS = [
    {"label": "1回目", "action": "warn",    "seconds": 0},
    {"label": "2回目", "action": "timeout", "seconds": 60},
    {"label": "3回目", "action": "timeout", "seconds": 180},
    {"label": "4回目", "action": "timeout", "seconds": 480},
    {"label": "5回目", "action": "timeout", "seconds": 600},
]

DEFAULTS = {
    "channel": "", "bot": "", "token": "",
    # 連投検知
    "exact_enabled":  True,
    "exact_lim":      "20",
    "sim_enabled":    True,
    "sim_lim":        "20",
    "sim_thr":        "70",
    # 速度検知
    "speed_enabled":  True,
    "speed_count":    "3",
    "speed_secs":     "5",
    # NGワード
    "ng_enabled":     True,
    # 段階的ペナルティ
    "penalty_enabled": True,
    "penalty_steps":   DEFAULT_STEPS,
    "penalty_final_ban": False,
    "penalty_reset_days": "1",
    # 警告メッセージ
    "warn_enabled":   True,
    "warn_msg_repeat": "[警告] @{user} 同じ内容のコメントを繰り返すことはご遠慮ください。",
    "warn_msg_speed":  "[警告] @{user} コメントの連続投稿はお控えください。",
    "warn_msg_ng":     "[警告] @{user} 不適切な発言はお控えください。",
    # 自動再接続
    "auto_reconnect": True,
    # 履歴
    "hist_limit":     "50",
    # リスト
    "whitelist": [], "ngwords": [],
    # AI設定
    "ai_enabled":         False,
    "ai_score_threshold": "60",
    "ai_warn_msg":        "[AI警告] @{user} コメントが有害と判断されました。",
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                for k, v in DEFAULTS.items():
                    d.setdefault(k, v)
                return d
        except Exception:
            pass
    return dict(DEFAULTS)

def save_config(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ─────────────────────────────────────────
#  TOAST
# ─────────────────────────────────────────
def toast_notify(title, msg):
    try:
        t, m = title.replace("'",""), msg.replace("'","")
        script = (
            "[Windows.UI.Notifications.ToastNotificationManager,"
            "Windows.UI.Notifications,ContentType=WindowsRuntime]|Out-Null;"
            "[Windows.Data.Xml.Dom.XmlDocument,"
            "Windows.Data.Xml.Dom.XmlDocument,ContentType=WindowsRuntime]|Out-Null;"
            "$xml=[Windows.UI.Notifications.ToastNotificationManager]::"
            "GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
            f"$xml.GetElementsByTagName('text')[0].InnerText='{t}';"
            f"$xml.GetElementsByTagName('text')[1].InnerText='{m}';"
            "$toast=[Windows.UI.Notifications.ToastNotification]::new($xml);"
            "[Windows.UI.Notifications.ToastNotificationManager]::"
            "CreateToastNotifier('AegisMod').Show($toast);"
        )
        subprocess.Popen(["powershell","-WindowStyle","Hidden","-Command",script],
                         creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception:
        pass

# ─────────────────────────────────────────
#  COLORS & FONTS
# ─────────────────────────────────────────
BG       = "#0e0e10"
SURFACE  = "#18181b"
SURFACE2 = "#1f1f23"
SURFACE3 = "#26262c"
TWITCH   = "#9147ff"
TWITCH_L = "#a970ff"
DANGER   = "#ff4040"
WARN     = "#ffb347"
SUCCESS  = "#00e676"
INFO     = "#5bc0de"
TEXT     = "#efeff1"
MUTED    = "#adadb8"
BORDER   = "#2d2d35"

FN  = ("Meiryo", 10)
FS  = ("Meiryo", 9)
FXS = ("Meiryo", 8)
FB  = ("Meiryo", 10, "bold")
FH  = ("Meiryo", 12, "bold")
FNU = ("Meiryo", 18, "bold")

# ─────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────
KNOWN_EMOTES = {
    "Kappa","PogChamp","LUL","OMEGALUL","PepeHands","monkaS","KEKW","Sadge",
    "peepoHappy","TriHard","BibleThump","ResidentSleeper","NotLikeThis","HeyGuys",
    "SeemsGood","4Head","FrankerZ","DansGame","PauseChamp","EZClap","Pog",
    "POGGERS","pepeLaugh","FeelsBadMan","FeelsGoodMan","catJAM",
}

def similarity(a, b):
    a, b = a.lower().strip(), b.lower().strip()
    if a == b: return 100
    if not a or not b: return 0
    return int(SequenceMatcher(None, a, b).ratio() * 100)

def is_emote_only(msg):
    words = msg.strip().split()
    return all(w in KNOWN_EMOTES for w in words) if words else False

def fmt_secs(s):
    if s < 60: return f"{s}秒"
    m = s // 60
    return f"{m}分" if s % 60 == 0 else f"{m}分{s%60}秒"


def center_on_parent(win, parent, w, h):
    """親ウィンドウの中央に子ウィンドウを配置・画面外補正"""
    parent.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    px = parent.winfo_x()
    py = parent.winfo_y()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    # 画面外補正
    x = max(0, min(x, sw - w))
    y = max(0, min(y, sh - h))
    win.geometry(f"{w}x{h}+{x}+{y}")

def restore_geometry(win, cfg, key, default_geo):
    """保存されたウィンドウ位置を復元・画面外なら補正"""
    geo = cfg.get(key, default_geo)
    try:
        # geometry文字列 "WxH+X+Y" をパース
        import re as _re
        m = _re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", geo)
        if m:
            w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            x = max(0, min(x, sw - w))
            y = max(0, min(y, sh - h))
            win.geometry(f"{w}x{h}+{x}+{y}")
        else:
            win.geometry(default_geo)
    except Exception:
        win.geometry(default_geo)

def apply_scrollbar_style(widget):
    style = ttk.Style(widget)
    style.theme_use("default")
    style.configure("Dark.Vertical.TScrollbar",
                     background=SURFACE3, troughcolor=SURFACE,
                     bordercolor=SURFACE, arrowcolor=SURFACE,
                     relief=tk.FLAT, arrowsize=0, width=6)
    style.map("Dark.Vertical.TScrollbar",
              background=[("active", TWITCH_L), ("pressed", TWITCH),
                           ("!active", SURFACE3)])

# ─────────────────────────────────────────
#  SYSTRAY
# ─────────────────────────────────────────
try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

def make_app_icon():
    """盾＋剣のアイコン画像を生成（トレイ・exe共通）"""
    img = Image.new("RGBA", (256, 256), color=(0,0,0,0))
    d = ImageDraw.Draw(img)
    # 背景円
    d.ellipse([4, 4, 252, 252], fill="#1a0a2e")
    # 盾本体
    scale = 4
    shield = [(32*scale,4*scale),(56*scale,14*scale),(56*scale,36*scale),
              (32*scale,60*scale),(8*scale,36*scale),(8*scale,14*scale)]
    d.polygon(shield, fill="#9147ff")
    # 盾のハイライト
    inner = [(32*scale,10*scale),(50*scale,18*scale),(50*scale,34*scale),
             (32*scale,54*scale),(14*scale,34*scale),(14*scale,18*scale)]
    d.polygon(inner, fill="#a970ff")
    # 盾の中央ライン（縦）
    d.rectangle([30*scale,14*scale,34*scale,50*scale], fill="#6020cc")
    # 剣
    d.line([40*scale,12*scale,24*scale,44*scale], fill="white", width=10)
    d.polygon([(24*scale,44*scale),(20*scale,52*scale),(28*scale,48*scale)], fill="white")
    d.rectangle([39*scale,8*scale,43*scale,16*scale], fill="#dddddd")
    return img

def make_tray_icon():
    """タスクトレイ用（64x64）"""
    return make_app_icon().resize((64, 64), Image.LANCZOS)

def save_ico_if_needed():
    """exeと同じディレクトリにaegismod.icoを生成（なければ）"""
    ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aegismod.ico")
    if not os.path.exists(ico_path):
        try:
            img = make_app_icon()
            img.save(ico_path, format="ICO",
                     sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])
        except Exception:
            pass
    return ico_path

# ─────────────────────────────────────────
#  IRC
# ─────────────────────────────────────────
class TwitchIRC:
    def __init__(self, app):
        self.app = app; self.sock = None; self.running = False

    def connect(self, channel, bot, token):
        self.running = True
        try:
            raw = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            ctx = ssl.create_default_context()
            self.sock = ctx.wrap_socket(raw, server_hostname="irc.chat.twitch.tv")
            self.sock.connect(("irc.chat.twitch.tv", 6697))
            self._send("CAP REQ :twitch.tv/tags twitch.tv/commands")
            self._send(f"PASS {token}")
            self._send(f"NICK {bot}")
            self._send(f"JOIN #{channel}")
            self._recv_loop(channel, bot)
        except Exception as ex:
            self.app.log(f"[ERROR] 接続エラー: {ex}", "error")
            self.app.log("  -> OAuthトークンがBotアカウントのものか確認", "error")
            self.app.set_status(False)

    def _send(self, msg):
        try: self.sock.send((msg+"\r\n").encode("utf-8"))
        except Exception: pass

    def send_pub(self, channel, msg):
        self._send(f"PRIVMSG #{channel} :{msg}")

    def _recv_loop(self, channel, bot):
        buf = ""
        while self.running:
            try:
                data = self.sock.recv(4096).decode("utf-8", errors="ignore")
                if not data: break
                buf += data
                while "\r\n" in buf:
                    line, buf = buf.split("\r\n", 1)
                    self._parse(line, channel, bot)
            except Exception:
                break
        if self.running:
            self.app.log("接続が切断されました", "warn")
            self.app.set_status(False)
            self.app.after(0, lambda: self.app._try_reconnect(channel, bot, self.app._cfg.get("token","")))

    def _parse(self, line, channel, bot):
        if line.startswith("PING"):
            self._send("PONG :tmi.twitch.tv"); return
        tags = {}; rest = line
        if line.startswith("@"):
            sp = line.index(" ")
            for t in line[1:sp].split(";"):
                i = t.find("="); tags[t[:i]] = t[i+1:]
            rest = line[sp+1:]
        if " 001 " in rest or " 376 " in rest:
            self.app.set_status(True)
            self.app._reconnect_count = 0
            self.app.log(f"[OK] #{channel} に接続！自動モデレーション開始。", "ok"); return
        if "Login authentication failed" in rest or "Improperly formatted" in rest:
            self.app.log("[ERROR] 認証失敗: twitchtokengenerator.com でACCESS TOKENを再取得してください","error")
            self.app.set_status(False); return
        m = re.match(r"^:([^!]+)![^\s]+ PRIVMSG #([^\s]+) :(.+)$", rest)
        if not m: return
        uid, _, msg = m.group(1), m.group(2), m.group(3)
        if uid.lower() == bot.lower(): return
        # チャットリアルタイム表示
        self.app.after(0, lambda u=uid, ms=msg: self.app.log_chat(u, ms))
        if uid.lower() in [w.lower() for w in self.app.whitelist]: return
        if self.app.var_ign_emotes.get() and is_emote_only(msg): return
        # !ranking !score !pena はコマンド除外の対象外
        _public_cmds = ("!ranking", "!score", "!pena")
        if self.app.var_ign_cmd.get() and msg.strip().startswith("!") and not any(msg.strip().lower().startswith(p) for p in _public_cmds): return
        if self.app.var_ign_mod.get():
            badges = tags.get("badges","")
            if "moderator" in badges or "broadcaster" in badges or tags.get("mod")=="1": return
        self.app.cnt_monitored += 1
        self.app.after(0, self.app.update_stats)
        self.app.process_message(uid, msg, channel)

    def disconnect(self):
        self.running = False
        try:
            if self.sock: self.sock.close()
        except Exception: pass
        self.sock = None

# ─────────────────────────────────────────
#  ペナルティ段階エディタ（設定ウィンドウ内）
# ─────────────────────────────────────────
class PenaltyStepsEditor(tk.Frame):
    """段階的ペナルティの段階をGUIで編集するウィジェット"""
    def __init__(self, parent, steps):
        super().__init__(parent, bg=SURFACE)
        self.steps = [dict(s) for s in steps]
        self._rows = []
        self._build_header()
        self._build_rows()

    def _build_header(self):
        hdr = tk.Frame(self, bg=SURFACE3)
        hdr.pack(fill=tk.X, pady=(0,2))
        for txt, w in [("段階",40),("処置",80),("時間(分)",80),("",30)]:
            tk.Label(hdr, text=txt, bg=SURFACE3, fg=MUTED, font=FXS,
                     width=w//7, anchor="w").pack(side=tk.LEFT, padx=4, pady=3)

    def _build_rows(self):
        for w in self._rows:
            w.destroy()
        self._rows = []
        for i, step in enumerate(self.steps):
            row = tk.Frame(self, bg=SURFACE2)
            row.pack(fill=tk.X, pady=1)
            self._rows.append(row)
            n = i + 1
            tk.Label(row, text=f"{n}回目", bg=SURFACE2, fg=MUTED, font=FXS,
                     width=5, anchor="w").pack(side=tk.LEFT, padx=6, pady=4)
            # 処置選択
            act_val = {"warn":"警告","timeout":"発言禁止"}.get(step["action"], step["action"])
            act_var = tk.StringVar(value=act_val)
            act_menu = ttk.Combobox(row, textvariable=act_var,
                                     values=["警告","発言禁止"],
                                     state="readonly", width=8, font=FN)
            act_menu.pack(side=tk.LEFT, padx=4, pady=3)
            # 秒数入力
            sec_val = str(round(step["seconds"] / 60, 1)) if step["seconds"] >= 60 else "0"
            sec_var = tk.StringVar(value=sec_val)
            sec_ent = tk.Entry(row, textvariable=sec_var, width=7, bg=SURFACE3,
                               fg=TEXT, insertbackground=TEXT, font=FN,
                               relief=tk.FLAT, highlightbackground=BORDER,
                               highlightthickness=1, highlightcolor=TWITCH)
            sec_ent.pack(side=tk.LEFT, padx=4, pady=3, ipady=3)
            tk.Label(row, text="分", bg=SURFACE2, fg=MUTED, font=FN).pack(side=tk.LEFT)
            # 削除ボタン
            idx = i
            tk.Button(row, text="✕", bg=SURFACE2, fg=DANGER, font=FS,
                      relief=tk.FLAT, cursor="hand2", bd=0,
                      command=lambda ii=idx: self._remove(ii)
                      ).pack(side=tk.RIGHT, padx=6)
            # 変数バインド
            act_var.trace_add("write", lambda *a, ii=idx, v=act_var: self._update_action(ii, v))
            sec_var.trace_add("write", lambda *a, ii=idx, v=sec_var: self._update_secs(ii, v))

        # 追加ボタン
        add_row = tk.Frame(self, bg=SURFACE)
        add_row.pack(fill=tk.X, pady=(4,0))
        self._rows.append(add_row)
        tk.Button(add_row, text="+ 段階を追加", bg=SURFACE3, fg=TWITCH_L,
                  font=FN, relief=tk.FLAT, cursor="hand2", bd=0,
                  command=self._add_step).pack(anchor="w", ipadx=6, ipady=3)

    def _update_action(self, i, var):
        if i < len(self.steps):
            jp2en = {"警告":"warn","発言禁止":"timeout"}
            self.steps[i]["action"] = jp2en.get(var.get(), var.get())

    def _update_secs(self, i, var):
        if i < len(self.steps):
            try: self.steps[i]["seconds"] = int(float(var.get()) * 60)
            except ValueError: pass

    def _remove(self, i):
        if len(self.steps) > 1:
            self.steps.pop(i)
            self._build_rows()

    def _add_step(self):
        n = len(self.steps) + 1
        self.steps.append({"label": f"{n}回目", "action": "timeout", "seconds": 600})
        self._build_rows()

    def get_steps(self):
        return [dict(s) for s in self.steps]

# ─────────────────────────────────────────
#  設定ウィンドウ
# ─────────────────────────────────────────
class SettingsWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("設定")
        self.minsize(460, 560)
        self.configure(bg=SURFACE)
        self.transient(app)
        apply_scrollbar_style(self)
        self._build()
        self._load()
        center_on_parent(self, app, 500, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _ent(self, p, width=None, **kw):
        kws = dict(bg=SURFACE2, fg=TEXT, insertbackground=TEXT, relief=tk.FLAT,
                   font=FN, highlightbackground=BORDER, highlightthickness=1,
                   highlightcolor=TWITCH)
        if width: kws["width"] = width
        kws.update(kw)
        e = tk.Entry(p, **kws)
        e.pack(fill=tk.X, pady=2, ipady=5)
        return e

    def _btn(self, p, text, cmd, bg=None, fg="white", **kw):
        return tk.Button(p, text=text, command=cmd, bg=bg or TWITCH, fg=fg,
                         font=FB, relief=tk.FLAT, cursor="hand2",
                         activebackground=TWITCH_L, activeforeground="white", **kw)

    def _sec(self, p, title, var=None):
        """セクションヘッダー（ON/OFFチェックボックス付き）"""
        r = tk.Frame(p, bg=SURFACE); r.pack(fill=tk.X, pady=(12,4), padx=12)
        tk.Label(r, text=f"── {title}", bg=SURFACE, fg=TWITCH_L, font=FB).pack(side=tk.LEFT)
        if var is not None:
            tk.Checkbutton(r, text="ON", variable=var, bg=SURFACE, fg=MUTED,
                           selectcolor=TWITCH, activebackground=SURFACE,
                           font=FXS).pack(side=tk.RIGHT)

    def _numrow(self, p, lbl, attr, width=7):
        r = tk.Frame(p, bg=SURFACE); r.pack(fill=tk.X, pady=2)
        tk.Label(r, text=lbl, bg=SURFACE, fg=MUTED, font=FS,
                 width=22, anchor="w").pack(side=tk.LEFT)
        v = tk.StringVar()
        setattr(self, attr, v)
        tk.Entry(r, textvariable=v, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 width=width, font=FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.LEFT, ipady=4, padx=4)
        return v

    def _build(self):
        canvas = tk.Canvas(self, bg=SURFACE, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview,
                            style="Dark.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        f = tk.Frame(canvas, bg=SURFACE)
        win = canvas.create_window((0,0), window=f, anchor="nw")
        f.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
        pad = dict(padx=12)

        # ── 認証 ──
        self._sec(f, "認証設定")
        af = tk.Frame(f, bg=SURFACE); af.pack(fill=tk.X, **pad)
        tk.Label(af, text="Channel Name", bg=SURFACE, fg=MUTED, font=FXS).pack(anchor="w", pady=(4,1))
        self.ent_channel = self._ent(af)
        tk.Label(af, text="Bot Username", bg=SURFACE, fg=MUTED, font=FXS).pack(anchor="w", pady=(4,1))
        self.ent_bot = self._ent(af)
        tk.Label(af, text="OAuth Token  ( ACCESS_TOKEN をそのまま貼り付けてください )", bg=SURFACE, fg=MUTED, font=FXS).pack(anchor="w", pady=(4,1))
        self.ent_token = self._ent(af, show="*")
        tk.Label(af, text="-> twitchtokengenerator.com でBotアカウントのACCESS TOKENを取得",
                 bg=SURFACE, fg=MUTED, font=FXS, wraplength=420, justify=tk.LEFT).pack(anchor="w")
        tk.Label(af, text="-> /mod Botアカウント名 でModに設定してください",
                 bg=SURFACE, fg=WARN, font=FXS).pack(anchor="w", pady=(2,0))

        # ── 段階的ペナルティ ──
        self.var_penalty_enabled = tk.BooleanVar()
        self._sec(f, "段階的ペナルティ", self.var_penalty_enabled)
        pf = tk.Frame(f, bg=SURFACE); pf.pack(fill=tk.X, **pad)

        self.penalty_editor = PenaltyStepsEditor(pf, DEFAULT_STEPS)
        self.penalty_editor.pack(fill=tk.X, pady=4)

        fr = tk.Frame(pf, bg=SURFACE); fr.pack(fill=tk.X, pady=2)
        self.var_final_ban = tk.BooleanVar()
        tk.Checkbutton(fr, text="最終段階の次はBAN（永久追放）",
                       variable=self.var_final_ban, bg=SURFACE, fg=TEXT,
                       selectcolor=TWITCH, activebackground=SURFACE, font=FN).pack(side=tk.LEFT)

        rr = tk.Frame(pf, bg=SURFACE); rr.pack(fill=tk.X, pady=2)
        tk.Label(rr, text="発言禁止後 ", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.LEFT)
        self.var_reset_days = tk.StringVar()
        tk.Entry(rr, textvariable=self.var_reset_days, width=5, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, font=FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.LEFT, ipady=4, padx=4)
        tk.Label(rr, text="日後に違反カウントをリセット", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.LEFT)

        # ── 連投検知 ──
        self.var_exact_enabled = tk.BooleanVar()
        self._sec(f, "完全一致 連投検知", self.var_exact_enabled)
        ef = tk.Frame(f, bg=SURFACE); ef.pack(fill=tk.X, **pad)
        self._numrow(ef, "制限回数", "var_exact_limit")

        self.var_sim_enabled = tk.BooleanVar()
        self._sec(f, "類似コメント 連投検知", self.var_sim_enabled)
        sf = tk.Frame(f, bg=SURFACE); sf.pack(fill=tk.X, **pad)
        self._numrow(sf, "制限回数", "var_sim_limit")
        self._numrow(sf, "類似度しきい値 (%)", "var_sim_thresh")

        # ── 速度検知 ──
        self.var_speed_enabled = tk.BooleanVar()
        self._sec(f, "連投速度検知", self.var_speed_enabled)
        vf = tk.Frame(f, bg=SURFACE); vf.pack(fill=tk.X, **pad)
        r1 = tk.Frame(vf, bg=SURFACE); r1.pack(fill=tk.X, pady=2)
        tk.Label(r1, text="秒以内に", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.RIGHT)
        self.var_speed_secs = tk.StringVar()
        tk.Entry(r1, textvariable=self.var_speed_secs, width=5, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, font=FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.RIGHT, ipady=4, padx=4)
        tk.Label(r1, text="設定時間:", bg=SURFACE, fg=MUTED, font=FS, width=12, anchor="w").pack(side=tk.LEFT)
        r2 = tk.Frame(vf, bg=SURFACE); r2.pack(fill=tk.X, pady=2)
        tk.Label(r2, text="回投稿でペナルティ", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.RIGHT)
        self.var_speed_count = tk.StringVar()
        tk.Entry(r2, textvariable=self.var_speed_count, width=5, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, font=FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.RIGHT, ipady=4, padx=4)
        tk.Label(r2, text="投稿数:", bg=SURFACE, fg=MUTED, font=FS, width=12, anchor="w").pack(side=tk.LEFT)

        # ── NGワード ──
        self.var_ng_enabled = tk.BooleanVar()
        self._sec(f, "NGワードリスト", self.var_ng_enabled)
        nf = tk.Frame(f, bg=SURFACE); nf.pack(fill=tk.X, **pad)
        tk.Label(nf, text="NGワードを含むコメントは段階的ペナルティカウントに加算されます",
                 bg=SURFACE, fg=MUTED, font=FXS, wraplength=420).pack(anchor="w", pady=(0,4))
        nr = tk.Frame(nf, bg=SURFACE); nr.pack(fill=tk.X)
        self.ent_ng = tk.Entry(nr, bg=SURFACE2, fg=TEXT, insertbackground=TEXT, font=FN,
                               relief=tk.FLAT, highlightbackground=BORDER,
                               highlightthickness=1, highlightcolor=TWITCH)
        self.ent_ng.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.ent_ng.bind("<Return>", lambda e: self._add_ng())
        self._btn(nr, "追加", self._add_ng).pack(side=tk.LEFT, padx=(4,0), ipady=5, ipadx=6)
        self.lst_ng_var = tk.StringVar()
        self.lst_ng = tk.Listbox(nf, listvariable=self.lst_ng_var, bg=SURFACE2, fg=TEXT,
                                  font=FN, selectbackground=TWITCH, relief=tk.FLAT,
                                  height=4, highlightthickness=0)
        self.lst_ng.pack(fill=tk.X, pady=2)
        self._btn(nf, "選択したNGワードを削除", self._remove_ng,
                  bg=SURFACE3, fg=DANGER).pack(fill=tk.X, ipady=4, pady=2)

        # ── ホワイトリスト ──
        self._sec(f, "ホワイトリスト")
        wlf = tk.Frame(f, bg=SURFACE); wlf.pack(fill=tk.X, **pad)
        wr = tk.Frame(wlf, bg=SURFACE); wr.pack(fill=tk.X)
        self.ent_wl = tk.Entry(wr, bg=SURFACE2, fg=TEXT, insertbackground=TEXT, font=FN,
                               relief=tk.FLAT, highlightbackground=BORDER,
                               highlightthickness=1, highlightcolor=TWITCH)
        self.ent_wl.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)
        self.ent_wl.bind("<Return>", lambda e: self._add_wl())
        self._btn(wr, "追加", self._add_wl).pack(side=tk.LEFT, padx=(4,0), ipady=5, ipadx=6)
        self.lst_wl_var = tk.StringVar()
        self.lst_wl = tk.Listbox(wlf, listvariable=self.lst_wl_var, bg=SURFACE2, fg=TEXT,
                                  font=FN, selectbackground=TWITCH, relief=tk.FLAT,
                                  height=3, highlightthickness=0)
        self.lst_wl.pack(fill=tk.X, pady=2)
        self._btn(wlf, "選択したユーザーを削除", self._remove_wl,
                  bg=SURFACE3, fg=DANGER).pack(fill=tk.X, ipady=4, pady=2)

        # ── 警告メッセージ ──
        self._sec(f, "警告メッセージ")
        wmf = tk.Frame(f, bg=SURFACE); wmf.pack(fill=tk.X, **pad)
        self.var_warn_enabled = tk.BooleanVar()
        tk.Checkbutton(wmf, text="ペナルティ実行前に警告メッセージを送信する",
                       variable=self.var_warn_enabled, bg=SURFACE, fg=TEXT,
                       selectcolor=TWITCH, activebackground=SURFACE, font=FN).pack(anchor="w", pady=(0,4))
        tk.Label(wmf, text="{user} = 対象ユーザー名に置換されます",
                 bg=SURFACE, fg=MUTED, font=FXS).pack(anchor="w", pady=(0,6))

        def warn_field(parent, label):
            tk.Label(parent, text=label, bg=SURFACE, fg=MUTED, font=FXS).pack(anchor="w", pady=(4,1))
            t = tk.Text(parent, height=2, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                        font=FN, relief=tk.FLAT, wrap=tk.WORD,
                        highlightbackground=BORDER, highlightthickness=1, highlightcolor=TWITCH)
            t.pack(fill=tk.X, pady=(0,4))
            return t

        self.ent_warn_repeat = warn_field(wmf, "完全一致・類似コメント違反時")
        self.ent_warn_speed  = warn_field(wmf, "連投速度違反時")
        self.ent_warn_ng     = warn_field(wmf, "NGワード違反時")

        # ── その他 ──
        self._sec(f, "その他")
        of = tk.Frame(f, bg=SURFACE); of.pack(fill=tk.X, **pad)

        self.var_ign_emotes = tk.BooleanVar()
        self.var_ign_cmd    = tk.BooleanVar()
        self.var_ign_mod    = tk.BooleanVar()
        self.var_auto_reconnect = tk.BooleanVar()
        for txt, var in [("スタンプのみのコメントを除外", self.var_ign_emotes),
                         ("コマンド(!)を除外",           self.var_ign_cmd),
                         ("Mod・配信者を除外",           self.var_ign_mod),
                         ("切断時に自動再接続（30秒×最大3回）", self.var_auto_reconnect)]:
            tk.Checkbutton(of, text=txt, variable=var, bg=SURFACE, fg=TEXT,
                           selectcolor=TWITCH, activebackground=SURFACE,
                           font=FN).pack(anchor="w", pady=1)

        hr = tk.Frame(of, bg=SURFACE); hr.pack(fill=tk.X, pady=4)
        tk.Label(hr, text="履歴保持件数 (ユーザーごと)", bg=SURFACE, fg=MUTED, font=FS,
                 width=24, anchor="w").pack(side=tk.LEFT)
        self.var_hist_limit = tk.StringVar()
        tk.Entry(hr, textvariable=self.var_hist_limit, width=6, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, font=FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.LEFT, ipady=4, padx=4)
        tk.Label(hr, text="件", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.LEFT)

        # ── 保存ボタン ──
        bf = tk.Frame(f, bg=SURFACE); bf.pack(fill=tk.X, padx=12, pady=(12,16))
        self._btn(bf, "設定を保存して閉じる", self._save_and_close).pack(fill=tk.X, ipady=7, pady=(0,4))
        self._btn(bf, "すべての設定をデフォルトに戻す", self._reset_to_defaults,
                  bg=SURFACE3, fg=WARN).pack(fill=tk.X, ipady=5)

    # ── リスト操作 ──
    def _add_ng(self):
        val = self.ent_ng.get().strip()
        if val and val not in self.app.ngwords:
            self.app.ngwords.append(val); self._refresh_ng()
        self.ent_ng.delete(0, tk.END)

    def _remove_ng(self):
        sel = self.lst_ng.curselection()
        if sel: self.app.ngwords.pop(sel[0]); self._refresh_ng()

    def _refresh_ng(self): self.lst_ng_var.set(self.app.ngwords)

    def _add_wl(self):
        val = self.ent_wl.get().strip().lower()
        if val and val not in self.app.whitelist:
            self.app.whitelist.append(val); self._refresh_wl()
        self.ent_wl.delete(0, tk.END)

    def _remove_wl(self):
        sel = self.lst_wl.curselection()
        if sel: self.app.whitelist.pop(sel[0]); self._refresh_wl()

    def _refresh_wl(self): self.lst_wl_var.set(self.app.whitelist)

    def _load(self):
        cfg = self.app._cfg
        def _set(e, v): e.delete(0, tk.END); e.insert(0, v)
        _set(self.ent_channel, cfg["channel"])
        _set(self.ent_bot,     cfg["bot"])
        _set(self.ent_token,   cfg["token"])
        self.var_penalty_enabled.set(cfg["penalty_enabled"])
        self.var_final_ban.set(cfg["penalty_final_ban"])
        self.var_reset_days.set(cfg.get("penalty_reset_days", cfg.get("penalty_reset_min", "1")))
        self.penalty_editor.steps = [dict(s) for s in cfg["penalty_steps"]]
        self.penalty_editor._build_rows()
        self.var_exact_enabled.set(cfg["exact_enabled"])
        self.var_exact_limit.set(cfg["exact_lim"])
        self.var_sim_enabled.set(cfg["sim_enabled"])
        self.var_sim_limit.set(cfg["sim_lim"])
        self.var_sim_thresh.set(cfg["sim_thr"])
        self.var_speed_enabled.set(cfg["speed_enabled"])
        self.var_speed_count.set(cfg["speed_count"])
        self.var_speed_secs.set(cfg["speed_secs"])
        self.var_ng_enabled.set(cfg["ng_enabled"])
        self.var_warn_enabled.set(cfg["warn_enabled"])
        self.ent_warn_repeat.delete("1.0", tk.END)
        self.ent_warn_repeat.insert("1.0", cfg.get("warn_msg_repeat", "[警告] @{user} 同じ内容のコメントを繰り返すことはご遠慮ください。"))
        self.ent_warn_speed.delete("1.0", tk.END)
        self.ent_warn_speed.insert("1.0", cfg.get("warn_msg_speed", "[警告] @{user} コメントの連続投稿はお控えください。"))
        self.ent_warn_ng.delete("1.0", tk.END)
        self.ent_warn_ng.insert("1.0", cfg.get("warn_msg_ng", "[警告] @{user} 不適切な発言はお控えください。"))
        self.var_ign_emotes.set(self.app.var_ign_emotes.get())
        self.var_ign_cmd.set(self.app.var_ign_cmd.get())
        self.var_ign_mod.set(self.app.var_ign_mod.get())
        self.var_auto_reconnect.set(cfg.get("auto_reconnect", True))
        self.var_hist_limit.set(cfg.get("hist_limit", "50"))
        self._refresh_ng(); self._refresh_wl()

    def _apply(self):
        cfg = self.app._cfg
        cfg["channel"]          = self.ent_channel.get().strip()
        cfg["bot"]              = self.ent_bot.get().strip()
        cfg["token"]            = self.ent_token.get().strip()
        cfg["penalty_enabled"]  = self.var_penalty_enabled.get()
        cfg["penalty_steps"]    = self.penalty_editor.get_steps()
        cfg["penalty_final_ban"]= self.var_final_ban.get()
        cfg["penalty_reset_days"] = self.var_reset_days.get()
        cfg["exact_enabled"]    = self.var_exact_enabled.get()
        cfg["exact_lim"]        = self.var_exact_limit.get()
        cfg["sim_enabled"]      = self.var_sim_enabled.get()
        cfg["sim_lim"]          = self.var_sim_limit.get()
        cfg["sim_thr"]          = self.var_sim_thresh.get()
        cfg["speed_enabled"]    = self.var_speed_enabled.get()
        cfg["speed_count"]      = self.var_speed_count.get()
        cfg["speed_secs"]       = self.var_speed_secs.get()
        cfg["ng_enabled"]       = self.var_ng_enabled.get()
        cfg["warn_enabled"]      = self.var_warn_enabled.get()
        cfg["warn_msg_repeat"]   = self.ent_warn_repeat.get("1.0", tk.END).strip()
        cfg["warn_msg_speed"]    = self.ent_warn_speed.get("1.0", tk.END).strip()
        cfg["warn_msg_ng"]       = self.ent_warn_ng.get("1.0", tk.END).strip()
        cfg["auto_reconnect"]   = self.var_auto_reconnect.get()
        cfg["hist_limit"]       = self.var_hist_limit.get()
        self.app.var_ign_emotes.set(self.var_ign_emotes.get())
        self.app.var_ign_cmd.set(self.var_ign_cmd.get())
        self.app.var_ign_mod.set(self.var_ign_mod.get())
        self.app._sync_conn_bar()

    def _reset_to_defaults(self):
        if not messagebox.askyesno("確認",
            "すべての設定をデフォルト値に戻します。\n"
            "チャンネル名・Bot名・OAuthトークンもリセットされます。\nよろしいですか？"):
            return
        # cfgをデフォルトで上書き
        for k, v in DEFAULTS.items():
            self.app._cfg[k] = v if not isinstance(v, list) else list(v)
        self.app.whitelist = []
        self.app.ngwords   = []
        self._load()
        self.app.log("[INFO] 設定をデフォルトに戻しました", "info")

    def _save_and_close(self):
        self._apply(); self.app._save_fields()
        self.destroy(); self.app._settings_win = None

    def _on_close(self):
        self._apply()
        # 変更があれば保存確認
        import json as _json
        current = _json.dumps(self.app._cfg, ensure_ascii=False, sort_keys=True)
        saved_cfg = load_config()
        saved = _json.dumps(saved_cfg, ensure_ascii=False, sort_keys=True)
        if current != saved:
            if messagebox.askyesno("保存確認", "設定が変更されています。保存しますか？"):
                self.app._save_fields()
        self.destroy(); self.app._settings_win = None

# ─────────────────────────────────────────
#  ログウィンドウ
# ─────────────────────────────────────────
class LogWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("モデレーションログ")
        self.minsize(620,360)
        self.configure(bg=BG)
        self.transient(app)
        apply_scrollbar_style(self)
        self._build()
        center_on_parent(self, app, 860, 560)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _btn(self, p, text, cmd, bg=None, fg="white", **kw):
        return tk.Button(p, text=text, command=cmd, bg=bg or TWITCH, fg=fg,
                         font=FB, relief=tk.FLAT, cursor="hand2",
                         activebackground=TWITCH_L, activeforeground="white", **kw)

    def _build(self):
        hdr = tk.Frame(self, bg=SURFACE, height=40)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="モデレーションログ", bg=SURFACE, fg=TWITCH_L, font=FB).pack(side=tk.LEFT, padx=12)
        ctrl = tk.Frame(hdr, bg=SURFACE); ctrl.pack(side=tk.RIGHT, padx=8)
        self._btn(ctrl, "CSV出力", self._export_csv).pack(side=tk.RIGHT, padx=4, pady=6, ipadx=6, ipady=2)
        self._btn(ctrl, "クリア", self._clear_log, bg=SURFACE3, fg=DANGER).pack(side=tk.RIGHT, padx=4, pady=6, ipadx=6, ipady=2)
        tk.Label(ctrl, text="検索:", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.RIGHT, padx=(8,2))
        self.ent_search = tk.Entry(ctrl, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                                    font=FN, relief=tk.FLAT, highlightbackground=BORDER,
                                    highlightthickness=1, highlightcolor=TWITCH, width=18)
        self.ent_search.pack(side=tk.RIGHT, ipady=4, pady=6)
        self.ent_search.bind("<KeyRelease>", lambda e: self.refresh())
        tk.Frame(self, bg=TWITCH, height=2).pack(fill=tk.X)
        tf = tk.Frame(self, bg=BG); tf.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        style = ttk.Style(self)
        style.configure("Log.Treeview", background=SURFACE2, foreground=TEXT,
                         fieldbackground=SURFACE2, font=FN, rowheight=26, borderwidth=0)
        style.configure("Log.Treeview.Heading", background=SURFACE3, foreground=MUTED, font=FS, relief=tk.FLAT)
        style.map("Log.Treeview", background=[("selected",SURFACE3)], foreground=[("selected",TEXT)])
        cols = ("時刻","対象者ID","コメント","回数","種別","処置")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse", style="Log.Treeview")
        for col, w in zip(cols, [70,110,300,45,80,100]):
            self.tree.heading(col, text=col); self.tree.column(col, width=w, minwidth=40)
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview, style="Dark.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("timeout", foreground=DANGER)
        self.tree.tag_configure("ban",     foreground="#ff8080")
        self.tree.tag_configure("warning", foreground=WARN)
        self.tree.tag_configure("ng",      foreground="#ff6060")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.refresh()

        # ── 取り消しパネル ──
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X, padx=0)
        rev = tk.Frame(self, bg=SURFACE, height=52)
        rev.pack(fill=tk.X); rev.pack_propagate(False)

        tk.Label(rev, text="取り消し対象:", bg=SURFACE, fg=MUTED, font=FS).pack(side=tk.LEFT, padx=(12,4))
        self.lbl_target = tk.Label(rev, text="(ログ行を選択してください)",
                                    bg=SURFACE, fg=MUTED, font=FN)
        self.lbl_target.pack(side=tk.LEFT, padx=(0,12))

        # 取り消しボタン群
        bf = tk.Frame(rev, bg=SURFACE); bf.pack(side=tk.RIGHT, padx=8)

        self.btn_reset_count = tk.Button(bf, text="違反カウントリセット",
                   command=self._reset_count,
                   bg=SURFACE3, fg=MUTED, font=FS, relief=tk.FLAT, cursor="hand2",
                   activebackground=SURFACE2, activeforeground=TEXT, bd=0, state=tk.DISABLED)
        self.btn_reset_count.pack(side=tk.LEFT, padx=4, ipady=4, ipadx=8)

        self.btn_untimeout = tk.Button(bf, text="発言禁止を解除",
                   command=self._untimeout,
                   bg=SURFACE3, fg=WARN, font=FS, relief=tk.FLAT, cursor="hand2",
                   activebackground=SURFACE2, activeforeground=WARN, bd=0, state=tk.DISABLED)
        self.btn_untimeout.pack(side=tk.LEFT, padx=4, ipady=4, ipadx=8)

        self.btn_unban = tk.Button(bf, text="BANを解除",
                   command=self._unban,
                   bg=SURFACE3, fg=DANGER, font=FS, relief=tk.FLAT, cursor="hand2",
                   activebackground=SURFACE2, activeforeground=DANGER, bd=0, state=tk.DISABLED)
        self.btn_unban.pack(side=tk.LEFT, padx=4, ipady=4, ipadx=8)

    def refresh(self):
        q = self.ent_search.get().strip().lower()
        rows = [r for r in self.app.log_data
                if not q or q in r["uid"].lower() or q in r["comment"].lower()]
        self.tree.delete(*self.tree.get_children())
        km = {"exact":"完全一致","similar":"類似","ng":"NGワード","speed":"速度違反"}
        for r in reversed(rows):
            tag = "ban" if r["action"]=="ban" else ("warning" if r["action"]=="warn" else "timeout")
            if r["kind"] == "ng": tag = "ng"
            self.tree.insert("", tk.END,
                values=(r["time"], r["uid"], r["comment"], r["count"],
                        km.get(r["kind"],r["kind"]), r["action_label"]),
                tags=(tag,))

    def _on_select(self, event=None):
        """ログ行選択時に対象ユーザーを取り消しパネルに表示"""
        sel = self.tree.selection()
        if not sel:
            self._clear_target(); return
        vals = self.tree.item(sel[0], "values")
        if not vals: self._clear_target(); return
        uid    = vals[1]
        action = vals[5]  # action_label
        self.lbl_target.config(text=f"@{uid}  [{action}]", fg=TEXT)
        self._selected_uid = uid
        # ボタンを有効化
        self.btn_reset_count.config(state=tk.NORMAL, fg=TWITCH_L)
        self.btn_untimeout.config(state=tk.NORMAL)
        self.btn_unban.config(state=tk.NORMAL)

    def _clear_target(self):
        self._selected_uid = None
        self.lbl_target.config(text="(ログ行を選択してください)", fg=MUTED)
        self.btn_reset_count.config(state=tk.DISABLED, fg=MUTED)
        self.btn_untimeout.config(state=tk.DISABLED)
        self.btn_unban.config(state=tk.DISABLED)

    def _get_uid(self):
        uid = getattr(self, "_selected_uid", None)
        if not uid:
            messagebox.showinfo("情報", "ログ行を選択してください")
        return uid

    def _untimeout(self):
        uid = self._get_uid()
        if not uid: return
        ch  = self.app._cfg.get("channel","").strip().lower()
        if not ch or not self.app.connected:
            messagebox.showwarning("未接続", "接続中のみ実行できます"); return
        self.app.irc.send_pub(ch, f"/untimeout {uid}")
        self.app.log(f"[取り消し] {uid} の発言禁止を解除しました", "ok")
        toast_notify("取り消し", f"{uid} の発言禁止を解除")

    def _unban(self):
        uid = self._get_uid()
        if not uid: return
        ch  = self.app._cfg.get("channel","").strip().lower()
        if not ch or not self.app.connected:
            messagebox.showwarning("未接続", "接続中のみ実行できます"); return
        if not messagebox.askyesno("確認", f"@{uid} のBANを解除しますか？"):
            return
        self.app.irc.send_pub(ch, f"/unban {uid}")
        self.app.log(f"[取り消し] {uid} のBANを解除しました", "ok")
        toast_notify("取り消し", f"{uid} のBAN解除")

    def _reset_count(self):
        uid = self._get_uid()
        if not uid: return
        if not messagebox.askyesno("確認", f"@{uid} の違反カウントをリセットしますか？\n（発言禁止・BANは解除されません）"):
            return
        self.app.penalty_count.pop(uid, None)
        self.app.penalty_reset_time.pop(uid, None)
        self.app.user_hist.pop(uid, None)
        self.app.speed_hist.pop(uid, None)
        self.app.log(f"[取り消し] {uid} の違反カウントをリセットしました", "ok")
        messagebox.showinfo("完了", f"@{uid} の違反カウントをリセットしました")

    def _clear_log(self):
        if messagebox.askyesno("確認","ログをすべて削除しますか？"):
            self.app.log_data.clear(); self.refresh()

    def _export_csv(self):
        if not self.app.log_data:
            messagebox.showinfo("情報","エクスポートするログがありません"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV files","*.csv")],
            initialfile=f"aegismod-{datetime.now().strftime('%Y%m%d')}.csv")
        if not path: return
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["time","uid","comment","count","kind","action","action_label"])
            w.writeheader(); w.writerows(self.app.log_data)
        messagebox.showinfo("完了",f"CSVをエクスポートしました:\n{path}")

    def _on_close(self):
        self.destroy(); self.app._log_win = None

# ─────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aegis Mod")
        self.minsize(560,460)
        self.configure(bg=BG); self.resizable(True, True)

        self._cfg              = load_config()
        self.irc               = TwitchIRC(self)
        self.connected         = False
        self.start_time        = None
        self.cnt_monitored     = 0
        self.cnt_timeouts      = 0
        self.cnt_warnings      = 0
        # user_hist: uid -> deque of msg strings
        self.user_hist         = {}
        # speed_hist: uid -> deque of timestamps
        self.speed_hist        = {}
        # penalty_count: uid -> int (違反カウント)
        self.penalty_count     = {}
        # penalty_reset_time: uid -> timestamp (最後のタイムアウト時刻)
        self.penalty_reset_time = {}
        self.whitelist         = list(self._cfg["whitelist"])
        self.ngwords           = list(self._cfg["ngwords"])
        self.log_data          = []
        self._uptime_job       = None
        self._tray             = None
        self._settings_win     = None
        self._log_win          = None
        self._reconnect_count  = 0
        self._reconnect_job    = None
        self._ai_engine        = get_engine() if HAS_AI else None
        self._cmd_history      = []   # コマンド履歴
        self._cmd_hist_idx     = -1   # 履歴カーソル
        self._recent_users     = []   # 最近BAN/発言禁止したユーザー
        self._suggest_popup    = None # サジェストポップアップ

        self.var_ign_emotes    = tk.BooleanVar(value=True)
        self.var_ign_cmd       = tk.BooleanVar(value=True)
        self.var_ign_mod       = tk.BooleanVar(value=True)

        apply_scrollbar_style(self)
        self._build_ui()
        self._sync_conn_bar()
        # ウィンドウアイコン設定
        try:
            ico = save_ico_if_needed()
            self.iconbitmap(ico)
        except Exception:
            pass
        restore_geometry(self, self._cfg, "main_geometry", "660x460+100+100")
        self._shortcuts = self._cfg.get("shortcuts", {})
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ══════════════════════════════════════
    #  UI構築
    # ══════════════════════════════════════
    def _build_ui(self):
        # ヘッダー
        hdr = tk.Frame(self, bg=SURFACE, height=46)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="Aegis Mod", bg=SURFACE, fg=TWITCH_L, font=FH).pack(side=tk.LEFT, padx=14)
        btn_area = tk.Frame(hdr, bg=SURFACE); btn_area.pack(side=tk.RIGHT, padx=8)
        tk.Button(btn_area, text="トレイに格納", command=self._minimize_to_tray,
                   bg=SURFACE3, fg=MUTED, font=FS, relief=tk.FLAT, cursor="hand2",
                   activebackground=SURFACE2, activeforeground=TEXT, bd=0
                   ).pack(side=tk.RIGHT, padx=4, pady=10, ipadx=6, ipady=2)
        tk.Button(btn_area, text="Aegis AI β", command=self._open_ai_window,
                   bg=SURFACE3, fg=TWITCH_L, font=FB, relief=tk.FLAT, cursor="hand2",
                   activebackground=TWITCH, activeforeground="white", bd=0
                   ).pack(side=tk.RIGHT, padx=4, pady=8, ipadx=8, ipady=2)
        tk.Button(btn_area, text="ログ", command=self._open_log,
                   bg=SURFACE3, fg=TEXT, font=FB, relief=tk.FLAT, cursor="hand2",
                   activebackground=TWITCH, activeforeground="white", bd=0
                   ).pack(side=tk.RIGHT, padx=4, pady=8, ipadx=8, ipady=2)
        tk.Button(btn_area, text="設定", command=self._open_settings,
                   bg=TWITCH, fg="white", font=FB, relief=tk.FLAT, cursor="hand2",
                   activebackground=TWITCH_L, activeforeground="white", bd=0
                   ).pack(side=tk.RIGHT, padx=4, pady=8, ipadx=8, ipady=2)
        self.lbl_status = tk.Label(hdr, text="● 未接続", bg=SURFACE, fg=MUTED, font=FS)
        self.lbl_status.pack(side=tk.RIGHT, padx=(0,8))
        tk.Frame(self, bg=TWITCH, height=2).pack(fill=tk.X)

        # 接続バー
        conn = tk.Frame(self, bg=SURFACE2, height=46)
        conn.pack(fill=tk.X); conn.pack_propagate(False)
        tk.Label(conn, text="Ch:", bg=SURFACE2, fg=MUTED, font=FS).pack(side=tk.LEFT, padx=(12,2))
        self.lbl_channel = tk.Label(conn, text="--", bg=SURFACE2, fg=TEXT, font=FB)
        self.lbl_channel.pack(side=tk.LEFT, padx=(0,10))
        tk.Label(conn, text="Bot:", bg=SURFACE2, fg=MUTED, font=FS).pack(side=tk.LEFT, padx=(0,2))
        self.lbl_bot = tk.Label(conn, text="--", bg=SURFACE2, fg=TEXT, font=FB)
        self.lbl_bot.pack(side=tk.LEFT)
        bf = tk.Frame(conn, bg=SURFACE2); bf.pack(side=tk.RIGHT, padx=8)
        self.btn_disconnect = tk.Button(bf, text="切断", command=self._disconnect,
                                         bg=SURFACE3, fg=DANGER, font=FB, relief=tk.FLAT,
                                         cursor="hand2", activebackground=SURFACE,
                                         activeforeground=DANGER, state=tk.DISABLED, bd=0)
        self.btn_disconnect.pack(side=tk.RIGHT, pady=8, ipadx=8, ipady=3)
        self.btn_connect = tk.Button(bf, text="接続", command=self._connect,
                                      bg=TWITCH, fg="white", font=FB, relief=tk.FLAT,
                                      cursor="hand2", activebackground=TWITCH_L,
                                      activeforeground="white", bd=0)
        self.btn_connect.pack(side=tk.RIGHT, padx=(0,6), pady=8, ipadx=12, ipady=3)
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X)

        # 統計バー
        stats = tk.Frame(self, bg=SURFACE, height=72)
        stats.pack(fill=tk.X); stats.pack_propagate(False)
        for i, (lbl, attr, color) in enumerate([
            ("監視数",       "var_sM", TEXT),
            ("発言禁止", "var_sT", DANGER),
            ("警告",         "var_sW", WARN),
            ("稼働時間",     "var_sU", SUCCESS),
        ]):
            if i > 0:
                tk.Frame(stats, bg=BORDER, width=1).pack(side=tk.LEFT, fill=tk.Y, pady=10)
            c = tk.Frame(stats, bg=SURFACE); c.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            v = tk.StringVar(value="0" if attr != "var_sU" else "--:--")
            setattr(self, attr, v)
            tk.Label(c, textvariable=v, bg=SURFACE, fg=color, font=FNU).pack(pady=(8,0))
            tk.Label(c, text=lbl,       bg=SURFACE, fg=MUTED,  font=FS).pack(pady=(0,8))
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X)

        # ライブモニター（チャット表示 + システムログ）- expand=True で可変
        mf = tk.Frame(self, bg=BG); mf.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6,0))
        mh = tk.Frame(mf, bg=BG);   mh.pack(fill=tk.X, pady=(0,4))
        tk.Label(mh, text="ライブモニター", bg=BG, fg=TWITCH_L, font=FB).pack(side=tk.LEFT)
        self.var_show_chat = tk.BooleanVar(value=True)
        tk.Checkbutton(mh, text="チャット表示", variable=self.var_show_chat,
                       bg=BG, fg=MUTED, selectcolor=TWITCH,
                       activebackground=BG, font=FXS).pack(side=tk.RIGHT, padx=8)
        tk.Button(mh, text="クリア", command=self._clear_mon,
                   bg=SURFACE3, fg=MUTED, font=FS, relief=tk.FLAT, cursor="hand2",
                   activebackground=SURFACE2, activeforeground=TEXT, bd=0
                   ).pack(side=tk.RIGHT, ipadx=6, ipady=2)
        mon_frame = tk.Frame(mf, bg=SURFACE2)
        mon_frame.pack(fill=tk.BOTH, expand=True)
        self.txt_mon = tk.Text(mon_frame, bg=SURFACE2, fg=MUTED, font=FN,
                               relief=tk.FLAT, state=tk.DISABLED,
                               insertbackground=TEXT, wrap=tk.WORD, height=4)
        mon_vsb = ttk.Scrollbar(mon_frame, orient="vertical",
                                 command=self.txt_mon.yview,
                                 style="Dark.Vertical.TScrollbar")
        self.txt_mon.configure(yscrollcommand=mon_vsb.set)
        mon_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_mon.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        for tag, color in [("ok",SUCCESS),("warn",WARN),("error",DANGER),
                            ("info",MUTED),("ts",TWITCH_L),("chat",TEXT),("chat_uid",INFO)]:
            self.txt_mon.tag_config(tag, foreground=color)

        # ── コマンドバー（常に表示・固定高さ）──
        tk.Frame(self, bg=BORDER, height=1).pack(fill=tk.X)
        cmd_bar = tk.Frame(self, bg=SURFACE2, height=42)
        cmd_bar.pack(fill=tk.X); cmd_bar.pack_propagate(False)
        tk.Label(cmd_bar, text="CMD:", bg=SURFACE2, fg=MUTED, font=FS).pack(side=tk.LEFT, padx=(10,4))
        self.ent_cmd = tk.Entry(cmd_bar, bg=SURFACE3, fg=TEXT, insertbackground=TEXT,
                                 font=FN, relief=tk.FLAT, highlightbackground=BORDER,
                                 highlightthickness=1, highlightcolor=TWITCH)
        self.ent_cmd.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, pady=6)
        tk.Button(cmd_bar, text="送信", command=self._send_cmd,
                   bg=TWITCH, fg="white", font=FB, relief=tk.FLAT, cursor="hand2",
                   activebackground=TWITCH_L, activeforeground="white", bd=0
                   ).pack(side=tk.LEFT, padx=(6,10), pady=6, ipadx=12, ipady=3)

        # キーバインド
        self.ent_cmd.bind("<Return>",   lambda e: self._send_cmd())
        self.ent_cmd.bind("<Up>",       self._cmd_hist_up)
        self.ent_cmd.bind("<Down>",     self._cmd_hist_down)
        self.ent_cmd.bind("<KeyRelease>", self._on_cmd_key)
        self.ent_cmd.bind("<FocusOut>", lambda e: self._hide_suggest())

    def _sync_conn_bar(self):
        ch  = self._cfg.get("channel","") or "--"
        bot = self._cfg.get("bot","")     or "--"
        self.lbl_channel.config(text=ch)
        self.lbl_bot.config(text=bot)

    # ══════════════════════════════════════
    #  ウィンドウ管理
    # ══════════════════════════════════════
    def _open_settings(self):
        if self._settings_win and self._settings_win.winfo_exists():
            self._settings_win.lift(); return
        self._settings_win = SettingsWindow(self)

    def _open_log(self):
        if self._log_win and self._log_win.winfo_exists():
            self._log_win.lift(); return
        self._log_win = LogWindow(self)

    # ══════════════════════════════════════
    #  CONNECT / RECONNECT
    # ══════════════════════════════════════
    def _connect(self):
        ch  = self._cfg.get("channel","").strip().lower()
        bot = self._cfg.get("bot","").strip().lower()
        tok = self._cfg.get("token","").strip()
        if not ch or not bot or not tok:
            messagebox.showwarning("入力エラー",
                "チャンネル名・Bot名・OAuthトークンが未設定です。\n「設定」ボタンから入力してください。")
            self._open_settings(); return
        if not tok.startswith("oauth:"):
            tok = "oauth:" + tok
            self._cfg["token"] = tok
        self._reconnect_count = 0
        self.log(f"#{ch} に接続中...", "info")
        self._save_fields()
        self.irc = TwitchIRC(self)
        threading.Thread(target=self.irc.connect, args=(ch, bot, tok), daemon=True).start()

    def _try_reconnect(self, channel, bot, token):
        if not self._cfg.get("auto_reconnect", True): return
        if self._reconnect_count >= 3:
            self.log("[INFO] 自動再接続を3回試みましたが失敗しました。手動で再接続してください。", "warn")
            return
        self._reconnect_count += 1
        self.log(f"[INFO] 30秒後に自動再接続します... ({self._reconnect_count}/3)", "info")
        self._reconnect_job = self.after(30000, lambda: self._do_reconnect(channel, bot, token))

    def _do_reconnect(self, channel, bot, token):
        self.log(f"[INFO] 自動再接続中... ({self._reconnect_count}/3)", "info")
        self.irc = TwitchIRC(self)
        threading.Thread(target=self.irc.connect, args=(channel, bot, token), daemon=True).start()

    def _disconnect(self):
        if self._reconnect_job:
            self.after_cancel(self._reconnect_job)
            self._reconnect_job = None
        self._reconnect_count = 3  # 自動再接続を停止
        self.irc.disconnect(); self.set_status(False); self.log("切断しました","info")

    def set_status(self, on):
        self.connected = on
        if on:
            self.lbl_status.config(text="● 監視中", fg=SUCCESS)
            self.btn_connect.config(state=tk.DISABLED)
            self.btn_disconnect.config(state=tk.NORMAL)
            self.start_time = time.time(); self._tick_uptime()
        else:
            self.lbl_status.config(text="● 未接続", fg=MUTED)
            self.btn_connect.config(state=tk.NORMAL)
            self.btn_disconnect.config(state=tk.DISABLED)
            if self._uptime_job: self.after_cancel(self._uptime_job)
            self.var_sU.set("--:--")

    def _tick_uptime(self):
        if not self.connected: return
        s = int(time.time()-self.start_time)
        h, r = divmod(s, 3600); m, sc = divmod(r, 60)
        self.var_sU.set(f"{h:02d}:{m:02d}:{sc:02d}" if h else f"{m:02d}:{sc:02d}")
        self._uptime_job = self.after(1000, self._tick_uptime)

    # ══════════════════════════════════════
    #  ペナルティ判定
    # ══════════════════════════════════════
    def _get_penalty_step(self, uid):
        """現在の違反カウントに対応するステップを返す"""
        cfg = self._cfg
        steps = cfg.get("penalty_steps", DEFAULT_STEPS)
        count = self.penalty_count.get(uid, 0)  # 0-indexed: 0=初回
        if count < len(steps):
            return steps[count]
        # 最終段階を超えた場合
        if cfg.get("penalty_final_ban", False):
            return {"action": "ban", "seconds": 0}
        return steps[-1]  # 最終ステップを繰り返す

    def _check_cooldown(self, uid):
        """クールダウン経過していればカウントリセット"""
        if uid not in self.penalty_reset_time:
            return
        try:    reset_days = float(self._cfg.get("penalty_reset_days", self._cfg.get("penalty_reset_min", "1")))
        except: reset_days = 1.0
        elapsed = time.time() - self.penalty_reset_time[uid]
        if elapsed >= reset_days * 86400:
            self.penalty_count[uid] = 0
            del self.penalty_reset_time[uid]

    def _apply_penalty(self, uid, msg, kind, channel):
        """共通ペナルティ実行"""
        cfg = self._cfg
        kind_jp = {"exact":"完全一致","similar":"類似","ng":"NGワード","speed":"速度違反","ai":"AI判定"}.get(kind, kind)

        # 段階的ペナルティOFF → 検知通知のみ・処置なし
        if not cfg.get("penalty_enabled", True):
            self.log(f"[検知] {uid} ({kind_jp}) 「{msg[:40]}」 ※段階的ペナルティOFF・処置なし", "info")
            self._add_log(uid, msg, 0, kind, "none", "処置なし（ペナルティOFF）")
            return

        self._check_cooldown(uid)
        step = self._get_penalty_step(uid)
        self.penalty_count[uid] = self.penalty_count.get(uid, 0) + 1
        count = self.penalty_count[uid]
        action = step["action"]
        secs   = step.get("seconds", 0)

        if action == "warn":
            self._send_warn_msg(uid, channel, kind)
            label = "警告"
            self._add_log(uid, msg, count, kind, "warn", "警告のみ")
            self.cnt_warnings += 1
            self.after(0, self.update_stats)
            self.log(f"[WARN] {uid} に警告 ({count}回目) 「{msg[:40]}」", "warn")
            toast_notify("警告", f"{uid} に警告 ({count}回目)")
        elif action == "ban":
            self._send_warn_msg(uid, channel, kind)
            self.irc.send_pub(channel, f"/ban {uid}")
            label = "永久BAN"
            self._add_log(uid, msg, count, kind, "ban", "永久BAN")
            self.user_hist[uid] = deque()
            self.cnt_timeouts += 1
            self.after(0, self.update_stats)
            self.log(f"[BAN] {uid} を永久BAN ({count}回目) 「{msg[:40]}」", "error")
            toast_notify("永久BAN", f"{uid} ({count}回目)")
        else:  # timeout
            self._send_warn_msg(uid, channel, kind)
            self.irc.send_pub(channel, f"/timeout {uid} {secs}")
            label = fmt_secs(secs) + "発言禁止"
            self._add_log(uid, msg, count, kind, "timeout", label)
            self.penalty_reset_time[uid] = time.time()
            self.user_hist[uid] = deque()
            self._add_recent_user(uid)
            self.cnt_timeouts += 1
            self.after(0, self.update_stats)
            kind_str = {"exact":"完全一致","similar":"類似","ng":"NGワード","speed":"速度違反"}.get(kind, kind)
            self.log(f"[BAN] {uid} を{label} ({count}回目/{kind_str}) 「{msg[:40]}」", "error")
            toast_notify("発言禁止", f"{uid} を{label} ({count}回目)")

    # ══════════════════════════════════════
    #  MOD LOGIC
    # ══════════════════════════════════════
    def process_message(self, uid, msg, channel):
        cfg = self._cfg
        # ── リスナーコマンド ──
        cmd = msg.strip().lower()
        if cmd == "!ranking":
            self.after(0, lambda: self._cmd_ranking(channel)); return
        if cmd == "!score":
            self.after(0, lambda u=uid: self._cmd_score(u, channel)); return
        if cmd == "!pena":
            self.after(0, lambda u=uid: self._cmd_pena(u, channel)); return

        try:    hist_limit = int(cfg.get("hist_limit","50"))
        except: hist_limit = 50

        # 履歴初期化
        if uid not in self.user_hist:
            self.user_hist[uid] = deque(maxlen=hist_limit)
        if uid not in self.speed_hist:
            self.speed_hist[uid] = deque()

        # NGワード検知
        if cfg.get("ng_enabled", True):
            for ng in self.ngwords:
                if ng.lower() in msg.lower():
                    self.after(0, lambda u=uid,m=msg: self._apply_penalty(u, m, "ng", channel))
                    return

        # 連投速度検知
        if cfg.get("speed_enabled", True):
            try:    spd_secs  = float(cfg.get("speed_secs","5"))
            except: spd_secs  = 5.0
            try:    spd_count = int(cfg.get("speed_count","3"))
            except: spd_count = 3
            now = time.time()
            sh = self.speed_hist[uid]
            sh.append(now)
            # 古いタイムスタンプを削除
            while sh and now - sh[0] > spd_secs:
                sh.popleft()
            if len(sh) >= spd_count:
                self.speed_hist[uid] = deque()
                self.after(0, lambda u=uid,m=msg: self._apply_penalty(u, m, "speed", channel))
                return

        # 完全一致・類似 連投検知
        self.user_hist[uid].append(msg)
        hist = list(self.user_hist[uid])

        if cfg.get("exact_enabled", True):
            try:    el = int(cfg.get("exact_lim","20"))
            except: el = 20
            exact_count = sum(1 for h in hist if h.lower() == msg.lower())
            if exact_count >= el:
                self.after(0, lambda u=uid,m=msg,c=exact_count: self._apply_penalty(u, m, "exact", channel))
                return

        if cfg.get("sim_enabled", True):
            try:    sl = int(cfg.get("sim_lim","20"))
            except: sl = 20
            try:    st = int(cfg.get("sim_thr","70"))
            except: st = 70
            # AIエンベディングが使えればそちらで類似度計算
            if self._ai_engine and cfg.get("ai_enabled") and HAS_AI:
                sim_scores = [self._ai_engine.embedding_similarity(h, msg)
                              for h in hist if h.lower() != msg.lower()]
                sim_count  = sum(1 for s in sim_scores if s >= st) + 1
            else:
                sim_count = len([h for h in hist if h.lower()!=msg.lower() and similarity(h,msg)>=st]) + 1
            if sim_count >= sl:
                self.after(0, lambda u=uid,m=msg,c=sim_count: self._apply_penalty(u, m, "similar", channel))
                return

        # ── AIスコアリング（通常チャット含む全コメント）──
        if self._ai_engine and cfg.get("ai_enabled") and HAS_AI:
            use_oa  = cfg.get("ai_use_openai", False)
            api_key = cfg.get("ai_openai_key", "")
            try:    thresh = int(cfg.get("ai_score_threshold", "60"))
            except: thresh = 60
            def _ai_check(u=uid, m=msg):
                result = self._ai_engine.calc_score(u, m, "chat", api_key, use_oa)
                score  = result["score"]
                # ライブモニターにスコア表示
                if score >= 30:
                    color = "error" if score >= thresh else "warn"
                    self.after(0, lambda sc=score, um=u: self.log(
                        f"[AI] {um} スコア:{sc}/100", color))
                # しきい値超えでペナルティ
                if score >= thresh:
                    self.after(0, lambda u2=u, m2=m: self._apply_ai_penalty(u2, m2, score))
            threading.Thread(target=_ai_check, daemon=True).start()

    def _send_warn_msg(self, uid, channel, kind="repeat"):
        cfg = self._cfg
        if not cfg.get("warn_enabled", True): return
        key = {"speed": "warn_msg_speed", "ng": "warn_msg_ng"}.get(kind, "warn_msg_repeat")
        tmpl = cfg.get(key, "")
        if tmpl:
            self.irc.send_pub(channel, tmpl.replace("{user}", uid))

    # ══════════════════════════════════════
    #  LOG & MONITOR
    # ══════════════════════════════════════
    def _add_log(self, uid, comment, count, kind, action, action_label):
        self.log_data.append(dict(
            time=datetime.now().strftime("%H:%M:%S"),
            uid=uid, comment=comment, count=count,
            kind=kind, action=action, action_label=action_label))
        if self._log_win and self._log_win.winfo_exists():
            self._log_win.refresh()

    def log(self, msg, kind="info"):
        self.txt_mon.config(state=tk.NORMAL)
        self.txt_mon.insert(tk.END, datetime.now().strftime("%H:%M:%S")+"  ", "ts")
        self.txt_mon.insert(tk.END, msg+"\n", kind)
        self.txt_mon.see(tk.END)
        self.txt_mon.config(state=tk.DISABLED)

    def log_chat(self, uid, msg):
        """チャットのリアルタイム表示"""
        if not self.var_show_chat.get(): return
        self.txt_mon.config(state=tk.NORMAL)
        self.txt_mon.insert(tk.END, datetime.now().strftime("%H:%M:%S")+"  ", "ts")
        self.txt_mon.insert(tk.END, f"{uid}: ", "chat_uid")
        self.txt_mon.insert(tk.END, msg+"\n", "chat")
        self.txt_mon.see(tk.END)
        self.txt_mon.config(state=tk.DISABLED)

    def _clear_mon(self):
        self.txt_mon.config(state=tk.NORMAL)
        self.txt_mon.delete("1.0", tk.END)
        self.txt_mon.config(state=tk.DISABLED)

    def update_stats(self):
        self.var_sM.set(str(self.cnt_monitored))
        self.var_sT.set(str(self.cnt_timeouts))
        self.var_sW.set(str(self.cnt_warnings))

    # ══════════════════════════════════════
    #  TRAY
    # ══════════════════════════════════════
    def _minimize_to_tray(self):
        if not HAS_TRAY:
            messagebox.showinfo("情報",
                "タスクトレイ機能を使うには pystray と Pillow が必要です。\n"
                "build-run.bat でexeをビルドすると自動インストールされます。"); return
        self.withdraw()
        if self._tray is None:
            menu = pystray.Menu(
                pystray.MenuItem("開く", self._tray_show),
                pystray.MenuItem("終了", self._tray_quit),
            )
            self._tray = pystray.Icon("AegisMod", make_tray_icon(), "Aegis Mod 監視中", menu)
            threading.Thread(target=self._tray.run, daemon=True).start()

    def _tray_show(self): self.deiconify(); self.lift()

    def _tray_quit(self):
        self._save_fields(); self.irc.disconnect()
        if self._tray: self._tray.stop()
        self.destroy()

    # ══════════════════════════════════════
    #  SAVE / LOAD
    # ══════════════════════════════════════
    def _save_fields(self):
        self._cfg["whitelist"] = self.whitelist
        self._cfg["ngwords"]   = self.ngwords
        save_config(self._cfg)
        self.log("[OK] 設定を保存しました","ok")


    # ══════════════════════════════════════
    #  コマンド入力
    # ══════════════════════════════════════
    # Twitchモデレーターコマンド候補
    CMD_SUGGESTIONS = [
        "/ban ","/unban ","/timeout ","/untimeout ",
        "/clear","/slow ","/slowoff",
        "/subscribers","/subscribersoff",
        "/emoteonly","/emoteonlyoff",
        "/followers","/followersoff",
        "/mod ","/unmod ","/vip ","/unvip ",
        "/color ","/commercial ","/host ","/unhost ",
    ]

    def _send_cmd(self):
        txt = self.ent_cmd.get().strip()
        if not txt: return
        ch = self._cfg.get("channel","").strip().lower()
        if not ch or not self.connected:
            self.log("[WARN] 接続中でないとコマンドを送信できません", "warn"); return
        self.irc.send_pub(ch, txt)
        self.log(f"[CMD] {txt}", "info")
        # 履歴に追加（重複なし・最大50件）
        if not self._cmd_history or self._cmd_history[-1] != txt:
            self._cmd_history.append(txt)
            if len(self._cmd_history) > 50:
                self._cmd_history.pop(0)
        self._cmd_hist_idx = -1
        self.ent_cmd.delete(0, tk.END)
        self._hide_suggest()

    def _cmd_hist_up(self, event):
        if not self._cmd_history: return
        if self._cmd_hist_idx == -1:
            self._cmd_hist_idx = len(self._cmd_history) - 1
        elif self._cmd_hist_idx > 0:
            self._cmd_hist_idx -= 1
        self._set_cmd(self._cmd_history[self._cmd_hist_idx])

    def _cmd_hist_down(self, event):
        if self._cmd_hist_idx == -1: return
        if self._cmd_hist_idx < len(self._cmd_history) - 1:
            self._cmd_hist_idx += 1
            self._set_cmd(self._cmd_history[self._cmd_hist_idx])
        else:
            self._cmd_hist_idx = -1
            self.ent_cmd.delete(0, tk.END)

    def _set_cmd(self, txt):
        self.ent_cmd.delete(0, tk.END)
        self.ent_cmd.insert(0, txt)
        self.ent_cmd.icursor(tk.END)

    def _on_cmd_key(self, event):
        if event.keysym in ("Up","Down","Return","Escape"): return
        txt = self.ent_cmd.get()
        if not txt:
            self._hide_suggest(); return
        # サジェスト候補を絞り込む
        candidates = []
        if txt.startswith("/"):
            # コマンド補完
            candidates = [s for s in self.CMD_SUGGESTIONS if s.startswith(txt)]
        # ユーザー名補完：コマンドの2番目トークンが未完成
        parts = txt.split(" ")
        if len(parts) == 2 and parts[0] in ("/ban","/unban","/timeout","/untimeout","/mod","/unmod","/vip","/unvip"):
            partial = parts[1].lower()
            user_cands = [u for u in self._recent_users if u.lower().startswith(partial)]
            candidates = [f"{parts[0]} {u}" for u in user_cands] + candidates
        if candidates:
            self._show_suggest(candidates[:8])
        else:
            self._hide_suggest()

    def _show_suggest(self, items):
        self._hide_suggest()
        # 入力欄の位置に合わせてポップアップ
        self.ent_cmd.update_idletasks()
        x = self.ent_cmd.winfo_rootx()
        y = self.ent_cmd.winfo_rooty() - len(items) * 24 - 4
        pop = tk.Toplevel(self)
        pop.wm_overrideredirect(True)
        pop.geometry(f"+{x}+{y}")
        pop.configure(bg=BORDER)
        self._suggest_popup = pop
        for item in items:
            btn = tk.Button(pop, text=item, bg=SURFACE3, fg=TEXT, font=FN,
                            relief=tk.FLAT, anchor="w", cursor="hand2",
                            activebackground=TWITCH, activeforeground="white",
                            command=lambda v=item: self._pick_suggest(v))
            btn.pack(fill=tk.X, padx=1, pady=1, ipadx=8, ipady=2)

    def _pick_suggest(self, val):
        self._set_cmd(val)
        self._hide_suggest()
        self.ent_cmd.focus_set()

    def _hide_suggest(self, event=None):
        if self._suggest_popup:
            try: self._suggest_popup.destroy()
            except Exception: pass
            self._suggest_popup = None

    def _add_recent_user(self, uid):
        if uid in self._recent_users:
            self._recent_users.remove(uid)
        self._recent_users.insert(0, uid)
        if len(self._recent_users) > 20:
            self._recent_users.pop()

    # ── ショートカット管理ウィンドウ ──
    def _open_shortcuts(self):
        win = tk.Toplevel(self)
        win.title("コマンドショートカット")
        win.configure(bg=SURFACE)
        win.transient(self)
        center_on_parent(win, self, 360, 280)
        win.resizable(False, False)

        tk.Label(win, text="F1〜F4 にコマンドを登録できます",
                 bg=SURFACE, fg=MUTED, font=FXS).pack(pady=(12,6), padx=16, anchor="w")
        tk.Label(win, text="接続中に F キーを押すと即送信されます",
                 bg=SURFACE, fg=MUTED, font=FXS).pack(padx=16, anchor="w", pady=(0,10))

        entries = {}
        for key in ["F1","F2","F3","F4"]:
            r = tk.Frame(win, bg=SURFACE); r.pack(fill=tk.X, padx=16, pady=3)
            tk.Label(r, text=f"{key}:", bg=SURFACE, fg=TWITCH_L, font=FB,
                     width=4, anchor="w").pack(side=tk.LEFT)
            e = tk.Entry(r, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                         font=FN, relief=tk.FLAT, highlightbackground=BORDER,
                         highlightthickness=1, highlightcolor=TWITCH)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5, padx=(4,0))
            e.insert(0, self._shortcuts.get(key,""))
            entries[key] = e

        def save():
            for key, e in entries.items():
                v = e.get().strip()
                if v: self._shortcuts[key] = v
                elif key in self._shortcuts: del self._shortcuts[key]
            self._cfg["shortcuts"] = self._shortcuts
            win.destroy()
            self.log("[OK] ショートカットを保存しました", "ok")

        tk.Button(win, text="保存して閉じる", command=save,
                   bg=TWITCH, fg="white", font=FB, relief=tk.FLAT,
                   cursor="hand2", activebackground=TWITCH_L,
                   activeforeground="white").pack(fill=tk.X, padx=16, pady=(12,16), ipady=6)

    def _run_shortcut(self, key):
        cmd = self._shortcuts.get(key,"")
        if not cmd: return
        ch = self._cfg.get("channel","").strip().lower()
        if not ch or not self.connected: return
        self.irc.send_pub(ch, cmd)
        self.log(f"[CMD/{key}] {cmd}", "info")


    # ══════════════════════════════════════
    #  AI機能
    # ══════════════════════════════════════
    def _apply_ai_penalty(self, uid, msg, score):
        """AIスコアによる警告・発言禁止（BAN除外・段階的ペナルティ完全踏襲）"""
        cfg = self._cfg
        ch  = cfg.get("channel","").strip().lower()
        if not ch or not self.connected: return
        if uid.lower() in [w.lower() for w in self.whitelist]: return

        # 段階的ペナルティOFF → 検知通知のみ・処置なし
        if not cfg.get("penalty_enabled", True):
            self.log(f"[検知] {uid} (AI score:{score}) 「{msg[:40]}」 ※段階的ペナルティOFF・処置なし", "info")
            self._add_log(uid, msg, 0, "ai", "none", "処置なし（ペナルティOFF）")
            return

        self._check_cooldown(uid)

        # BANは除外: banステップを見つけたらその手前のtimeoutステップを使う
        steps = cfg.get("penalty_steps", DEFAULT_STEPS)
        count = self.penalty_count.get(uid, 0)
        # 使うステップを決定（BAN除外）
        if count < len(steps):
            step = steps[count]
        else:
            step = steps[-1]
        # BANのとき -> ban以外の最後のtimeoutステップを探す
        if step.get("action") == "ban":
            timeout_steps = [s for s in steps if s.get("action") != "ban"]
            step = timeout_steps[-1] if timeout_steps else {"action": "timeout", "seconds": 600}
        self.penalty_count[uid] = self.penalty_count.get(uid, 0) + 1
        count  = self.penalty_count[uid]
        action = step["action"]
        secs   = step.get("seconds", 0)

        warn_tmpl = cfg.get("ai_warn_msg", "[AI警告] @{user} コメントが有害と判断されました。")
        if cfg.get("warn_enabled", True) and warn_tmpl:
            self.irc.send_pub(ch, warn_tmpl.replace("{user}", uid))

        if action == "warn":
            self._add_log(uid, msg, count, "ai", "warn", f"AI警告(score:{score})")
            self.cnt_warnings += 1
            self.after(0, self.update_stats)
            self.log(f"[AI] {uid} に警告 score:{score} ({count}回目) 「{msg[:40]}」", "warn")
            toast_notify("AI警告", f"{uid} score:{score}")
        else:
            self.irc.send_pub(ch, f"/timeout {uid} {secs}")
            label = fmt_secs(secs) + "発言禁止"
            self._add_log(uid, msg, count, "ai", "timeout", f"AI {label}(score:{score})")
            self.penalty_reset_time[uid] = time.time()
            self.user_hist[uid] = deque()
            self._add_recent_user(uid)
            self.cnt_timeouts += 1
            self.after(0, self.update_stats)
            self.log(f"[AI] {uid} を{label} score:{score} ({count}回目) 「{msg[:40]}」", "error")
            toast_notify("AI発言禁止", f"{uid} score:{score}")

        # フィードバック記録（後からCSV化できるよう）
        if self._ai_engine:
            self._ai_engine.add_feedback(uid, msg, score, "ai", True)

    # ══════════════════════════════════════
    #  リスナーコマンド
    # ══════════════════════════════════════
    def _cmd_ranking(self, channel):
        """!ranking -> チャットにランキングトップ3を投稿"""
        if not self._ai_engine:
            self.irc.send_pub(channel, "AI機能が無効です。"); return
        rows = self._ai_engine.get_ranking(3)
        if not rows:
            self.irc.send_pub(channel, "[Aegis] ランキングデータがまだありません。"); return
        medals = ["🥇","🥈","🥉"]
        lines  = ["[Aegis] 酷すぎるコメントランキング TOP3"]
        for i, r in enumerate(rows):
            lines.append(f"{medals[i]} {r['uid']}  合計:{r['total_score']}pt  平均:{r['avg_score']}pt  最高:{r['max_score']}pt")
        mvp = self._ai_engine.get_mvp()
        if mvp:
            lines.append(f"最強コメンテーター賞: @{mvp['uid']} おめでとう！")
        for line in lines:
            self.irc.send_pub(channel, line)
            time.sleep(0.5)
        self.log(f"[CMD] !ranking 発表", "info")

    def _cmd_score(self, uid, channel):
        """!score -> 自分の直近AIスコアをチャットに返信"""
        if not self._ai_engine:
            self.irc.send_pub(channel, f"@{uid} AI機能が無効です。"); return
        stats = self._ai_engine.get_user_stats(uid)
        if stats["count"] == 0:
            self.irc.send_pub(channel, f"@{uid} スコアデータがありません。")
        else:
            self.irc.send_pub(channel,
                f"@{uid} あなたのAIスコア -> "
                f"直近{stats['count']}件: 平均{stats['avg']}pt / 最高{stats['max']}pt")
        self.log(f"[CMD] !score {uid}", "info")

    def _cmd_pena(self, uid, channel):
        """!pena -> 自分の段階・ペナルティ状況をチャットに返信"""
        self._check_cooldown(uid)
        count = self.penalty_count.get(uid, 0)
        steps = self._cfg.get("penalty_steps", DEFAULT_STEPS)

        if count == 0:
            self.irc.send_pub(channel, f"@{uid} 現在ペナルティ記録はありません。")
            self.log(f"[CMD] !pena {uid} -> 記録なし", "info")
            return

        # 現在の段階
        cur_idx = min(count - 1, len(steps) - 1)
        cur     = steps[cur_idx]
        action_jp = {"warn": "警告", "timeout": "発言禁止"}.get(cur.get("action",""), cur.get("action",""))
        cur_label = action_jp if cur.get("action") == "warn" else f'{action_jp} {fmt_secs(cur.get("seconds",0))}'

        # 次のステップ
        next_idx = min(count, len(steps) - 1)
        nxt      = steps[next_idx]
        nxt_action = nxt.get("action","")
        if nxt_action == "ban":
            # BAN除外: ban以外の最後を使う
            timeout_steps = [s for s in steps if s.get("action") != "ban"]
            nxt = timeout_steps[-1] if timeout_steps else {"action":"timeout","seconds":600}
            nxt_action = nxt.get("action","")
        nxt_jp    = {"warn": "警告のみ", "timeout": "発言禁止"}.get(nxt_action, nxt_action)
        nxt_label = nxt_jp if nxt_action == "warn" else f'{nxt_jp} {fmt_secs(nxt.get("seconds",0))}'

        msg_out = (
            f"@{uid} {count}回目の違反です（現在の処置: {cur_label}）。"
            f"次回のペナルティは「{nxt_label}」の処置です。"
        )
        self.irc.send_pub(channel, msg_out)
        self.log(f"[CMD] !pena {uid} -> {msg_out}", "info")

    def _open_ai_window(self):
        if not HAS_AI:
            messagebox.showinfo("AI機能",
                "AI機能を使うには以下をインストールしてください:\n"
                "pip install sentence-transformers openai\n\n"
                "exeビルド時は build-run.bat が自動インストールします。")
            return
        win = tk.Toplevel(self)
        win.title("AI モデレーション β")
        win.configure(bg=SURFACE)
        win.transient(self)
        center_on_parent(win, self, 560, 620)
        apply_scrollbar_style(win)
        self._build_ai_window(win)

    def _build_ai_window(self, win):
        S_FN = FN; S_FS = FS; S_FB = FB; S_FXS = FXS

        # ヘッダー
        hdr = tk.Frame(win, bg=SURFACE, height=44); hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="AI モデレーション β", bg=SURFACE, fg=TWITCH_L, font=S_FB).pack(side=tk.LEFT, padx=12)
        tk.Label(hdr, text="β版・動作は保証しません", bg=SURFACE, fg=WARN, font=S_FXS).pack(side=tk.RIGHT, padx=12)
        tk.Frame(win, bg=TWITCH, height=2).pack(fill=tk.X)

        canvas = tk.Canvas(win, bg=SURFACE, highlightthickness=0)
        vsb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview, style="Dark.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        f = tk.Frame(canvas, bg=SURFACE)
        wid = canvas.create_window((0,0), window=f, anchor="nw")
        f.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        pad = dict(padx=12)

        def sec(title):
            r = tk.Frame(f, bg=SURFACE); r.pack(fill=tk.X, pady=(12,4), **pad)
            tk.Label(r, text=f"── {title}", bg=SURFACE, fg=TWITCH_L, font=S_FB).pack(side=tk.LEFT)

        def numrow(parent, lbl, var, width=7):
            r = tk.Frame(parent, bg=SURFACE); r.pack(fill=tk.X, pady=2)
            tk.Label(r, text=lbl, bg=SURFACE, fg=MUTED, font=S_FS, width=22, anchor="w").pack(side=tk.LEFT)
            tk.Entry(r, textvariable=var, width=width, bg=SURFACE2, fg=TEXT,
                     insertbackground=TEXT, font=S_FN, relief=tk.FLAT,
                     highlightbackground=BORDER, highlightthickness=1,
                     highlightcolor=TWITCH).pack(side=tk.LEFT, ipady=4, padx=4)

        # ── ON/OFF ──
        sec("AI モデレーション設定")
        af = tk.Frame(f, bg=SURFACE); af.pack(fill=tk.X, **pad)
        var_en = tk.BooleanVar(value=self._cfg.get("ai_enabled", False))
        tk.Checkbutton(af, text="AI モデレーションを有効にする（β）",
                       variable=var_en, bg=SURFACE, fg=TEXT,
                       selectcolor=TWITCH, activebackground=SURFACE, font=S_FN).pack(anchor="w")
        tk.Label(af, text="有効にすると全コメントをリアルタイムにAI解析します。処理負荷が増加します。",
                 bg=SURFACE, fg=MUTED, font=S_FXS, wraplength=500).pack(anchor="w", pady=(2,0))

        # ── しきい値 ──
        sec("スコアしきい値")
        sf = tk.Frame(f, bg=SURFACE); sf.pack(fill=tk.X, **pad)
        var_thresh = tk.StringVar(value=self._cfg.get("ai_score_threshold", "60"))
        tr = tk.Frame(sf, bg=SURFACE); tr.pack(fill=tk.X, pady=2)
        tk.Label(tr, text="ペナルティ実行スコア", bg=SURFACE, fg=MUTED, font=S_FS, anchor="w").pack(side=tk.LEFT)
        tk.Entry(tr, textvariable=var_thresh, width=6, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, font=S_FN, relief=tk.FLAT,
                 highlightbackground=BORDER, highlightthickness=1,
                 highlightcolor=TWITCH).pack(side=tk.LEFT, ipady=4, padx=8)
        tk.Label(tr, text="/ 100", bg=SURFACE, fg=MUTED, font=S_FS).pack(side=tk.LEFT)
        tk.Label(sf, text="しきい値以上でペナルティ実行",
                 bg=SURFACE, fg=MUTED, font=S_FXS).pack(anchor="w", pady=(2,0))


        # ── AI警告メッセージ ──
        sec("AI警告メッセージ")
        wf = tk.Frame(f, bg=SURFACE); wf.pack(fill=tk.X, **pad)
        tk.Label(wf, text="{user} = ユーザー名に置換", bg=SURFACE, fg=MUTED, font=S_FXS).pack(anchor="w", pady=(0,4))
        var_wmsg = tk.StringVar(value=self._cfg.get("ai_warn_msg", "[AI警告] @{user} コメントが有害と判断されました。"))
        tk.Entry(wf, textvariable=var_wmsg, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                 font=S_FN, relief=tk.FLAT, highlightbackground=BORDER,
                 highlightthickness=1, highlightcolor=TWITCH).pack(fill=tk.X, ipady=5)

        # ── フィードバック統計 ──
        sec("学習データ統計")
        lf = tk.Frame(f, bg=SURFACE); lf.pack(fill=tk.X, **pad)
        stats = self._ai_engine.get_feedback_stats() if self._ai_engine else {"total":0,"correct":0,"accuracy":0}
        tk.Label(lf, text=f"フィードバック件数: {stats['total']}件  正解率: {stats['accuracy']}%",
                 bg=SURFACE, fg=TEXT, font=S_FN).pack(anchor="w", pady=2)

        def export_csv():
            from tkinter import filedialog as fd
            path = fd.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV","*.csv")],
                initialfile=f"aegismod-training-{datetime.now().strftime('%Y%m%d')}.csv")
            if path and self._ai_engine:
                result = self._ai_engine.export_training_csv(path)
                if result:
                    messagebox.showinfo("完了", f"CSVエクスポート完了:\n{result}")
                else:
                    messagebox.showinfo("情報", "フィードバックデータがありません")

        tk.Button(lf, text="学習データをCSVエクスポート", command=export_csv,
                   bg=SURFACE3, fg=TWITCH_L, font=S_FN, relief=tk.FLAT,
                   cursor="hand2", activebackground=SURFACE2).pack(anchor="w", ipady=4, ipadx=8, pady=4)

        # ── ランキング ──
        sec("酷すぎるコメント ランキング")
        rf = tk.Frame(f, bg=SURFACE); rf.pack(fill=tk.X, **pad)

        rank_frame = tk.Frame(rf, bg=SURFACE2)
        rank_frame.pack(fill=tk.X, pady=4)

        def refresh_ranking():
            for w in rank_frame.winfo_children():
                w.destroy()
            if not self._ai_engine:
                tk.Label(rank_frame, text="AI無効", bg=SURFACE2, fg=MUTED, font=S_FXS).pack(padx=8, pady=4)
                return
            rows = self._ai_engine.get_ranking(10)
            if not rows:
                tk.Label(rank_frame, text="データなし（配信中にコメントが蓄積されます）",
                         bg=SURFACE2, fg=MUTED, font=S_FXS).pack(padx=8, pady=6)
                return
            # ヘッダー
            hf = tk.Frame(rank_frame, bg=SURFACE3); hf.pack(fill=tk.X)
            for txt, w in [("順位",35),("ユーザー",110),("合計",55),("平均",55),("最高",55),("最悪コメント",160)]:
                tk.Label(hf, text=txt, bg=SURFACE3, fg=MUTED, font=S_FXS,
                         width=w//7, anchor="w").pack(side=tk.LEFT, padx=4, pady=3)
            for i, row in enumerate(rows):
                bg = SURFACE2 if i % 2 == 0 else SURFACE
                rf2 = tk.Frame(rank_frame, bg=bg); rf2.pack(fill=tk.X)
                medal = ["🥇","🥈","🥉"][i] if i < 3 else f"{i+1}."
                fg_c  = [WARN, MUTED, MUTED][i] if i < 3 else MUTED
                for txt, w in [
                    (medal,                      35),
                    (row["uid"][:14],            110),
                    (str(row["total_score"]),     55),
                    (str(row["avg_score"]),       55),
                    (str(row["max_score"]),       55),
                    (row["worst_msg"][:22],      160),
                ]:
                    tk.Label(rf2, text=txt, bg=bg, fg=fg_c if w==35 else TEXT,
                             font=S_FXS, anchor="w").pack(side=tk.LEFT, padx=4, pady=3)

            # MVP
            mvp = self._ai_engine.get_mvp()
            if mvp:
                mf2 = tk.Frame(rank_frame, bg=SURFACE3); mf2.pack(fill=tk.X, pady=(4,0))
                tk.Label(mf2, text=f"最強酷すぎるコメンテーター賞: {mvp['uid']}  (合計スコア: {mvp['total_score']})",
                         bg=SURFACE3, fg=WARN, font=S_FB).pack(padx=8, pady=6)

        refresh_ranking()
        tk.Button(rf, text="ランキングを更新", command=refresh_ranking,
                   bg=SURFACE3, fg=TWITCH_L, font=S_FN, relief=tk.FLAT,
                   cursor="hand2", activebackground=SURFACE2).pack(anchor="w", ipady=4, ipadx=8, pady=(4,0))

        # ── セッションリセット ──
        sec("セッション")
        ssf = tk.Frame(f, bg=SURFACE); ssf.pack(fill=tk.X, **pad)
        def clear_session():
            if messagebox.askyesno("確認", "このセッションのAIスコア履歴をリセットしますか？"):
                if self._ai_engine: self._ai_engine.clear_session()
                refresh_ranking()
                self.log("[AI] セッションスコアをリセットしました", "info")
        tk.Button(ssf, text="セッションスコアをリセット", command=clear_session,
                   bg=SURFACE3, fg=DANGER, font=S_FN, relief=tk.FLAT,
                   cursor="hand2").pack(anchor="w", ipady=4, ipadx=8)

        # ── 保存 ──
        bf = tk.Frame(f, bg=SURFACE); bf.pack(fill=tk.X, padx=12, pady=(12,16))
        def save_ai():
            self._cfg["ai_enabled"]         = var_en.get()
            self._cfg["ai_score_threshold"] = var_thresh.get()
            self._cfg["ai_warn_msg"]        = var_wmsg.get()
            self._save_fields()
            win.destroy()
        tk.Button(bf, text="設定を保存して閉じる", command=save_ai,
                   bg=TWITCH, fg="white", font=S_FB, relief=tk.FLAT,
                   cursor="hand2", activebackground=TWITCH_L).pack(fill=tk.X, ipady=7)

    def _on_close(self):
        self._cfg["main_geometry"] = self.geometry()
        self._save_fields(); self.irc.disconnect()
        if self._tray: self._tray.stop()
        self.destroy()

# ─────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
