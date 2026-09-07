# -*- coding: utf-8 -*-
"""串口工具：设备号–ICCID–卡号自动提取与进度看板

数据源：串口直连（pyserial）/ 导入日志 TXT / 实时监听日志 TXT
解析规则：cur device_num 与 iccid 相邻 ≤10 行配对（经 SaveWindows 日志回归验证）
运行环境：C:\Python311（依赖 tkinter、openpyxl、pyserial）
"""
import bisect
import csv
import datetime
import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import urllib.request
import webbrowser
from tkinter import ttk, filedialog, messagebox

import openpyxl

try:
    import serial
    from serial.tools import list_ports
except ImportError:
    serial = None

APP_VERSION = "1.1.0"
GITEE_REPO = "darling5/serial-card-tool"
GITHUB_REPO = "darling5/serial-card-tool"

RESERVED_FIELDS = {"idx", "ts", "dev", "iccid", "card", "status", "device_num"}
FIXED_COLS = ["idx", "ts", "dev", "iccid", "card", "status"]
COL_LABELS = {"idx": "序号", "ts": "时间", "dev": "cur device_num", "iccid": "iccid",
              "card": "卡号", "status": "状态"}
COL_WIDTHS = {"idx": 50, "ts": 90, "dev": 130, "iccid": 190, "card": 130, "status": 130}

# 打包后 __file__ 指向临时解压目录，存档须落在 exe 旁边
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(APP_DIR, "serial_card_tool_config.json")


def load_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
        return (cfg.get("custom_fields") or [], cfg.get("column_order") or [])
    except Exception:
        return [], []


def save_config(fields, order):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"custom_fields": fields, "column_order": order},
                      f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ---------------- 解析引擎（规则与离线验证版 extract_pair.py 逐条一致） ----------------

ICCID_RE = re.compile(r"iccid[:：]\s*([0-9A-Fa-fxX]{20})\b")
DEV_RE = re.compile(r"(?:device_num|_num)[:：]\s*([0-9]+)")
DEV_EMPTY_RE = re.compile(r"cur device_num[:：]\s*$")
TAIL_RE = re.compile(r"([0-9]{10,})\s*$")
TS_RE = re.compile(r"\[(\d{2}:\d{2}:\d{2}\.\d{3})\]")

PAIR_WINDOW = 10
DEV_PREFIX = "785"


class Record:
    __slots__ = ("dev", "iccid", "dev_line", "iccid_line", "ts", "fields")

    def __init__(self, dev, iccid, dev_line, iccid_line, ts, fields=None):
        self.dev = dev
        self.iccid = iccid
        self.dev_line = dev_line
        self.iccid_line = iccid_line
        self.ts = ts
        self.fields = fields or {}


class ParserEngine:
    """增量提取（devices/iccids/自定义字段随行缓存）+ 全量重配对（bisect 窗口）。
    device_num 与 iccid 双向就近配对（≤10 行，与离线验证版一致）；
    自定义字段按"字段名:值"归属其后最近的设备行（≤10 行）。"""

    def __init__(self):
        self.lines = []
        self.devices = []          # (行号, 设备号)，按行号有序
        self.iccids = []           # (行号, iccid)，按行号有序
        self._pending = None       # (行号, 部分值 or None)，等下一行做续行合并
        self.custom_fields = []    # 自定义字段名列表
        self._field_res = {}       # 字段名 -> 编译正则
        self.field_hits = {}       # 字段名 -> [(行号, 值)]
        self._records = []
        self._key = None
        self.version = 0

    def set_fields(self, names):
        """设置自定义字段集（去重、过滤保留名），并对已有行重扫。"""
        names = [n for n in dict.fromkeys(names) if n and n not in RESERVED_FIELDS]
        self.custom_fields = names
        self._field_res = {n: re.compile(re.escape(n) + r"[:：]\s*([0-9A-Za-z_.\-]+)")
                           for n in names}
        self.field_hits = {n: [] for n in names}
        for i, line in enumerate(self.lines):
            self._scan_fields(i, line)
        self._key = None
        return self.repair()

    def _scan_fields(self, i, line):
        hit = False
        for n, rx in self._field_res.items():
            m = rx.search(line)
            if m:
                self.field_hits[n].append((i, m.group(1)))
                hit = True
        return hit

    def feed_lines(self, lines):
        trig = False
        for line in lines:
            if self._feed_one(line):
                trig = True
        if trig:
            return self.repair()
        return False

    def flush(self):
        """流结束（导入完成）时结算未完成的续行设备。"""
        if self._pending is None:
            return False
        pidx, pval = self._pending
        self._pending = None
        if pval is not None and pval.startswith(DEV_PREFIX) and len(pval) >= 11:
            self.devices.append((pidx, pval))
        return self.repair()

    def records(self):
        return self._records

    def _feed_one(self, line):
        idx = len(self.lines)
        self.lines.append(line)
        trig = False
        if self._pending is not None:
            pidx, pval = self._pending
            self._pending = None
            if not DEV_RE.search(line):
                tm = TAIL_RE.search(line)
                if tm:
                    pval = tm.group(1) if pval is None else pval + tm.group(1)
            if pval is not None and pval.startswith(DEV_PREFIX) and len(pval) >= 11:
                self.devices.append((pidx, pval))
            trig = True
        m = ICCID_RE.search(line)
        if m:
            self.iccids.append((idx, m.group(1)))
            trig = True
        m = DEV_RE.search(line)
        if m:
            val = m.group(1)
            if len(val) < 12:
                self._pending = (idx, val)
            elif val.startswith(DEV_PREFIX) and len(val) >= 11:
                self.devices.append((idx, val))
            trig = True
        elif DEV_EMPTY_RE.search(line.strip()):
            self._pending = (idx, None)
            trig = True
        if self._field_res and self._scan_fields(idx, line):
            trig = True
        return trig

    def repair(self):
        icc_lines = [ii for ii, _ in self.iccids]
        pairs = {}
        for di, dv in self.devices:
            lo = bisect.bisect_left(icc_lines, di - PAIR_WINDOW)
            hi = bisect.bisect_right(icc_lines, di + PAIR_WINDOW)
            best = None
            for k in range(lo, hi):
                ii = icc_lines[k]
                if best is None or abs(ii - di) < abs(icc_lines[best] - di):
                    best = k
            if best is not None:
                pairs.setdefault((dv, self.iccids[best][1]),
                                 (di, self.iccids[best][0]))

        dev_map = {}
        for (dv, iv), (di, ii) in pairs.items():
            dev_map.setdefault(dv, []).append((iv, di, ii))
        for di, dv in self.devices:
            if dv not in dev_map:
                dev_map[dv] = [(None, di, None)]

        # 自定义字段：值归属其后最近的设备行（AT 块内字段均在 device_num 之后，≤10 行）
        dev_lines = [di for di, _ in self.devices]
        field_vals = {}
        for n in self.custom_fields:
            for h, v in self.field_hits[n]:
                pos = bisect.bisect_right(dev_lines, h) - 1
                if pos < 0 or h - dev_lines[pos] > PAIR_WINDOW:
                    continue
                field_vals.setdefault((n, dev_lines[pos]), v)

        records = []
        for dv, items in dev_map.items():
            first = min(di for _, di, _ in items)
            ts = self._ts_before(first)
            for iv, di, ii in items:
                fields = {n: field_vals.get((n, di), "") for n in self.custom_fields}
                records.append(Record(dv, iv, di, ii, ts, fields))
        records.sort(key=lambda r: (r.dev_line, r.iccid_line if r.iccid_line is not None else 0))

        key = [(r.dev, r.iccid, r.dev_line, r.iccid_line,
                tuple(r.fields.get(n, "") for n in self.custom_fields)) for r in records]
        if key != self._key:
            self._records = records
            self._key = key
            self.version += 1
            return True
        return False

    def _ts_before(self, idx):
        for j in range(idx, max(-1, idx - 300), -1):
            m = TS_RE.search(self.lines[j])
            if m:
                return m.group(1)
        return ""


# ---------------- 卡号表加载 ----------------

def load_card_table(path):
    """支持 Excel（表头含 ICCID/卡号，或按 20 位值嗅探）与 CSV，返回 iccid -> 卡号。"""
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext in (".xlsx", ".xlsm", ".xls"):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                rows.append(["" if c is None else str(c).strip() for c in row])
        wb.close()
    else:
        for enc in ("utf-8-sig", "gbk"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    rows = [[c.strip() for c in r] for r in csv.reader(f)]
                break
            except UnicodeDecodeError:
                continue

    iccid_col = card_col = None
    for r in rows[:10]:
        for ci, c in enumerate(r):
            if c and "iccid" in c.lower():
                iccid_col = ci
            if c and "卡号" in c:
                card_col = ci
    if iccid_col is None:
        for r in rows:
            for ci, c in enumerate(r):
                if len(c) == 20 and all(ch in "0123456789abcdefABCDEFxX" for ch in c):
                    iccid_col = ci
                    for cj, cc in enumerate(r):
                        if cj != ci and cc and cc.replace(".", "").isdigit():
                            card_col = cj
                            break
                    break
            if iccid_col is not None:
                break
    if iccid_col is None:
        raise ValueError("未找到 ICCID 列（需表头含 ICCID 或出现 20 位卡值）")

    card_map = {}
    for r in rows:
        if len(r) <= iccid_col:
            continue
        icc = r[iccid_col]
        if len(icc) != 20:
            continue
        card = r[card_col] if (card_col is not None and len(r) > card_col) else ""
        card_map[icc] = card
    if not card_map:
        raise ValueError("卡号表为空或格式不识别")
    return card_map


# ---------------- 数据源线程 ----------------

class SerialReader(threading.Thread):
    def __init__(self, port, baud, out_q, archive_path=None):
        super().__init__(daemon=True)
        self.port = port
        self.baud = baud
        self.q = out_q
        self.archive_path = archive_path
        self.stop_evt = threading.Event()

    def stop(self):
        self.stop_evt.set()

    def run(self):
        ser = None
        f = None
        buf = b""
        try:
            ser = serial.Serial(self.port, self.baud, timeout=0.2)
            if self.archive_path:
                f = open(self.archive_path, "ab")
            self.q.put(("status", "串口 %s 已连接 @ %d" % (self.port, self.baud)))
            while not self.stop_evt.is_set():
                n = ser.in_waiting
                data = ser.read(n if n else 1)
                if not data:
                    continue
                if f:
                    f.write(data)
                    f.flush()
                buf += data
                while True:
                    nl = buf.find(b"\n")
                    cr = buf.find(b"\r")
                    if nl == -1 and cr == -1:
                        break
                    if cr == -1 or (nl != -1 and nl < cr):
                        idx, ln = nl, 1
                    else:
                        idx, ln = cr, 1
                        if buf[idx + 1: idx + 2] == b"\n":
                            ln = 2
                    self.q.put(("line", buf[:idx].decode("utf-8", errors="replace")))
                    buf = buf[idx + ln:]
        except Exception as e:
            self.q.put(("status", "串口错误：%s" % e))
        finally:
            if buf.strip():
                self.q.put(("line", buf.decode("utf-8", errors="replace")))
            if f:
                f.close()
            if ser:
                try:
                    ser.close()
                except Exception:
                    pass
            self.q.put(("status", "串口已断开"))


class FileImporter(threading.Thread):
    def __init__(self, path, out_q):
        super().__init__(daemon=True)
        self.path = path
        self.q = out_q

    def run(self):
        total = 0
        try:
            with open(self.path, encoding="utf-8", errors="replace") as f:
                batch = []
                for line in f:
                    batch.append(line.rstrip("\r\n"))
                    total += 1
                    if len(batch) >= 2000:
                        self.q.put(("lines", batch))
                        self.q.put(("status", "已读取 %d 行…" % total))
                        batch = []
                if batch:
                    self.q.put(("lines", batch))
            self.q.put(("status", "导入完成：%s（共 %d 行）" % (os.path.basename(self.path), total)))
        except Exception as e:
            self.q.put(("status", "读取失败：%s" % e))


class FileTail(threading.Thread):
    def __init__(self, path, out_q):
        super().__init__(daemon=True)
        self.path = path
        self.q = out_q
        self.stop_evt = threading.Event()

    def stop(self):
        self.stop_evt.set()

    def run(self):
        f = None
        offset = 0
        buf = b""
        try:
            f = open(self.path, "rb")
            self.q.put(("status", "正在监听：%s（先解析已有内容）" % os.path.basename(self.path)))
            while not self.stop_evt.is_set():
                try:
                    size = os.path.getsize(self.path)
                except OSError:
                    time.sleep(1)
                    continue
                if size < offset:
                    offset = 0
                    buf = b""
                    f.close()
                    f = open(self.path, "rb")
                if size > offset:
                    f.seek(offset)
                    buf += f.read(size - offset)
                    offset = size
                    while True:
                        idx = buf.find(b"\n")
                        if idx == -1:
                            break
                        self.q.put(("line", buf[:idx].rstrip(b"\r").decode("utf-8", errors="replace")))
                        buf = buf[idx + 1:]
                time.sleep(1)
        except Exception as e:
            self.q.put(("status", "监听失败：%s" % e))
        finally:
            if f:
                f.close()
            self.q.put(("status", "已停止监听"))


# ---------------- 版本更新检查（Gitee 优先，GitHub 回退） ----------------

def _ver_tuple(v):
    parts = []
    for seg in v.split("."):
        num = ""
        for ch in seg:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    return tuple(parts)


def fetch_latest_version():
    """返回 (最新版本号, 下载页URL) 或 None（两平台都失败）。"""
    sources = [
        ("https://gitee.com/api/v5/repos/%s/releases/latest" % GITEE_REPO,
         "https://gitee.com/%s/releases/latest" % GITEE_REPO),
        ("https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO,
         "https://github.com/%s/releases/latest" % GITHUB_REPO),
    ]
    for api, page in sources:
        try:
            req = urllib.request.Request(api, headers={"User-Agent": "SerialCardTool"})
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read().decode("utf-8"))
            tag = (data.get("tag_name") or "").lstrip("vV")
            if tag:
                return tag, page
        except Exception:
            continue
    return None


# ---------------- 字段与列设置对话框 ----------------

class FieldSettingsDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("字段与列设置")
        self.transient(app.root)
        self.resizable(False, False)
        self._fields = list(app.custom_fields)
        self._order = list(app.column_order)

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)

        left = ttk.LabelFrame(body, text=" 自定义提取字段（日志中“字段名:值”的前缀） ", padding=8)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.field_list = tk.Listbox(left, width=30, height=12, exportselection=False)
        self.field_list.pack(fill="both", expand=True)
        for f in self._fields:
            self.field_list.insert("end", f)
        row = ttk.Frame(left)
        row.pack(fill="x", pady=(6, 0))
        self.name_var = tk.StringVar()
        ent = ttk.Entry(row, textvariable=self.name_var)
        ent.pack(side="left", fill="x", expand=True, padx=(0, 4))
        ent.bind("<Return>", lambda e: self._add_field())
        ent.focus_set()
        ttk.Button(row, text="添加", width=6, command=self._add_field).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="删除", width=6, command=self._del_field).pack(side="left")

        right = ttk.LabelFrame(body, text=" 表格列顺序（上移/下移/删除列） ", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        self.col_list = tk.Listbox(right, width=30, height=12, exportselection=False)
        self.col_list.pack(fill="both", expand=True)
        row = ttk.Frame(right)
        row.pack(fill="x", pady=(6, 0))
        ttk.Button(row, text="上移", width=6,
                   command=lambda: self._move(-1)).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="下移", width=6,
                   command=lambda: self._move(1)).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="删除列", width=8,
                   command=self._del_col).pack(side="left")
        self._refresh_cols()

        btns = ttk.Frame(body)
        btns.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="恢复默认列",
                   command=self._reset_default).pack(side="left", padx=4)
        ttk.Button(btns, text="应用", command=self._apply).pack(side="left", padx=4)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left")

        self.grab_set()

    def _refresh_cols(self):
        self.col_list.delete(0, "end")
        for k in self._order:
            self.col_list.insert("end", COL_LABELS.get(k, k))

    def _add_field(self):
        name = self.name_var.get().strip()
        if not name:
            return
        if name in RESERVED_FIELDS:
            messagebox.showwarning("提示", "“%s”是内置字段，无需添加" % name, parent=self)
            return
        if any(ch in name for ch in ":："):
            messagebox.showwarning("提示", "字段名不能包含冒号", parent=self)
            return
        if name in self._fields:
            self.name_var.set("")
            return
        self._fields.append(name)
        self.field_list.insert("end", name)
        self.field_list.see("end")
        self.name_var.set("")
        if "status" in self._order:
            self._order.insert(self._order.index("status"), name)
        else:
            self._order.append(name)
        self._refresh_cols()

    def _del_field(self):
        sel = self.field_list.curselection()
        if not sel:
            return
        name = self.field_list.get(sel[0])
        self.field_list.delete(sel[0])
        self._fields.remove(name)
        if name in self._order:
            self._order.remove(name)
        self._refresh_cols()

    def _move(self, d):
        sel = self.col_list.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + d
        if not (0 <= j < len(self._order)):
            return
        self._order[i], self._order[j] = self._order[j], self._order[i]
        self._refresh_cols()
        self.col_list.selection_set(j)
        self.col_list.see(j)

    def _del_col(self):
        sel = self.col_list.curselection()
        if not sel:
            return
        if len(self._order) <= 1:
            messagebox.showwarning("提示", "至少需要保留一列", parent=self)
            return
        i = sel[0]
        key = self._order.pop(i)
        if key in self._fields:
            self._fields.remove(key)
            self.field_list.delete(0, "end")
            for f in self._fields:
                self.field_list.insert("end", f)
        self._refresh_cols()
        if self._order:
            j = min(i, len(self._order) - 1)
            self.col_list.selection_set(j)
            self.col_list.see(j)

    def _reset_default(self):
        self._order = ["idx", "ts", "dev", "iccid", "card"] + self._fields + ["status"]
        self._refresh_cols()

    def _apply(self):
        if not self._order:
            self._order = ["idx", "ts", "dev", "iccid", "card"] + self._fields + ["status"]
        save_config(self._fields, self._order)
        self.app._apply_field_settings(self._fields, self._order)
        self.destroy()


# ---------------- GUI ----------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("串口工具 · 设备号-ICCID-卡号提取")
        root.geometry("1200x800")

        fields, order = load_config()
        self.custom_fields = [f for f in fields if f and f not in RESERVED_FIELDS]
        self.engine = ParserEngine()
        self.column_order = self._merge_order(order)
        self.engine.set_fields(self.custom_fields)
        self.card_map = {}
        self.card_ver = 0
        self.exclusions = set()
        self.line_q = queue.Queue()
        self.serial_thread = None
        self.tail_thread = None
        self._status_text = "就绪"
        self._shown_state = None

        self._build_toolbar()
        self._build_middle()
        self._build_table()
        self._build_bottom()

        self.statusbar.config(text="解析行数 0 ｜ 就绪")
        root.after(200, self._poll)

    # ---- 布局 ----

    def _build_toolbar(self):
        bar = ttk.Frame(self.root, padding=(8, 6))
        bar.grid(row=0, column=0, sticky="ew")

        ttk.Label(bar, text="串口:").pack(side="left")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(bar, textvariable=self.port_var, width=8, state="readonly")
        self.port_combo.pack(side="left", padx=(2, 4))
        ttk.Button(bar, text="刷新", width=5, command=self._refresh_ports).pack(side="left", padx=(0, 8))

        ttk.Label(bar, text="波特率:").pack(side="left")
        self.baud_var = tk.StringVar(value="115200")
        ttk.Combobox(bar, textvariable=self.baud_var, width=8, state="readonly",
                     values=["9600", "19200", "38400", "57600", "115200"]).pack(side="left", padx=(2, 8))

        self.serial_btn = ttk.Button(bar, text="连接", width=8, command=self._toggle_serial)
        self.serial_btn.pack(side="left", padx=(0, 10))

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4)
        ttk.Button(bar, text="打开日志文件…", command=self._open_log).pack(side="left", padx=4)
        self.tail_btn = ttk.Button(bar, text="实时监听日志…", command=self._toggle_tail)
        self.tail_btn.pack(side="left", padx=4)

        self.archive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="原始数据存档", variable=self.archive_var).pack(side="left", padx=10)

        self.update_btn = ttk.Button(bar, text="检查版本更新", command=self._check_update)
        self.update_btn.pack(side="right", padx=(10, 0))

        self._refresh_ports()

    def _build_middle(self):
        mid = ttk.Frame(self.root, padding=(8, 4))
        mid.grid(row=1, column=0, sticky="ew")
        mid.columnconfigure(1, weight=1)

        # 左：配置
        cfg = ttk.LabelFrame(mid, text=" 配置 ", padding=8)
        cfg.grid(row=0, column=0, sticky="ns", padx=(0, 8))

        row = ttk.Frame(cfg)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="本次总数:").pack(side="left")
        self.total_var = tk.StringVar(value="500")
        ttk.Spinbox(row, from_=1, to=100000, textvariable=self.total_var, width=8).pack(side="left", padx=6)

        row = ttk.Frame(cfg)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="起始设备号:").pack(side="left")
        self.start_var = tk.StringVar(value="785000060001")
        ttk.Entry(row, textvariable=self.start_var, width=16).pack(side="left", padx=6)

        ttk.Button(cfg, text="导入卡号表（Excel/CSV）…", command=self._import_cards).pack(fill="x", pady=(8, 2))
        self.card_label = ttk.Label(cfg, text="已加载卡号表：未加载", foreground="#666")
        self.card_label.pack(anchor="w", pady=(0, 6))

        ttk.Label(cfg, text="排除 ICCID（每行一个，可粘贴）:").pack(anchor="w")
        self.excl_text = tk.Text(cfg, width=34, height=8, font=("Consolas", 10))
        self.excl_text.pack(fill="x", pady=2)
        self.excl_label = ttk.Label(cfg, text="已识别排除：0 个", foreground="#666")
        self.excl_label.pack(anchor="w", pady=(0, 6))

        ttk.Button(cfg, text="字段与列设置…",
                   command=self._open_field_settings).pack(fill="x", pady=(0, 6))
        ttk.Button(cfg, text="清空已解析数据", command=self._clear_data).pack(fill="x")

        # 右：进度看板
        panel = ttk.LabelFrame(mid, text=" 进度看板 ", padding=10)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.columnconfigure(0, weight=1)
        panel.columnconfigure(1, weight=1)
        panel.columnconfigure(2, weight=1)

        self.big_label = tk.Label(panel, text="已读 0 / 500",
                                  font=("Microsoft YaHei", 24, "bold"), fg="#2f5bea")
        self.big_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        self.progress = ttk.Progressbar(panel, maximum=100, value=0)
        self.progress.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        self.stat_vars = {}
        items = [
            ("remaining", "剩余未读取", "#d9433b"), ("paired", "配对成功（设备）", "#1a7f37"),
            ("matched", "匹配卡号", "#1a7f37"), ("fail", "读取失败", "#b00"),
            ("swap", "换卡设备", "#9a6700"), ("nonpool", "非池内卡", "#9a6700"),
            ("missing_dev", "范围内缺失设备", "#b00"), ("missing_card", "池内缺失卡号", "#b00"),
        ]
        for i, (key, name, color) in enumerate(items):
            r, c = divmod(i, 2)
            cell = ttk.Frame(panel)
            cell.grid(row=2 + r, column=c, sticky="w", padx=(0, 20), pady=3)
            var = tk.StringVar(value="0")
            ttk.Label(cell, text=name + "：").pack(side="left")
            tk.Label(cell, textvariable=var, font=("Microsoft YaHei", 13, "bold"),
                     fg=color).pack(side="left")
            self.stat_vars[key] = var

    def _build_table(self):
        frame = ttk.LabelFrame(self.root, text=" 明细（双击单元格复制） ", padding=4)
        frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=4)
        self.root.rowconfigure(2, weight=1)
        self.root.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.table_frame = frame
        self._rebuild_table()

    def _rebuild_table(self):
        for w in self.table_frame.winfo_children():
            w.destroy()
        cols = tuple(self.column_order)
        self.tree = ttk.Treeview(self.table_frame, columns=cols, show="headings")
        for cid in cols:
            self.tree.heading(cid, text=COL_LABELS.get(cid, cid))
            self.tree.column(cid, width=COL_WIDTHS.get(cid, 120), anchor="w",
                             stretch=(cid in ("iccid", "status") or cid in self.custom_fields))
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for tag, bg, fg in [("fail", "#ffe5e5", "#b00000"), ("excl", "#f3eaff", "#6a4bd6"),
                            ("oldcard", "#fff6e0", "#9a6700"), ("swap", "#fffbe6", "#9a6700")]:
            self.tree.tag_configure(tag, background=bg, foreground=fg)

        self.tree.bind("<Double-1>", self._copy_cell)

    def _build_bottom(self):
        bar = ttk.Frame(self.root, padding=(8, 4))
        bar.grid(row=3, column=0, sticky="ew")
        ttk.Button(bar, text="导出明细 CSV", command=self._export_detail).pack(side="left", padx=4)
        ttk.Button(bar, text="导出缺失设备 CSV", command=self._export_missing_devices).pack(side="left", padx=4)
        ttk.Button(bar, text="导出缺失卡号 CSV", command=self._export_missing_cards).pack(side="left", padx=4)

        bottom = ttk.Frame(self.root)
        bottom.grid(row=4, column=0, sticky="ew")
        self.statusbar = ttk.Label(bottom, anchor="w", relief="sunken", padding=(8, 3))
        self.statusbar.pack(side="left", fill="x", expand=True)
        ttk.Label(bottom, text="智环未来(深圳)科技有限公司", anchor="e",
                  relief="sunken", padding=(8, 3), foreground="#666"
                  ).pack(side="right")

    # ---- 数据源控制 ----

    def _check_update(self):
        self.update_btn.config(state="disabled")
        self._status_text = "正在检查版本更新…"

        def worker():
            result = fetch_latest_version()
            self.root.after(0, lambda: self._update_result(result))

        threading.Thread(target=worker, daemon=True).start()

    def _update_result(self, result):
        try:
            self.update_btn.config(state="normal")
        except tk.TclError:
            return
        if result is None:
            messagebox.showwarning(
                "检查更新", "无法获取版本信息（Gitee 与 GitHub 均未响应），\n请检查网络后重试。")
            self._status_text = "版本检查失败"
            return
        tag, page = result
        if _ver_tuple(tag) > _ver_tuple(APP_VERSION):
            self._status_text = "发现新版本 v%s" % tag
            if messagebox.askyesno(
                    "发现新版本",
                    "当前版本 v%s，最新版本 v%s。\n\n是否前往下载？" % (APP_VERSION, tag)):
                webbrowser.open(page)
        else:
            self._status_text = "已是最新版本 v%s" % APP_VERSION
            messagebox.showinfo("检查更新", "当前已是最新版本（v%s）。" % APP_VERSION)

    def _refresh_ports(self):
        if serial is None:
            return
        ports = [p.device for p in list_ports.comports()]
        self.port_combo["values"] = ports
        if ports and not self.port_var.get():
            self.port_var.set(ports[0])
        elif not ports:
            self.port_var.set("")

    def _toggle_serial(self):
        if self.serial_thread and self.serial_thread.is_alive():
            self.serial_thread.stop()
            self.serial_btn.config(text="连接")
            return
        if serial is None:
            messagebox.showerror("缺少依赖",
                                 "未安装 pyserial，请运行：\nC:\\Python311\\python.exe -m pip install pyserial")
            return
        port = self.port_var.get()
        if not port:
            messagebox.showwarning("提示", "未检测到串口或未选择串口")
            return
        archive = None
        if self.archive_var.get():
            d = os.path.join(APP_DIR, "serial_log")
            os.makedirs(d, exist_ok=True)
            archive = os.path.join(
                d, datetime.datetime.now().strftime("serial_%Y%m%d_%H%M%S.txt"))
        try:
            self.serial_thread = SerialReader(port, int(self.baud_var.get()),
                                              self.line_q, archive)
        except ValueError:
            messagebox.showwarning("提示", "波特率无效")
            return
        self.serial_thread.start()
        self.serial_btn.config(text="断开")

    def _open_log(self):
        path = filedialog.askopenfilename(
            title="打开串口日志文件",
            filetypes=[("日志文件", "*.txt *.log"), ("所有文件", "*.*")])
        if not path:
            return
        FileImporter(path, self.line_q).start()

    def _toggle_tail(self):
        if self.tail_thread and self.tail_thread.is_alive():
            self.tail_thread.stop()
            self.tail_btn.config(text="实时监听日志…")
            return
        path = filedialog.askopenfilename(
            title="选择要监听的日志文件（先解析已有内容，再实时跟踪）",
            filetypes=[("日志文件", "*.txt *.log"), ("所有文件", "*.*")])
        if not path:
            return
        self.tail_thread = FileTail(path, self.line_q)
        self.tail_thread.start()
        self.tail_btn.config(text="停止监听")

    # ---- 卡号表 / 排除 ----

    def _import_cards(self):
        path = filedialog.askopenfilename(
            title="导入卡号与 ICCID 匹配表",
            filetypes=[("卡号表", "*.xlsx *.csv"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            self.card_map = load_card_table(path)
            self.card_ver += 1
            self.card_label.config(
                text="已加载卡号表：%d 张（%s）" % (len(self.card_map), os.path.basename(path)),
                foreground="#1a7f37")
            self._shown_state = None
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def _parse_exclusions(self):
        excl = set()
        for ln in self.excl_text.get("1.0", "end").splitlines():
            v = ln.strip()
            if len(v) > 20:
                v = v[:20]
            if v:
                excl.add(v)
        return excl

    def _clear_data(self):
        self.engine = ParserEngine()
        self.engine.set_fields(self.custom_fields)
        self._shown_state = None
        self.tree.delete(*self.tree.get_children())
        self._status_text = "已清空"
        self._update_stats()

    # ---- 字段与列设置 ----

    def _merge_order(self, saved):
        valid = set(FIXED_COLS) | set(self.custom_fields)
        order, seen = [], set()
        for k in saved:
            if k in valid and k not in seen:
                order.append(k)
                seen.add(k)
        if not order:
            order = ["idx", "ts", "dev", "iccid", "card"] + self.custom_fields + ["status"]
        return order

    def _open_field_settings(self):
        FieldSettingsDialog(self)

    def _apply_field_settings(self, fields, order):
        self.custom_fields = list(fields)
        self.column_order = list(order)
        self.engine.set_fields(self.custom_fields)
        self._rebuild_table()
        self._shown_state = None
        self._status_text = "字段设置已应用（%d 个自定义字段）" % len(self.custom_fields)

    # ---- 主循环 ----

    def _poll(self):
        try:
            lines = []
            status = []
            try:
                while True:
                    kind, payload = self.line_q.get_nowait()
                    if kind == "line":
                        lines.append(payload)
                    elif kind == "lines":
                        lines.extend(payload)
                    elif kind == "status":
                        status.append(payload)
                    elif kind == "done":
                        self.engine.flush()
            except queue.Empty:
                pass
            if status:
                self._status_text = " ｜ ".join(status[-2:])

            new_excl = self._parse_exclusions()
            if new_excl != self.exclusions:
                self.exclusions = new_excl
                self.excl_label.config(text="已识别排除：%d 个" % len(self.exclusions))
                self._shown_state = None

            if lines:
                self.engine.feed_lines(lines)

            state = (self.engine.version, self.card_ver, tuple(sorted(self.exclusions)))
            if state != self._shown_state:
                self._refresh_table()
                self._shown_state = state
            self._update_stats()
            self.statusbar.config(text="解析行数 %d ｜ %s" % (len(self.engine.lines), self._status_text))
        except tk.TclError:
            return
        self.root.after(200, self._poll)

    def _classify(self, r, swap_devs):
        if r.iccid is None:
            return "", "读取失败", "fail"
        if r.iccid in self.exclusions:
            return self.card_map.get(r.iccid, ""), "已排除", "excl"
        if r.iccid in self.card_map:
            if swap_devs.get(r.dev, 0) >= 2:
                return self.card_map[r.iccid], "已匹配（换卡）", "swap"
            return self.card_map[r.iccid], "已匹配", "ok"
        return "", "旧卡/非池内", "oldcard"

    def _cell_value(self, key, n, r, card, status):
        if key == "idx":
            return n
        if key == "ts":
            return r.ts or ""
        if key == "dev":
            return r.dev
        if key == "iccid":
            return r.iccid or ""
        if key == "card":
            return card
        if key == "status":
            return status
        return r.fields.get(key, "")

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        swap_devs = {}
        for r in self.engine.records():
            if r.iccid:
                swap_devs.setdefault(r.dev, set()).add(r.iccid)
        swap_devs = {d: len(v) for d, v in swap_devs.items()}
        for n, r in enumerate(self.engine.records(), 1):
            card, status, tag = self._classify(r, swap_devs)
            vals = []
            for key in self.column_order:
                v = self._cell_value(key, n, r, card, status)
                vals.append(v if v != "" else "—")
            self.tree.insert("", "end", iid=str(n), tags=(tag,), values=vals)

    def _update_stats(self):
        recs = self.engine.records()
        devs = {}
        for r in recs:
            devs.setdefault(r.dev, []).append(r)
        read = len(devs)
        try:
            total = int(self.total_var.get())
        except ValueError:
            total = 0
        remaining = max(0, total - read) if total else 0
        paired_devs = sum(1 for rs in devs.values() if any(r.iccid for r in rs))
        fail = read - paired_devs
        swap = sum(1 for rs in devs.values()
                   if len({r.iccid for r in rs if r.iccid}) >= 2)
        matched = {r.iccid for r in recs
                   if r.iccid and r.iccid in self.card_map and r.iccid not in self.exclusions}
        nonpool = {r.iccid for r in recs if r.iccid and r.iccid not in self.card_map}

        self.big_label.config(text="已读 %d / %s" % (read, total if total else "?"))
        self.progress.config(value=(read * 100.0 / total) if total else 0)
        self.stat_vars["remaining"].set(str(remaining))
        self.stat_vars["paired"].set(str(paired_devs))
        self.stat_vars["matched"].set(str(len(matched)))
        self.stat_vars["fail"].set(str(fail))
        self.stat_vars["swap"].set(str(swap))
        self.stat_vars["nonpool"].set(str(len(nonpool)))

        try:
            start = int(self.start_var.get())
        except ValueError:
            start = 0
        if start and total:
            miss_d = sum(1 for d in range(start, start + total) if str(d) not in devs)
            self.stat_vars["missing_dev"].set(str(miss_d))
        else:
            self.stat_vars["missing_dev"].set("—")
        if self.card_map:
            appeared = {r.iccid for r in recs if r.iccid}
            miss_c = sum(1 for icc in self.card_map
                         if icc not in appeared and icc not in self.exclusions)
            self.stat_vars["missing_card"].set(str(miss_c))
        else:
            self.stat_vars["missing_card"].set("—")

    # ---- 导出 ----

    def _export_path(self, default_name):
        return filedialog.asksaveasfilename(defaultextension=".csv",
                                            initialfile=default_name,
                                            filetypes=[("CSV 文件", "*.csv")])

    def _export_detail(self):
        path = self._export_path("设备卡号明细.csv")
        if not path:
            return
        swap_devs = {}
        for r in self.engine.records():
            if r.iccid:
                swap_devs.setdefault(r.dev, set()).add(r.iccid)
        swap_devs = {d: len(v) for d, v in swap_devs.items()}
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow([COL_LABELS.get(k, k) for k in self.column_order])
            for n, r in enumerate(self.engine.records(), 1):
                card, status, _ = self._classify(r, swap_devs)
                w.writerow([self._cell_value(k, n, r, card, status)
                            for k in self.column_order])
        messagebox.showinfo("完成", "已导出：%s" % path)

    def _export_missing_devices(self):
        try:
            start = int(self.start_var.get())
            total = int(self.total_var.get())
        except ValueError:
            messagebox.showwarning("提示", "起始设备号或本次总数无效")
            return
        if not start or not total:
            messagebox.showwarning("提示", "起始设备号或本次总数无效")
            return
        devs = {r.dev for r in self.engine.records()}
        missing = [d for d in range(start, start + total) if str(d) not in devs]
        path = self._export_path("缺失设备.csv")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["cur device_num", "状态"])
            for d in missing:
                w.writerow([d, "缺失（日志中无记录）"])
        messagebox.showinfo("完成", "共 %d 台缺失，已导出：%s" % (len(missing), path))

    def _export_missing_cards(self):
        if not self.card_map:
            messagebox.showwarning("提示", "请先导入卡号表")
            return
        appeared = {r.iccid for r in self.engine.records() if r.iccid}
        rows = []
        for icc, card in self.card_map.items():
            if icc in appeared:
                continue
            if icc in self.exclusions:
                rows.append((card, icc, "已排除（指定ICCID）"))
            else:
                rows.append((card, icc, "缺失（日志中未出现）"))
        rows.sort(key=lambda x: x[0])
        path = self._export_path("缺失卡号.csv")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["卡号(144)", "iccid", "状态"])
            w.writerows(rows)
        messagebox.showinfo("完成", "共 %d 张，已导出：%s" % (len(rows), path))

    # ---- 其他 ----

    def _copy_cell(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if item and col:
            val = self.tree.set(item, col)
            if val and val != "—":
                self.root.clipboard_clear()
                self.root.clipboard_append(val)
                self._status_text = "已复制：%s" % val


def main():
    import sys
    root = tk.Tk()
    App(root)
    if "--smoke" in sys.argv:
        root.withdraw()
        root.after(1500, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
