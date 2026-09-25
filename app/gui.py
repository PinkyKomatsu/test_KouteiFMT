"""tkinter / ttk の GUI（日本語 UI）。

タブ：報告書作成 / フォーマット設定 / 過去事例 / 設定
"""
from __future__ import annotations

import bisect
import os
import re
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import analyzer, cases as cases_mod, generator, storage
from . import textutil as tu
from . import xlsx_writer as xw

APP_TITLE = "Excel報告書作成アシスタント"
FONT_CANDIDATES = ("Yu Gothic UI", "Meiryo UI", "Meiryo", "MS UI Gothic")

COLOR_REUSE = "#fff3b0"   # 流用（要確認）
COLOR_AUTO = "#dff5e1"    # 自動入力
COLOR_LABEL = "#dbe8ff"   # プレビュー：ラベル
COLOR_FIELD = "#fff3b0"   # プレビュー：記入欄
COLOR_DISABLED = "#eeeeee"

MEMO_HINT = ("1行に1項目。「課題：〜」のように欄名を書くとその欄に入ります。"
             "「【今後の対応】」だけの行を書くと、以降の行はその欄に入ります。"
             "表は「内容｜担当｜期限｜状態」の形で書きます。")

DATE_FORMATS = [
    ("auto", "自動（過去事例に合わせる）"),
    ("{Y}/{MM}/{DD}", None),
    ("{Y}/{M}/{D}", None),
    ("{Y}年{M}月{D}日", None),
    ("{Y}年{M}月{D}日（{W}）", None),
    ("{Y}-{MM}-{DD}", None),
]

KIND_BY_LABEL = {v: k for k, v in analyzer.KIND_LABELS.items()}


def pick_font_family(root) -> str:
    families = set(tkfont.families(root))
    for name in FONT_CANDIDATES:
        if name in families:
            return name
    return tkfont.nametofont("TkDefaultFont").actual("family")


class ScrollFrame(ttk.Frame):
    """縦スクロールできるフレーム（self.inner に子ウィジェットを置く）。"""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vbar.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))

    @staticmethod
    def on_wheel(e):
        """App 全体の MouseWheel から呼ぶ。ポインタの下の ScrollFrame をスクロールする。"""
        try:
            w = e.widget.winfo_containing(e.x_root, e.y_root)
        except (tk.TclError, AttributeError, KeyError):
            return
        # テキスト欄・一覧・プレビューの上ではそのウィジェット自身のスクロールに任せる
        if isinstance(w, (tk.Text, ttk.Treeview)) or w is None:
            return
        while w is not None and not isinstance(w, ScrollFrame):
            w = w.master
        if w is not None:
            w.canvas.yview_scroll(int(-e.delta / 120), "units")

    def clear(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self.canvas.yview_moveto(0)


# ---------------------------------------------------------------------------
# 欄ごとの編集ウィジェット
# ---------------------------------------------------------------------------

class TextEditor:
    def __init__(self, parent, lines: int, font):
        self.widget = tk.Text(parent, height=lines, wrap="char", undo=True, font=font,
                              relief="solid", bd=1, padx=4, pady=3)
        self.widget.bind("<<Modified>>", self._on_modified)

    def set(self, text: str, color: str):
        self.widget.delete("1.0", "end")
        self.widget.insert("1.0", text or "")
        self.widget.configure(bg=color)
        self.widget.edit_modified(False)
        self.widget.edit_reset()

    def _on_modified(self, _):
        if self.widget.edit_modified():
            self.widget.configure(bg="white")  # 手で編集したら白に戻す
            self.widget.edit_modified(False)

    def get(self) -> str:
        return self.widget.get("1.0", "end-1c")


class EntryEditor:
    def __init__(self, parent, font):
        self.var = tk.StringVar()
        self.widget = tk.Entry(parent, textvariable=self.var, font=font, relief="solid", bd=1)
        self._setting = False
        self.var.trace_add("write", self._on_write)

    def set(self, text: str, color: str):
        self._setting = True
        self.var.set(text or "")
        self.widget.configure(bg=color)
        self._setting = False

    def _on_write(self, *_):
        if not self._setting:
            self.widget.configure(bg="white")

    def get(self) -> str:
        return self.var.get()


class TableEditor:
    """表の欄（列ごとの field）を Entry のグリッドで編集する。"""

    def __init__(self, parent, columns: list[dict], font):
        self.widget = ttk.Frame(parent)
        self.columns = columns
        n_rows = min(len(c["cells"]) for c in columns)
        self.entries: dict[str, list[EntryEditor]] = {}
        for j, col in enumerate(columns):
            ttk.Label(self.widget, text=col["name"]).grid(row=0, column=j, sticky="w", padx=1)
            self.widget.columnconfigure(j, weight=0 if generator._column_role(col["name"]) == "seq" else 1)
            eds = []
            for i in range(n_rows):
                ed = EntryEditor(self.widget, font)
                width = 4 if generator._column_role(col["name"]) == "seq" else 10
                ed.widget.configure(width=width)
                ed.widget.grid(row=i + 1, column=j, sticky="ew", padx=1, pady=1)
                eds.append(ed)
            self.entries[col["id"]] = eds

    def set_column(self, fid: str, text: str, color: str):
        lines = (text or "").split("\n")
        for i, ed in enumerate(self.entries.get(fid, [])):
            ed.set(lines[i] if i < len(lines) else "", color)

    def get_column(self, fid: str) -> str:
        return "\n".join(ed.get() for ed in self.entries.get(fid, []))


# ---------------------------------------------------------------------------
# 報告書作成タブ
# ---------------------------------------------------------------------------

class CreateTab(ttk.Frame):
    def __init__(self, master, app: "App"):
        super().__init__(master, padding=8)
        self.app = app
        self.hits: list[tuple[float, dict]] = []
        self.checked: set[str] = set()
        self.editors: dict[str, tuple[object, str]] = {}  # field_id -> (editor, 種類)
        self.notes: dict[str, ttk.Label] = {}

        self.guide = ttk.Label(
            self, anchor="center", justify="center", font=(app.font_family, 12),
            text="まずフォーマットを取り込んでください。\n\n"
                 "「フォーマット設定」タブの［Excelフォーマットを取り込む］から始めます。")
        self.body = ttk.Frame(self)

        # 下部：出力
        bottom = ttk.Frame(self.body)
        bottom.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Button(bottom, text="Excelに出力", command=self.export, style="Accent.TButton").pack(side="right")
        self.open_var = tk.BooleanVar(value=app.settings.get("open_after_export", True))
        self.register_var = tk.BooleanVar(value=app.settings.get("register_output", True))
        ttk.Checkbutton(bottom, text="出力後に開く", variable=self.open_var).pack(side="right", padx=8)
        ttk.Checkbutton(bottom, text="過去事例に登録", variable=self.register_var).pack(side="right", padx=8)
        self.status = ttk.Label(bottom, text="", foreground="#555")
        self.status.pack(side="left")

        paned = ttk.PanedWindow(self.body, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        right = ttk.Frame(paned)
        paned.add(left, weight=2)
        paned.add(right, weight=3)

        # 左側：入力と参考事例
        ttk.Label(left, text="トピック").pack(anchor="w")
        self.topic_var = tk.StringVar()
        ttk.Entry(left, textvariable=self.topic_var, font=app.text_font).pack(fill="x")
        ttk.Label(left, text="要点メモ").pack(anchor="w", pady=(8, 0))
        self.memo = tk.Text(left, height=9, wrap="char", undo=True, font=app.text_font, relief="solid", bd=1)
        self.memo.pack(fill="x")
        ttk.Label(left, text=MEMO_HINT, foreground="#666", wraplength=420, justify="left").pack(anchor="w", pady=(2, 0))
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="類似事例を検索", command=self.search).pack(side="left")
        ttk.Button(btns, text="下書きを作成", command=self.generate).pack(side="left", padx=6)

        ttk.Label(left, text="参考事例（☑ の上位3件から書き方を学習します。クリックで切替）").pack(anchor="w")
        tree_frame = ttk.Frame(left)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=("use", "score", "subject"), show="headings",
                                 height=6, selectmode="browse")
        self.tree.heading("use", text="使用")
        self.tree.heading("score", text="類似度")
        self.tree.heading("subject", text="件名")
        self.tree.column("use", width=44, anchor="center", stretch=False)
        self.tree.column("score", width=64, anchor="e", stretch=False)
        self.tree.column("subject", width=260)
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        ttk.Label(left, text="事例のプレビュー").pack(anchor="w", pady=(6, 0))
        self.preview = tk.Text(left, height=9, wrap="char", state="disabled", bg="#f7f7f7",
                               font=app.text_font, relief="solid", bd=1)
        self.preview.pack(fill="both", expand=True)

        # 右側：欄ごとの編集フォーム
        legend = ttk.Frame(right)
        legend.pack(fill="x")
        ttk.Label(legend, text="下書き（自由に編集できます）").pack(side="left")
        tk.Label(legend, text=" 流用・要確認 ", bg=COLOR_REUSE).pack(side="right", padx=2)
        tk.Label(legend, text=" 自動入力 ", bg=COLOR_AUTO).pack(side="right", padx=2)
        self.scroll = ScrollFrame(right)
        self.scroll.pack(fill="both", expand=True, pady=(4, 0))

    # ---- 表示の切り替え ----
    def refresh(self):
        fmt = self.app.fmt
        if fmt is None:
            self.body.pack_forget()
            self.guide.pack(fill="both", expand=True)
            return
        self.guide.pack_forget()
        self.body.pack(fill="both", expand=True)
        self.build_editors()
        self.hits, self.checked = [], set()
        self.tree.delete(*self.tree.get_children())
        self._set_preview("")

    def build_editors(self):
        self.scroll.clear()
        self.editors.clear()
        self.notes.clear()
        fmt = self.app.fmt
        font = self.app.text_font
        done_tables = set()
        for f in fmt["fields"]:
            if not f.get("enabled"):
                continue
            if f["direction"] == "table":
                if f["table_id"] in done_tables:
                    continue
                done_tables.add(f["table_id"])
                columns = [c for c in fmt["fields"] if c.get("table_id") == f["table_id"] and c.get("enabled")]
                box = self._field_box("表：" + "／".join(c["name"] for c in columns),
                                      f"{f['sheet']}!{columns[0]['cells'][0]}〜", "表", f["id"])
                ed = TableEditor(box, columns, font)
                ed.widget.pack(fill="x")
                for c in columns:
                    self.editors[c["id"]] = (ed, "table")
                    self.notes[c["id"]] = self.notes[f["id"]]
                continue
            box = self._field_box(f["name"], f"{f['sheet']}!{f['cells'][0]}",
                                  analyzer.KIND_LABELS.get(f["kind"], f["kind"]), f["id"])
            lines = max(1, min(18, round((f.get("height_px") or 20) / 20)))
            if f["kind"] == "body":
                lines = max(3, lines)
            if lines <= 1:
                ed = EntryEditor(box, font)
            else:
                ed = TextEditor(box, lines, font)
            ed.widget.pack(fill="x")
            self.editors[f["id"]] = (ed, "single")

    def _field_box(self, title, cell, kind_label, fid):
        box = ttk.Frame(self.scroll.inner, padding=(2, 6, 12, 2))
        box.pack(fill="x", expand=True)
        head = ttk.Frame(box)
        head.pack(fill="x")
        ttk.Label(head, text=title, font=self.app.bold_font).pack(side="left")
        ttk.Label(head, text=f"  {cell}  〔{kind_label}〕", foreground="#777").pack(side="left")
        note = ttk.Label(box, text="", foreground="#777", wraplength=560, justify="left")
        note.pack(fill="x")
        self.notes[fid] = note
        return box

    # ---- 類似事例 ----
    def _query(self) -> str:
        return self.topic_var.get() + "\n" + self.memo.get("1.0", "end-1c")

    def search(self):
        if not self.app.cases:
            self.status.configure(text="過去事例がありません（「過去事例」タブで取り込めます）")
            self.hits = []
            self.tree.delete(*self.tree.get_children())
            return
        if not self._query().strip():
            messagebox.showinfo(APP_TITLE, "トピックかメモを入力してから検索してください。")
            return
        self.hits = [h for h in self.app.index.search(self._query(), top=10)]
        self.checked = {c["id"] for s, c in self.hits[:3] if s > 0}
        self.tree.delete(*self.tree.get_children())
        for score, c in self.hits:
            self.tree.insert("", "end", iid=c["id"], values=(
                "☑" if c["id"] in self.checked else "☐", f"{score:.2f}", cases_mod.subject_of(self.app.fmt, c)))
        if self.hits:
            self.tree.selection_set(self.hits[0][1]["id"])
        self.status.configure(text=f"{len(self.hits)}件の事例が見つかりました")

    def _on_tree_click(self, e):
        if self.tree.identify_region(e.x, e.y) != "cell" or self.tree.identify_column(e.x) != "#1":
            return
        iid = self.tree.identify_row(e.y)
        if not iid:
            return
        if iid in self.checked:
            self.checked.discard(iid)
        else:
            self.checked.add(iid)
        self.tree.set(iid, "use", "☑" if iid in self.checked else "☐")

    def _on_tree_select(self, _):
        sel = self.tree.selection()
        if not sel:
            return
        case = next((c for _, c in self.hits if c["id"] == sel[0]), None)
        if case:
            self._set_preview(cases_mod.case_preview(self.app.fmt, case))

    def _set_preview(self, text):
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.configure(state="disabled")

    # ---- 下書き ----
    def generate(self):
        if self.app.fmt is None:
            return
        if not self.topic_var.get().strip() and not self.memo.get("1.0", "end-1c").strip():
            messagebox.showinfo(APP_TITLE, "トピックとメモを入力してください。")
            return
        if not self.hits and self.app.cases:
            self.search()
        refs = [c for _, c in self.hits if c["id"] in self.checked][:3]
        try:
            results = generator.generate(self.app.fmt, self.topic_var.get(), self.memo.get("1.0", "end-1c"),
                                         refs, self.app.cases, self.app.settings)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"下書きの作成に失敗しました。\n{e}")
            return
        highlight = self.app.settings.get("highlight", True)
        for fid, (ed, kind) in self.editors.items():
            r = results.get(fid, {"text": "", "status": "empty", "note": ""})
            color = "white"
            if highlight and r["status"] == "reuse":
                color = COLOR_REUSE
            elif highlight and r["status"] == "auto":
                color = COLOR_AUTO
            if kind == "table":
                ed.set_column(fid, r["text"], color)
            else:
                ed.set(r["text"], color)
            if fid in self.notes:
                self.notes[fid].configure(text=r.get("note", ""))
        n_reuse = sum(1 for r in results.values() if r["status"] == "reuse")
        msg = f"下書きを作成しました（参考事例 {len(refs)}件）"
        if n_reuse:
            msg += f"。黄色の {n_reuse} 欄は過去事例の流用です。内容を確認してください"
        self.status.configure(text=msg)

    def collect(self) -> dict[str, str]:
        out = {}
        for fid, (ed, kind) in self.editors.items():
            out[fid] = ed.get_column(fid) if kind == "table" else ed.get()
        return out

    # ---- 出力 ----
    def export(self):
        fmt = self.app.fmt
        if fmt is None:
            return
        template = storage.template_path(fmt)
        if not template.exists():
            messagebox.showerror(APP_TITLE, "フォーマットの原本が見つかりません。フォーマットを取り込み直してください。")
            return
        settings = self.app.settings
        today = date.today()
        name = generator.make_filename(settings.get("filename_pattern"), today,
                                       self.topic_var.get(), settings.get("reporter", ""))
        initial_dir = settings.get("output_dir") or str(Path.home() / "Documents")
        path = filedialog.asksaveasfilename(
            title="Excelに出力", initialdir=initial_dir, initialfile=name + template.suffix,
            defaultextension=template.suffix, filetypes=[("Excel ブック", "*" + template.suffix)])
        if not path:
            return
        values = generator.to_cell_values(fmt, self.collect())
        try:
            warns = xw.write_cells(template, path, values)
        except PermissionError:
            messagebox.showerror(APP_TITLE, "保存できませんでした。\n同じ名前のファイルを Excel で開いていないか確認してください。")
            return
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"出力に失敗しました。\n{e}")
            return
        msg = f"出力しました：{Path(path).name}"
        if self.register_var.get():
            log = self.app.register_output(path)
            if log:
                msg += "（過去事例に登録）"
        self.status.configure(text=msg)
        if warns:
            messagebox.showwarning(APP_TITLE, "\n".join(warns))
        if self.open_var.get():
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except (AttributeError, OSError) as e:
                messagebox.showwarning(APP_TITLE, f"ファイルを開けませんでした。\n{e}")


# ---------------------------------------------------------------------------
# フォーマット設定タブ
# ---------------------------------------------------------------------------

class FormatTab(ttk.Frame):
    HEAD_W, HEAD_H = 34, 20

    def __init__(self, master, app: "App"):
        super().__init__(master, padding=8)
        self.app = app
        self.layout = None
        self.selected_ref: str | None = None
        self._loading = False
        self._redraw_job = None

        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Button(top, text="Excelフォーマットを取り込む", command=self.import_format).pack(side="left")
        ttk.Button(top, text="このフォーマットを削除", command=self.delete_format).pack(side="left", padx=6)
        ttk.Label(top, text="シート").pack(side="left", padx=(16, 4))
        self.sheet_var = tk.StringVar()
        self.sheet_box = ttk.Combobox(top, textvariable=self.sheet_var, state="readonly", width=20)
        self.sheet_box.pack(side="left")
        self.sheet_box.bind("<<ComboboxSelected>>", lambda e: self.load_layout())
        tk.Label(top, text=" 記入欄 ", bg=COLOR_FIELD).pack(side="right", padx=2)
        tk.Label(top, text=" ラベル ", bg=COLOR_LABEL).pack(side="right", padx=2)

        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, pady=(8, 0))

        # 左：シートのプレビュー
        left = ttk.Frame(paned)
        self.canvas = tk.Canvas(left, bg="white", highlightthickness=0)
        xbar = ttk.Scrollbar(left, orient="horizontal", command=self.canvas.xview)
        ybar = ttk.Scrollbar(left, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.sel_label = ttk.Label(left, text="セルをクリックして選択", foreground="#555")
        self.sel_label.grid(row=2, column=0, sticky="w", pady=(4, 0))

        # 右：記入欄の一覧と編集
        right = ttk.Frame(paned, padding=(8, 0, 0, 0))
        paned.add(left, weight=3)
        paned.add(right, weight=2)
        self.tree = ttk.Treeview(right, columns=("name", "kind", "dir", "cells", "use"), show="headings",
                                 selectmode="browse", height=14)
        for col, text, width in (("name", "欄名", 120), ("kind", "種別", 70), ("dir", "方向", 60),
                                 ("cells", "セル", 110), ("use", "使用", 44)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=col in ("name", "cells"))
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        form = ttk.LabelFrame(right, text="選択した記入欄", padding=8)
        form.pack(fill="x", pady=8)
        self.name_var = tk.StringVar()
        self.kind_var = tk.StringVar()
        self.cells_var = tk.StringVar()
        self.prefix_var = tk.StringVar()
        self.use_var = tk.BooleanVar()
        ttk.Label(form, text="欄名").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.name_var).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Label(form, text="種別").grid(row=1, column=0, sticky="w")
        ttk.Combobox(form, textvariable=self.kind_var, state="readonly",
                     values=list(analyzer.KIND_LABELS.values())).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Label(form, text="セル").grid(row=2, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.cells_var).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Label(form, text="接頭辞").grid(row=3, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.prefix_var).grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Checkbutton(form, text="この欄に記入する", variable=self.use_var).grid(row=4, column=1, sticky="w")
        self.form_msg = ttk.Label(form, text="", foreground="#c00")
        self.form_msg.grid(row=5, column=0, columnspan=2, sticky="w")
        form.columnconfigure(1, weight=1)
        for v in (self.name_var, self.kind_var, self.cells_var, self.prefix_var, self.use_var):
            v.trace_add("write", self._on_edit)

        btns = ttk.Frame(right)
        btns.pack(fill="x")
        ttk.Button(btns, text="選択セルを記入欄に追加", command=self.add_field).pack(side="left")
        ttk.Button(btns, text="削除", command=self.remove_field).pack(side="left", padx=4)
        ttk.Button(btns, text="上へ", command=lambda: self.move(-1)).pack(side="left")
        ttk.Button(btns, text="下へ", command=lambda: self.move(1)).pack(side="left", padx=4)
        ttk.Button(btns, text="定義を保存", command=self.save, style="Accent.TButton").pack(side="right")

    # ---- 取り込み・削除 ----
    def import_format(self):
        path = filedialog.askopenfilename(
            title="Excelフォーマットを選択",
            filetypes=[("Excel ブック", "*.xlsx *.xlsm"), ("Excel 97-2003", "*.xls"), ("すべて", "*.*")])
        if not path:
            return
        try:
            xw.check_supported(path)
        except (xw.XlsNotSupported, xw.WriteError) as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return
        name = simpledialog.askstring(APP_TITLE, "このフォーマットの名前を入力してください。",
                                      initialvalue=Path(path).stem, parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        if name in storage.list_formats() and not messagebox.askyesno(
                APP_TITLE, f"「{name}」はすでにあります。フォーマットを置き換えますか？\n（過去事例は残ります）"):
            return
        try:
            fmt = analyzer.analyze(path, name)
            storage.save_format(fmt, template_src=path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"フォーマットを解析できませんでした。\n{e}")
            return
        self.app.refresh_formats(select=name)
        n = sum(1 for f in fmt["fields"] if f.get("enabled"))
        messagebox.showinfo(APP_TITLE, f"{len(fmt['fields'])}個の欄を検出しました（記入する欄 {n}個）。\n"
                                       "プレビューと一覧を確認し、必要なら修正して［定義を保存］を押してください。")

    def delete_format(self):
        fmt = self.app.fmt
        if fmt is None:
            return
        if not messagebox.askyesno(APP_TITLE, f"フォーマット「{fmt['name']}」と、その過去事例を削除しますか？"):
            return
        storage.delete_format(fmt["name"])
        self.app.refresh_formats()

    # ---- 表示 ----
    def refresh(self):
        fmt = self.app.fmt
        self.tree.delete(*self.tree.get_children())
        self.canvas.delete("all")
        self.layout = None
        if fmt is None:
            self.sheet_box.configure(values=[])
            self.sheet_var.set("")
            return
        sheets = fmt.get("sheets") or sorted({f["sheet"] for f in fmt["fields"]})
        self.sheet_box.configure(values=sheets)
        if self.sheet_var.get() not in sheets:
            self.sheet_var.set(sheets[0] if sheets else "")
        self.fill_tree()
        self.load_layout()

    def fill_tree(self, select: str | None = None):
        self.tree.delete(*self.tree.get_children())
        for f in self.app.fmt["fields"]:
            self.tree.insert("", "end", iid=f["id"], values=self._row(f))
        if select and self.tree.exists(select):
            self.tree.selection_set(select)
            self.tree.see(select)

    def _row(self, f):
        cells = ",".join(f["cells"]) if len(f["cells"]) <= 2 else f"{f['cells'][0]}〜{f['cells'][-1]}"
        return (f["name"], analyzer.KIND_LABELS.get(f["kind"], f["kind"]),
                analyzer.DIRECTION_LABELS.get(f["direction"], f["direction"]), cells,
                "○" if f.get("enabled") else "−")

    def load_layout(self):
        fmt = self.app.fmt
        if fmt is None:
            return
        try:
            self.layout = analyzer.sheet_layout(storage.template_path(fmt), self.sheet_var.get() or None)
        except Exception as e:
            self.layout = None
            self.canvas.delete("all")
            self.canvas.create_text(10, 10, anchor="nw", text=f"プレビューを表示できません：{e}")
            return
        self.draw()

    def schedule_draw(self):
        if self._redraw_job:
            self.after_cancel(self._redraw_job)
        self._redraw_job = self.after(120, self.draw)

    def draw(self):
        self._redraw_job = None
        lay = self.layout
        c = self.canvas
        c.delete("all")
        if not lay:
            return
        s = self.app.px_scale
        ox, oy = self.HEAD_W, self.HEAD_H
        self._xs = [0]
        for w in lay["col_px"]:
            self._xs.append(self._xs[-1] + w * s)
        self._ys = [0]
        for h in lay["row_px"]:
            self._ys.append(self._ys[-1] + h * s)

        sheet = lay["sheet"]
        field_at, label_at = {}, set()
        for f in self.app.fmt["fields"]:
            if f["sheet"] != sheet:
                continue
            for ref in f["cells"]:
                field_at[ref] = f
            if f.get("label_cell"):
                label_at.add(f["label_cell"])

        small = (self.app.font_family, 8)
        # 列・行の見出し
        for j in range(lay["n_cols"]):
            x1, x2 = ox + self._xs[j], ox + self._xs[j + 1]
            c.create_rectangle(x1, 0, x2, oy, fill="#f0f0f0", outline="#d0d0d0")
            c.create_text((x1 + x2) / 2, oy / 2, text=xw.index_to_col(j + 1), font=small, fill="#666")
        for i in range(lay["n_rows"]):
            y1, y2 = oy + self._ys[i], oy + self._ys[i + 1]
            c.create_rectangle(0, y1, ox, y2, fill="#f0f0f0", outline="#d0d0d0")
            c.create_text(ox / 2, (y1 + y2) / 2, text=str(i + 1), font=small, fill="#666")

        cell_font = (self.app.font_family, max(7, int(9 * s)))
        for cell in lay["cells"]:
            x1 = ox + self._xs[cell["c1"] - 1]
            x2 = ox + self._xs[cell["c2"]]
            y1 = oy + self._ys[cell["r1"] - 1]
            y2 = oy + self._ys[cell["r2"]]
            if x2 - x1 < 1 or y2 - y1 < 1:
                continue
            ref = cell["ref"]
            f = field_at.get(ref)
            text = cell["text"]
            if f is not None:
                fill = COLOR_FIELD if f.get("enabled") else COLOR_DISABLED
                if f["direction"] == "inline":
                    fill = COLOR_LABEL
                    text = text + f"  ［{f['name']}］"
                else:
                    text = f"［{f['name']}］"
            elif ref in label_at:
                fill = COLOR_LABEL
            else:
                fill = "white"
            outline = "#444" if cell["border"] else "#e4e4e4"
            c.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline)
            if text:
                c.create_text(x1 + 3, y1 + 2, anchor="nw", text=text.replace("\n", " "), font=cell_font,
                              width=max(10, x2 - x1 - 6), fill="#8a6d00" if f is not None else "#222")
        if self.selected_ref:
            box = self._cell_box(self.selected_ref)
            if box:
                c.create_rectangle(*box, outline="#e00", width=2)
        c.configure(scrollregion=(0, 0, ox + self._xs[-1] + 2, oy + self._ys[-1] + 2))

    def _cell_box(self, ref):
        for cell in self.layout["cells"]:
            if cell["ref"] == ref:
                return (self.HEAD_W + self._xs[cell["c1"] - 1], self.HEAD_H + self._ys[cell["r1"] - 1],
                        self.HEAD_W + self._xs[cell["c2"]], self.HEAD_H + self._ys[cell["r2"]])
        return None

    def _on_canvas_click(self, e):
        if not self.layout:
            return
        x = self.canvas.canvasx(e.x) - self.HEAD_W
        y = self.canvas.canvasy(e.y) - self.HEAD_H
        if x < 0 or y < 0:
            return
        col = bisect.bisect_right(self._xs, x)
        row = bisect.bisect_right(self._ys, y)
        if col < 1 or row < 1 or col > self.layout["n_cols"] or row > self.layout["n_rows"]:
            return
        ref = self.layout["owner"].get(f"{row},{col}", xw.make_ref(col, row))
        self.selected_ref = ref
        self.sel_label.configure(text=f"選択セル：{self.layout['sheet']}!{ref}")
        for f in self.app.fmt["fields"]:
            if f["sheet"] == self.layout["sheet"] and (ref in f["cells"] or ref == f.get("label_cell")):
                self.tree.selection_set(f["id"])
                self.tree.see(f["id"])
                break
        self.draw()

    # ---- 編集 ----
    def _current(self) -> dict | None:
        sel = self.tree.selection()
        if not sel or self.app.fmt is None:
            return None
        return next((f for f in self.app.fmt["fields"] if f["id"] == sel[0]), None)

    def _on_select(self, _):
        f = self._current()
        if f is None:
            return
        self._loading = True
        self.name_var.set(f["name"])
        self.kind_var.set(analyzer.KIND_LABELS.get(f["kind"], f["kind"]))
        self.cells_var.set(", ".join(f["cells"]))
        self.prefix_var.set(f.get("prefix", ""))
        self.use_var.set(bool(f.get("enabled")))
        self.form_msg.configure(text="")
        self._loading = False

    def _on_edit(self, *_):
        if self._loading:
            return
        f = self._current()
        if f is None:
            return
        f["name"] = self.name_var.get().strip() or f["name"]
        kind = KIND_BY_LABEL.get(self.kind_var.get(), f["kind"])
        if kind != f["kind"] and kind == "skip":
            self._loading = True
            self.use_var.set(False)
            self._loading = False
        f["kind"] = kind
        f["prefix"] = self.prefix_var.get()
        f["enabled"] = bool(self.use_var.get())
        refs = [r.strip().upper().replace("$", "") for r in re.split(r"[,\s、]+", self.cells_var.get()) if r.strip()]
        try:
            for r in refs:
                xw.split_ref(r)
            if refs:
                f["cells"] = refs
            self.form_msg.configure(text="")
        except xw.WriteError:
            self.form_msg.configure(text="セル番地が正しくありません（例：B3 または B23, B24）")
        self.tree.item(f["id"], values=self._row(f))
        self.schedule_draw()

    def add_field(self):
        fmt = self.app.fmt
        if fmt is None or not self.layout:
            return
        if not self.selected_ref:
            messagebox.showinfo(APP_TITLE, "左のプレビューで、記入欄にしたいセルをクリックしてください。")
            return
        sheet = self.layout["sheet"]
        for f in fmt["fields"]:
            if f["sheet"] == sheet and self.selected_ref in f["cells"]:
                messagebox.showinfo(APP_TITLE, f"{self.selected_ref} はすでに「{f['name']}」の記入欄です。")
                return
        try:
            field = analyzer.field_from_cell(storage.template_path(fmt), sheet, self.selected_ref, fmt["fields"])
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"記入欄を追加できませんでした。\n{e}")
            return
        fmt["fields"].append(field)
        self.fill_tree(select=field["id"])
        self.draw()

    def remove_field(self):
        f = self._current()
        if f is None:
            return
        self.app.fmt["fields"].remove(f)
        self.fill_tree()
        self.draw()

    def move(self, delta):
        f = self._current()
        if f is None:
            return
        fields = self.app.fmt["fields"]
        i = fields.index(f)
        j = max(0, min(len(fields) - 1, i + delta))
        fields.insert(j, fields.pop(i))
        self.fill_tree(select=f["id"])

    def save(self):
        fmt = self.app.fmt
        if fmt is None:
            return
        storage.save_format(fmt)
        self.app.on_format_saved()
        if self.app.cases and messagebox.askyesno(
                APP_TITLE, "定義を保存しました。\n過去事例を新しい定義で読み込み直しますか？"):
            self.app.reload_cases()


# ---------------------------------------------------------------------------
# 過去事例タブ
# ---------------------------------------------------------------------------

class CasesTab(ttk.Frame):
    def __init__(self, master, app: "App"):
        super().__init__(master, padding=8)
        self.app = app
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Button(top, text="ファイルを追加", command=self.add_files).pack(side="left")
        ttk.Button(top, text="フォルダを追加", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(top, text="再読込", command=self.app.reload_cases).pack(side="left")
        ttk.Button(top, text="削除", command=self.remove).pack(side="left", padx=4)
        self.count = ttk.Label(top, text="", foreground="#555")
        self.count.pack(side="right")

        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, pady=(8, 0))
        upper = ttk.PanedWindow(paned, orient="horizontal")
        list_frame = ttk.Frame(upper)
        self.tree = ttk.Treeview(list_frame, columns=("subject", "date", "file", "origin", "warn"),
                                 show="headings", selectmode="extended")
        for col, text, width in (("subject", "件名", 240), ("date", "日付", 90), ("file", "ファイル", 200),
                                 ("origin", "登録元", 60), ("warn", "警告", 44)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, stretch=col in ("subject", "file"))
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.preview = tk.Text(upper, wrap="char", state="disabled", bg="#f7f7f7", font=app.text_font,
                               relief="solid", bd=1)
        upper.add(list_frame, weight=3)
        upper.add(self.preview, weight=2)
        log_frame = ttk.LabelFrame(paned, text="取り込みログ", padding=4)
        self.log = tk.Text(log_frame, height=6, wrap="char", state="disabled", font=app.text_font)
        self.log.pack(fill="both", expand=True)
        paned.add(upper, weight=4)
        paned.add(log_frame, weight=1)

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        fmt = self.app.fmt
        if fmt is None:
            self.count.configure(text="フォーマットが未登録です")
            return
        for c in self.app.cases:
            self.tree.insert("", "end", iid=c["id"], values=(
                cases_mod.subject_of(fmt, c), cases_mod.date_of(fmt, c), Path(c["source"]).name,
                "出力" if c.get("origin") == "output" else "取込", len(c.get("warnings", [])) or ""))
        self.count.configure(text=f"{len(self.app.cases)}件")

    def write_log(self, lines):
        if not lines:
            return
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        for l in lines:
            self.log.insert("end", f"[{stamp}] {l}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def add_files(self):
        if not self._need_format():
            return
        paths = filedialog.askopenfilenames(title="過去の報告書を選択",
                                            filetypes=[("Excel ブック", "*.xlsx *.xlsm"), ("すべて", "*.*")])
        if paths:
            self.app.import_cases(paths)

    def add_folder(self):
        if not self._need_format():
            return
        path = filedialog.askdirectory(title="過去の報告書のフォルダを選択（サブフォルダも含みます）")
        if path:
            self.app.import_cases([path])

    def remove(self):
        sel = set(self.tree.selection())
        if not sel:
            return
        if not messagebox.askyesno(APP_TITLE, f"選択した {len(sel)}件の事例を削除しますか？\n（元の Excel ファイルは削除しません）"):
            return
        self.app.cases[:] = [c for c in self.app.cases if c["id"] not in sel]
        self.app.save_cases()

    def _need_format(self):
        if self.app.fmt is None:
            messagebox.showinfo(APP_TITLE, "先に「フォーマット設定」タブでフォーマットを取り込んでください。")
            return False
        return True

    def _on_select(self, _):
        sel = self.tree.selection()
        case = next((c for c in self.app.cases if sel and c["id"] == sel[0]), None)
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        if case:
            self.preview.insert("1.0", f"{case['source']}\n\n" + cases_mod.case_preview(self.app.fmt, case))
        self.preview.configure(state="disabled")


# ---------------------------------------------------------------------------
# 設定タブ
# ---------------------------------------------------------------------------

class SettingsTab(ttk.Frame):
    def __init__(self, master, app: "App"):
        super().__init__(master, padding=16)
        self.app = app
        s = app.settings
        self.vars = {
            "reporter": tk.StringVar(value=s.get("reporter", "")),
            "department": tk.StringVar(value=s.get("department", "")),
            "output_dir": tk.StringVar(value=s.get("output_dir", "")),
            "filename_pattern": tk.StringVar(value=s.get("filename_pattern", "{date}_{topic}")),
            "reuse": tk.BooleanVar(value=s.get("reuse", True)),
            "highlight": tk.BooleanVar(value=s.get("highlight", True)),
            "register_output": tk.BooleanVar(value=s.get("register_output", True)),
            "open_after_export": tk.BooleanVar(value=s.get("open_after_export", True)),
        }
        today = date.today()
        self.date_choices = [(k, label or tu.format_date(today, k)) for k, label in DATE_FORMATS]
        current = next((label for k, label in self.date_choices if k == s.get("date_format", "auto")),
                       self.date_choices[0][1])
        self.date_var = tk.StringVar(value=current)

        row = 0

        def add(label, widget, hint=None):
            nonlocal row
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 12))
            widget.grid(row=row, column=1, sticky="ew", pady=4)
            if hint:
                ttk.Label(self, text=hint, foreground="#777").grid(row=row, column=2, sticky="w", padx=8)
            row += 1

        add("報告者", ttk.Entry(self, textvariable=self.vars["reporter"], width=30), "空欄なら過去事例で最も多い値")
        add("部署", ttk.Entry(self, textvariable=self.vars["department"], width=30), "空欄なら過去事例で最も多い値")
        add("日付書式", ttk.Combobox(self, textvariable=self.date_var, state="readonly",
                                   values=[label for _, label in self.date_choices], width=28))
        out = ttk.Frame(self)
        ttk.Entry(out, textvariable=self.vars["output_dir"]).pack(side="left", fill="x", expand=True)
        ttk.Button(out, text="参照", command=self.choose_dir).pack(side="left", padx=4)
        add("出力先フォルダ", out, "空欄ならドキュメント")
        add("ファイル名パターン", ttk.Entry(self, textvariable=self.vars["filename_pattern"], width=30),
            "{date}（例 20260926）{topic} {reporter} が使えます")
        for key, text in (("reuse", "メモのない欄に過去事例の文面を流用する"),
                          ("highlight", "流用した欄（要確認）・自動入力の欄を色分けする"),
                          ("register_output", "出力した報告書を過去事例に登録する（既定）"),
                          ("open_after_export", "出力後に Excel で開く（既定）")):
            ttk.Checkbutton(self, text=text, variable=self.vars[key]).grid(row=row, column=1, sticky="w", pady=2)
            row += 1
        ttk.Button(self, text="設定を保存", command=self.save, style="Accent.TButton").grid(
            row=row, column=1, sticky="w", pady=(12, 0))
        row += 1
        ttk.Label(self, text=f"データの保存先：{storage.data_dir()}", foreground="#777").grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(24, 0))
        self.columnconfigure(1, weight=1)

    def choose_dir(self):
        d = filedialog.askdirectory(title="出力先フォルダ")
        if d:
            self.vars["output_dir"].set(d)

    def save(self):
        s = self.app.settings
        for k, v in self.vars.items():
            s[k] = v.get().strip() if isinstance(v, tk.StringVar) else bool(v.get())
        s["date_format"] = next((k for k, label in self.date_choices if label == self.date_var.get()), "auto")
        storage.save_settings(s)
        self.app.create_tab.open_var.set(s["open_after_export"])
        self.app.create_tab.register_var.set(s["register_output"])
        messagebox.showinfo(APP_TITLE, "設定を保存しました。")


# ---------------------------------------------------------------------------
# アプリ本体
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x820")
        self.minsize(960, 600)

        self.font_family = pick_font_family(self)
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont",
                     "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(name).configure(family=self.font_family, size=10)
            except tk.TclError:
                pass
        self.text_font = (self.font_family, 10)
        self.bold_font = (self.font_family, 10, "bold")
        self.px_scale = max(1.0, self.winfo_fpixels("1i") / 96.0)  # 高 DPI でもプレビューを実寸に近づける

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        line = tkfont.Font(family=self.font_family, size=10).metrics("linespace")
        style.configure("Treeview", rowheight=int(line * 1.5))
        style.configure("Accent.TButton", font=self.bold_font)

        self.settings = storage.load_settings()
        self.fmt: dict | None = None
        self.cases: list[dict] = []
        self.index = cases_mod.Index({"fields": []}, [])

        top = ttk.Frame(self, padding=(8, 8, 8, 0))
        top.pack(fill="x")
        ttk.Label(top, text="フォーマット").pack(side="left")
        self.format_var = tk.StringVar()
        self.format_box = ttk.Combobox(top, textvariable=self.format_var, state="readonly", width=32)
        self.format_box.pack(side="left", padx=6)
        self.format_box.bind("<<ComboboxSelected>>", lambda e: self.set_format(self.format_var.get()))

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)
        self.create_tab = CreateTab(self.notebook, self)
        self.format_tab = FormatTab(self.notebook, self)
        self.cases_tab = CasesTab(self.notebook, self)
        self.settings_tab = SettingsTab(self.notebook, self)
        self.notebook.add(self.create_tab, text="  報告書作成  ")
        self.notebook.add(self.format_tab, text="  フォーマット設定  ")
        self.notebook.add(self.cases_tab, text="  過去事例  ")
        self.notebook.add(self.settings_tab, text="  設定  ")

        self.bind_all("<MouseWheel>", ScrollFrame.on_wheel, add="+")
        self.refresh_formats(select=self.settings.get("last_format"))

    # ---- フォーマット ----
    def refresh_formats(self, select: str | None = None):
        names = storage.list_formats()
        self.format_box.configure(values=names)
        if select not in names:
            select = names[0] if names else None
        self.set_format(select)

    def set_format(self, name: str | None):
        self.fmt = storage.load_format(name) if name else None
        self.format_var.set(name or "")
        self.cases = storage.load_cases(name) if self.fmt else []
        self.rebuild_index()
        if self.fmt:
            self.settings["last_format"] = name
            storage.save_settings(self.settings)
        self.create_tab.refresh()
        self.format_tab.refresh()
        self.cases_tab.refresh()

    def on_format_saved(self):
        self.rebuild_index()
        self.create_tab.refresh()

    # ---- 過去事例 ----
    def rebuild_index(self):
        self.index = cases_mod.Index(self.fmt or {"fields": []}, self.cases)

    def save_cases(self):
        if self.fmt is None:
            return
        storage.save_cases(self.fmt["name"], self.cases)
        self.rebuild_index()
        self.cases_tab.refresh()

    def import_cases(self, paths):
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            log = cases_mod.import_paths(self.fmt, self.cases, paths)
        finally:
            self.config(cursor="")
        self.save_cases()
        self.cases_tab.write_log(log)

    def reload_cases(self):
        if self.fmt is None:
            return
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            log = cases_mod.reload_all(self.fmt, self.cases)
        finally:
            self.config(cursor="")
        self.save_cases()
        self.cases_tab.write_log(log or ["再読込する事例がありません"])

    def register_output(self, path) -> list[str]:
        if self.fmt is None:
            return []
        log = cases_mod.import_paths(self.fmt, self.cases, [path], origin="output")
        self.save_cases()
        self.cases_tab.write_log(log)
        return log


def run():
    app = App()
    app.mainloop()
