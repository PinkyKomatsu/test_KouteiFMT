"""報告書作成画面（メイン）：入力 → 類似事例 → 日英の文面生成 → 編集 → プレビュー → コピー・出力。"""
from __future__ import annotations

import os
from datetime import date
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QPlainTextEdit,
                               QProgressBar, QPushButton, QScrollArea, QSplitter, QStackedWidget, QTableWidget,
                               QTableWidgetItem, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

from .. import analyzer, generator_ja, report, storage
from .. import cases as cases_mod
from .. import xlsx_writer as xw
from .common import AppState, FormatSelector, entry_color, run_async, show_error
from .sheet_preview import SheetPreview

MEMO_HINT = ("1行に1項目。「課題：〜」のように欄名を書くとその欄に入ります。「【今後の予定】」だけの行を書くと、"
             "以降の行はその欄に入ります。表は「内容｜担当｜期限｜状態」の形で書きます。")


class TextBox(QPlainTextEdit):
    """背景色付きの編集欄。手で編集したら白に戻し、edited を出す。"""
    edited = Signal(str)

    def __init__(self, lines: int, parent=None):
        super().__init__(parent)
        self._setting = False
        self.setTabChangesFocus(True)
        self.setFixedHeight(int(self.fontMetrics().lineSpacing() * lines + 14))
        self.textChanged.connect(self._changed)

    def set_text(self, text: str, color: str):
        self._setting = True
        self.setPlainText(text or "")
        self.setStyleSheet(f"QPlainTextEdit {{ background: {color}; }}")
        self._setting = False

    def _changed(self):
        if self._setting:
            return
        self.setStyleSheet("QPlainTextEdit { background: #ffffff; }")
        self.edited.emit(self.toPlainText())


def _badge(label: QLabel, text: str, color: str):
    label.setText(text)
    label.setStyleSheet(f"background:{color}; border:1px solid #d0d0d0; border-radius:3px; padding:1px 4px; color:#333")
    label.setVisible(bool(text))


class FieldCard(QFrame):
    changed = Signal(str, str, str)   # 欄 id, "ja"/"en", 本文
    copyRequested = Signal(str, str)
    retranslateRequested = Signal(str)
    pickRequested = Signal(str)

    def __init__(self, fmt: dict, field: dict, parent=None):
        super().__init__(parent)
        self.field = field
        self.fid = field["id"]
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName("card")
        by_id = {f["id"]: f for f in fmt["fields"]}
        partner = by_id.get(field.get("pair") or "")

        head = QHBoxLayout()
        title = QLabel(f"<b>{field['name']}</b>")
        head.addWidget(title)
        where = f"{field['sheet']}!{field['cells'][0]}" if field.get("cells") else "記入先が未確定"
        if partner:
            where += f"　英語 → {partner['name']}（{partner['cells'][0] if partner.get('cells') else '未確定'}）"
        info = QLabel(f"{where}　〔{analyzer.KIND_LABELS.get(field['kind'], field['kind'])}・"
                      f"{analyzer.LANG_LABELS.get(field.get('lang', 'ja'))}欄〕")
        info.setStyleSheet("color:#666")
        head.addWidget(info)
        head.addStretch(1)
        if not field.get("cells") or field.get("review"):
            pick = QPushButton("プレビューで記入先を指定")
            pick.setStyleSheet("color:#b35c00")
            pick.clicked.connect(lambda: self.pickRequested.emit(self.fid))
            head.addWidget(pick)

        lines = max(1, min(10, round((field.get("height_px") or 20) / 20)))
        if field["kind"] in analyzer.TEXT_KINDS:
            lines = max(3, lines)
        self.ja = TextBox(lines)
        self.en = TextBox(lines)
        self.ja.edited.connect(lambda t: self.changed.emit(self.fid, "ja", t))
        self.en.edited.connect(lambda t: self.changed.emit(self.fid, "en", t))
        self.ja_note, self.en_note = QLabel(), QLabel()
        for n in (self.ja_note, self.en_note):
            n.setWordWrap(True)

        grid = QGridLayout()
        grid.addWidget(QLabel("日本語"), 0, 0)
        grid.addWidget(QLabel("English"), 0, 1)
        grid.addWidget(self.ja, 1, 0)
        grid.addWidget(self.en, 1, 1)
        grid.addWidget(self.ja_note, 2, 0)
        grid.addWidget(self.en_note, 2, 1)
        ja_btns, en_btns = QHBoxLayout(), QHBoxLayout()
        b = QPushButton("日本語をコピー")
        b.clicked.connect(lambda: self.copyRequested.emit(self.fid, "ja"))
        ja_btns.addWidget(b)
        ja_btns.addStretch(1)
        b = QPushButton("英語をコピー")
        b.clicked.connect(lambda: self.copyRequested.emit(self.fid, "en"))
        en_btns.addWidget(b)
        self.retranslate_btn = QPushButton("英訳を更新")
        self.retranslate_btn.clicked.connect(lambda: self.retranslateRequested.emit(self.fid))
        en_btns.addWidget(self.retranslate_btn)
        en_btns.addStretch(1)
        grid.addLayout(ja_btns, 3, 0)
        grid.addLayout(en_btns, 3, 1)

        layout = QVBoxLayout(self)
        layout.addLayout(head)
        layout.addLayout(grid)

    def set_entry(self, entry: dict):
        self.ja.set_text(entry.get("ja", ""), entry_color(entry.get("ja_status", "empty"), entry.get("review_ja", False)))
        self.en.set_text(entry.get("en", ""), entry_color(entry.get("en_status", "empty"), entry.get("review_en", False)))
        _badge(self.ja_note, _note(entry.get("ja_status"), entry.get("note_ja", ""), entry.get("review_ja")),
               entry_color(entry.get("ja_status", "empty"), entry.get("review_ja", False)))
        _badge(self.en_note, _note(entry.get("en_status"), entry.get("note_en", ""), entry.get("review_en")),
               entry_color(entry.get("en_status", "empty"), entry.get("review_en", False)))

    def focus_ja(self):
        self.ja.setFocus()


def _note(status: str | None, note: str, review: bool | None) -> str:
    """判定の根拠（メモを整形／事例を流用／機械翻訳／簡易訳 など）。要確認なら先頭に付ける。"""
    if not status or status == "empty":
        return note or ""
    label = report.SOURCE_LABELS.get(status, status)
    text = note if note and (note.startswith(label) or note.startswith("要確認")) else (f"{label}：{note}" if note else label)
    if review and status != "reuse" and not text.startswith("要確認"):
        text = "要確認・" + text
    return text

class TableCard(QFrame):
    """表の欄（列ごとの field）をまとめて編集するカード。"""
    changed = Signal(str, str, str)
    copyRequested = Signal(str, str)
    retranslateRequested = Signal(str)
    pickRequested = Signal(str)

    def __init__(self, fmt: dict, columns: list[dict], parent=None):
        super().__init__(parent)
        self.columns = columns
        self.fid = columns[0]["id"]
        self.setFrameShape(QFrame.StyledPanel)
        n_rows = min(len(c["cells"]) for c in columns) if all(c.get("cells") for c in columns) else 5
        head = QHBoxLayout()
        head.addWidget(QLabel("<b>表：" + "／".join(c["name"] for c in columns) + "</b>"))
        where = f"{columns[0]['sheet']}!{columns[0]['cells'][0]}〜" if columns[0].get("cells") else "記入先が未確定"
        info = QLabel(where)
        info.setStyleSheet("color:#666")
        head.addWidget(info)
        head.addStretch(1)
        self.grids = {}
        self._setting = False
        grid = QGridLayout()
        for col, lang in enumerate(("ja", "en")):
            t = QTableWidget(n_rows, len(columns))
            t.setHorizontalHeaderLabels([c["name"] if lang == "ja" else report.english_name(fmt, c) for c in columns])
            t.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            t.verticalHeader().setVisible(False)
            t.setFixedHeight(t.horizontalHeader().height() + n_rows * 26 + 6)
            t.itemChanged.connect(lambda item, lg=lang: self._changed(lg, item))
            self.grids[lang] = t
            grid.addWidget(QLabel("日本語" if lang == "ja" else "English"), 0, col)
            grid.addWidget(t, 1, col)
        self.ja_note, self.en_note = QLabel(), QLabel()
        grid.addWidget(self.ja_note, 2, 0)
        grid.addWidget(self.en_note, 2, 1)
        jb, eb = QHBoxLayout(), QHBoxLayout()
        b = QPushButton("日本語をコピー")
        b.clicked.connect(lambda: self.copyRequested.emit(self.fid, "ja"))
        jb.addWidget(b)
        jb.addStretch(1)
        b = QPushButton("英語をコピー")
        b.clicked.connect(lambda: self.copyRequested.emit(self.fid, "en"))
        eb.addWidget(b)
        b = QPushButton("英訳を更新")
        b.clicked.connect(lambda: self.retranslateRequested.emit(self.fid))
        eb.addWidget(b)
        eb.addStretch(1)
        grid.addLayout(jb, 3, 0)
        grid.addLayout(eb, 3, 1)
        layout = QVBoxLayout(self)
        layout.addLayout(head)
        layout.addLayout(grid)

    def set_entries(self, draft: dict):
        self._setting = True
        for lang, t in self.grids.items():
            for j, c in enumerate(self.columns):
                entry = draft.get(c["id"], {})
                lines = (entry.get(lang) or "").split("\n")
                color = entry_color(entry.get(f"{lang}_status", "empty"), entry.get(f"review_{lang}", False))
                for i in range(t.rowCount()):
                    item = QTableWidgetItem(lines[i] if i < len(lines) else "")
                    item.setBackground(_brush(color))
                    t.setItem(i, j, item)
        first = draft.get(self.columns[0]["id"], {})
        _badge(self.ja_note, first.get("note_ja", ""), "#ffffff")
        notes = [draft.get(c["id"], {}).get("note_en", "") for c in self.columns]
        _badge(self.en_note, next((n for n in notes if "要確認" in n), next((n for n in notes if n), "")), "#ffffff")
        self._setting = False

    def column_text(self, fid: str, lang: str) -> str:
        t = self.grids[lang]
        j = [c["id"] for c in self.columns].index(fid)
        vals = [(t.item(i, j).text() if t.item(i, j) else "") for i in range(t.rowCount())]
        while vals and not vals[-1]:
            vals.pop()
        return "\n".join(vals)

    def _changed(self, lang, item):
        if self._setting:
            return
        item.setBackground(_brush("#ffffff"))
        fid = self.columns[item.column()]["id"]
        self.changed.emit(fid, lang, self.column_text(fid, lang))

    def focus_ja(self):
        self.grids["ja"].setFocus()


def _brush(hex_color: str):
    from PySide6.QtGui import QBrush, QColor
    return QBrush(QColor(hex_color))


class ComposePage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.draft: dict[str, dict] = {}
        self.cards: dict[str, QWidget] = {}
        self.hits: list[tuple[float, dict]] = []
        self.pick_field: str | None = None
        self.busy = False
        self.last_export: Path | None = None

        # ---- 左：入力と参考事例 ----
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 6, 0)
        row = QHBoxLayout()
        row.addWidget(QLabel("フォーマット"))
        row.addWidget(FormatSelector(state), 1)
        ll.addLayout(row)
        ll.addWidget(QLabel("トピック（必須）"))
        self.topic = QLineEdit()
        self.topic.setPlaceholderText("例：移行リハーサル（第2回）の結果報告")
        ll.addWidget(self.topic)
        row = QHBoxLayout()
        row.addWidget(QLabel("工程"))
        self.phase = QComboBox()
        self.phase.setEditable(True)
        row.addWidget(self.phase, 1)
        row.addWidget(QLabel("進捗"))
        self.status_box = QComboBox()
        self.status_box.addItems(storage.STATUSES)
        row.addWidget(self.status_box)
        ll.addLayout(row)
        ll.addWidget(QLabel("要点メモ（1行1項目）"))
        self.memo = QPlainTextEdit()
        self.memo.setPlaceholderText(MEMO_HINT)
        ll.addWidget(self.memo, 2)
        hint = QLabel(MEMO_HINT)
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#777; font-size:9pt")
        ll.addWidget(hint)
        row = QHBoxLayout()
        self.search_btn = QPushButton("類似事例を検索")
        self.search_btn.clicked.connect(lambda: self.search())
        self.gen_btn = QPushButton("文面を生成")
        self.gen_btn.setStyleSheet("font-weight:bold")
        self.gen_btn.clicked.connect(lambda: self.generate())
        row.addWidget(self.search_btn)
        row.addWidget(self.gen_btn)
        ll.addLayout(row)
        ll.addWidget(QLabel("参考事例（☑ の上位3件から書き方を学習）"))
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["使用", "類似度", "件名", "工程"])
        self.tree.setRootIsDecorated(False)
        self.tree.header().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tree.itemSelectionChanged.connect(self._show_case)
        ll.addWidget(self.tree, 2)
        self.case_preview = QPlainTextEdit()
        self.case_preview.setReadOnly(True)
        ll.addWidget(self.case_preview, 2)

        # ---- 中央：欄ごとのカード ----
        center = QWidget()
        cl = QVBoxLayout(center)
        cl.setContentsMargins(0, 0, 0, 0)
        legend = QLabel("<span style='background:#fff3b0'>&nbsp;事例を流用&nbsp;</span> "
                        "<span style='background:#ffd8a8'>&nbsp;要確認・簡易訳&nbsp;</span> "
                        "<span style='background:#e8f0fe'>&nbsp;機械翻訳&nbsp;</span> "
                        "<span style='background:#e6f4ea'>&nbsp;自動・定型文・翻訳メモリ&nbsp;</span> "
                        "（手で編集すると白に戻ります）")
        cl.addWidget(legend)
        self.card_area = QScrollArea()
        self.card_area.setWidgetResizable(True)
        self.card_host = QWidget()
        self.card_layout = QVBoxLayout(self.card_host)
        self.card_layout.addStretch(1)
        self.card_area.setWidget(self.card_host)
        cl.addWidget(self.card_area, 1)

        # ---- 右：記入後の Excel プレビュー ----
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(6, 0, 0, 0)
        row = QHBoxLayout()
        row.addWidget(QLabel("記入後のプレビュー（セルをクリックすると欄へ移動）"))
        row.addStretch(1)
        self.en_preview = QCheckBox("英語版で表示")
        self.en_preview.toggled.connect(lambda _: self.refresh_preview())
        row.addWidget(self.en_preview)
        rl.addLayout(row)
        self.preview = SheetPreview()
        self.preview.cellClicked.connect(self._on_preview_click)
        rl.addWidget(self.preview, 1)
        self.pick_label = QLabel("")
        self.pick_label.setStyleSheet("color:#b35c00; font-weight:bold")
        rl.addWidget(self.pick_label)

        split = QSplitter(Qt.Horizontal)
        for w, s in ((left, 3), (center, 5), (right, 4)):
            split.addWidget(w)
        split.setSizes([360, 640, 520])

        # ---- 下部 ----
        bottom = QHBoxLayout()
        self.copy_ja_btn = QPushButton("日本語を全欄コピー")
        self.copy_ja_btn.clicked.connect(lambda: self.copy_all("ja"))
        self.copy_en_btn = QPushButton("英語を全欄コピー")
        self.copy_en_btn.clicked.connect(lambda: self.copy_all("en"))
        bottom.addWidget(self.copy_ja_btn)
        bottom.addWidget(self.copy_en_btn)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setMaximumWidth(120)
        self.progress.setVisible(False)
        bottom.addWidget(self.progress)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#555")
        bottom.addWidget(self.status, 1)
        self.open_after = QCheckBox("出力後に開く")
        self.open_after.setChecked(bool(state.settings.get("open_after_export", True)))
        self.register = QCheckBox("過去事例に登録")
        self.register.setChecked(bool(state.settings.get("register_output", True)))
        bottom.addWidget(self.open_after)
        bottom.addWidget(self.register)
        self.export_en_btn = QPushButton("英語版を別ブックで出力")
        self.export_en_btn.clicked.connect(lambda: self.export(english_only=True))
        self.export_btn = QPushButton("Excelに出力")
        self.export_btn.setStyleSheet("font-weight:bold")
        self.export_btn.clicked.connect(lambda: self.export())
        bottom.addWidget(self.export_en_btn)
        bottom.addWidget(self.export_btn)

        main = QWidget()
        ml = QVBoxLayout(main)
        ml.setContentsMargins(0, 0, 0, 0)
        ml.addWidget(split, 1)
        ml.addLayout(bottom)

        guide = QLabel("まずフォーマットを取り込んでください。\n\n「フォーマット設定」タブの［Excelフォーマットを取り込む］から始めます。")
        guide.setAlignment(Qt.AlignCenter)
        guide.setStyleSheet("font-size:14pt; color:#555")
        self.stack = QStackedWidget()
        self.stack.addWidget(guide)
        self.stack.addWidget(main)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self.stack)

        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self.refresh_preview)

        state.formatChanged.connect(self.on_format_changed)
        state.settingsChanged.connect(self._load_phases)
        self._load_phases()
        self._current_format = None
        self.on_format_changed()

    # ---- 状態の反映 ----
    def _load_phases(self):
        current = self.phase.currentText()
        self.phase.clear()
        self.phase.addItems(self.state.settings.get("phases") or storage.DEFAULT_PHASES)
        if current:
            self.phase.setCurrentText(current)

    def on_format_changed(self):
        fmt = self.state.fmt
        self.stack.setCurrentIndex(1 if fmt else 0)
        same = fmt is not None and self._current_format == fmt["name"]
        self._current_format = fmt["name"] if fmt else None
        if not same:
            self.draft = {}
            self.hits = []
            self.tree.clear()
            self.case_preview.clear()
        self.build_cards()
        try:
            self.preview.load(self.state.template_path(), fmt.get("sheets") if fmt else None)
        except Exception as e:
            show_error(self, "プレビューを表示できません", e)
        self.refresh_preview()

    def build_cards(self):
        while self.card_layout.count() > 1:
            w = self.card_layout.takeAt(0).widget()
            if w:
                w.deleteLater()
        self.cards.clear()
        fmt = self.state.fmt
        if not fmt:
            return
        done_tables = set()
        for f in generator_ja.logical_fields(fmt):
            if f.get("table_id"):
                if f["table_id"] in done_tables:
                    continue
                done_tables.add(f["table_id"])
                columns = [c for c in generator_ja.logical_fields(fmt) if c.get("table_id") == f["table_id"]]
                card = TableCard(fmt, columns)
                for c in columns:
                    self.cards[c["id"]] = card
            else:
                card = FieldCard(fmt, f)
                self.cards[f["id"]] = card
            card.changed.connect(self._on_edit)
            card.copyRequested.connect(self.copy_field)
            card.retranslateRequested.connect(self.retranslate)
            card.pickRequested.connect(self._start_pick)
            self.card_layout.insertWidget(self.card_layout.count() - 1, card)
        self._show_draft()

    def _show_draft(self):
        shown = set()
        for fid, card in self.cards.items():
            if id(card) in shown:
                continue
            shown.add(id(card))
            if isinstance(card, TableCard):
                card.set_entries(self.draft)
            else:
                card.set_entry(self.draft.get(fid, report.new_entry()))

    # ---- 入力 ----
    def set_inputs(self, topic: str, phase: str, status: str, memo: str):
        self.topic.setText(topic)
        self.phase.setCurrentText(phase)
        self.status_box.setCurrentText(status)
        self.memo.setPlainText(memo)

    def _query(self) -> str:
        return self.topic.text() + "\n" + self.memo.toPlainText()

    def search(self):
        if not self.state.fmt:
            return []
        if not self.state.cases:
            self.status.setText("過去事例がありません（「過去事例」タブで取り込めます）")
            return []
        self.hits = self.state.index.search(self._query(), top=10, phase=self.phase.currentText().strip() or None)
        self.tree.clear()
        for i, (score, c) in enumerate(self.hits):
            item = QTreeWidgetItem(["", f"{score:.2f}", cases_mod.subject_of(self.state.fmt, c),
                                    cases_mod.phase_of(self.state.fmt, c)])
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(0, Qt.Checked if i < 3 and score > 0 else Qt.Unchecked)
            item.setData(0, Qt.UserRole, c["id"])
            self.tree.addTopLevelItem(item)
        if self.hits:
            self.tree.setCurrentItem(self.tree.topLevelItem(0))
        self.status.setText(f"類似事例 {len(self.hits)}件")
        return self.hits

    def checked_cases(self) -> list[dict]:
        ids = [self.tree.topLevelItem(i).data(0, Qt.UserRole) for i in range(self.tree.topLevelItemCount())
               if self.tree.topLevelItem(i).checkState(0) == Qt.Checked]
        by_id = {c["id"]: c for _, c in self.hits}
        return [by_id[i] for i in ids if i in by_id][:3]

    def _show_case(self):
        item = self.tree.currentItem()
        if not item:
            return
        case = next((c for _, c in self.hits if c["id"] == item.data(0, Qt.UserRole)), None)
        if case:
            self.case_preview.setPlainText(cases_mod.case_preview(self.state.fmt, case))

    # ---- 生成 ----
    def generate(self, sync: bool = False):
        if not self.state.fmt or self.busy:
            return
        topic = self.topic.text().strip()
        if not topic:
            QMessageBox.information(self, "文面を生成", "トピックを入力してください。")
            return
        if not self.hits and self.state.cases:
            self.search()
        fmt, settings = self.state.fmt, self.state.settings_copy()
        args = (fmt, topic, self.phase.currentText().strip(), self.status_box.currentText(),
                self.memo.toPlainText(), self.checked_cases(), list(self.state.cases), settings,
                self.state.translator(), date.today())

        def work():
            return report.build_draft(*args)

        if sync:
            self._generated(work())
            return
        self._set_busy(True, "文面を生成しています（英訳を含む）…")
        run_async(work, self._generated, self._failed)

    def _generated(self, draft):
        self._set_busy(False)
        self.draft = draft
        self._show_draft()
        self.refresh_preview()
        n_reuse = sum(1 for e in draft.values() if e["ja_status"] == "reuse")
        n_review = sum(1 for e in draft.values() if e["review_ja"] or e["review_en"])
        t = self.state.translator()
        msg = "文面を作成しました"
        if n_reuse or n_review:
            msg += f"（流用 {n_reuse}欄・要確認 {n_review}欄。色の付いた欄を確認してください）"
        if not t.available:
            msg += "　※翻訳モデルがないため英語は簡易訳です"
        self.status.setText(msg)

    def _failed(self, err):
        self._set_busy(False)
        show_error(self, "処理に失敗しました", err)

    def _set_busy(self, busy: bool, message: str = ""):
        self.busy = busy
        self.progress.setVisible(busy)
        for b in (self.gen_btn, self.search_btn, self.export_btn, self.export_en_btn):
            b.setEnabled(not busy)
        if message:
            self.status.setText(message)

    # ---- 編集 ----
    def _entry(self, fid: str) -> dict:
        return self.draft.setdefault(fid, report.new_entry())

    def _on_edit(self, fid: str, lang: str, text: str):
        e = self._entry(fid)
        e[lang] = text
        e[f"{lang}_status"] = "edited"
        e[f"review_{lang}"] = False
        self._preview_timer.start(300)

    def retranslate(self, fid: str, sync: bool = False):
        fmt = self.state.fmt
        card = self.cards.get(fid)
        fids = [c["id"] for c in card.columns] if isinstance(card, TableCard) else [fid]
        fields = {f["id"]: f for f in fmt["fields"]}
        translator, settings = self.state.translator(), self.state.settings_copy()

        def work():
            out = {}
            for i in fids:
                e = dict(self._entry(i))
                out[i] = report.translate_entry(fmt, fields[i], e, translator, settings)
            return out

        def done(result):
            self._set_busy(False)
            self.draft.update(result)
            self._show_draft()
            self.refresh_preview()
            self.status.setText("英訳を更新しました")

        if sync:
            done(work())
            return
        self._set_busy(True, "英訳を更新しています…")
        run_async(work, done, self._failed)

    # ---- プレビュー ----
    def refresh_preview(self):
        fmt = self.state.fmt
        if not fmt:
            return
        values = report.preview_values(fmt, self.draft, english_only=self.en_preview.isChecked()) if self.draft else {}
        review = {fid for fid, e in self.draft.items() if e.get("review_ja") or e.get("review_en")}
        self.preview.apply(fmt, values if self.draft else None, review)

    def _on_preview_click(self, sheet: str, ref: str):
        fmt = self.state.fmt
        if not fmt:
            return
        if self.pick_field:
            self._assign_pick(sheet, ref)
            return
        f = report.field_at(fmt, sheet, ref)
        card = self.cards.get(f["id"]) if f else None
        if card:
            self.card_area.ensureWidgetVisible(card)
            card.focus_ja()

    def _start_pick(self, fid: str):
        self.pick_field = fid
        name = next((f["name"] for f in self.state.fmt["fields"] if f["id"] == fid), fid)
        self.pick_label.setText(f"「{name}」の記入先を、プレビューのセルをクリックして指定してください（Esc で中止）")

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape and self.pick_field:
            self.pick_field = None
            self.pick_label.setText("")
            return
        super().keyPressEvent(e)

    def _assign_pick(self, sheet: str, ref: str):
        fmt = self.state.fmt
        field = next((f for f in fmt["fields"] if f["id"] == self.pick_field), None)
        self.pick_field = None
        self.pick_label.setText("")
        if field is None:
            return
        field["sheet"] = sheet
        analyzer.assign_cells(self.state.template_path(), field, [ref])
        self.state.save_format(fmt)            # 定義にも保存する（下書きは保たれる）
        self.status.setText(f"「{field['name']}」の記入先を {sheet}!{ref} にしました")

    # ---- コピー ----
    def copy_field(self, fid: str, lang: str):
        card = self.cards.get(fid)
        if isinstance(card, TableCard):
            text = report._table_block(self.state.fmt, card.columns, self.draft, lang) or ""
        else:
            field = next(f for f in self.state.fmt["fields"] if f["id"] == fid)
            text = report.field_clip(self.state.fmt, field, self._entry(fid), lang)
        QGuiApplication.clipboard().setText(text)
        self.status.setText(("日本語" if lang == "ja" else "英語") + "をコピーしました")
        return text

    def copy_all(self, lang: str) -> str:
        if not self.state.fmt:
            return ""
        text = report.clipboard_text(self.state.fmt, self.draft, lang)
        QGuiApplication.clipboard().setText(text)
        self.status.setText(("日本語" if lang == "ja" else "英語") + "を全欄コピーしました（【欄名】本文 の形式）")
        return text

    # ---- 出力 ----
    def export(self, english_only: bool = False, path: str | None = None, open_after: bool | None = None,
               register: bool | None = None) -> Path | None:
        fmt = self.state.fmt
        if not fmt:
            return None
        if not self.draft:
            QMessageBox.information(self, "Excelに出力", "先に［文面を生成］を押してください。")
            return None
        template = self.state.template_path()
        if path is None:
            s = self.state.settings
            name = report.make_filename(s.get("filename_pattern"), date.today(), self.phase.currentText().strip(),
                                        self.topic.text(), s.get("reporter", ""))
            if english_only:
                name += "_EN"
            folder = s.get("output_dir") or str(Path.home() / "Documents")
            path, _ = QFileDialog.getSaveFileName(self, "Excelに出力", str(Path(folder) / (name + template.suffix)),
                                                  f"Excel ブック (*{template.suffix})")
            if not path:
                return None
        values = report.cell_values(fmt, self.draft, english_only=english_only)
        try:
            warns = xw.write_cells(template, path, values)
        except PermissionError:
            QMessageBox.critical(self, "保存できません", "保存できませんでした。同じ名前のファイルを Excel で開いていないか確認してください。")
            return None
        except Exception as e:
            show_error(self, "出力に失敗しました", e)
            return None
        self.last_export = Path(path)
        msg = f"出力しました：{Path(path).name}"
        register = self.register.isChecked() if register is None else register
        if register and not english_only:
            pages = self.window()
            cases_page = getattr(pages, "cases_page", None)
            if cases_page is not None:
                cases_page.import_paths([path], policy="mask", origin="output")
            else:
                cases_mod.import_paths(fmt, self.state.cases, [path], origin="output")
                self.state.save_cases()
            msg += "（過去事例に登録）"
        self.status.setText(msg)
        if warns:
            QMessageBox.warning(self, "出力", "\n".join(warns))
        if self.open_after.isChecked() if open_after is None else open_after:
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except (AttributeError, OSError) as e:
                QMessageBox.warning(self, "出力", f"ファイルを開けませんでした。\n{e}")
        return Path(path)
