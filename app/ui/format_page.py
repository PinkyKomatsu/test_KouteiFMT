"""フォーマット設定画面：Excel の取り込み・解析結果の確認と修正（R1・R5・R7）。"""
from __future__ import annotations

import copy
import re
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QHeaderView,
                               QInputDialog, QLabel, QMessageBox, QPushButton, QSplitter, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from .. import analyzer, storage
from .. import xlsx_writer as xw
from .common import AppState, FormatSelector, show_error
from .sheet_preview import SheetPreview

KIND_ORDER = ["subject", "date", "phase", "status", "reporter", "dept", "body", "issue", "plan", "request", "list", "skip"]
LANG_ORDER = ["ja", "en", "both"]
COLS = ["欄名", "種別", "言語", "セル", "確信度", "使用", "要確認"]


class FormatPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.fmt: dict | None = None       # 編集中の定義（保存するまで state には反映しない）
        self.source: str | None = None     # 新しく取り込んだ原本（保存時にコピーする）
        self.template: Path | None = None
        self.dirty = False
        self._filling = False

        top = QHBoxLayout()
        self.import_btn = QPushButton("Excelフォーマットを取り込む")
        self.import_btn.clicked.connect(self.choose_file)
        top.addWidget(self.import_btn)
        top.addWidget(QLabel("フォーマット"))
        self.selector = FormatSelector(state)
        top.addWidget(self.selector)
        self.delete_btn = QPushButton("このフォーマットを削除")
        self.delete_btn.clicked.connect(self.delete_format)
        top.addWidget(self.delete_btn)
        top.addStretch(1)
        legend = QLabel("<span style='background:#dbe8ff'>&nbsp;ラベル&nbsp;</span> "
                        "<span style='background:#fff3b0'>&nbsp;記入欄&nbsp;</span> "
                        "<span style='background:#efe6ff'>&nbsp;英語欄&nbsp;</span> "
                        "<span style='border:2px solid #f08c00'>&nbsp;？ 要確認&nbsp;</span>")
        top.addWidget(legend)
        self.save_btn = QPushButton("定義を保存")
        self.save_btn.setStyleSheet("font-weight:bold")
        self.save_btn.clicked.connect(lambda: self.save())
        top.addWidget(self.save_btn)

        self.preview = SheetPreview()
        self.preview.cellClicked.connect(self._on_cell_clicked)
        self.preview.selectionChanged.connect(self._on_selection)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        self.table = QTableWidget(0, len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(self._on_item_changed)
        self.table.itemSelectionChanged.connect(self._on_row_selected)
        rl.addWidget(self.table)
        self.reason = QLabel("")
        self.reason.setWordWrap(True)
        self.reason.setStyleSheet("color:#b35c00")
        rl.addWidget(self.reason)
        self.sel_label = QLabel("プレビューでセルをクリック（ドラッグで範囲）して選びます")
        self.sel_label.setStyleSheet("color:#555")
        rl.addWidget(self.sel_label)
        btns = QHBoxLayout()
        self.assign_btn = QPushButton("この欄を記入先にする")
        self.assign_btn.setToolTip("一覧で選んだ欄の記入先を、プレビューで選んだセルに変えます")
        self.assign_btn.clicked.connect(lambda: self.assign_selected())
        self.add_btn = QPushButton("記入欄として追加")
        self.add_btn.setToolTip("プレビューで選んだセルを新しい記入欄にします（ラベルは左→上の順に探します）")
        self.add_btn.clicked.connect(lambda: self.add_selected())
        self.remove_btn = QPushButton("記入欄から外す")
        self.remove_btn.clicked.connect(lambda: self.remove_selected())
        for b in (self.assign_btn, self.add_btn, self.remove_btn):
            btns.addWidget(b)
        rl.addLayout(btns)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.preview)
        split.addWidget(right)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setSizes([900, 640])

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(split, 1)
        self.status = QLabel("")
        layout.addWidget(self.status)

        state.formatChanged.connect(self.load_from_state)
        self.load_from_state()

    # ---- 読み込み ----
    def load_from_state(self):
        if self.dirty and self.fmt and self.state.fmt and self.fmt["name"] == self.state.fmt["name"]:
            return
        self.fmt = copy.deepcopy(self.state.fmt) if self.state.fmt else None
        self.source = None
        self.template = self.state.template_path()
        self.dirty = False
        self.refresh(reload_preview=True)

    def refresh(self, reload_preview=False):
        if reload_preview:
            try:
                self.preview.load(self.template, self.fmt.get("sheets") if self.fmt else None)
            except Exception as e:
                show_error(self, "プレビューを表示できません", e)
        self.fill_table()
        self.preview.apply(self.fmt)
        self.delete_btn.setEnabled(self.fmt is not None and self.source is None)
        n_review = len(analyzer.review_fields(self.fmt)) if self.fmt else 0
        if self.fmt is None:
            self.status.setText("［Excelフォーマットを取り込む］から始めます。")
        else:
            state = "（未保存）" if self.dirty else ""
            self.status.setText(f"{self.fmt['name']}{state}：記入欄 {len(self.fmt['fields'])}個"
                                + (f"、要確認 {n_review}個" if n_review else ""))

    def fill_table(self):
        self._filling = True
        self.table.setRowCount(0)
        for f in (self.fmt or {}).get("fields", []):
            row = self.table.rowCount()
            self.table.insertRow(row)
            name = QTableWidgetItem(f["name"])
            name.setData(Qt.UserRole, f["id"])
            self.table.setItem(row, 0, name)
            kind = QComboBox()
            for k in KIND_ORDER:
                kind.addItem(analyzer.KIND_LABELS[k], k)
            kind.setCurrentIndex(KIND_ORDER.index(f["kind"]) if f["kind"] in KIND_ORDER else 6)
            kind.currentIndexChanged.connect(lambda _i, fid=f["id"], w=kind: self._set(fid, "kind", w.currentData()))
            self.table.setCellWidget(row, 1, kind)
            lang = QComboBox()
            for k in LANG_ORDER:
                lang.addItem(analyzer.LANG_LABELS[k], k)
            lang.setCurrentIndex(LANG_ORDER.index(f.get("lang", "ja")))
            lang.currentIndexChanged.connect(lambda _i, fid=f["id"], w=lang: self._set(fid, "lang", w.currentData()))
            self.table.setCellWidget(row, 2, lang)
            cells = ", ".join(f["cells"]) if len(f["cells"]) <= 3 else f"{f['cells'][0]}:{f['cells'][-1]}（{len(f['cells'])}行）"
            self.table.setItem(row, 3, QTableWidgetItem(cells or "（未指定）"))
            conf = QTableWidgetItem(f"{f.get('confidence', 1):.2f}")
            conf.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(row, 4, conf)
            use = QCheckBox()
            use.setChecked(bool(f.get("enabled")))
            use.toggled.connect(lambda v, fid=f["id"]: self._set(fid, "enabled", v))
            self.table.setCellWidget(row, 5, use)
            review = bool(f.get("review") or not f.get("cells"))
            rv = QTableWidgetItem("？ " + "／".join(f.get("reasons") or ["記入先が未確定"]) if review else "")
            rv.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            if review:
                rv.setForeground(QColor("#b35c00"))
            self.table.setItem(row, 6, rv)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._filling = False

    # ---- 取り込み ----
    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excelフォーマットを選択", "",
                                              "Excel ブック (*.xlsx *.xlsm);;Excel 97-2003 (*.xls);;すべて (*.*)")
        if not path:
            return
        name, ok = QInputDialog.getText(self, "フォーマット名", "このフォーマットの名前を入力してください。",
                                        text=Path(path).stem)
        if ok and name.strip():
            self.import_file(path, name.strip())

    def import_file(self, path, name: str) -> bool:
        """Excel を解析して編集中の定義にする（保存するまで登録しない）。"""
        try:
            xw.check_supported(path)
            fmt = analyzer.analyze(path, name)
        except xw.XlsNotSupported as e:
            QMessageBox.warning(self, "取り込めません", str(e))
            return False
        except Exception as e:
            show_error(self, "フォーマットを解析できませんでした", e)
            return False
        self.fmt, self.source, self.template, self.dirty = fmt, str(path), Path(path), True
        self.refresh(reload_preview=True)
        n_review = len(analyzer.review_fields(fmt))
        msg = f"{len(fmt['fields'])}個の欄を検出しました。"
        if n_review:
            msg += f"\n要確認の欄が {n_review}個あります（オレンジの枠）。プレビューでセルを選び、［この欄を記入先にする］で指定できます。"
        self.status.setText(msg.replace("\n", " "))
        return True

    # ---- 編集 ----
    def _field(self, fid) -> dict | None:
        return next((f for f in (self.fmt or {}).get("fields", []) if f["id"] == fid), None)

    def _current_field(self) -> dict | None:
        row = self.table.currentRow()
        item = self.table.item(row, 0) if row >= 0 else None
        return self._field(item.data(Qt.UserRole)) if item else None

    def _set(self, fid, key, value):
        f = self._field(fid)
        if f is None or self._filling:
            return
        f[key] = value
        if key == "kind" and value == "skip":
            f["enabled"] = False
        if key in ("kind", "lang"):
            analyzer.detect_pairs(self.fmt["fields"])
        self.dirty = True
        self.refresh()

    def _on_item_changed(self, item):
        if self._filling or self.fmt is None:
            return
        f = self._field(self.table.item(item.row(), 0).data(Qt.UserRole))
        if f is None:
            return
        if item.column() == 0 and item.text().strip():
            f["name"] = item.text().strip()
        elif item.column() == 3:
            refs = [r.strip().upper().replace("$", "") for r in re.split(r"[,\s、]+", item.text()) if r.strip()]
            try:
                for r in refs:
                    xw.split_ref(r)
                if refs:
                    analyzer.assign_cells(self.template, f, refs)
            except xw.WriteError:
                QMessageBox.warning(self, "セル番地", "セル番地が正しくありません（例：B3 または B24, B25）")
        self.dirty = True
        self.refresh()

    def _on_row_selected(self):
        f = self._current_field()
        if f is None:
            self.reason.setText("")
            return
        self.reason.setText("／".join(f.get("reasons") or []))
        ref = (f["cells"] or [f.get("label_cell")])[0]
        if ref:
            self.preview.focus(f["sheet"], ref)

    def _on_cell_clicked(self, sheet, ref):
        for row in range(self.table.rowCount()):
            f = self._field(self.table.item(row, 0).data(Qt.UserRole))
            if f and f["sheet"] == sheet and (ref in f["cells"] or ref == f.get("label_cell")):
                self.table.blockSignals(True)
                self.table.selectRow(row)
                self.table.blockSignals(False)
                self.reason.setText("／".join(f.get("reasons") or []))
                break

    def _on_selection(self, sheet, refs):
        self.sel_label.setText(f"選択中：{sheet}!{', '.join(refs[:6])}{' …' if len(refs) > 6 else ''}" if refs
                               else "プレビューでセルをクリック（ドラッグで範囲）して選びます")

    def assign_selected(self, field: dict | None = None, refs: list[str] | None = None):
        field = field or self._current_field()
        sheet, sel = self.preview.selected_refs()
        refs = refs or sel
        if field is None or not refs:
            QMessageBox.information(self, "記入先の指定", "右の一覧で欄を選び、プレビューで記入先のセルを選んでください。")
            return
        field["sheet"] = sheet or field["sheet"]
        analyzer.assign_cells(self.template, field, refs)
        self.dirty = True
        self.refresh()

    def add_selected(self, sheet: str | None = None, refs: list[str] | None = None):
        if self.fmt is None:
            return
        if refs is None:
            sheet, refs = self.preview.selected_refs()
        if not refs:
            QMessageBox.information(self, "記入欄の追加", "プレビューで記入欄にしたいセルを選んでください。")
            return
        used = {r for f in self.fmt["fields"] if f["sheet"] == sheet for r in f["cells"]}
        if any(r in used for r in refs):
            QMessageBox.information(self, "記入欄の追加", "選んだセルはすでに記入欄です。［この欄を記入先にする］を使ってください。")
            return
        field = analyzer.field_from_cells(self.template, sheet, refs, self.fmt["fields"])
        self.fmt["fields"].append(field)
        analyzer.detect_pairs(self.fmt["fields"])
        self.dirty = True
        self.refresh()

    def remove_selected(self):
        if self.fmt is None:
            return
        sheet, refs = self.preview.selected_refs()
        field = None
        if refs:
            for f in self.fmt["fields"]:
                if f["sheet"] == sheet and any(r in f["cells"] for r in refs):
                    field = f
                    f["cells"] = [c for c in f["cells"] if c not in refs]
                    break
        else:
            field = self._current_field()
            if field:
                field["cells"] = []
        if field is None:
            return
        if not field["cells"]:
            self.fmt["fields"].remove(field)
        analyzer.detect_pairs(self.fmt["fields"])
        self.dirty = True
        self.refresh()

    # ---- 保存・削除 ----
    def save(self, confirm: bool = True) -> bool:
        if self.fmt is None:
            return False
        pending = analyzer.review_fields(self.fmt)
        if pending and confirm:
            names = "、".join(f["name"] for f in pending[:5])
            ans = QMessageBox.question(self, "要確認の欄があります",
                                       f"要確認の欄が {len(pending)}個残っています（{names}）。\nこのまま保存しますか？")
            if ans != QMessageBox.Yes:
                return False
        if self.source and self.fmt["name"] in storage.list_formats() and confirm:
            ans = QMessageBox.question(self, "上書きの確認", f"フォーマット「{self.fmt['name']}」を置き換えますか？（過去事例は残ります）")
            if ans != QMessageBox.Yes:
                return False
        try:
            source = self.source
            self.dirty = False
            self.state.save_format(self.fmt, template_src=source)
        except Exception as e:
            show_error(self, "保存できませんでした", e)
            return False
        if self.state.cases and confirm:
            if QMessageBox.question(self, "過去事例", "定義を保存しました。過去事例を新しい定義で読み込み直しますか？") == QMessageBox.Yes:
                self.window().cases_page.reload()
        return True

    def delete_format(self):
        if self.state.fmt is None:
            return
        name = self.state.fmt["name"]
        if QMessageBox.question(self, "削除の確認", f"フォーマット「{name}」と、その過去事例を削除しますか？") != QMessageBox.Yes:
            return
        storage.delete_format(name)
        names = self.state.format_names()
        self.state.formatsChanged.emit()
        self.state.set_format(names[0] if names else None)
