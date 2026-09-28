"""過去事例画面：過去の報告書の取り込み・再読込・削除（R4）。個人情報が見つかったら扱いを選ばせる。"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QFileDialog, QHBoxLayout, QHeaderView, QLabel,
                               QMessageBox, QPlainTextEdit, QPushButton, QSplitter, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget)

from .. import cases as cases_mod
from .common import AppState, FormatSelector, show_error


class CasesPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.pii_policy: str | None = None   # "mask" / "skip" / None（毎回たずねる）

        top = QHBoxLayout()
        top.addWidget(QLabel("フォーマット"))
        top.addWidget(FormatSelector(state))
        for text, slot in (("ファイルを追加", self.add_files), ("フォルダを追加", self.add_folder),
                           ("再読込", self.reload), ("削除", self.remove)):
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, s=slot: s())
            top.addWidget(b)
        top.addStretch(1)
        self.count = QLabel("")
        top.addWidget(self.count)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["件名", "工程", "日付", "ファイル", "登録元", "注意"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.itemSelectionChanged.connect(self._show_preview)
        self.preview = QPlainTextEdit()
        self.preview.setReadOnly(True)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(2000)

        upper = QSplitter(Qt.Horizontal)
        upper.addWidget(self.table)
        upper.addWidget(self.preview)
        upper.setStretchFactor(0, 3)
        upper.setStretchFactor(1, 2)
        split = QSplitter(Qt.Vertical)
        split.addWidget(upper)
        logbox = QWidget()
        ll = QVBoxLayout(logbox)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.addWidget(QLabel("取り込みログ"))
        ll.addWidget(self.log)
        split.addWidget(logbox)
        split.setStretchFactor(0, 4)
        split.setStretchFactor(1, 1)

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(split, 1)
        state.casesChanged.connect(self.refresh)
        self.refresh()

    def refresh(self):
        fmt = self.state.fmt
        self.table.setRowCount(0)
        if fmt is None:
            self.count.setText("フォーマットが未登録です")
            return
        for c in self.state.cases:
            row = self.table.rowCount()
            self.table.insertRow(row)
            notes = []
            if c.get("masked"):
                notes.append("伏せ字あり")
            if c.get("warnings"):
                notes.append(f"警告{len(c['warnings'])}")
            values = [cases_mod.subject_of(fmt, c), cases_mod.phase_of(fmt, c), cases_mod.date_of(fmt, c),
                      Path(c["source"]).name, "出力" if c.get("origin") == "output" else "取込", "・".join(notes)]
            for col, v in enumerate(values):
                item = QTableWidgetItem(v)
                item.setData(Qt.UserRole, c["id"])
                self.table.setItem(row, col, item)
        self.count.setText(f"{len(self.state.cases)}件（対訳 {len(cases_mod.translation_memory(self.state.cases))}組）")

    def write_log(self, lines):
        stamp = datetime.now().strftime("%H:%M:%S")
        for l in lines:
            self.log.appendPlainText(f"[{stamp}] {l}")

    # ---- 個人情報 ----
    def ask_pii(self, path: Path, findings) -> str:
        if self.pii_policy:
            return self.pii_policy
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle("個人情報らしき記載があります")
        detail = "\n".join("・" + f.describe() for f in findings[:10]) + ("\n…" if len(findings) > 10 else "")
        box.setText(f"「{path.name}」に個人情報らしき記載が {len(findings)}件あります。\n\n{detail}\n\nどうしますか？")
        mask = box.addButton("伏せ字にして取り込む", QMessageBox.AcceptRole)
        skip = box.addButton("この報告書を取り込まない", QMessageBox.RejectRole)
        same = QCheckBox("以降のファイルも同じようにする")
        box.setCheckBox(same)
        box.exec()
        choice = "skip" if box.clickedButton() is skip else "mask"
        if same.isChecked():
            self.pii_policy = choice
        return choice

    # ---- 操作 ----
    def _need_format(self) -> bool:
        if self.state.fmt is None:
            QMessageBox.information(self, "過去事例", "先に「フォーマット設定」でフォーマットを取り込んで保存してください。")
            return False
        return True

    def add_files(self):
        if not self._need_format():
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "過去の報告書を選択", "", "Excel ブック (*.xlsx *.xlsm);;すべて (*.*)")
        if paths:
            self.import_paths(paths)

    def add_folder(self):
        if not self._need_format():
            return
        d = QFileDialog.getExistingDirectory(self, "過去の報告書のフォルダを選択（サブフォルダも含みます）")
        if d:
            self.import_paths([d])

    def import_paths(self, paths, policy: str | None = None, origin: str = "import") -> list[str]:
        self.pii_policy = policy
        try:
            log = cases_mod.import_paths(self.state.fmt, self.state.cases, paths, origin=origin, on_pii=self.ask_pii)
        except Exception as e:
            show_error(self, "取り込みに失敗しました", e)
            return []
        finally:
            self.pii_policy = None
        self.state.save_cases()
        self.write_log(log)
        return log

    def reload(self):
        if self.state.fmt is None:
            return
        log = cases_mod.reload_all(self.state.fmt, self.state.cases, on_pii=self.ask_pii)
        self.state.save_cases()
        self.write_log(log or ["再読込する事例がありません"])

    def remove(self):
        ids = {self.table.item(i.row(), 0).data(Qt.UserRole) for i in self.table.selectedIndexes()}
        if not ids:
            return
        if QMessageBox.question(self, "削除の確認", f"選んだ {len(ids)}件の事例を削除しますか？（元の Excel は削除しません）") != QMessageBox.Yes:
            return
        self.state.cases[:] = [c for c in self.state.cases if c["id"] not in ids]
        self.state.save_cases()

    def _show_preview(self):
        rows = {i.row() for i in self.table.selectedIndexes()}
        if not rows:
            self.preview.clear()
            return
        cid = self.table.item(min(rows), 0).data(Qt.UserRole)
        case = next((c for c in self.state.cases if c["id"] == cid), None)
        if case:
            self.preview.setPlainText(f"{case['source']}\n\n" + cases_mod.case_preview(self.state.fmt, case))
