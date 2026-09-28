"""設定画面：報告者・日付書式・工程の候補・流用・出力先・用語集・翻訳モデル。"""
from __future__ import annotations

from PySide6.QtWidgets import (QCheckBox, QComboBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QPushButton,
                               QScrollArea, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from .. import storage, translator_en
from .common import AppState, run_async, show_error

DATE_FORMATS = [("auto", "自動（過去事例に合わせる）"), ("yyyy/MM/dd", "yyyy/MM/dd（2026/09/27）"),
                ("yyyy年M月d日", "yyyy年M月d日（2026年9月27日）"), ("yyyy/M/d", "yyyy/M/d（2026/9/27）")]
DATE_FORMATS_EN = [("MMM d, yyyy", "MMM d, yyyy（Sep 27, 2026）"), ("MMMM d, yyyy", "MMMM d, yyyy（September 27, 2026）"),
                   ("yyyy-MM-dd", "yyyy-MM-dd（2026-09-27）")]


class SettingsPage(QWidget):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        s = state.settings

        form = QFormLayout()
        self.reporter = QLineEdit(s.get("reporter", ""))
        self.reporter.setPlaceholderText("空欄なら過去事例で最も多い値")
        self.dept = QLineEdit(s.get("department", ""))
        self.date_fmt = QComboBox()
        for k, label in DATE_FORMATS:
            self.date_fmt.addItem(label, k)
        self.date_fmt.setCurrentIndex(max(0, [k for k, _ in DATE_FORMATS].index(s.get("date_format", "auto"))
                                          if s.get("date_format", "auto") in [k for k, _ in DATE_FORMATS] else 0))
        self.date_fmt_en = QComboBox()
        for k, label in DATE_FORMATS_EN:
            self.date_fmt_en.addItem(label, k)
        keys_en = [k for k, _ in DATE_FORMATS_EN]
        self.date_fmt_en.setCurrentIndex(keys_en.index(s.get("date_format_en")) if s.get("date_format_en") in keys_en else 0)
        self.phases = QPlainTextEdit("\n".join(s.get("phases") or storage.DEFAULT_PHASES))
        self.phases.setFixedHeight(120)
        self.reuse = QCheckBox("メモのない欄に過去事例の文面を流用する（要確認として黄色で表示）")
        self.reuse.setChecked(bool(s.get("reuse", True)))
        self.output_dir = QLineEdit(s.get("output_dir", ""))
        self.output_dir.setPlaceholderText("空欄ならドキュメント")
        out_row = QHBoxLayout()
        out_row.addWidget(self.output_dir)
        browse = QPushButton("参照")
        browse.clicked.connect(self._choose_output)
        out_row.addWidget(browse)
        self.pattern = QLineEdit(s.get("filename_pattern", "{date}_{phase}_{topic}"))
        self.register = QCheckBox("出力した報告書を過去事例に登録する（既定）")
        self.register.setChecked(bool(s.get("register_output", True)))
        self.open_after = QCheckBox("出力後に Excel で開く（既定）")
        self.open_after.setChecked(bool(s.get("open_after_export", True)))

        form.addRow("報告者", self.reporter)
        form.addRow("部署", self.dept)
        form.addRow("日付書式（日本語）", self.date_fmt)
        form.addRow("日付書式（英語）", self.date_fmt_en)
        form.addRow("工程の候補（1行に1つ）", self.phases)
        form.addRow("", self.reuse)
        form.addRow("出力先フォルダ", out_row)
        form.addRow("ファイル名パターン", self.pattern)
        form.addRow("", QLabel("{date}（例 20260927）・{phase}・{topic}・{reporter} が使えます"))
        form.addRow("", self.register)
        form.addRow("", self.open_after)

        # 翻訳モデル
        model_box = QGroupBox("翻訳モデル（CTranslate2 / opus-mt-ja-en int8）")
        ml = QVBoxLayout(model_box)
        row = QHBoxLayout()
        self.model_path = QLineEdit(s.get("model_path", ""))
        self.model_path.setPlaceholderText(f"空欄なら {storage.default_model_dir()}")
        row.addWidget(self.model_path)
        mb = QPushButton("参照")
        mb.clicked.connect(self._choose_model)
        row.addWidget(mb)
        self.test_btn = QPushButton("動作確認")
        self.test_btn.clicked.connect(self.test_model)
        row.addWidget(self.test_btn)
        ml.addLayout(row)
        self.model_status = QLabel("")
        self.model_status.setWordWrap(True)
        ml.addWidget(self.model_status)

        # 用語集
        gl_box = QGroupBox("医事会計の用語集（翻訳時に優先して使います）")
        gl = QVBoxLayout(gl_box)
        self.glossary = QTableWidget(0, 2)
        self.glossary.setHorizontalHeaderLabels(["日本語", "英語"])
        self.glossary.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.glossary.verticalHeader().setVisible(False)
        self.glossary.setMinimumHeight(260)
        gl.addWidget(self.glossary)
        gb = QHBoxLayout()
        add = QPushButton("行を追加")
        add.clicked.connect(self._add_term)
        rm = QPushButton("選んだ行を削除")
        rm.clicked.connect(self._remove_terms)
        save_gl = QPushButton("用語集を保存")
        save_gl.clicked.connect(self.save_glossary)
        gb.addWidget(add)
        gb.addWidget(rm)
        gb.addStretch(1)
        gb.addWidget(save_gl)
        gl.addLayout(gb)
        self._fill_glossary()

        save = QPushButton("設定を保存")
        save.setStyleSheet("font-weight:bold")
        save.clicked.connect(self.save)
        info = QLabel(f"データの保存先：{storage.data_dir()}\nログ：{storage.log_dir() / 'app.log'}")
        info.setStyleSheet("color:#666")

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.addLayout(form)
        layout.addWidget(save)
        layout.addWidget(model_box)
        layout.addWidget(gl_box)
        layout.addWidget(info)
        layout.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    # ---- 設定 ----
    def _choose_output(self):
        d = QFileDialog.getExistingDirectory(self, "出力先フォルダ")
        if d:
            self.output_dir.setText(d)

    def _choose_model(self):
        d = QFileDialog.getExistingDirectory(self, "翻訳モデルのフォルダ（model.bin があるフォルダ）")
        if d:
            self.model_path.setText(d)

    def save(self):
        s = self.state.settings
        s["reporter"] = self.reporter.text().strip()
        s["department"] = self.dept.text().strip()
        s["date_format"] = self.date_fmt.currentData()
        s["date_format_en"] = self.date_fmt_en.currentData()
        s["phases"] = [p.strip() for p in self.phases.toPlainText().splitlines() if p.strip()] or list(storage.DEFAULT_PHASES)
        s["reuse"] = self.reuse.isChecked()
        s["output_dir"] = self.output_dir.text().strip()
        s["filename_pattern"] = self.pattern.text().strip() or "{date}_{phase}_{topic}"
        s["register_output"] = self.register.isChecked()
        s["open_after_export"] = self.open_after.isChecked()
        s["model_path"] = self.model_path.text().strip()
        try:
            self.state.save_settings()
        except Exception as e:
            show_error(self, "設定を保存できませんでした", e)
            return
        QMessageBox.information(self, "設定", "設定を保存しました。")

    def test_model(self):
        path = self.model_path.text().strip() or str(storage.default_model_dir())
        t = translator_en.Translator(path, self.state.terms)
        self.test_btn.setEnabled(False)
        self.model_status.setText("確認しています…")

        def done(msg):
            self.test_btn.setEnabled(True)
            self.model_status.setText(msg)

        def failed(err):
            self.test_btn.setEnabled(True)
            self.model_status.setText(f"確認できませんでした：{err}")

        run_async(t.self_test, done, failed)

    # ---- 用語集 ----
    def _fill_glossary(self):
        self.glossary.setRowCount(0)
        for ja, en in self.state.terms:
            row = self.glossary.rowCount()
            self.glossary.insertRow(row)
            self.glossary.setItem(row, 0, QTableWidgetItem(ja))
            self.glossary.setItem(row, 1, QTableWidgetItem(en))

    def _add_term(self):
        row = self.glossary.rowCount()
        self.glossary.insertRow(row)
        self.glossary.setItem(row, 0, QTableWidgetItem(""))
        self.glossary.setItem(row, 1, QTableWidgetItem(""))
        self.glossary.editItem(self.glossary.item(row, 0))

    def _remove_terms(self):
        for row in sorted({i.row() for i in self.glossary.selectedIndexes()}, reverse=True):
            self.glossary.removeRow(row)

    def save_glossary(self):
        terms = []
        for row in range(self.glossary.rowCount()):
            ja = (self.glossary.item(row, 0) or QTableWidgetItem("")).text().strip()
            en = (self.glossary.item(row, 1) or QTableWidgetItem("")).text().strip()
            if ja and en:
                terms.append((ja, en))
        try:
            self.state.save_glossary(terms)
        except Exception as e:
            show_error(self, "用語集を保存できませんでした", e)
            return
        self._fill_glossary()
        QMessageBox.information(self, "用語集", f"用語集を保存しました（{len(terms)}語）。")
