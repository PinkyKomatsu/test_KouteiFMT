"""メインウィンドウ：報告書作成 / フォーマット設定 / 過去事例 / 設定 の 4 タブ。"""
from __future__ import annotations

from PySide6.QtWidgets import QLabel, QMainWindow, QTabWidget

from .. import storage
from .cases_page import CasesPage
from .common import AppState
from .compose_page import ComposePage
from .format_page import FormatPage
from .settings_page import SettingsPage

APP_TITLE = "医事会計システム更新 工程報告書作成"


class MainWindow(QMainWindow):
    def __init__(self, state: AppState | None = None):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1500, 920)
        self.state = state or AppState()
        self.tabs = QTabWidget()
        self.compose_page = ComposePage(self.state)
        self.format_page = FormatPage(self.state)
        self.cases_page = CasesPage(self.state)
        self.settings_page = SettingsPage(self.state)
        self.tabs.addTab(self.compose_page, "報告書作成")
        self.tabs.addTab(self.format_page, "フォーマット設定")
        self.tabs.addTab(self.cases_page, "過去事例")
        self.tabs.addTab(self.settings_page, "設定")
        self.setCentralWidget(self.tabs)
        self.model_label = QLabel("")
        self.statusBar().addPermanentWidget(self.model_label)
        self.statusBar().showMessage(f"データ：{storage.data_dir()}")
        self._update_model_label()
        self.state.settingsChanged.connect(self._update_model_label)

        names = self.state.format_names()
        last = self.state.settings.get("last_format")
        if names:
            self.state.set_format(last if last in names else names[0])
        else:
            self.tabs.setCurrentWidget(self.compose_page)

    def _update_model_label(self):
        model = storage.model_dir(self.state.settings)
        ok = (model / "model.bin").exists()
        self.model_label.setText("翻訳モデル：あり（オフライン）" if ok else "翻訳モデル：なし（簡易訳で動作）")
        self.model_label.setStyleSheet("color:#2b7a3d" if ok else "color:#b35c00")

    def closeEvent(self, e):
        page = self.format_page
        if page.dirty and page.fmt is not None:
            from PySide6.QtWidgets import QMessageBox
            ans = QMessageBox.question(self, APP_TITLE, "フォーマットの定義が保存されていません。閉じてもよいですか？")
            if ans != QMessageBox.Yes:
                e.ignore()
                return
        super().closeEvent(e)
