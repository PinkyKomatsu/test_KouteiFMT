"""画面で共有する状態・バックグラウンド処理・色。"""
from __future__ import annotations

import copy
import logging
import traceback

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal
from PySide6.QtWidgets import QComboBox, QMessageBox, QWidget

from .. import cases as cases_mod
from .. import glossary, storage, translator_en

log = logging.getLogger("medacct")

# 背景色：流用＝黄、要確認＝オレンジ、手で編集＝白
COLORS = {
    "reuse": "#fff3b0", "review": "#ffd8a8", "simple": "#ffd8a8",
    "auto": "#e6f4ea", "template": "#e6f4ea", "tm": "#e6f4ea",
    "mt": "#e8f0fe", "memo": "#ffffff", "edited": "#ffffff", "empty": "#ffffff", "copy": "#ffffff",
}
PREVIEW_LABEL = "#dbe8ff"   # プレビュー：ラベル（青）
PREVIEW_FIELD = "#fff3b0"   # プレビュー：記入欄（黄）
PREVIEW_EN = "#efe6ff"      # プレビュー：英語欄
PREVIEW_OFF = "#eeeeee"     # プレビュー：使用しない欄
REVIEW_PEN = "#f08c00"      # 要確認の枠（オレンジ）


def entry_color(status: str, review: bool) -> str:
    if status == "reuse":
        return COLORS["reuse"]
    if status == "edited":
        return COLORS["edited"]
    if review or status == "simple":
        return COLORS["review"]
    return COLORS.get(status, "#ffffff")


def show_error(parent: QWidget | None, title: str, err) -> None:
    log.error("%s: %s", title, err if isinstance(err, str) else "".join(traceback.format_exception(err)))
    QMessageBox.critical(parent, title, str(err))


class AppState(QObject):
    """アプリ全体の状態（設定・選択中のフォーマット・過去事例・翻訳器）。"""
    formatsChanged = Signal()
    formatChanged = Signal()
    casesChanged = Signal()
    settingsChanged = Signal()

    def __init__(self):
        super().__init__()
        self.settings = storage.load_settings()
        self.fmt: dict | None = None
        self.cases: list[dict] = []
        self.index = cases_mod.Index({"fields": []}, [])
        self.terms = glossary.load()
        self._translator: translator_en.Translator | None = None

    # ---- フォーマット ----
    def format_names(self) -> list[str]:
        return storage.list_formats()

    def set_format(self, name: str | None) -> None:
        self.fmt = storage.load_format(name) if name else None
        self.cases = storage.load_cases(name) if self.fmt else []
        if self.fmt:
            self.settings["last_format"] = self.fmt["name"]
            storage.save_settings(self.settings)
        self._rebuild()
        self.formatChanged.emit()
        self.casesChanged.emit()

    def save_format(self, fmt: dict, template_src=None) -> None:
        storage.save_format(fmt, template_src=template_src)
        self.formatsChanged.emit()
        self.set_format(fmt["name"])

    def template_path(self):
        return storage.template_path(self.fmt) if self.fmt else None

    # ---- 過去事例 ----
    def save_cases(self) -> None:
        if self.fmt is None:
            return
        storage.save_cases(self.fmt["name"], self.cases)
        self._rebuild()
        self.casesChanged.emit()

    def _rebuild(self) -> None:
        self.index = cases_mod.Index(self.fmt or {"fields": []}, self.cases)
        if self._translator is not None:
            self._translator.set_memory(cases_mod.translation_memory(self.cases))

    # ---- 設定・翻訳 ----
    def save_settings(self) -> None:
        storage.save_settings(self.settings)
        self._translator = None   # モデルのパスが変わっている可能性がある
        self.settingsChanged.emit()

    def save_glossary(self, terms) -> None:
        glossary.save(terms)
        self.terms = glossary.load()
        self._translator = None

    def translator(self) -> translator_en.Translator:
        if self._translator is None:
            self._translator = translator_en.Translator(storage.model_dir(self.settings), self.terms,
                                                        cases_mod.translation_memory(self.cases))
        return self._translator

    def settings_copy(self) -> dict:
        return copy.deepcopy(self.settings)


# ---------------------------------------------------------------------------
# バックグラウンド処理（翻訳などで UI を固めない）
# ---------------------------------------------------------------------------

class _Signals(QObject):
    done = Signal(object)
    failed = Signal(object)


class Worker(QRunnable):
    def __init__(self, fn):
        super().__init__()
        self.fn = fn
        self.signals = _Signals()

    def run(self):
        try:
            result = self.fn()
        except Exception as e:  # 画面側でダイアログを出す
            log.exception("バックグラウンド処理でエラー")
            self.signals.failed.emit(e)
        else:
            self.signals.done.emit(result)


_active: set[Worker] = set()


def run_async(fn, on_done, on_failed) -> Worker:
    w = Worker(fn)
    _active.add(w)

    def finish(result):
        _active.discard(w)
        on_done(result)

    def fail(err):
        _active.discard(w)
        on_failed(err)

    w.signals.done.connect(finish)
    w.signals.failed.connect(fail)
    QThreadPool.globalInstance().start(w)
    return w


class FormatSelector(QComboBox):
    """選択中のフォーマットを切り替えるコンボボックス（どの画面でも同じ状態を指す）。"""

    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.state = state
        self.setMinimumWidth(220)
        self._updating = False
        self.currentTextChanged.connect(self._changed)
        state.formatsChanged.connect(self.refresh)
        state.formatChanged.connect(self.refresh)
        self.refresh()

    def refresh(self):
        self._updating = True
        self.clear()
        names = self.state.format_names()
        self.addItems(names)
        if self.state.fmt and self.state.fmt["name"] in names:
            self.setCurrentText(self.state.fmt["name"])
        self.setEnabled(bool(names))
        self._updating = False

    def _changed(self, name):
        if self._updating or not name:
            return
        if not self.state.fmt or self.state.fmt["name"] != name:
            self.state.set_format(name)
