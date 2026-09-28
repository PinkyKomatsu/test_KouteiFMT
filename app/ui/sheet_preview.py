"""Excel 風のプレビュー（QTableWidget）。

列幅・行高・結合セル・背景色・太字・罫線の有無を再現する。シートが複数あればタブで切り替える。
ラベルは青、記入欄は黄（欄名または記入予定の値を表示）、要確認の欄はオレンジの枠と「？」で示す。
クリック・ドラッグで選んだセルを cellClicked / selectionChanged で通知する。
"""
from __future__ import annotations

from PySide6.QtCore import QRect, Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QPen
from PySide6.QtWidgets import (QAbstractItemView, QStyledItemDelegate, QTableWidget, QTableWidgetItem,
                               QTabWidget)

from .. import analyzer
from .. import xlsx_writer as xw
from .common import PREVIEW_EN, PREVIEW_FIELD, PREVIEW_LABEL, PREVIEW_OFF, REVIEW_PEN

BORDER_ROLE = Qt.UserRole + 1
REVIEW_ROLE = Qt.UserRole + 2
REF_ROLE = Qt.UserRole + 3


class _CellDelegate(QStyledItemDelegate):
    """罫線（枠あり＝濃い線、なし＝薄い格子）と要確認の印を描く。"""

    def paint(self, painter, option, index):
        super().paint(painter, option, index)
        r = option.rect.adjusted(0, 0, -1, -1)
        painter.save()
        painter.setPen(QPen(QColor("#555555" if index.data(BORDER_ROLE) else "#e4e4e4"), 1))
        painter.drawRect(r)
        if index.data(REVIEW_ROLE):
            painter.setPen(QPen(QColor(REVIEW_PEN), 2))
            painter.drawRect(r.adjusted(1, 1, -1, -1))
            f = QFont(option.font)
            f.setBold(True)
            painter.setFont(f)
            painter.drawText(QRect(r.right() - 18, r.top() + 1, 16, 16), Qt.AlignCenter, "？")
        painter.restore()


class SheetView(QTableWidget):
    refClicked = Signal(str)

    def __init__(self, layout: dict, parent=None):
        super().__init__(layout["n_rows"], layout["n_cols"], parent)
        self.layout_info = layout
        self.sheet = layout["sheet"]
        self.setItemDelegate(_CellDelegate(self))
        self.setShowGrid(False)
        self.setWordWrap(True)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self.setHorizontalHeaderLabels([xw.index_to_col(i + 1) for i in range(layout["n_cols"])])
        self.setVerticalHeaderLabels([str(i + 1) for i in range(layout["n_rows"])])
        self.verticalHeader().setDefaultAlignment(Qt.AlignCenter)
        for i, px in enumerate(layout["col_px"]):
            self.setColumnWidth(i, max(px, 1))
            self.setColumnHidden(i, px == 0)
        for i, px in enumerate(layout["row_px"]):
            self.setRowHeight(i, max(px, 1))
            self.setRowHidden(i, px == 0)

        self.pos_of: dict[str, tuple[int, int]] = {}
        self.owner: dict[tuple[int, int], str] = {}
        self.base: dict[str, dict] = {}
        for c in layout["cells"]:
            r, col = c["r1"] - 1, c["c1"] - 1
            item = QTableWidgetItem(c["text"])
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            item.setData(BORDER_ROLE, c["border"])
            item.setData(REF_ROLE, c["ref"])
            if c["fill"]:
                item.setBackground(QBrush(QColor(c["fill"])))
            font = QFont(self.font())
            font.setBold(c["bold"])
            if c["size"] >= 14:
                font.setPointSizeF(min(c["size"] * 0.8, 16))
            item.setFont(font)
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop if c["r2"] > c["r1"] or "\n" in c["text"]
                                  else Qt.AlignLeft | Qt.AlignVCenter)
            self.setItem(r, col, item)
            rows, cols = c["r2"] - c["r1"] + 1, c["c2"] - c["c1"] + 1
            if rows > 1 or cols > 1:
                self.setSpan(r, col, rows, cols)
            self.pos_of[c["ref"]] = (r, col)
            for rr in range(c["r1"], c["r2"] + 1):
                for cc in range(c["c1"], c["c2"] + 1):
                    self.owner[(rr - 1, cc - 1)] = c["ref"]
            self.base[c["ref"]] = {"text": c["text"], "fill": c["fill"]}
        self.cellClicked.connect(self._clicked)

    def _clicked(self, row, col):
        ref = self.owner.get((row, col))
        if ref:
            self.refClicked.emit(ref)

    def item_at_ref(self, ref: str) -> QTableWidgetItem | None:
        pos = self.pos_of.get(ref)
        return self.item(*pos) if pos else None

    def reset(self):
        for ref, b in self.base.items():
            item = self.item_at_ref(ref)
            if item is None:
                continue
            item.setText(b["text"])
            item.setBackground(QBrush(QColor(b["fill"])) if b["fill"] else QBrush())
            item.setData(REVIEW_ROLE, False)
            item.setToolTip("")
            item.setForeground(QBrush(QColor("#000000")))

    def mark(self, ref: str, text: str | None = None, fill: str | None = None, review: bool = False,
             tooltip: str = "", color: str | None = None):
        item = self.item_at_ref(ref)
        if item is None:
            return
        if text is not None:
            item.setText(text)
        if fill:
            item.setBackground(QBrush(QColor(fill)))
        if color:
            item.setForeground(QBrush(QColor(color)))
        item.setData(REVIEW_ROLE, review)
        if tooltip:
            item.setToolTip(tooltip)

    def selected_refs(self) -> list[str]:
        refs = []
        for idx in self.selectedIndexes():
            ref = self.owner.get((idx.row(), idx.column()))
            if ref and ref not in refs:
                refs.append(ref)
        return sorted(refs, key=lambda r: (xw.split_ref(r)[1], xw.split_ref(r)[0]))

    def focus_ref(self, ref: str):
        pos = self.pos_of.get(ref)
        if pos:
            self.setCurrentCell(*pos)
            self.scrollToItem(self.item(*pos), QAbstractItemView.PositionAtCenter)


class SheetPreview(QTabWidget):
    cellClicked = Signal(str, str)        # シート名, セル
    selectionChanged = Signal(str, list)  # シート名, セルの一覧

    def __init__(self, parent=None):
        super().__init__(parent)
        self.views: dict[str, SheetView] = {}
        self.template = None

    def load(self, template, sheets: list[str] | None = None):
        self.clear()
        self.views.clear()
        self.template = template
        if template is None:
            return
        names = sheets or analyzer.load_book(template).sheetnames
        for name in names:
            view = SheetView(analyzer.sheet_layout(template, name))
            view.refClicked.connect(lambda ref, s=name: self.cellClicked.emit(s, ref))
            view.itemSelectionChanged.connect(lambda s=name, v=view: self.selectionChanged.emit(s, v.selected_refs()))
            self.views[name] = view
            self.addTab(view, name)
        self.tabBar().setVisible(len(names) > 1)

    def current_sheet(self) -> str | None:
        w = self.currentWidget()
        return w.sheet if isinstance(w, SheetView) else None

    def selected_refs(self) -> tuple[str | None, list[str]]:
        w = self.currentWidget()
        return (w.sheet, w.selected_refs()) if isinstance(w, SheetView) else (None, [])

    def focus(self, sheet: str, ref: str):
        view = self.views.get(sheet)
        if view:
            self.setCurrentWidget(view)
            view.focus_ref(ref)

    def apply(self, fmt: dict | None, values: dict[str, dict[str, str]] | None = None,
              review_ids: set[str] | None = None):
        """ラベル・記入欄を色分けする。values を渡すと記入予定の値を重ねて表示する（報告書作成画面）。"""
        for v in self.views.values():
            v.reset()
        if not fmt:
            return
        review_ids = review_ids or set()
        for f in fmt["fields"]:
            view = self.views.get(f["sheet"])
            if view is None:
                continue
            review = bool(f.get("review") or not f.get("cells")) and f.get("enabled")
            review = review or f["id"] in review_ids
            tip = f"{f['name']}（{analyzer.KIND_LABELS.get(f['kind'], f['kind'])}・" \
                  f"{analyzer.LANG_LABELS.get(f.get('lang', 'ja'))}・確信度 {f.get('confidence', 1):.2f}）"
            if f.get("reasons"):
                tip += "\n" + "\n".join(f["reasons"])
            if f.get("label_cell") and f["direction"] != "inline":
                view.mark(f["label_cell"], fill=PREVIEW_LABEL, review=review and not f.get("cells"), tooltip=tip)
            english = f.get("lang") == "en" or bool(f.get("partner_of"))
            fill = PREVIEW_OFF if not f.get("enabled") else (PREVIEW_EN if english else PREVIEW_FIELD)
            for i, ref in enumerate(f.get("cells", [])):
                if values is not None:
                    text = values.get(f["sheet"], {}).get(ref)
                    view.mark(ref, text=text if text is not None else "", fill=fill, review=review, tooltip=tip)
                else:
                    label = f"［{f['name']}］" if i == 0 or f["direction"] == "table" else ""
                    if f["direction"] == "inline":
                        label = (f.get("prefix") or "") + f"［{f['name']}］"
                    view.mark(ref, text=label, fill=fill, review=review, tooltip=tip, color="#7a5c00")
