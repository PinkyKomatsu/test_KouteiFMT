"""フォーマット解析：Excel を読み、ラベルと記入欄を推定してフォーマット定義を作る。

openpyxl は読み取りだけに使う（保存はしない）。結合セルは左上のアンカーセルに正規化する。
"""
from __future__ import annotations

import io
import re
import warnings
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl
from lxml import etree

from . import xlsx_writer as xw

KIND_LABELS = {
    "date": "日付",
    "subject": "件名",
    "reporter": "報告者",
    "dept": "部署",
    "body": "本文",
    "list": "表",
    "skip": "記入しない",
}
DIRECTION_LABELS = {"right": "右", "down": "下", "inline": "同じセル", "table": "表", "manual": "手動"}

# 上から順に判定（「承認者氏名」は記入しないにする）
KIND_RULES = [
    ("skip", ("承認", "検印", "上長", "宛先")),
    ("date", ("日付", "報告日", "作成日")),
    ("subject", ("件名", "表題", "テーマ")),
    ("reporter", ("報告者", "担当者", "氏名")),
    ("dept", ("部署", "所属")),
]

PLACEHOLDER_RE = re.compile(
    r"^\s*(|〇+|○+|◯+|[xXｘＸ×]{2,}|[＿_]+|[-－]+|（ここに記入）|\(ここに記入\)|ここに記入|記入欄|未記入)\s*$"
)
INLINE_RE = re.compile(
    r"^(?P<label>.{1,20}?)\s*(?P<colon>[：:])\s*(?:〇+|○+|◯+|[xXｘＸ×]{2,}|[＿_]+|（ここに記入）|\(ここに記入\))\s*$"
)
_NAME_STRIP_RE = re.compile(r"^[\s【\[［■●◆・]+|[\s】\]］：:]+$")

MAX_ROWS = 400
MAX_COLS = 60


def classify(name: str) -> str:
    for kind, words in KIND_RULES:
        if any(w in name for w in words):
            return kind
    return "body"


def clean_name(text: str) -> str:
    s = _NAME_STRIP_RE.sub("", (text or "").strip())
    return s or (text or "").strip()


def is_placeholder(value) -> bool:
    if value is None:
        return True
    return isinstance(value, str) and bool(PLACEHOLDER_RE.match(value))


# ---------------------------------------------------------------------------
# 読み込み（図形・画像は解析に不要なので外してから openpyxl に渡す）
# ---------------------------------------------------------------------------

def _without_drawings(path: Path) -> io.BytesIO:
    """シートの rels から drawing の参照を外したコピーをメモリ上に作る。

    openpyxl は画像を読むのに Pillow を必要とするため、画像を含む原本でも
    確実に読めるようにする（原本ファイル自体は変更しない）。
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if re.match(r"^xl/worksheets/_rels/[^/]+\.rels$", info.filename):
                root = xw.parse_xml(data)
                for r in list(root):
                    if (r.get("Type") or "").endswith("/drawing"):
                        root.remove(r)
                data = xw.serialize_xml(root)
            zout.writestr(info.filename, data)
    buf.seek(0)
    return buf


def load_book(path):
    """解析用に Excel を開く（値は数式の計算結果を使う）。"""
    path = Path(path)
    xw.check_supported(path)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(_without_drawings(path), data_only=True)


# ---------------------------------------------------------------------------
# グリッド（結合セルをアンカーに正規化）
# ---------------------------------------------------------------------------

@dataclass
class Anchor:
    r1: int
    c1: int
    r2: int
    c2: int
    value: object
    border: bool
    merged: bool

    @property
    def ref(self) -> str:
        return xw.make_ref(self.c1, self.r1)

    @property
    def text(self) -> str:
        return self.value.strip() if isinstance(self.value, str) else ""

    @property
    def empty(self) -> bool:
        return is_placeholder(self.value)

    @property
    def fillable(self) -> bool:
        """記入欄になり得る空欄（罫線も結合もない空白は余白とみなす）。"""
        return self.empty and (self.border or self.merged)


def _has_border(cell) -> bool:
    try:
        b = cell.border
    except AttributeError:
        return False
    if b is None:
        return False
    return any(getattr(getattr(b, side, None), "style", None) for side in ("left", "right", "top", "bottom"))


def col_px(ws, c: int) -> int:
    dim = ws.column_dimensions.get(xw.index_to_col(c))
    width = dim.width if dim is not None and dim.width else None
    if dim is not None and dim.hidden:
        return 0
    return int(round(width * 7 + 5)) if width else 64


def row_px(ws, r: int) -> int:
    dim = ws.row_dimensions.get(r)
    if dim is not None and dim.hidden:
        return 0
    height = dim.height if dim is not None and dim.height else None
    return int(round(height * 96 / 72)) if height else 20


class Grid:
    def __init__(self, ws):
        self.ws = ws
        self.owner: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for mr in ws.merged_cells.ranges:
            rng = (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    self.owner[(r, c)] = rng
        keys = list(ws._cells.keys()) + list(self.owner.keys())
        self.max_row = min(max((k[0] for k in keys), default=1), MAX_ROWS)
        self.max_col = min(max((k[1] for k in keys), default=1), MAX_COLS)
        self._cache: dict[tuple[int, int], Anchor] = {}
        self.col_px = [0] + [col_px(ws, c) for c in range(1, self.max_col + 2)]
        self.row_px = [0] + [row_px(ws, r) for r in range(1, self.max_row + 2)]

    def anchor_at(self, r: int, c: int) -> Anchor:
        rng = self.owner.get((r, c))
        r1, c1, r2, c2 = rng if rng else (r, c, r, c)
        key = (r1, c1)
        a = self._cache.get(key)
        if a is None:
            cell = self.ws._cells.get((r1, c1))
            value = cell.value if cell is not None else None
            border = False
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    cl = self.ws._cells.get((rr, cc))
                    if cl is not None and _has_border(cl):
                        border = True
                        break
                if border:
                    break
            a = Anchor(r1, c1, r2, c2, value, border, rng is not None)
            self._cache[key] = a
        return a

    def anchor_ref(self, ref: str) -> Anchor:
        c, r = xw.split_ref(ref)
        return self.anchor_at(r, c)

    def anchors(self):
        """範囲内のアンカーを行・列の順に返す。"""
        for r in range(1, self.max_row + 1):
            for c in range(1, self.max_col + 1):
                rng = self.owner.get((r, c))
                if rng and (rng[0], rng[1]) != (r, c):
                    continue
                yield self.anchor_at(r, c)

    def width(self, a: Anchor) -> int:
        return sum(self.col_px[c] for c in range(a.c1, min(a.c2, self.max_col + 1) + 1))

    def height(self, a: Anchor) -> int:
        return sum(self.row_px[r] for r in range(a.r1, min(a.r2, self.max_row + 1) + 1))

    def area(self, a: Anchor) -> int:
        return self.width(a) * self.height(a)

    def sheet_width(self) -> int:
        return sum(self.col_px[1:self.max_col + 1])


def is_label(a: Anchor) -> bool:
    if not isinstance(a.value, str):
        return False
    s = a.value.strip()
    return 0 < len(s) <= 30 and not s.endswith("。") and not is_placeholder(s)


# ---------------------------------------------------------------------------
# 記入欄の推定
# ---------------------------------------------------------------------------

def _field(fid, name, label: Anchor | None, sheet, cells, kind, direction, grid: Grid, prefix="", table_id=None):
    first = grid.anchor_ref(cells[0])
    return {
        "id": fid,
        "name": name,
        "label_text": label.text if label else "",
        "label_cell": label.ref if label else "",
        "sheet": sheet,
        "cells": cells,
        "kind": kind,
        "direction": direction,
        "prefix": prefix,
        "height_px": grid.height(first),
        "width_px": grid.width(first),
        "table_id": table_id,
        "enabled": kind != "skip",
    }


def _detect_sheet(ws, grid: Grid, next_id, table_counter) -> list[dict]:
    fields: list[dict] = []
    used: set[tuple[int, int]] = set()
    labels = [a for a in grid.anchors() if is_label(a)]

    def key(a: Anchor):
        return (a.r1, a.c1)

    # 1. 表：同じ行に隣接するラベルが 3 つ以上、それぞれの直下が空欄
    by_row: dict[int, list[Anchor]] = {}
    for a in labels:
        by_row.setdefault(a.r1, []).append(a)
    for r, row_labels in sorted(by_row.items()):
        row_labels.sort(key=lambda a: a.c1)
        runs, run = [], [row_labels[0]]
        for a in row_labels[1:]:
            if a.c1 == run[-1].c2 + 1 and a.r2 == run[-1].r2:
                run.append(a)
            else:
                runs.append(run)
                run = [a]
        runs.append(run)
        for run in runs:
            if len(run) < 3:
                continue
            columns = []
            for h in run:
                cells, rr = [], h.r2 + 1
                while rr <= grid.max_row and len(cells) < 200:
                    b = grid.anchor_at(rr, h.c1)
                    if b.c1 != h.c1 or b.r1 != rr or not b.empty or not b.border:
                        break
                    cells.append(b)
                    rr = b.r2 + 1
                columns.append(cells)
            n = min(len(c) for c in columns)
            if n == 0:
                continue
            table_counter[0] += 1
            tid = f"t{table_counter[0]}"
            for h, cells in zip(run, columns):
                fields.append(_field(next_id(), clean_name(h.text), h, ws.title,
                                     [b.ref for b in cells[:n]], "list", "table", grid, table_id=tid))
                used.add(key(h))
                used.update(key(b) for b in cells[:n])

    # 2. 同じセル内に記入：「件名：〇〇」
    for a in labels:
        if key(a) in used:
            continue
        m = INLINE_RE.match(a.text)
        if m:
            prefix = a.text[: m.end("colon")]
            name = clean_name(m.group("label"))
            fields.append(_field(next_id(), name, a, ws.title, [a.ref], classify(name), "inline", grid, prefix=prefix))
            used.add(key(a))

    # 3. 右か下の空欄：先に全ラベルの「右」を確定させ、その後で「下」を割り当てる
    pending = [a for a in labels if key(a) not in used]
    sheet_w = grid.sheet_width() or 1

    def down_of(a: Anchor):
        b = grid.anchor_at(a.r2 + 1, a.c1) if a.r2 + 1 <= grid.max_row else None
        if b and b.c1 == a.c1 and b.r1 == a.r2 + 1 and b.empty and b.border and key(b) not in used:
            return b
        return None

    chosen: dict[tuple[int, int], tuple[Anchor, str, Anchor]] = {}
    for a in pending:
        if a.c2 + 1 > grid.max_col:
            continue
        right = grid.anchor_at(a.r1, a.c2 + 1)
        if right.c1 != a.c2 + 1 or right.r1 != a.r1 or not right.fillable or key(right) in used:
            continue
        if grid.width(a) >= sheet_w * 0.6:
            continue  # 幅の広い見出しは下の欄を採る
        down = down_of(a)
        if down and not right.merged and down.merged and grid.area(down) >= grid.area(right) * 2:
            continue  # 右が単独セルで、下が 2 倍以上大きな結合欄なら下を採る
        chosen[key(a)] = (a, "right", right)
        used.add(key(right))
    for a in pending:
        if key(a) in chosen:
            continue
        down = down_of(a)
        if down:
            chosen[key(a)] = (a, "down", down)
            used.add(key(down))

    for a, direction, target in chosen.values():
        name = clean_name(a.text)
        fields.append(_field(next_id(), name, a, ws.title, [target.ref], classify(name), direction, grid))
        used.add(key(a))
    return fields


def _sort_key(sheet_order, field):
    c, r = xw.split_ref(field["cells"][0])
    return (sheet_order.get(field["sheet"], 0), r, c)


def _dedupe_names(fields: list[dict]) -> None:
    seen: dict[str, int] = {}
    for f in fields:
        base = f["name"]
        if base in seen:
            seen[base] += 1
            f["name"] = f"{base}({seen[base]})"
        else:
            seen[base] = 1


def analyze(path, name: str | None = None) -> dict:
    """Excel を解析してフォーマット定義（dict）を返す。"""
    path = Path(path)
    wb = load_book(path)
    counter = [0]
    table_counter = [0]

    def next_id():
        counter[0] += 1
        return f"f{counter[0]}"

    fields: list[dict] = []
    sheets = []
    for ws in wb.worksheets:
        sheets.append(ws.title)
        fields.extend(_detect_sheet(ws, Grid(ws), next_id, table_counter))
    order = {s: i for i, s in enumerate(sheets)}
    fields.sort(key=lambda f: _sort_key(order, f))
    _dedupe_names(fields)
    return {
        "version": 1,
        "name": name or path.stem,
        "template": "template" + path.suffix.lower(),
        "source": str(path),
        "created": datetime.now().isoformat(timespec="seconds"),
        "sheets": sheets,
        "fields": fields,
    }


def new_field_id(fields: list[dict]) -> str:
    nums = [int(f["id"][1:]) for f in fields if re.match(r"^f\d+$", f.get("id", ""))]
    return f"f{max(nums, default=0) + 1}"


def field_from_cell(template, sheet: str, ref: str, fields: list[dict]) -> dict:
    """選択セルを記入欄にする。ラベルは左 → 上の順に探す。"""
    wb = load_book(template)
    ws = wb[sheet]
    grid = Grid(ws)
    target = grid.anchor_ref(ref)
    label = None
    for c in range(target.c1 - 1, max(0, target.c1 - 4), -1):
        a = grid.anchor_at(target.r1, c)
        if is_label(a):
            label = a
            break
        if not a.empty:
            break
    if label is None:
        for r in range(target.r1 - 1, max(0, target.r1 - 4), -1):
            a = grid.anchor_at(r, target.c1)
            if is_label(a):
                label = a
                break
            if not a.empty:
                break
    name = clean_name(label.text) if label else target.ref
    field = _field(new_field_id(fields), name, label, sheet, [target.ref], classify(name), "manual", grid)
    names = {f["name"] for f in fields}
    base, n = field["name"], 1
    while field["name"] in names:
        n += 1
        field["name"] = f"{base}({n})"
    return field


# ---------------------------------------------------------------------------
# GUI プレビュー用
# ---------------------------------------------------------------------------

def display_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{value.year}/{value.month}/{value.day}"
    if isinstance(value, date):
        return f"{value.year}/{value.month}/{value.day}"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def sheet_layout(template, sheet: str | None = None, max_rows: int = 80, max_cols: int = 40) -> dict:
    """Canvas でシートを再現するための情報（列幅・行高・結合・文字列）。"""
    wb = load_book(template)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    grid = Grid(ws)
    n_rows = min(grid.max_row, max_rows)
    n_cols = min(grid.max_col, max_cols)
    cells = []
    for a in grid.anchors():
        if a.r1 > n_rows or a.c1 > n_cols:
            continue
        cells.append({
            "ref": a.ref, "r1": a.r1, "c1": a.c1,
            "r2": min(a.r2, n_rows), "c2": min(a.c2, n_cols),
            "text": display_text(a.value), "border": a.border, "label": is_label(a),
        })
    return {
        "sheet": ws.title,
        "sheets": wb.sheetnames,
        "n_rows": n_rows,
        "n_cols": n_cols,
        "col_px": grid.col_px[1:n_cols + 1],
        "row_px": grid.row_px[1:n_rows + 1],
        "cells": cells,
        "owner": {f"{r},{c}": xw.make_ref(rng[1], rng[0]) for (r, c), rng in grid.owner.items()
                  if r <= n_rows and c <= n_cols},
    }
