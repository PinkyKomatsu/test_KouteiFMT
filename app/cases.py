"""過去事例：報告書 Excel から欄ごとの値を取り出し、類似検索する。

値はラベル基準で探す。フォーマット定義での「ラベル → 記入セル」の相対位置を覚えておき、
過去の報告書で同じラベル文字列を見つけたら、その位置からの相対位置で読む。
旧版で行がずれていても読める。ラベルが見つからないときだけセル番地で読み、警告を出す。
"""
from __future__ import annotations

import math
import os
import uuid
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from . import analyzer
from . import textutil as tu
from . import xlsx_writer as xw

EXCEL_EXT = (".xlsx", ".xlsm")


# ---------------------------------------------------------------------------
# 値の取り出し
# ---------------------------------------------------------------------------

def _date_text(d, number_format: str) -> str:
    nf = (number_format or "").lower()
    if "年" in nf:
        return f"{d.year}年{d.month}月{d.day}日"
    if "mm" in nf and "dd" in nf:
        sep = "-" if "-" in nf else "/"
        return f"{d.year}{sep}{d.month:02d}{sep}{d.day:02d}"
    return f"{d.year}/{d.month}/{d.day}"


def cell_text(grid: analyzer.Grid, ref: str) -> str:
    a = grid.anchor_ref(ref)
    v = a.value
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        cell = grid.ws._cells.get((a.r1, a.c1))
        return _date_text(v, getattr(cell, "number_format", ""))
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    text = str(v).replace("\r\n", "\n").strip()
    return "" if analyzer.is_placeholder(text) else text


def _norm(s: str) -> str:
    return "".join(tu.nfkc(s or "").split())


def _find_label(grid: analyzer.Grid, field: dict):
    """定義と同じラベル文字列のアンカーを探す。複数あれば元の位置に最も近いもの。"""
    label = field.get("label_text") or ""
    if not label:
        return None
    target = _norm(label)
    prefix = _norm(field.get("prefix") or "")
    oc, orow = xw.split_ref(field["label_cell"]) if field.get("label_cell") else (1, 1)
    best, best_d = None, None
    for a in grid.anchors():
        if not isinstance(a.value, str):
            continue
        t = _norm(a.value)
        if field["direction"] == "inline":
            ok = bool(prefix) and t.startswith(prefix)
        else:
            ok = t == target
        if ok:
            d = abs(a.r1 - orow) + abs(a.c1 - oc)
            if best is None or d < best_d:
                best, best_d = a, d
    return best


def _shift(ref: str, dr: int, dc: int) -> str:
    c, r = xw.split_ref(ref)
    return xw.make_ref(max(1, c + dc), max(1, r + dr))


def extract(fmt: dict, path) -> tuple[dict, list[str]]:
    """報告書 1 件から {field_id: 値} を取り出す。表の欄は行ごとのリスト。"""
    wb = analyzer.load_book(path)
    grids: dict[str, analyzer.Grid] = {}
    values: dict[str, object] = {}
    warns: list[str] = []
    for f in fmt["fields"]:
        sheet = f["sheet"] if f["sheet"] in wb.sheetnames else wb.sheetnames[0]
        if sheet != f["sheet"]:
            warns.append(f"シート「{f['sheet']}」がないため「{sheet}」から読みました")
        grid = grids.get(sheet) or grids.setdefault(sheet, analyzer.Grid(wb[sheet]))

        found = _find_label(grid, f)
        if found is not None:
            lc, lr = xw.split_ref(f["label_cell"])
            dr, dc = found.r1 - lr, found.c1 - lc
            refs = [_shift(r, dr, dc) for r in f["cells"]]
        else:
            refs = list(f["cells"])
            if f.get("label_text"):
                warns.append(f"「{f['name']}」のラベルが見つからないため、セル番地 {refs[0]} から読みました")

        if f["direction"] == "inline":
            raw = cell_text(grid, refs[0])
            prefix = f.get("prefix") or ""
            if raw.startswith(prefix):
                raw = raw[len(prefix):]
            elif "：" in raw or ":" in raw:
                raw = raw.split("：", 1)[-1] if "：" in raw else raw.split(":", 1)[-1]
            raw = raw.strip()
            values[f["id"]] = "" if analyzer.is_placeholder(raw) else raw
        elif f["direction"] == "table":
            values[f["id"]] = [cell_text(grid, r) for r in refs]
        else:
            values[f["id"]] = cell_text(grid, refs[0])
    return values, warns


def field_value_text(value) -> str:
    if isinstance(value, list):
        return "\n".join(v for v in value if v)
    return value or ""


def subject_of(fmt: dict, case: dict) -> str:
    for f in fmt["fields"]:
        if f["kind"] == "subject":
            s = field_value_text(case["values"].get(f["id"]))
            if s:
                return s
    return Path(case.get("source", "")).stem


def date_of(fmt: dict, case: dict) -> str:
    for f in fmt["fields"]:
        if f["kind"] == "date":
            return field_value_text(case["values"].get(f["id"]))
    return ""


def case_preview(fmt: dict, case: dict) -> str:
    lines = []
    for f in fmt["fields"]:
        v = case["values"].get(f["id"])
        if isinstance(v, list):
            if not any(v):
                continue
            lines.append(f"■ {f['name']}")
            lines.extend(f"  {i + 1}. {x}" for i, x in enumerate(v) if x)
        elif v:
            lines.append(f"■ {f['name']}")
            lines.append(v)
    if case.get("warnings"):
        lines.append("")
        lines.append("（警告）")
        lines.extend(case["warnings"])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 取り込み
# ---------------------------------------------------------------------------

def collect_files(paths) -> tuple[list[Path], list[str]]:
    """ファイル・フォルダ（サブフォルダを含む）から対象の Excel を集める。"""
    files, log = [], []
    for p in map(Path, paths):
        items = sorted(p.rglob("*")) if p.is_dir() else [p]
        for f in items:
            if not f.is_file() or f.name.startswith("~$"):
                continue
            ext = f.suffix.lower()
            if ext == ".xls":
                log.append(f"スキップ: {f.name}（.xls は xlsx で保存し直してください）")
            elif ext in EXCEL_EXT:
                files.append(f)
    return files, log


def _same_file(a: str, b: str) -> bool:
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except (OSError, ValueError):
        return a == b


def make_case(fmt: dict, path, origin: str = "import") -> dict:
    values, warns = extract(fmt, path)
    p = Path(path)
    return {
        "id": uuid.uuid4().hex,
        "source": str(p.resolve()),
        "mtime": p.stat().st_mtime,
        "imported_at": datetime.now().isoformat(timespec="seconds"),
        "origin": origin,
        "values": values,
        "warnings": warns,
    }


def import_paths(fmt: dict, cases: list[dict], paths, origin: str = "import") -> list[str]:
    """ファイル・フォルダを取り込み、cases を更新する。ログ行を返す。"""
    files, log = collect_files(paths)
    if not files and not log:
        log.append("取り込める Excel ファイル（.xlsx / .xlsm）が見つかりませんでした")
    for f in files:
        try:
            case = make_case(fmt, f, origin)
        except Exception as e:  # 壊れたファイルなどは記録して続行
            log.append(f"失敗: {f.name}（{e}）")
            continue
        for i, old in enumerate(cases):
            if _same_file(old["source"], case["source"]):
                case["id"] = old["id"]
                case["origin"] = old.get("origin", origin)
                cases[i] = case
                log.append(f"更新: {f.name}")
                break
        else:
            cases.append(case)
            log.append(f"追加: {f.name}")
        for w in case["warnings"]:
            log.append(f"  警告: {w}")
    return log


def reload_all(fmt: dict, cases: list[dict]) -> list[str]:
    """定義を変更したときなどに、取込元ファイルから抽出し直す。"""
    log = []
    for i, c in enumerate(cases):
        src = Path(c["source"])
        if not src.exists():
            log.append(f"見つかりません（以前の内容を残します）: {src}")
            continue
        try:
            new = make_case(fmt, src, c.get("origin", "import"))
        except Exception as e:
            log.append(f"失敗: {src.name}（{e}）")
            continue
        new["id"] = c["id"]
        new["imported_at"] = c.get("imported_at", new["imported_at"])
        cases[i] = new
        log.append(f"再読込: {src.name}")
        for w in new["warnings"]:
            log.append(f"  警告: {w}")
    return log


# ---------------------------------------------------------------------------
# 類似検索（文字 2/3-gram の TF-IDF + コサイン類似度）
# ---------------------------------------------------------------------------

def _case_grams(fmt: dict, case: dict) -> Counter:
    grams: Counter = Counter()
    for f in fmt["fields"]:
        if f["kind"] in ("date", "reporter", "dept", "skip"):
            continue
        text = field_value_text(case["values"].get(f["id"]))
        if not text:
            continue
        g = tu.ngrams(text)
        weight = 2 if f["kind"] == "subject" else 1
        for k, v in g.items():
            grams[k] += v * weight
    return grams


class Index:
    def __init__(self, fmt: dict, cases: list[dict]):
        self.cases = cases
        docs = [_case_grams(fmt, c) for c in cases]
        n = len(docs)
        df: Counter = Counter()
        for d in docs:
            df.update(d.keys())
        self.idf = {k: math.log((n + 1) / (v + 1)) + 1 for k, v in df.items()}
        self.vectors = [self._weigh(d) for d in docs]

    def _weigh(self, grams: Counter) -> dict:
        return {k: v * self.idf.get(k, 1.0) for k, v in grams.items()}

    def search(self, query: str, top: int = 10) -> list[tuple[float, dict]]:
        q = self._weigh(tu.ngrams(query))
        scored = [(tu.cosine(q, v), c) for v, c in zip(self.vectors, self.cases)]
        scored.sort(key=lambda x: -x[0])
        return scored[:top]
