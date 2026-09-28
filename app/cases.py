"""過去事例：報告書 Excel から欄ごとの値を取り出し、類似検索する。

- 値はラベル基準で探す。フォーマット定義での「ラベル → 記入セル」の相対位置を覚えておき、
  過去の報告書で同じラベル文字列を見つけたら、その位置から読む（旧版で行がずれていても読める）。
  ラベルが見つからないときだけセル番地で読み、警告を出す。
- 英語欄がある報告書は、日本語欄との対訳ペアを翻訳メモリとして保存する。
- 取り込み時に個人情報を検出し、伏せ字にするか取り込みを中止するかを呼び出し元に選ばせる。
"""
from __future__ import annotations

import math
import os
import uuid
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from . import analyzer, privacy
from . import textutil as tu
from . import xlsx_writer as xw

EXCEL_EXT = (".xlsx", ".xlsm")
PHASE_BONUS = 0.1


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
    label = field.get("label_text") or ""
    if not label or not field.get("label_cell"):
        return None
    target = _norm(label)
    prefix = _norm(field.get("prefix") or "")
    oc, orow = xw.split_ref(field["label_cell"])
    best, best_d = None, None
    for a in grid.anchors():
        if not isinstance(a.value, str):
            continue
        t = _norm(a.value)
        ok = (bool(prefix) and t.startswith(prefix)) if field["direction"] == "inline" else t == target
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
        if not f.get("cells"):
            continue
        sheet = f["sheet"] if f["sheet"] in wb.sheetnames else wb.sheetnames[0]
        if sheet != f["sheet"]:
            warns.append(f"シート「{f['sheet']}」がないため「{sheet}」から読みました")
        grid = grids.get(sheet) or grids.setdefault(sheet, analyzer.Grid(wb[sheet]))

        found = _find_label(grid, f)
        if found is not None:
            lc, lr = xw.split_ref(f["label_cell"])
            refs = [_shift(r, found.r1 - lr, found.c1 - lc) for r in f["cells"]]
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
        elif f["direction"] == "table" or len(refs) > 1:
            values[f["id"]] = [cell_text(grid, r) for r in refs]
        else:
            values[f["id"]] = cell_text(grid, refs[0])
    return values, warns


def field_value_text(value) -> str:
    if isinstance(value, list):
        return "\n".join(v for v in value if v)
    return value or ""


def _value_of_kind(fmt: dict, case: dict, kind: str) -> str:
    for f in fmt["fields"]:
        if f["kind"] == kind and not f.get("partner_of"):
            s = field_value_text(case["values"].get(f["id"]))
            if s:
                return s
    return ""


def subject_of(fmt: dict, case: dict) -> str:
    return _value_of_kind(fmt, case, "subject") or Path(case.get("source", "")).stem


def date_of(fmt: dict, case: dict) -> str:
    return _value_of_kind(fmt, case, "date")


def phase_of(fmt: dict, case: dict) -> str:
    return _value_of_kind(fmt, case, "phase")


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
        lines += ["", "（警告）"] + case["warnings"]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# 対訳（翻訳メモリ）
# ---------------------------------------------------------------------------

def bilingual_pairs(fmt: dict, values: dict) -> list[list[str]]:
    """日本語欄と対になる英語欄から、対訳ペア [日本語, 英語] を作る。
    行数が同じなら行ごと、違えば欄全体を 1 組にする。"""
    by_id = {f["id"]: f for f in fmt["fields"]}
    pairs = []
    for f in fmt["fields"]:
        en_id = f.get("pair")
        if not en_id or en_id not in by_id:
            continue
        ja = field_value_text(values.get(f["id"]))
        en = field_value_text(values.get(en_id))
        if not ja or not en:
            continue
        ja_lines = [l.strip() for l in ja.splitlines() if l.strip()]
        en_lines = [l.strip() for l in en.splitlines() if l.strip()]
        if len(ja_lines) == len(en_lines):
            pairs.extend([j, e] for j, e in zip(ja_lines, en_lines))
        else:
            pairs.append([ja, en])
    return pairs


def translation_memory(cases: list[dict]) -> list[tuple[str, str]]:
    seen, out = set(), []
    for c in cases:
        for ja, en in c.get("tm", []):
            if ja not in seen:
                seen.add(ja)
                out.append((ja, en))
    return out


# ---------------------------------------------------------------------------
# 取り込み
# ---------------------------------------------------------------------------

# 個人情報が見つかったときの判断：(ファイル, 見つかった箇所) -> "mask"（伏せ字にして取り込む） / "skip"（取り込まない）
PiiHandler = Callable[[Path, list], str]


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
                log.append(f"スキップ: {f.name}（.xls は Excel で .xlsx として保存し直してください）")
            elif ext in EXCEL_EXT:
                files.append(f)
    return files, log


def _same_file(a: str, b: str) -> bool:
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except (OSError, ValueError):
        return a == b


def make_case(fmt: dict, path, origin: str = "import", on_pii: PiiHandler | None = None) -> tuple[dict | None, list[str]]:
    """1 件を取り込む。個人情報があり、中止が選ばれたら (None, ログ) を返す。"""
    values, warns = extract(fmt, path)
    log = []
    findings = privacy.scan_values(fmt, values)
    masked = False
    if findings:
        choice = on_pii(Path(path), findings) if on_pii else "mask"
        if choice == "skip":
            return None, [f"中止: {Path(path).name}（個人情報らしき記載 {len(findings)}件）"]
        values = privacy.mask_values(values)
        masked = True
        log.append(f"  伏せ字: {Path(path).name}（" + "、".join(f.describe() for f in findings[:3])
                   + ("…" if len(findings) > 3 else "") + "）")
    p = Path(path)
    case = {
        "id": uuid.uuid4().hex,
        "source": str(p.resolve()),
        "mtime": p.stat().st_mtime,
        "imported_at": datetime.now().isoformat(timespec="seconds"),
        "origin": origin,
        "values": values,
        "tm": bilingual_pairs(fmt, values),
        "warnings": warns,
        "masked": masked,
    }
    return case, log


def import_paths(fmt: dict, cases: list[dict], paths, origin: str = "import",
                 on_pii: PiiHandler | None = None) -> list[str]:
    """ファイル・フォルダを取り込み、cases を更新する。ログ行を返す。"""
    files, log = collect_files(paths)
    if not files and not log:
        log.append("取り込める Excel ファイル（.xlsx / .xlsm）が見つかりませんでした")
    for f in files:
        try:
            case, extra = make_case(fmt, f, origin, on_pii)
        except Exception as e:  # 壊れたファイルなどは記録して続行
            log.append(f"失敗: {f.name}（{e}）")
            continue
        if case is None:
            log.extend(extra)
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
        log.extend(extra)
        log.extend(f"  警告: {w}" for w in case["warnings"])
    return log


def reload_all(fmt: dict, cases: list[dict], on_pii: PiiHandler | None = None) -> list[str]:
    """定義を変更したときなどに、取込元ファイルから抽出し直す。"""
    log = []
    for i, c in enumerate(list(cases)):
        src = Path(c["source"])
        if not src.exists():
            log.append(f"見つかりません（以前の内容を残します）: {src}")
            continue
        try:
            new, extra = make_case(fmt, src, c.get("origin", "import"), on_pii)
        except Exception as e:
            log.append(f"失敗: {src.name}（{e}）")
            continue
        if new is None:
            log.extend(extra)
            continue
        new["id"] = c["id"]
        new["imported_at"] = c.get("imported_at", new["imported_at"])
        cases[i] = new
        log.append(f"再読込: {src.name}")
        log.extend(extra)
        log.extend(f"  警告: {w}" for w in new["warnings"])
    return log


# ---------------------------------------------------------------------------
# 類似検索（文字 2/3-gram の TF-IDF + コサイン類似度。同じ工程に加点）
# ---------------------------------------------------------------------------

def _case_grams(fmt: dict, case: dict) -> Counter:
    grams: Counter = Counter()
    for f in fmt["fields"]:
        if f["kind"] in ("date", "reporter", "dept", "skip", "status") or f.get("lang") == "en":
            continue
        text = field_value_text(case["values"].get(f["id"]))
        if not text:
            continue
        weight = 2 if f["kind"] == "subject" else 1
        for k, v in tu.ngrams(text).items():
            grams[k] += v * weight
    return grams


class Index:
    def __init__(self, fmt: dict, cases: list[dict]):
        self.fmt = fmt
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

    def search(self, query: str, top: int = 10, phase: str | None = None) -> list[tuple[float, dict]]:
        q = self._weigh(tu.ngrams(query))
        scored = []
        for v, c in zip(self.vectors, self.cases):
            s = tu.cosine(q, v)
            if phase and s > 0 and phase_of(self.fmt, c) == phase:
                s += PHASE_BONUS
            scored.append((round(s, 4), c))
        scored.sort(key=lambda x: -x[0])
        return scored[:top]
