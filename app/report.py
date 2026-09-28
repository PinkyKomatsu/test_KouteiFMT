"""報告書の下書き（日本語と英語）を組み立て、Excel のセル・コピー用の文字列に変換する。

UI から切り離した関数だけを置く（テストしやすくするため）。

    draft = build_draft(fmt, topic, phase, status, memo, refs, all_cases, settings, translator)
    values = cell_values(fmt, draft)                 # write_cells に渡す {(シート, セル): 値}
    text_ja = clipboard_text(fmt, draft, "ja")       # 【欄名】本文 の形式
"""
from __future__ import annotations

import re
from datetime import date

from . import generator_ja, translator_en
from . import textutil as tu

# 背景色の区分（UI で使う）
SOURCE_LABELS = {
    "memo": "メモを整形", "reuse": "事例を流用", "auto": "自動入力", "empty": "空欄",
    "template": "定型文", "tm": "翻訳メモリ", "mt": "機械翻訳", "simple": "簡易訳", "copy": "そのまま",
    "edited": "手で編集",
}


def new_entry(ja="", ja_status="empty", note_ja="", review_ja=False) -> dict:
    return {"ja": ja, "en": "", "ja_status": ja_status, "en_status": "empty", "note_ja": note_ja,
            "note_en": "", "review_ja": review_ja, "review_en": False}


# ---------------------------------------------------------------------------
# 英訳
# ---------------------------------------------------------------------------

def _table_columns(fmt: dict) -> dict[str, list[dict]]:
    tables: dict[str, list[dict]] = {}
    for f in generator_ja.logical_fields(fmt):
        if f.get("table_id"):
            tables.setdefault(f["table_id"], []).append(f)
    return tables


def translate_entry(fmt: dict, field: dict, entry: dict, translator: translator_en.Translator,
                    settings: dict, today: date | None = None) -> dict:
    """1 欄の日本語から英語を作り、entry を更新して返す。"""
    today = today or date.today()
    ja = entry.get("ja") or ""
    kind = field["kind"]
    en, method, review, notes = "", "empty", False, []
    if not ja.strip():
        pass
    elif kind == "date":
        en, method = tu.format_pattern(today, settings.get("date_format_en") or "MMM d, yyyy"), "template"
        if tu.find_dates(ja) and not ja.startswith(tu.format_date(today, "{Y}")):
            en, method = ja, "copy"   # 本日以外の日付を手で入れた場合はそのまま
    elif kind == "phase" and ja.strip() in translator_en.PHASES:
        en, method = translator_en.PHASES[ja.strip()], "template"
    elif kind == "status" and ja.strip() in translator_en.STATUSES:
        en, method = translator_en.STATUSES[ja.strip()], "template"
    elif kind == "reporter":
        en, method = ja, "copy"       # 人名はそのまま（ローマ字表記は用語集に登録すると使われます）
        g = translator.glossary.lookup(ja.strip())
        if g:
            en, method = g, "template"
    elif kind == "list" and field.get("table_id"):
        role = generator_ja.column_role(field["name"])
        cells = ja.split("\n")
        out = []
        for c in cells:
            if not c.strip() or role in ("seq", "due", "owner"):
                out.append(c)
                continue
            r = translator.translate_text(c)
            out.append(r.text)
            review |= r.review
            notes += [n for n in r.notes if n not in notes]
            method = r.method if method in ("empty", "copy", "template") else method
        en = "\n".join(out)
        if method == "empty":
            method = "copy"
    else:
        r = translator.translate_text(ja)
        en, method, review, notes = r.text, r.method, r.review, r.notes
    if method == "simple":
        review = True
    entry.update(en=en, en_status=method, review_en=review,
                 note_en="／".join([translator_en.METHOD_LABELS.get(method, method)] + notes) if ja.strip() else "")
    return entry


def build_draft(fmt: dict, topic: str, phase: str, status: str, memo: str, ref_cases: list[dict],
                all_cases: list[dict], settings: dict, translator: translator_en.Translator,
                today: date | None = None) -> dict[str, dict]:
    """日本語を生成し、英語に訳した下書き {欄 id: entry} を返す。"""
    today = today or date.today()
    ja = generator_ja.generate(fmt, topic, phase, status, memo, ref_cases, all_cases, settings, today)
    draft: dict[str, dict] = {}
    for f in generator_ja.logical_fields(fmt):
        r = ja.get(f["id"], {"ja": "", "status": "empty", "note": ""})
        entry = new_entry(r["ja"], r["status"], r.get("note", ""), r.get("review", False))
        draft[f["id"]] = translate_entry(fmt, f, entry, translator, settings, today)
    return draft


# ---------------------------------------------------------------------------
# Excel のセル
# ---------------------------------------------------------------------------

def _text_for(field: dict, entry: dict, english_only: bool) -> str:
    if english_only:
        return entry.get("en", "")
    lang = field.get("lang", "ja")
    if lang == "en":
        return entry.get("en", "")
    if lang == "both":
        ja, en = entry.get("ja", ""), entry.get("en", "")
        return ja + ("\n" + en if en and en != ja else "")
    return entry.get("ja", "")


def _put(values: dict, field: dict, text: str) -> None:
    if not field.get("cells"):
        return
    if field["direction"] == "table" or len(field["cells"]) > 1:
        lines = (text or "").split("\n")
        for i, ref in enumerate(field["cells"]):
            values[(field["sheet"], ref)] = lines[i] if i < len(lines) else ""
    elif field["direction"] == "inline":
        values[(field["sheet"], field["cells"][0])] = (field.get("prefix") or "") + (text or "")
    else:
        values[(field["sheet"], field["cells"][0])] = text or ""


def cell_values(fmt: dict, draft: dict[str, dict], english_only: bool = False) -> dict[tuple[str, str], str]:
    """write_cells に渡す値。日本語欄には日本語、英語欄（対になる欄を含む）には英語を入れる。
    english_only=True なら全欄に英語を入れる（英語版を別ブックで出力するとき）。"""
    by_id = {f["id"]: f for f in fmt["fields"]}
    values: dict[tuple[str, str], str] = {}
    for f in generator_ja.logical_fields(fmt):
        entry = draft.get(f["id"])
        if entry is None:
            continue
        _put(values, f, _text_for(f, entry, english_only))
        partner = by_id.get(f.get("pair") or "")
        if partner and partner.get("enabled"):
            _put(values, partner, entry.get("en", ""))
    return values


def preview_values(fmt: dict, draft: dict[str, dict], english_only: bool = False) -> dict[str, dict[str, str]]:
    """プレビュー用：{シート: {セル: 値}}。"""
    out: dict[str, dict[str, str]] = {}
    for (sheet, ref), v in cell_values(fmt, draft, english_only).items():
        out.setdefault(sheet, {})[ref] = v
    return out


def field_at(fmt: dict, sheet: str, ref: str) -> dict | None:
    """セルが属する欄（英語版の欄なら対になる日本語欄）を返す。"""
    by_id = {f["id"]: f for f in fmt["fields"]}
    for f in fmt["fields"]:
        if f["sheet"] == sheet and (ref in f.get("cells", []) or ref == f.get("label_cell")):
            if f.get("partner_of") and f["partner_of"] in by_id:
                return by_id[f["partner_of"]]
            return f
    return None


# ---------------------------------------------------------------------------
# コピー（R9）
# ---------------------------------------------------------------------------

def english_name(fmt: dict, field: dict) -> str:
    by_id = {f["id"]: f for f in fmt["fields"]}
    partner = by_id.get(field.get("pair") or "")
    if partner and partner.get("label_text"):
        return partner["name"]
    if field.get("lang") == "en":
        return field["name"]
    if field.get("lang") == "both" and re.search(r"[/／]", field["name"]):
        return re.split(r"\s*[/／]\s*", field["name"])[-1]
    if field.get("table_id"):
        role = generator_ja.column_role(field["name"])
        if role:
            return translator_en.COLUMN_NAMES[role]
    if field["kind"] not in ("body",) and field["kind"] in translator_en.FIELD_NAMES:
        return translator_en.FIELD_NAMES[field["kind"]]
    return translator_en.HEADINGS.get(field["name"]) or translator_en.FIELD_NAMES.get(field["kind"], field["name"])


def _table_block(fmt: dict, columns: list[dict], draft: dict, lang: str) -> str | None:
    key = "en" if lang == "en" else "ja"
    cols = [(draft.get(c["id"], {}).get(key) or "").split("\n") for c in columns]
    n = max((len(c) for c in cols), default=0)
    rows = []
    for i in range(n):
        cells = [c[i] if i < len(c) else "" for c in cols]
        if any(x.strip() for x in cells[1:] if x) and any(cells):
            rows.append((" | " if lang == "en" else "｜").join(cells))
    if not rows:
        return None
    names = [english_name(fmt, c) if lang == "en" else c["name"] for c in columns]
    head = (" / " if lang == "en" else "／").join(names)
    title = f"[{head}]" if lang == "en" else f"【{head}】"
    return title + "\n" + "\n".join(rows)


def field_clip(fmt: dict, field: dict, entry: dict, lang: str) -> str:
    """1 欄分のコピー文字列（欄ごとの［日本語をコピー］［英語をコピー］）。本文だけを返す。"""
    return entry.get("en" if lang == "en" else "ja", "") or ""


def clipboard_text(fmt: dict, draft: dict[str, dict], lang: str) -> str:
    """全欄をまとめたコピー文字列。「【欄名】本文」（英語は「[Name] text」）の形式で、欄の間は空行。"""
    blocks = []
    done_tables = set()
    tables = _table_columns(fmt)
    for f in generator_ja.logical_fields(fmt):
        tid = f.get("table_id")
        if tid:
            if tid not in done_tables:
                done_tables.add(tid)
                block = _table_block(fmt, tables[tid], draft, lang)
                if block:
                    blocks.append(block)
            continue
        entry = draft.get(f["id"])
        text = (entry or {}).get("en" if lang == "en" else "ja", "") or ""
        if not text.strip():
            continue
        name = english_name(fmt, f) if lang == "en" else f["name"]
        head = f"[{name}]" if lang == "en" else f"【{name}】"
        blocks.append(head + ("\n" if "\n" in text else " " if lang == "en" else "") + text)
    return "\n\n".join(blocks)


# ---------------------------------------------------------------------------
# ファイル名
# ---------------------------------------------------------------------------

_FORBIDDEN_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def make_filename(pattern: str, today: date, phase: str = "", topic: str = "", reporter: str = "") -> str:
    """ファイル名（拡張子なし）。禁止文字は「_」に置き換える。"""
    name = pattern or "{date}_{phase}_{topic}"
    name = (name.replace("{date}", f"{today.year}{today.month:02d}{today.day:02d}")
            .replace("{phase}", phase or "").replace("{topic}", topic.strip() or "報告書")
            .replace("{reporter}", reporter or ""))
    name = _FORBIDDEN_RE.sub("_", name).replace("\n", " ")
    name = re.sub(r"_{2,}", "_", name).strip(" ._")
    return name[:120] or "報告書"
