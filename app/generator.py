"""ルールベースの下書き生成。

1. メモを 1 行 1 項目として各欄に振り分ける
2. 参考事例（上位 3 件）から欄ごとの書き方を学習する
3. メモ行を学習した書き方に整形する
4. 件名・日付・報告者・表・メモのない欄を埋める

各欄の結果は {"text", "status", "note"}。status は memo / reuse / auto / empty。
"""
from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass, field as dc_field
from datetime import date

from . import textutil as tu
from .cases import field_value_text

MAIN_KEYWORDS = ("内容", "概要", "報告")
LOW_SCORE = 0.15
NONE_TEXT = "特になし"

LABELED_RE = re.compile(r"^\s*(?:【\s*(?P<h>.+?)\s*】\s*(?P<hrest>.*)|(?P<n>[^：:\s]{1,15})\s*[：:]\s*(?P<nrest>.*))$")
TABLE_SPLIT_RE = re.compile(r"\s*[｜|]\s*")
FORMULAIC_RE = re.compile(r"以上|報告いたします|ご報告|お疲れ様|お世話になって|下記|以下の|について報告")

# 表の列見出しの役割
TABLE_COLUMN_ROLES = [
    ("seq", ("No", "NO", "no", "№", "番号", "項番", "#")),
    ("content", ("内容", "アクション", "項目", "タスク", "対応", "作業", "件名")),
    ("owner", ("担当", "責任者", "氏名", "実施者")),
    ("due", ("期限", "期日", "予定日", "日付", "納期", "完了日")),
    ("status", ("状態", "ステータス", "進捗", "状況")),
]
MEMO_TABLE_ORDER = ["content", "owner", "due", "status"]

# 見出し名から推定する役割
HEADING_ROLES = {
    "result": ("状況", "現状", "結果", "経過", "実績", "概要", "内容"),
    "plan": ("対応", "今後", "予定", "対策", "計画"),
    "issue": ("課題", "問題", "懸念"),
    "request": ("依頼", "お願い", "相談"),
}


@dataclass
class Style:
    bullet: str | None = None
    bulleted: bool = False
    ending: str | None = None
    period: bool | None = None
    headings: list[str] = dc_field(default_factory=list)
    opening: str | None = None
    closing: str | None = None
    sections: dict[str, list[str]] = dc_field(default_factory=dict)  # 見出し -> 過去の記載行
    samples: int = 0

    def describe(self) -> str:
        if not self.samples:
            return "参考事例なし（メモのまま）"
        parts = [f"事例{self.samples}件から学習"]
        if self.bulleted and self.bullet:
            parts.append("箇条書き「" + tu.make_bullet(self.bullet, 1).strip() + "」")
        elif not self.bulleted:
            parts.append("文章形式")
        if self.ending:
            parts.append({"desu": "です・ます調", "dearu": "だ・である調", "taigen": "体言止め"}[self.ending])
        if self.headings:
            parts.append("見出し " + "".join(self.headings))
        if self.closing:
            parts.append("結び「" + self.closing + "」")
        return "、".join(parts)


# ---------------------------------------------------------------------------
# 学習
# ---------------------------------------------------------------------------

def _lines(text: str) -> list[str]:
    return [l.rstrip() for l in (text or "").splitlines() if l.strip()]


def _is_plain(line: str) -> bool:
    return tu.heading_name(line) is None and tu.detect_bullet(line)[0] is None


def _formulaic(line: str, count: int, n: int) -> bool:
    """冒頭文・結び文として採用するか（半数以上の事例で同じ位置に現れる）。"""
    if n >= 2:
        return count >= max(2, math.ceil(n / 2))
    return count >= 1 and bool(FORMULAIC_RE.search(line)) and len(line) <= 40


def learn_style(texts: list[str]) -> Style:
    texts = [t for t in texts if t and t.strip()]
    style = Style(samples=len(texts))
    if not texts:
        return style
    n = len(texts)
    all_lines = [_lines(t) for t in texts]

    firsts = Counter(ls[0] for ls in all_lines if ls and _is_plain(ls[0]) and len(ls) > 1)
    lasts = Counter(ls[-1] for ls in all_lines if ls and _is_plain(ls[-1]) and len(ls) > 1)
    if firsts:
        line, cnt = firsts.most_common(1)[0]
        if _formulaic(line, cnt, n):
            style.opening = line
    if lasts:
        line, cnt = lasts.most_common(1)[0]
        if _formulaic(line, cnt, n) and line != style.opening:
            style.closing = line

    # 見出し：半数以上の事例に現れるもの（順序は上位の事例に従う）
    heading_count: Counter = Counter()
    order: list[str] = []
    for ls in all_lines:
        hs = [l.strip() for l in ls if tu.heading_name(l)]
        heading_count.update(set(hs))
        order.extend(h for h in hs if h not in order)
    style.headings = [h for h in order if heading_count[h] >= math.ceil(n / 2)]

    # 本文行の書き方（多数決）
    content: list[str] = []
    for ls in all_lines:
        current = None
        for l in ls:
            s = l.strip()
            if s in (style.opening, style.closing):
                continue
            if tu.heading_name(s):
                current = s
                continue
            content.append(s)
            if current:
                style.sections.setdefault(current, []).append(tu.strip_bullet(s))
    if content:
        bullets = Counter(tu.detect_bullet(l)[0] for l in content if tu.detect_bullet(l)[0])
        bullet_lines = sum(bullets.values())
        style.bulleted = bullet_lines * 2 >= len(content)
        if bullets:
            style.bullet = bullets.most_common(1)[0][0]
        bodies = [tu.strip_bullet(l).strip() for l in content]
        bodies = [b for b in bodies if b and b != NONE_TEXT]
        if bodies:
            style.ending = Counter(tu.ending_of(b) for b in bodies).most_common(1)[0][0]
            style.period = sum(tu.has_period(b) for b in bodies) * 2 >= len(bodies)
    return style


# ---------------------------------------------------------------------------
# 整形
# ---------------------------------------------------------------------------

def format_item(line: str, style: Style, index: int) -> str:
    text = tu.strip_bullet(line).strip()
    if text != NONE_TEXT and style.samples:
        if style.ending:
            text = tu.convert_ending(text, style.ending)
        else:
            text = tu.split_tail(text)[0]
        if style.period and not text.endswith(("。", "？", "！", "?", "!")):
            text += "。"
    elif text == NONE_TEXT:
        text = NONE_TEXT
    if style.bulleted and style.bullet:
        text = tu.make_bullet(style.bullet, index) + text
    return text


def _heading_roles(heading: str) -> set[str]:
    name = tu.heading_name(heading) or heading
    return {r for r, kws in HEADING_ROLES.items() if any(k in name for k in kws)}


HEADING_ROLE_BONUS = 0.15  # 見出しの振り分けでは、過去の記載との類似度を優先する


def _choose_heading(item: str, style: Style) -> str:
    role = tu.role_of(item)
    best, best_score = None, -1.0
    for h in style.headings:
        past = style.sections.get(h, [])
        s = max((tu.sim(item, p) for p in past), default=0.0)
        if role and role in _heading_roles(h):
            s += HEADING_ROLE_BONUS
        if s > best_score:
            best, best_score = h, s
    if best_score <= 0:
        # 手がかりがなければ「状況」系の見出し、なければ最初の見出し
        for h in style.headings:
            if "result" in _heading_roles(h):
                return h
        return style.headings[0]
    return best


def compose(items: list[tuple[str, str | None]], style: Style) -> str:
    """メモ項目を学習した書き方で 1 つの文章にする。items は (本文, 見出しの指定)。"""
    out: list[str] = []
    if style.opening:
        out.append(style.opening)
    if style.headings:
        buckets: dict[str, list[str]] = {h: [] for h in style.headings}
        for text, hint in items:
            h = hint if hint in buckets else _choose_heading(text, style)
            buckets[h].append(text)
        for h in style.headings:
            out.append(h)
            lines = buckets[h] or [NONE_TEXT]
            out.extend(format_item(l, style, i + 1) for i, l in enumerate(lines))
    else:
        out.extend(format_item(t, style, i + 1) for i, (t, _) in enumerate(items))
    if style.closing:
        out.append(style.closing)
    return "\n".join(out)


# ---------------------------------------------------------------------------
# メモの振り分け
# ---------------------------------------------------------------------------

def body_fields(fmt: dict) -> list[dict]:
    return [f for f in fmt["fields"] if f.get("enabled") and f["kind"] == "body"]


def main_field(fmt: dict) -> dict | None:
    bodies = body_fields(fmt)
    for f in bodies:
        if any(k in f["name"] for k in MAIN_KEYWORDS):
            return f
    return max(bodies, key=lambda f: f.get("height_px", 0), default=None)


def _match_field(name: str, fields: list[dict]) -> dict | None:
    key = tu.compact(name)
    if not key:
        return None
    for f in fields:
        if tu.compact(f["name"]) == key:
            return f
    for f in fields:
        fk = tu.compact(f["name"])
        if len(key) >= 2 and len(fk) >= 2 and (key in fk or fk in key):
            return f
    return None


def _heading_line(name: str, headings: list[str]) -> str | None:
    key = tu.compact(name)
    for h in headings:
        if tu.compact(tu.heading_name(h) or h) == key:
            return h
    return None


@dataclass
class Assignment:
    items: dict[str, list[tuple[str, str | None]]]  # field_id -> [(本文, 見出し)]
    how: dict[str, Counter]                         # field_id -> 振り分け方法の件数
    overrides: dict[str, str]                       # 件名・日付などの欄名指定
    table_lines: list[str]


def assign_memo(fmt: dict, memo: str, all_cases: list[dict], styles: dict[str, Style]) -> Assignment:
    fields = [f for f in fmt["fields"] if f.get("enabled")]
    bodies = body_fields(fmt)
    main = main_field(fmt)
    has_table = any(f["direction"] == "table" for f in fields)
    main_headings = styles[main["id"]].headings if main else []

    # 各本文欄の過去記載（行単位）
    past: dict[str, list[str]] = {}
    for f in bodies:
        lines = []
        for c in all_cases:
            for l in _lines(field_value_text(c["values"].get(f["id"]))):
                if tu.heading_name(l) is None:
                    lines.append(tu.strip_bullet(l))
        past[f["id"]] = lines

    items = {f["id"]: [] for f in bodies}
    how = {f["id"]: Counter() for f in bodies}
    overrides: dict[str, str] = {}
    table_lines: list[str] = []
    current: dict | None = None
    current_hint: str | None = None

    for raw in (memo or "").splitlines():
        line = raw.strip()
        if not line:
            continue
        if has_table and TABLE_SPLIT_RE.search(line):
            table_lines.append(line)
            continue

        m = LABELED_RE.match(line)
        if m:
            name = m.group("h") or m.group("n")
            rest = (m.group("hrest") if m.group("h") else m.group("nrest")) or ""
            target = _match_field(name, fields)
            heading = _heading_line(name, main_headings) if main else None
            if target is not None and target["kind"] == "body":
                if rest:
                    items[target["id"]].append((rest, None))
                    how[target["id"]]["欄名指定"] += 1
                else:
                    current, current_hint = target, None
                continue
            if target is not None and target["kind"] in ("subject", "date", "reporter", "dept"):
                if rest:
                    overrides[target["id"]] = rest
                continue
            if heading is not None:
                if rest:
                    items[main["id"]].append((rest, heading))
                    how[main["id"]]["見出し指定"] += 1
                else:
                    current, current_hint = main, heading
                continue
            # 欄名でなければ普通の行として扱う（「10:00 発生」など）

        if current is not None:
            items[current["id"]].append((line, current_hint))
            how[current["id"]]["見出しの下"] += 1
            continue
        if not bodies:
            continue

        role = tu.role_of(line)
        best, best_score = None, -1.0
        for f in bodies:
            score = 0.7 * max((tu.sim(line, p) for p in past[f["id"]]), default=0.0)
            if role and role in tu.roles_of_name(f["name"]):
                score += 0.35
            if score > best_score:
                best, best_score = f, score
        if best_score < LOW_SCORE and main is not None:
            best = main
            how[best["id"]]["本文へ"] += 1
        else:
            how[best["id"]]["類似度"] += 1
        items[best["id"]].append((line, None))

    return Assignment(items, how, overrides, table_lines)


# ---------------------------------------------------------------------------
# 欄ごとの記入内容
# ---------------------------------------------------------------------------

def _column_role(name: str) -> str | None:
    for role, words in TABLE_COLUMN_ROLES:
        if role == "seq":
            if name.strip() in words or name.strip().lower() in ("no", "no.", "#"):
                return role
        elif any(w in name for w in words):
            return role
    return None


def _fill_table(columns: list[dict], lines: list[str]) -> dict[str, dict]:
    rows_available = min(len(c["cells"]) for c in columns)
    rows = [TABLE_SPLIT_RE.split(l.strip().strip("｜|")) for l in lines]
    note_extra = ""
    if len(rows) > rows_available:
        note_extra = f"（{len(rows)}行中 {rows_available}行まで記入。残りは表に入りません）"
        rows = rows[:rows_available]

    roles = [_column_role(c["name"]) for c in columns]
    values = {c["id"]: [] for c in columns}
    for i, parts in enumerate(rows):
        by_role = {r: (parts[k] if k < len(parts) else "") for k, r in enumerate(MEMO_TABLE_ORDER)}
        extra = parts[len(MEMO_TABLE_ORDER):]
        used_roles = set()
        for c, role in zip(columns, roles):
            if role == "seq":
                values[c["id"]].append(str(i + 1))
            elif role in by_role and role not in used_roles:
                values[c["id"]].append(by_role[role])
                used_roles.add(role)
            else:
                values[c["id"]].append(extra.pop(0) if extra else "")
    out = {}
    for c in columns:
        text = "\n".join(values[c["id"]])
        if lines:
            out[c["id"]] = {"text": text, "status": "memo",
                            "note": "メモの「内容｜担当｜期限｜状態」から記入" + note_extra}
        else:
            out[c["id"]] = {"text": "", "status": "empty",
                            "note": "メモに「内容｜担当｜期限｜状態」の行がありません"}
    return out


def _most_common(cases: list[dict], fid: str) -> str:
    vals = Counter(field_value_text(c["values"].get(fid)) for c in cases)
    vals.pop("", None)
    return vals.most_common(1)[0][0] if vals else ""


def resolve_date_format(fmt: dict, field: dict, settings: dict, cases: list[dict]) -> str:
    chosen = settings.get("date_format") or "auto"
    if chosen != "auto":
        return chosen
    samples = [field_value_text(c["values"].get(field["id"])) for c in cases]
    return tu.detect_date_format(samples) or "{Y}/{MM}/{DD}"


def _reuse_text(case: dict, fid: str, subject_old: str, topic: str, today: date) -> str:
    text = field_value_text(case["values"].get(fid))
    if subject_old and topic and subject_old in text:
        text = text.replace(subject_old, topic)
    return tu.replace_dates(text, today)


def generate(fmt: dict, topic: str, memo: str, ref_cases: list[dict], all_cases: list[dict],
             settings: dict, today: date | None = None) -> dict[str, dict]:
    """下書きを作る。ref_cases は書き方を学習する参考事例（類似度順、最大 3 件を使う）。"""
    from .cases import subject_of  # 循環 import を避ける

    today = today or date.today()
    ref_cases = list(ref_cases)[:3]
    fields = [f for f in fmt["fields"] if f.get("enabled")]
    styles = {f["id"]: learn_style([field_value_text(c["values"].get(f["id"])) for c in ref_cases])
              for f in body_fields(fmt)}
    assignment = assign_memo(fmt, memo, all_cases, styles)
    source_cases = ref_cases or all_cases
    results: dict[str, dict] = {}

    # 表（table_id ごと。最初の表にメモの表行を入れる）
    tables: dict[str, list[dict]] = {}
    for f in fields:
        if f["direction"] == "table":
            tables.setdefault(f["table_id"], []).append(f)
    for i, (tid, columns) in enumerate(tables.items()):
        results.update(_fill_table(columns, assignment.table_lines if i == 0 else []))

    top_case = ref_cases[0] if ref_cases else None
    old_subject = subject_of(fmt, top_case) if top_case else ""

    for f in fields:
        fid = f["id"]
        if f["direction"] == "table":
            continue
        kind = f["kind"]
        if fid in assignment.overrides:
            results[fid] = {"text": assignment.overrides[fid], "status": "memo", "note": "メモの欄名指定から記入"}
        elif kind == "subject":
            results[fid] = {"text": topic.strip(), "status": "memo" if topic.strip() else "empty",
                            "note": "トピックを記入"}
        elif kind == "date":
            fmt_str = resolve_date_format(fmt, f, settings, source_cases)
            how = "設定の書式" if (settings.get("date_format") or "auto") != "auto" else "過去事例の書式"
            results[fid] = {"text": tu.format_date(today, fmt_str), "status": "auto",
                            "note": f"本日の日付（{how}）"}
        elif kind in ("reporter", "dept"):
            key = "reporter" if kind == "reporter" else "department"
            if settings.get(key):
                results[fid] = {"text": settings[key], "status": "auto", "note": "設定の値"}
            else:
                v = _most_common(all_cases, fid)
                results[fid] = {"text": v, "status": "auto" if v else "empty",
                                "note": "過去事例で最も多い値" if v else "設定・過去事例に値がありません"}
        elif kind == "body":
            items = assignment.items.get(fid, [])
            style = styles[fid]
            if items:
                how = "、".join(f"{k}{v}行" for k, v in assignment.how[fid].items())
                results[fid] = {"text": compose(items, style), "status": "memo",
                                "note": f"メモ{len(items)}行（{how}）／{style.describe()}"}
            elif settings.get("reuse", True) and top_case and field_value_text(top_case["values"].get(fid)):
                results[fid] = {"text": _reuse_text(top_case, fid, old_subject, topic, today), "status": "reuse",
                                "note": f"要確認：事例「{old_subject}」の文面を流用（件名・日付を置換）"}
            else:
                results[fid] = {"text": "", "status": "empty", "note": "該当するメモがありません"}
        else:
            results[fid] = {"text": "", "status": "empty", "note": ""}
    return results


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------

def to_cell_values(fmt: dict, texts: dict[str, str]) -> dict[tuple[str, str], str]:
    """{field_id: 文字列} を write_cells に渡す {(シート, セル): 値} にする。"""
    values: dict[tuple[str, str], str] = {}
    for f in fmt["fields"]:
        if not f.get("enabled") or f["id"] not in texts:
            continue
        text = texts[f["id"]] or ""
        if f["direction"] == "table":
            lines = text.split("\n")
            for i, ref in enumerate(f["cells"]):
                values[(f["sheet"], ref)] = lines[i] if i < len(lines) else ""
        elif f["direction"] == "inline":
            values[(f["sheet"], f["cells"][0])] = (f.get("prefix") or "") + text
        else:
            values[(f["sheet"], f["cells"][0])] = text
    return values


_FILE_FORBIDDEN_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def make_filename(pattern: str, today: date, topic: str, reporter: str = "") -> str:
    """ファイル名（拡張子なし）。禁止文字は「_」に置き換える。"""
    name = (pattern or "{date}_{topic}")
    name = name.replace("{date}", f"{today.year}{today.month:02d}{today.day:02d}")
    name = name.replace("{topic}", topic.strip() or "報告書").replace("{reporter}", reporter or "")
    name = _FILE_FORBIDDEN_RE.sub("_", name).replace("\n", " ").strip(" .")
    return name[:120] or "報告書"
