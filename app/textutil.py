"""日本語テキスト処理：正規化・n-gram 類似度・箇条書き・文末変換・役割判定・日付書式。

すべて標準ライブラリだけで動くルールベースの処理です。
"""
from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from datetime import date

# ---------------------------------------------------------------------------
# 正規化・n-gram
# ---------------------------------------------------------------------------

# 類似度計算で無視する空白・記号（長音「ー」はカタカナ語の一部なので残す）
_STRIP_RE = re.compile(
    r"[\s　、。，．,.・:：;；!！?？「」『』（）()\[\]［］【】<>＜＞〈〉\-－―~〜|｜/／*＊■□●○◆◇※]+"
)


def nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")


def compact(text: str) -> str:
    """比較用：NFKC 正規化し、小文字化して空白と記号を除く。"""
    return _STRIP_RE.sub("", nfkc(text).lower())


def ngrams(text: str, sizes=(2, 3)) -> Counter:
    """文字 n-gram の出現回数（既定は 2-gram と 3-gram）。"""
    t = compact(text)
    grams: Counter = Counter()
    for n in sizes:
        for i in range(len(t) - n + 1):
            grams[t[i:i + n]] += 1
    if len(t) == 1:
        grams[t] += 1
    return grams


def cosine(a: Counter | dict, b: Counter | dict) -> float:
    if not a or not b:
        return 0.0
    if len(a) > len(b):
        a, b = b, a
    dot = sum(v * b.get(k, 0) for k, v in a.items())
    if not dot:
        return 0.0
    na = math.sqrt(sum(v * v for v in a.values()))
    nb = math.sqrt(sum(v * v for v in b.values()))
    return dot / (na * nb)


def sim(a: str, b: str) -> float:
    """2 つの文字列の 2/3-gram コサイン類似度（0〜1）。"""
    return cosine(ngrams(a), ngrams(b))


# ---------------------------------------------------------------------------
# 見出し・箇条書き
# ---------------------------------------------------------------------------

HEADING_RE = re.compile(r"^\s*([【［\[〈＜<])\s*(.+?)\s*([】］\]〉＞>])\s*$")

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"

# (種類, 正規表現)。上から順に判定する。
_BULLET_RES = [
    ("circled", re.compile(r"^\s*([①-⑳])(\s*)")),
    ("paren", re.compile(r"^\s*([（(])\s*(\d{1,2})\s*([）)])(\s*)")),
    ("dot", re.compile(r"^\s*(\d{1,2})([.．])(?!\d)(\s*)")),
    ("close", re.compile(r"^\s*(\d{1,2})([)）])(\s*)")),
    ("symbol", re.compile(r"^\s*([・■□●○◆◇▪▫\-－＊*＞>])(\s*)")),
]


def heading_name(line: str) -> str | None:
    """「【背景】」のような見出しだけの行なら中身（背景）を返す。"""
    m = HEADING_RE.match(line or "")
    return m.group(2) if m else None


def detect_bullet(line: str):
    """箇条書き記号を判定し (テンプレート, 記号を除いた本文) を返す。記号がなければ (None, 行)。

    テンプレートは付け直しに使う文字列で、連番の場合は {n} を含む。
    例: "・", "■ ", "{n}. ", "（{n}）", "circled"
    """
    for kind, rx in _BULLET_RES:
        m = rx.match(line)
        if not m:
            continue
        rest = line[m.end():]
        if kind == "circled":
            return "circled" + m.group(2), rest
        if kind == "paren":
            return f"{m.group(1)}{{n}}{m.group(3)}{m.group(4)}", rest
        if kind == "dot":
            return f"{{n}}{m.group(2)}{m.group(3)}", rest
        if kind == "close":
            return f"{{n}}{m.group(2)}{m.group(3)}", rest
        return m.group(1) + m.group(2), rest
    return None, line


def strip_bullet(line: str) -> str:
    return detect_bullet(line)[1]


def make_bullet(template: str | None, n: int) -> str:
    """テンプレートから n 番目の記号を作る。"""
    if not template:
        return ""
    if template.startswith("circled"):
        space = template[len("circled"):]
        return (CIRCLED[n - 1] if 1 <= n <= len(CIRCLED) else f"({n})") + space
    if "{n}" in template:
        return template.replace("{n}", str(n))
    return template


# ---------------------------------------------------------------------------
# 文末
# ---------------------------------------------------------------------------

_TAIL_PUNCT_RE = re.compile(r"[。．.！!]+$")

_DESU_RE = re.compile(
    r"(です|ます|ました|でした|ません|ませんでした|ましょう|ください|下さい|いたします|致します|ございます)$"
)
_DEARU_RE = re.compile(r"(だ|である|だった|であった|た|ない|る|う|く|す|つ|ぬ|ぶ|む|ぐ|い)$")


def split_tail(text: str) -> tuple[str, str]:
    """本文と末尾の句点類に分ける。"""
    text = text.rstrip()
    m = _TAIL_PUNCT_RE.search(text)
    if m:
        return text[:m.start()], m.group(0)
    return text, ""


def ending_of(text: str) -> str:
    """文末の調子：'desu'（です・ます調） / 'dearu'（だ・である調） / 'taigen'（体言止め）。"""
    body, _ = split_tail(text)
    body = body.rstrip("）)」』 ")
    if _DESU_RE.search(body):
        return "desu"
    if _DEARU_RE.search(body):
        return "dearu"
    return "taigen"


def has_period(text: str) -> bool:
    return text.rstrip().endswith(("。", "．"))


# 変換表（長い接尾辞から順に照合する）
TO_DESU = [
    ("していなかった", "していませんでした"), ("できなかった", "できませんでした"),
    ("しなかった", "しませんでした"), ("ていなかった", "ていませんでした"),
    ("していない", "していません"), ("していた", "していました"), ("している", "しています"),
    ("できない", "できません"), ("できた", "できました"), ("できる", "できます"),
    ("しない", "しません"), ("した", "しました"), ("する", "します"),
    ("ていない", "ていません"), ("ていた", "ていました"), ("ている", "ています"), ("てある", "てあります"),
    ("であった", "でした"), ("だった", "でした"), ("である", "です"),
    ("がない", "がありません"), ("はない", "はありません"), ("もない", "もありません"),
    ("問題ない", "問題ありません"),
    ("なった", "なりました"), ("なる", "なります"), ("あった", "ありました"), ("ある", "あります"),
    ("れた", "れました"), ("れる", "れます"), ("せた", "せました"), ("せる", "せます"),
    ("きた", "きました"), ("いた", "いました"), ("いる", "います"),
    ("だ", "です"),
]

TO_DEARU = [
    ("していませんでした", "していなかった"), ("できませんでした", "できなかった"),
    ("しませんでした", "しなかった"), ("ていませんでした", "ていなかった"),
    ("していません", "していない"), ("していました", "していた"), ("しています", "している"),
    ("できません", "できない"), ("できました", "できた"), ("できます", "できる"),
    ("しません", "しない"), ("しました", "した"), ("します", "する"),
    ("ていません", "ていない"), ("ていました", "ていた"), ("ています", "ている"), ("てあります", "てある"),
    ("ありませんでした", "なかった"), ("ありません", "ない"),
    ("なりました", "なった"), ("なります", "なる"), ("ありました", "あった"), ("あります", "ある"),
    ("れました", "れた"), ("れます", "れる"), ("せました", "せた"), ("せます", "せる"),
    ("きました", "きた"), ("いました", "いた"), ("います", "いる"),
    ("いたします", "する"), ("致します", "する"),
    ("でした", "だった"), ("です", "である"),
]

TO_TAIGEN = [
    ("しています", "中"), ("している", "中"),
    ("しました", ""), ("した", ""),
    ("する予定です", "予定"), ("する予定", "予定"),
    ("します", "予定"), ("する", "予定"),
    ("です", ""), ("である", ""), ("だ", ""),
]


def _apply_table(body: str, table) -> str:
    for src, dst in table:
        if body.endswith(src):
            return body[: len(body) - len(src)] + dst
    return body


def convert_ending(text: str, target: str, add_desu: bool = True) -> str:
    """文末を target（'desu' / 'dearu' / 'taigen'）にそろえる。句点は外した状態で返す。"""
    body, _ = split_tail(text)
    if not body or not target:
        return body
    current = ending_of(body)
    if target == "desu":
        if current == "desu":
            return body
        if current == "taigen":
            return body + "です" if add_desu else body
        return _apply_table(body, TO_DESU)
    if target == "dearu":
        if current == "desu":
            return _apply_table(body, TO_DEARU)
        return body
    if target == "taigen":
        if current == "taigen":
            return body
        out = _apply_table(body, TO_TAIGEN)
        return out or body
    return body


# ---------------------------------------------------------------------------
# 役割（依頼・予定・課題・実績）
# ---------------------------------------------------------------------------

ROLE_RULES = [
    ("request", re.compile(r"お願い|ご判断|ご確認ください|ご検討|ご承認|ご相談")),
    ("plan", re.compile(r"予定|までに|今後|次回|(する|します)$")),
    ("issue", re.compile(r"課題|懸念|未確認|ていない|問題|リスク|できない")),
    ("result", re.compile(r"(た|ました)$|完了|原因|発生|済み|判明")),
]

ROLE_LABELS = {"request": "依頼", "plan": "予定", "issue": "課題", "result": "実績"}

# 欄名・見出し名から役割を推定するキーワード
FIELD_ROLE_KEYWORDS = {
    "issue": ("課題", "懸念", "問題", "リスク"),
    "plan": ("今後", "予定", "対応", "対策", "次回", "アクション", "計画"),
    "request": ("依頼", "お願い", "相談", "判断", "連絡事項"),
    "result": ("内容", "概要", "報告", "状況", "実績", "結果", "経過", "現状"),
}


def role_of(line: str) -> str | None:
    body, _ = split_tail(strip_bullet(line).strip())
    for role, rx in ROLE_RULES:
        if rx.search(body):
            return role
    return None


def roles_of_name(name: str) -> set[str]:
    return {role for role, kws in FIELD_ROLE_KEYWORDS.items() if any(k in (name or "") for k in kws)}


# ---------------------------------------------------------------------------
# 日付（Windows の strftime は %-m 非対応のため自前で整形）
# ---------------------------------------------------------------------------

WEEKDAYS = "月火水木金土日"

_DATE_KANJI_RE = re.compile(r"(\d{4})年(\d{1,2})月(\d{1,2})日(\s*[（(][月火水木金土日][）)])?")
_DATE_SEP_RE = re.compile(r"(?<!\d)(\d{4})([/\-.])(\d{1,2})\2(\d{1,2})(?!\d)(\s*[（(][月火水木金土日][）)])?")


def format_date(d: date, fmt: str) -> str:
    """{Y} {MM} {M} {DD} {D} {W} を置き換える。"""
    return (
        fmt.replace("{Y}", str(d.year))
        .replace("{MM}", f"{d.month:02d}")
        .replace("{M}", str(d.month))
        .replace("{DD}", f"{d.day:02d}")
        .replace("{D}", str(d.day))
        .replace("{W}", WEEKDAYS[d.weekday()])
    )


def _weekday_suffix(raw: str | None) -> str:
    if not raw:
        return ""
    raw = raw.strip()
    return raw[0] + "{W}" + raw[-1]


def _padded(month: str, day: str) -> bool | None:
    if month.startswith("0") or day.startswith("0"):
        return True
    if len(month) == 1 or len(day) == 1:
        return False
    return None  # 10/15 などは判別できない


def date_format_of(match) -> tuple[str, bool | None]:
    """日付の正規表現マッチから書式テンプレートを作る。"""
    if match.re is _DATE_KANJI_RE:
        y, m, d, w = match.groups()
        pad = _padded(m, d)
        mm, dd = ("{MM}", "{DD}") if pad else ("{M}", "{D}")
        return f"{{Y}}年{mm}月{dd}日" + _weekday_suffix(w), pad
    y, sep, m, d, w = match.groups()
    pad = _padded(m, d)
    mm, dd = ("{MM}", "{DD}") if pad in (True, None) else ("{M}", "{D}")
    return f"{{Y}}{sep}{mm}{sep}{dd}" + _weekday_suffix(w), pad


def find_dates(text: str):
    """文字列中の日付マッチを出現順に返す。"""
    found = list(_DATE_KANJI_RE.finditer(text or "")) + list(_DATE_SEP_RE.finditer(text or ""))
    return sorted(found, key=lambda m: m.start())


def detect_date_format(samples) -> str | None:
    """過去の日付文字列から最も多い書式を推定する。"""
    votes: Counter = Counter()
    for s in samples:
        for m in find_dates(str(s or ""))[:1]:
            fmt, _ = date_format_of(m)
            votes[fmt] += 1
    return votes.most_common(1)[0][0] if votes else None


def replace_dates(text: str, today: date) -> str:
    """文中の日付を、それぞれ元と同じ書式で本日の日付に置き換える。"""
    if not text:
        return text
    out, pos = [], 0
    for m in find_dates(text):
        if m.start() < pos:
            continue
        fmt, _ = date_format_of(m)
        out.append(text[pos:m.start()])
        out.append(format_date(today, fmt))
        pos = m.end()
    out.append(text[pos:])
    return "".join(out)
