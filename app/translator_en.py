"""英語の生成（オフライン）。

- 定型部分（冒頭文・結び文・見出し・工程名・進捗・「特になし」など）は対訳テンプレートで訳す。
- 自由記述は次の順で訳す。
    1. 翻訳メモリ（過去事例の対訳）に類似度 0.9 以上の文があれば、それを使う
    2. 用語集の語を記号（T1, T2 …）に置き換えて保護し、CTranslate2 で機械翻訳する
       ※ 仕様例の「⟦T1⟧」は opus-mt の語彙にない文字のため訳文で壊れる。
         英数字の記号 T1 なら訳文にそのまま残るため、この形にしている。
    3. 記号を英語の用語に戻し、上長向けの報告文（We を主語、短縮形なし）に整える
- 日付・数値・時刻・英数字の固有名詞が訳文に残っているか検証し、一致しなければ要確認にする。
- 翻訳モデルがなければ、用語集とテンプレートだけで簡易訳を作る（要確認（簡易訳））。

実行時に torch / transformers は使わない。通信もしない。
"""
from __future__ import annotations

import re
import threading
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from . import textutil as tu
from .glossary import Glossary

TM_THRESHOLD = 0.9

# 行全体の定型文
LINE_TEMPLATES = {
    "お疲れ様です。": "Please find my report below.",
    "お疲れさまです。": "Please find my report below.",
    "お疲れ様です": "Please find my report below.",
    "以上、ご報告いたします。": "That concludes my report.",
    "以上、ご報告します。": "That concludes my report.",
    "以上、報告いたします。": "That concludes my report.",
    "以上です。": "That is all.",
    "以上": "That is all.",
    "よろしくお願いいたします。": "Thank you for your continued support.",
    "よろしくお願いします。": "Thank you for your continued support.",
    "ご確認のほど、よろしくお願いいたします。": "Your review would be greatly appreciated.",
    "ご確認をお願いいたします。": "Your review would be greatly appreciated.",
    "特になし": "None",
    "特になし。": "None.",
    "なし": "None",
}

HEADINGS = {
    "背景": "Background", "経緯": "Background", "状況": "Status", "現状": "Current Status",
    "対応": "Actions", "対策": "Countermeasures", "今後の予定": "Next Steps", "今後の対応": "Next Steps",
    "予定": "Plan", "課題": "Issues", "結果": "Results", "概要": "Overview", "原因": "Cause",
    "影響": "Impact", "依頼事項": "Requests", "報告内容": "Report", "進捗": "Progress",
    "実績": "Achievements", "所感": "Remarks", "備考": "Notes", "詳細": "Details",
}

PHASES = {
    "要件定義": "Requirements definition", "基本設計": "Basic design", "詳細設計": "Detailed design",
    "設計": "Design", "開発": "Development", "単体テスト": "Unit testing", "結合テスト": "Integration testing",
    "総合テスト": "System integration testing", "運用テスト": "Operational testing",
    "移行リハーサル": "Migration rehearsal", "本番移行": "Production cutover",
    "稼働後フォロー": "Post-go-live support", "稼働判定": "Go-live decision",
}

STATUSES = {
    "予定どおり": "On schedule", "予定通り": "On schedule", "遅延": "Delayed", "完了": "Completed",
    "中止": "Cancelled", "対応中": "In progress", "未着手": "Not started", "保留": "On hold",
    "確認中": "Under review", "実施中": "In progress", "済": "Done", "済み": "Done",
}

FIELD_NAMES = {
    "date": "Date", "subject": "Subject", "phase": "Phase", "status": "Status", "reporter": "Reporter",
    "dept": "Department", "body": "Report", "issue": "Issues", "plan": "Next Steps", "request": "Requests",
}
COLUMN_NAMES = {"seq": "No.", "content": "Action", "owner": "Owner", "due": "Due", "status": "Status"}

_NUMBER_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "fifteen": 15, "twenty": 20, "thirty": 30,
}
_CONTRACTIONS = [
    (r"\bI'm\b", "I am"), (r"\bI've\b", "I have"), (r"\bI'll\b", "I will"), (r"\bI'd\b", "I would"),
    (r"\bwe're\b", "we are"), (r"\bwe've\b", "we have"), (r"\bwe'll\b", "we will"), (r"\bwe'd\b", "we would"),
    (r"\byou're\b", "you are"), (r"\byou've\b", "you have"), (r"\byou'll\b", "you will"),
    (r"\bit's\b", "it is"), (r"\bthat's\b", "that is"), (r"\bthere's\b", "there is"),
    (r"\bcan't\b", "cannot"), (r"\bwon't\b", "will not"), (r"\bdon't\b", "do not"), (r"\bdoesn't\b", "does not"),
    (r"\bdidn't\b", "did not"), (r"\bisn't\b", "is not"), (r"\baren't\b", "are not"), (r"\bwasn't\b", "was not"),
    (r"\bweren't\b", "were not"), (r"\bhasn't\b", "has not"), (r"\bhaven't\b", "have not"),
]
_PRONOUN_FIXES = [
    (r"^(?:I|You) am\b", "We are"), (r"^(?:I|You) was\b", "We were"), (r"^(?:I|You) have\b", "We have"),
    (r"^(?:I|You) has\b", "We have"), (r"^(?:I|You) will\b", "We will"), (r"^(?:I|You)\b", "We"),
]
_JA_DATE_FULL_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_JA_DATE_MD_RE = re.compile(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_DIGITS_RE = re.compile(r"\d+(?:\.\d+)?")
_ASCII_WORD_RE = re.compile(r"[A-Za-z][A-Za-z0-9.\-]*[A-Za-z0-9]|[A-Za-z]{2,}")
_PLACEHOLDER_RE = re.compile(r"\bT(\d{1,2})\b")
_JA_CHAR_RE = re.compile(r"[぀-ヿ㐀-鿿]")
_ORDINAL_RE = re.compile(r"(?P<open>[（(]\s*)?第\s*(?P<n>\d{1,3})\s*(?P<unit>回|次|期|版|弾)(?(open)\s*[）)])")
_ORDINAL_UNITS = {"回": "round", "次": "round", "期": "term", "版": "edition", "弾": "wave"}
_DATE_PREP_RE = re.compile(r"\b(?:at|in)\s+(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\b")


def ordinal(n: int) -> str:
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


@dataclass
class LineResult:
    text: str
    method: str                 # template / tm / mt / simple / copy
    review: bool = False
    notes: list[str] = field(default_factory=list)


@dataclass
class TextResult:
    text: str
    method: str
    review: bool
    notes: list[str]

    @property
    def status(self) -> str:
        return self.method


METHOD_LABELS = {"template": "定型文", "tm": "翻訳メモリ", "mt": "機械翻訳", "simple": "簡易訳", "copy": "そのまま", "empty": "空欄"}


def normalize_ja(text: str) -> str:
    """翻訳前の正規化：全角英数を半角に、「9月25日」を「9/25」に。"""
    s = tu.nfkc(text)
    s = _JA_DATE_FULL_RE.sub(lambda m: f"{m.group(1)}/{m.group(2)}/{m.group(3)}", s)
    return _JA_DATE_MD_RE.sub(lambda m: f"{m.group(1)}/{m.group(2)}", s)


_MONTH_RE = re.compile(r"\b(" + "|".join(m[:3] for m in tu.MONTHS_EN) + r")[a-z]*\.?\b", re.I)


def numbers_in(text: str) -> Counter:
    """数値の出現回数。先頭のゼロは無視し（09 と 9 は同じ）、英語の月名は月の数字として数える。"""
    nums = Counter()
    for n in _DIGITS_RE.findall(text or ""):
        nums[n if "." in n else str(int(n))] += 1
    for m in _MONTH_RE.finditer(text or ""):
        nums[str([x[:3].lower() for x in tu.MONTHS_EN].index(m.group(1).lower()) + 1)] += 1
    return nums


def check_consistency(ja: str, en: str) -> list[str]:
    """日付・数値・英数字の固有名詞が訳文に残っているか。問題があれば説明を返す。"""
    problems = []
    missing = numbers_in(normalize_ja(ja)) - numbers_in(en)
    if missing:
        problems.append("数値・日付が一致しません（" + "、".join(sorted(missing)) + "）")
    en_low = (en or "").lower()
    words = {w for w in _ASCII_WORD_RE.findall(normalize_ja(ja)) if not re.fullmatch(r"T\d+", w)}
    lost = [w for w in words if w.lower() not in en_low]
    if lost:
        problems.append("英数字の固有名詞が訳文にありません（" + "、".join(sorted(lost)) + "）")
    return problems


def _fix_number_words(ja: str, en: str) -> str:
    """機械翻訳が「3件」を "Three" のように言い換えたとき、数字に戻す。"""
    missing = numbers_in(normalize_ja(ja)) - numbers_in(en)
    for num in list(missing):
        for word, value in _NUMBER_WORDS.items():
            if str(value) == num:
                en, n = re.subn(rf"\b{word}\b", num, en, count=1, flags=re.I)
                if n:
                    break
    return en


def polish(en: str, ja: str) -> str:
    """上長向けの報告文に整える（短縮形を使わない・主語は We・文頭大文字・句点）。"""
    s = " ".join((en or "").split())
    for pat, rep in _CONTRACTIONS:
        s = re.sub(pat, rep, s, flags=re.I)
    for pat, rep in _PRONOUN_FIXES:
        s, n = re.subn(pat, rep, s)
        if n:
            break
    s = _fix_number_words(ja, s)
    s = _DATE_PREP_RE.sub(r"on \1", s)      # 日付の前置詞は on（at 9/26 → on 9/26）
    if s:
        s = s[0].upper() + s[1:]
    ja_body = tu.split_tail(ja)
    if ja_body[1] and not s.endswith((".", "!", "?")):
        s += "."
    if not ja_body[1] and s.endswith(".") and len(s.split()) <= 6:
        s = s[:-1]   # 体言止め（短い語句）には句点を付けない
    return s


class Translator:
    """日本語 → 英語。モデルは初回使用時に読み込む（スレッドセーフ）。"""

    def __init__(self, model_dir, terms=(), tm_pairs=(), threads: int = 4):
        self.model_dir = Path(model_dir) if model_dir else None
        self.glossary = Glossary(terms)
        self.threads = threads
        self._lock = threading.Lock()
        self._loaded = False
        self._translator = None
        self._sp_src = None
        self._sp_tgt = None
        self.error: str | None = None
        self.set_memory(tm_pairs)

    # ---- 準備 ----
    def set_memory(self, tm_pairs) -> None:
        self.tm = []
        for ja, en in tm_pairs or ():
            key = tu.strip_bullet(ja.strip()).strip()
            if key:
                self.tm.append((key, tu.strip_bullet(en.strip()).strip(), tu.ngrams(key)))

    def load(self) -> bool:
        with self._lock:
            if self._loaded:
                return self._translator is not None
            self._loaded = True
            try:
                if self.model_dir is None or not (self.model_dir / "model.bin").exists():
                    raise FileNotFoundError(f"翻訳モデルが見つかりません: {self.model_dir}")
                import ctranslate2
                import sentencepiece as spm
                self._translator = ctranslate2.Translator(str(self.model_dir), device="cpu", compute_type="int8",
                                                          intra_threads=self.threads)
                # sentencepiece は Windows で日本語を含むパスを開けないため、バイト列で渡す
                self._sp_src = spm.SentencePieceProcessor(model_proto=(self.model_dir / "source.spm").read_bytes())
                self._sp_tgt = spm.SentencePieceProcessor(model_proto=(self.model_dir / "target.spm").read_bytes())
            except Exception as e:  # モデルがない・壊れている → 簡易訳で続ける
                self.error = str(e)
                self._translator = None
            return self._translator is not None

    @property
    def available(self) -> bool:
        return self.load()

    # ---- 1 行の定型・メモリ ----
    def _template(self, body: str) -> str | None:
        b = body.strip()
        if b in LINE_TEMPLATES:
            return LINE_TEMPLATES[b]
        core = tu.split_tail(b)[0]
        for table in (STATUSES, PHASES):
            if core in table:
                return table[core]
        return None

    def _heading(self, line: str) -> str | None:
        name = tu.heading_name(line)
        if name is None:
            return None
        en = HEADINGS.get(name) or self.glossary.lookup(name) or PHASES.get(name)
        if en is None:
            en = self.translate_phrase(name)
        return f"[{en}]"

    def _memory(self, body: str) -> str | None:
        if not self.tm:
            return None
        q = tu.ngrams(body)
        best, best_s = None, 0.0
        for ja, en, grams in self.tm:
            s = 1.0 if ja == body else tu.cosine(q, grams)
            if s > best_s:
                best, best_s = en, s
        return best if best_s >= TM_THRESHOLD else None

    # ---- 機械翻訳 ----
    def _protect(self, text: str) -> tuple[str, list[str], list[str]]:
        """用語（と「第2回」のような回数）を T1, T2 … に置き換える。(保護後の文, 英語の用語, 日本語の用語)"""
        found = self.glossary.find(text)
        for m in _ORDINAL_RE.finditer(text):
            s, e = m.span()
            if not any(not (e <= fs or s >= fe) for fs, fe, _, _ in found):
                n = int(m.group("n"))
                en = f"{ordinal(n)} {_ORDINAL_UNITS.get(m.group('unit'), 'round')}"
                found.append((s, e, m.group(0), f"({en})" if m.group("open") else en))
        found.sort()
        # 隣り合う用語（間が空白だけ）は 1 つの記号にまとめる（T1T2 のような並びは訳文で壊れやすい）
        merged: list[tuple[int, int, str, str]] = []
        for item in found:
            if merged and not text[merged[-1][1]:item[0]].strip():
                ps, _, pja, pen = merged[-1]
                merged[-1] = (ps, item[1], text[ps:item[1]], f"{pen} {item[3]}")
            else:
                merged.append(item)
        out, pos, ens, jas = [], 0, [], []
        for s, e, ja, en in merged:
            out.append(text[pos:s])
            ens.append(en)
            jas.append(ja)
            out.append(f"T{len(ens)}")
            pos = e
        out.append(text[pos:])
        return "".join(out), ens, jas

    def _mt(self, sources: list[str]) -> list[LineResult]:
        prepared = [self._protect(normalize_ja(s)) for s in sources]
        batch = [self._sp_src.encode(p[0], out_type=str) + ["</s>"] for p in prepared]
        with self._lock:
            outputs = self._translator.translate_batch(batch, beam_size=4, max_decoding_length=256,
                                                       repetition_penalty=1.1)
        results = []
        for src, (protected, ens, jas), out in zip(sources, prepared, outputs):
            en = self._sp_tgt.decode(out.hypotheses[0])
            notes = []
            used = set()

            def restore(m):
                i = int(m.group(1)) - 1
                if 0 <= i < len(ens):
                    used.add(i)
                    return ens[i]
                return m.group(0)

            en = _PLACEHOLDER_RE.sub(restore, en)
            lost = [jas[i] for i in range(len(ens)) if i not in used]
            if lost:
                notes.append("用語が訳文に残りませんでした（" + "、".join(lost) + "）")
                en = en.rstrip(".") + " (" + ", ".join(ens[i] for i in range(len(ens)) if i not in used) + ")."
            en = polish(en, src)
            notes += check_consistency(src, en)
            results.append(LineResult(en, "mt", bool(notes), notes))
        return results

    def _simple(self, source: str) -> LineResult:
        """モデルがないときの簡易訳：用語集で置き換えられる語だけ英語にする。"""
        s = self.glossary.replace(normalize_ja(source))
        return LineResult(s, "simple", True, ["要確認（簡易訳）：翻訳モデルがないため用語集だけで置き換えました"])

    # ---- 公開 API ----
    def translate_phrase(self, text: str) -> str:
        """短い語句（見出し名など）を訳す。"""
        return self.translate_text(text).text

    def translate_text(self, text: str) -> TextResult:
        """複数行の文章を行ごとに訳す。箇条書き記号は日本語側と同じものを使う。"""
        lines = (text or "").split("\n")
        results: list[LineResult | None] = [None] * len(lines)
        prefixes = [""] * len(lines)
        pending: list[int] = []
        for i, line in enumerate(lines):
            if not line.strip():
                results[i] = LineResult("", "copy")
                continue
            heading = self._heading(line) if tu.heading_name(line) else None
            if heading:
                results[i] = LineResult(heading, "template")
                continue
            template, body = tu.detect_bullet(line)
            prefixes[i] = tu.make_bullet(template, 1) if template else ""
            if template and "{n}" in template or (template or "").startswith("circled"):
                prefixes[i] = line[: len(line) - len(body)]
            body = body.strip()
            hit = self._template(body)
            if hit is not None:
                results[i] = LineResult(hit, "template")
                continue
            hit = self._memory(body)
            if hit is not None:
                results[i] = LineResult(hit, "tm")
                continue
            if not _JA_CHAR_RE.search(body):
                results[i] = LineResult(body, "copy")
                continue
            pending.append(i)

        if pending:
            sources = [tu.detect_bullet(lines[i])[1].strip() for i in pending]
            if self.available:
                outs = self._mt(sources)
            else:
                outs = [self._simple(s) for s in sources]
            for i, r in zip(pending, outs):
                results[i] = r

        out_lines, notes, review = [], [], False
        methods = Counter()
        for prefix, r in zip(prefixes, results):
            out_lines.append(prefix + r.text if r.text else r.text)
            methods[r.method] += 1
            review |= r.review
            for n in r.notes:
                if n not in notes:
                    notes.append(n)
        for m in ("simple", "mt", "tm", "template", "copy"):
            if methods.get(m):
                method = m
                break
        else:
            method = "empty"
        return TextResult("\n".join(out_lines), method, review, notes)

    def self_test(self) -> str:
        """設定画面の［動作確認］用。"""
        import time
        t0 = time.perf_counter()
        ok = self.load()
        r = self.translate_text("レセプトの点検で3件のエラーが見つかりました。")
        dt = time.perf_counter() - t0
        if not ok:
            return f"翻訳モデルを読み込めませんでした（簡易訳で動作します）。\n{self.error}\n簡易訳：{r.text}"
        return f"翻訳モデルは正常です（{dt:.1f}秒）。\n訳：{r.text}"
