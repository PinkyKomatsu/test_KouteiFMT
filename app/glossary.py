"""医事会計の用語集（日本語 → 英語）。data/glossary.csv に保存し、画面から編集できる。"""
from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path

from . import storage

DEFAULT_TERMS = [
    ("医事会計システム", "medical accounting system"),
    ("レセプト", "medical claim (receipt)"),
    ("診療報酬改定", "medical fee revision"),
    ("点数マスタ", "fee schedule master"),
    ("算定", "fee calculation"),
    ("窓口会計", "front-desk billing"),
    ("未収金", "outstanding patient payments"),
    ("電子カルテ連携", "EHR interface"),
    ("DPC", "DPC (Diagnosis Procedure Combination)"),
    ("移行リハーサル", "migration rehearsal"),
    ("本番移行", "production cutover"),
    ("稼働判定", "go-live decision"),
    ("切り戻し", "rollback"),
    ("並行稼働", "parallel run"),
    ("総合テスト", "system integration test"),
    ("保険者", "insurer"),
    ("公費", "public expense programs"),
]


def load(path: Path | None = None) -> list[tuple[str, str]]:
    """用語集を読む。ファイルがなければ同梱の初期用語集（なければ既定値）から作る。"""
    path = Path(path) if path else storage.glossary_path()
    if not path.exists():
        bundled = storage.bundled_glossary_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        if bundled.exists() and bundled.resolve() != path.resolve():
            shutil.copy2(bundled, path)
        else:
            save(DEFAULT_TERMS, path)
    terms = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            if len(row) < 2 or not row[0].strip() or row[0].strip().lower() == "ja":
                continue
            terms.append((row[0].strip(), row[1].strip()))
    return terms


def save(terms, path: Path | None = None) -> None:
    path = Path(path) if path else storage.glossary_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ja", "en"])
        for ja, en in terms:
            if ja.strip() and en.strip():
                w.writerow([ja.strip(), en.strip()])
    tmp.replace(path)


class Glossary:
    def __init__(self, terms):
        # 長い語から照合する（「医事会計システム」を「医事会計」より優先）
        self.terms = sorted({ja: en for ja, en in terms if ja and en}.items(), key=lambda t: -len(t[0]))

    def find(self, text: str) -> list[tuple[int, int, str, str]]:
        """文中の用語を (開始, 終了, 日本語, 英語) で返す（重なりなし）。"""
        found: list[tuple[int, int, str, str]] = []
        for ja, en in self.terms:
            start = 0
            while True:
                i = text.find(ja, start)
                if i < 0:
                    break
                j = i + len(ja)
                if not any(not (j <= s or i >= e) for s, e, _, _ in found):
                    found.append((i, j, ja, en))
                start = j
        return sorted(found)

    def replace(self, text: str) -> str:
        """用語を英語に置き換える（簡易訳用）。"""
        out, pos = [], 0
        for s, e, _, en in self.find(text):
            out.append(text[pos:s])
            out.append(f" {en} ")
            pos = e
        out.append(text[pos:])
        return re.sub(r" {2,}", " ", "".join(out)).strip()

    def lookup(self, ja: str) -> str | None:
        for j, en in self.terms:
            if j == ja:
                return en
        return None
