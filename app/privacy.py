"""個人情報の検出と伏せ字（過去事例の取り込み時に使う）。

対象：患者 ID、氏名らしき文字列、生年月日、電話番号。
報告書に残してよい社内の担当者名（表の「担当」など）は対象にしない。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

MASK = "●"

# 「○○様」のうち、人名ではない一般的な呼び方
_NOT_NAMES = {"患者", "お客", "客", "皆", "各", "利用者", "関係者", "担当者", "受診者", "ご家族", "家族", "先生"}

_PATTERNS = [
    ("患者ID", re.compile(
        r"(?P<label>(?:患者|カルテ|診察券|受診者)\s*(?:ID|ＩＤ|番号|No\.?|Ｎｏ\.?)\s*[:：]?\s*)(?P<value>\d{5,10})(?!\d)")),
    ("患者ID", re.compile(r"(?P<label>\bID\s*[:：]\s*)(?P<value>\d{6,10})(?!\d)")),
    ("氏名", re.compile(
        r"(?P<label>(?:患者名|患者氏名|氏名|お名前)\s*[:：]?\s*)(?P<value>[一-龥々]{1,4}[ 　]?[一-龥々ぁ-んァ-ヶ]{1,4})")),
    # 「鈴木 一郎様」のような敬称付きの人名（「仕様」「同様」などの語は除く）
    ("氏名", re.compile(r"(?<![一-龥々])(?P<value>[一-龥々]{2,4}(?:[ 　][一-龥々]{1,3})?)(?<![仕同模多異一各有態])(?P<label>様)")),
    ("生年月日", re.compile(
        r"(?P<label>(?:生年月日|誕生日|DOB)\s*[:：]?\s*)"
        r"(?P<value>(?:(?:明治|大正|昭和|平成|令和|[MTSHR])\s*\d{1,2}|\d{4})\s*[年/.\-]\s*\d{1,2}\s*[月/.\-]\s*\d{1,2}\s*日?)")),
    ("生年月日", re.compile(r"(?P<value>(?:明治|大正|昭和|平成)\s*\d{1,2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)(?P<label>\s*生まれ?)?")),
    ("電話番号", re.compile(
        r"(?<![\d/])(?P<value>0\d{1,4}[-‐－(（]\d{1,4}[-‐－)）]\d{3,4}|0[5-9]0\d{8})(?![\d/])")),
]


@dataclass
class Finding:
    kind: str       # 患者ID / 氏名 / 生年月日 / 電話番号
    value: str      # 見つかった文字列（伏せる部分）
    start: int
    end: int
    field: str = ""  # 欄名（scan_values のとき）

    def describe(self) -> str:
        where = f"「{self.field}」の" if self.field else ""
        return f"{where}{self.kind}：{self.value}"


def find(text: str) -> list[Finding]:
    """文字列中の個人情報らしき箇所を返す（重なりは先に見つかったものを優先）。"""
    found: list[Finding] = []
    for kind, rx in _PATTERNS:
        for m in rx.finditer(text or ""):
            value = m.group("value")
            if kind == "氏名" and m.group("label") == "様" and value.replace(" ", "").replace("　", "") in _NOT_NAMES:
                continue
            s, e = m.start("value"), m.end("value")
            if any(not (e <= f.start or s >= f.end) for f in found):
                continue
            found.append(Finding(kind, value, s, e))
    return sorted(found, key=lambda f: f.start)


def mask(text: str, findings: list[Finding] | None = None) -> str:
    """見つかった箇所を ● で伏せる（文字数は保つ）。"""
    findings = find(text) if findings is None else findings
    out = text
    for f in sorted(findings, key=lambda f: -f.start):
        out = out[:f.start] + MASK * (f.end - f.start) + out[f.end:]
    return out


def scan_values(fmt: dict, values: dict) -> list[Finding]:
    """取り出した欄の値（文字列または行のリスト）をすべて調べる。"""
    names = {f["id"]: f["name"] for f in fmt["fields"]}
    out = []
    for fid, v in values.items():
        for text in (v if isinstance(v, list) else [v]):
            for f in find(text or ""):
                f.field = names.get(fid, fid)
                out.append(f)
    return out


def mask_values(values: dict) -> dict:
    return {fid: ([mask(x or "") for x in v] if isinstance(v, list) else mask(v or "")) for fid, v in values.items()}
