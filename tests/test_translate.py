"""英語の生成（オフライン翻訳）のテスト（受け入れテスト 5 の英語部分・8）。"""
from __future__ import annotations

import re
import time

import pytest

from app import generator_ja, glossary, report, storage, translator_en
from tests.conftest import MEMO, PHASE, STATUS, TODAY, TOPIC, model_available

needs_model = pytest.mark.skipif(not model_available(), reason="翻訳モデル（models/ja-en）がありません")
JA_RE = re.compile(r"[぀-ゟ゠-ヺー-ヿ]")   # ひらがな・カタカナ（訳し残し。記号「・」は除く）


def by_name(fmt, name):
    return next(f for f in fmt["fields"] if f["name"] == name)


@needs_model
def test_all_english_fields_filled(fmt, draft):
    for fid, entry in draft.items():
        if entry["ja"].strip():
            assert entry["en"].strip(), fid
            assert entry["en_status"] in ("template", "tm", "mt", "copy"), (fid, entry["en_status"])


@needs_model
def test_glossary_terms_used(fmt, draft):
    en = draft[by_name(fmt, "報告内容")["id"]]["en"]
    assert "medical claim" in en and "reception" not in en.lower()
    assert "migration rehearsal" in en.lower()
    assert "fee schedule master" in draft[by_name(fmt, "課題")["id"]]["en"]
    assert "production cutover" in draft[by_name(fmt, "今後の予定")["id"]]["en"]


@needs_model
def test_dates_and_numbers_match(fmt, draft):
    for fid, entry in draft.items():
        for ja_line, en_line in zip(entry["ja"].split("\n"), entry["en"].split("\n")):
            assert translator_en.check_consistency(ja_line, en_line) == [], (ja_line, en_line)
        assert not entry["review_en"], (fid, entry["note_en"])


@needs_model
def test_style_templates_and_bullets(fmt, draft):
    ja_lines = draft[by_name(fmt, "報告内容")["id"]]["ja"].split("\n")
    en_lines = draft[by_name(fmt, "報告内容")["id"]]["en"].split("\n")
    assert len(ja_lines) == len(en_lines)
    assert en_lines[0] == "Please find my report below."
    assert en_lines[-1] == "That concludes my report."
    assert [l for l in en_lines if l.startswith("[")] == ["[Background]", "[Status]", "[Actions]"]
    for j, e in zip(ja_lines, en_lines):
        if j.startswith("・"):
            assert e.startswith("・"), "箇条書き記号は日本語側とそろえる"
            assert not JA_RE.search(e), e
    assert not re.search(r"\b(I|I'm|I've|you)\b", " ".join(en_lines)), "上長向けの報告文（We を主語）"


@needs_model
def test_translation_memory_preferred(translator):
    r = translator.translate_text("・データ移行は予定の時間内に完了しました。")
    assert r.method == "tm"
    assert r.text == "・The data migration was completed within the planned time."


@needs_model
def test_glossary_protection_multiple_terms(translator):
    r = translator.translate_text("点数マスタと算定ロジックの確認が完了しました。")
    assert "fee schedule master" in r.text and "fee calculation" in r.text, r.text
    assert not r.review, r.notes


@needs_model
def test_model_in_japanese_folder(tmp_path):
    """日本語を含むフォルダ（例：ユーザー名が日本語）に置いたモデルでも翻訳できること。"""
    import shutil
    dest = tmp_path / "別のPC" / "モデル"
    shutil.copytree(storage.default_model_dir(), dest)
    t = translator_en.Translator(dest, glossary.load())
    assert t.available, t.error
    assert "medical claim" in t.translate_text("レセプトの点検が完了しました。").text


def test_templates_without_model():
    t = translator_en.Translator(None, glossary.DEFAULT_TERMS)
    assert t.translate_text("以上、ご報告いたします。").text == "That concludes my report."
    assert t.translate_text("【今後の予定】").text == "[Next Steps]"
    assert t.translate_text("・特になし").text == "・None"


def test_consistency_check_detects_mismatch():
    assert translator_en.check_consistency("3件のエラー", "Three errors")
    assert translator_en.check_consistency("9月26日に実施", "It was conducted on 9/26.") == []
    assert translator_en.check_consistency("DPCの設定", "The settings were checked.")


def test_without_model_falls_back_to_simple(fmt, imported, refs, tmp_path):
    """models/ がなくても生成が完了し、英語欄が「要確認（簡易訳）」になる（受け入れテスト 8）。"""
    t = translator_en.Translator(tmp_path / "no-model", glossary.load())
    assert not t.available and t.error
    d = report.build_draft(fmt, TOPIC, PHASE, STATUS, MEMO, refs, imported["cases"],
                           dict(storage.DEFAULT_SETTINGS), t, TODAY)
    main = d[by_name(fmt, "報告内容")["id"]]
    assert main["en_status"] == "simple" and main["review_en"]
    assert "要確認（簡易訳）" in main["note_en"]
    assert "medical claim" in main["en"], "用語集の語は英語になる"
    assert d[by_name(fmt, "工程")["id"]]["en"] == "Migration rehearsal", "定型部分はテンプレートで訳す"


@needs_model
def test_generation_under_10_seconds(fmt, imported, refs):
    t = translator_en.Translator(storage.default_model_dir(), glossary.load())
    start = time.perf_counter()
    report.build_draft(fmt, TOPIC, PHASE, STATUS, MEMO, refs, imported["cases"],
                       dict(storage.DEFAULT_SETTINGS), t, TODAY)
    assert time.perf_counter() - start < 10


def test_runtime_has_no_network_or_torch_imports():
    """実行時のコードが通信系・torch/transformers を import していないこと。"""
    from pathlib import Path
    banned = r"^\s*(?:import|from)\s+({})\b"
    names = ["socket", "urllib", "http", "requests", "ftplib", "smtplib", "ssl", "torch", "transformers"]
    root = Path(__file__).resolve().parent.parent
    for py in (root / "app").rglob("*.py"):
        assert not re.search(banned.format("|".join(names)), py.read_text(encoding="utf-8"), re.M), py.name
    # main.py は自己診断で通信を「禁止する」ためだけに socket を使う
    main_src = (root / "main.py").read_text(encoding="utf-8")
    assert not re.search(banned.format("|".join(n for n in names if n != "socket")), main_src, re.M)
    assert "socket.create_connection = deny" in main_src


def test_socket_is_blocked_during_tests():
    import socket
    with pytest.raises(RuntimeError):
        socket.create_connection(("example.com", 80), timeout=1)
