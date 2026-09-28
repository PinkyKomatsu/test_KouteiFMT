"""日本語の生成・Excel のセル・コピー用文字列のテスト（受け入れテスト 5 の日本語部分・6・7）。"""
from __future__ import annotations

import re
import zipfile

import pytest

from app import analyzer, generator_ja, report, storage
from app import textutil as tu
from app import xlsx_writer as xw
from tests import make_samples
from tests.conftest import MEMO, PHASE, SHEET, STATUS, TODAY, TOPIC

ENDING_RE = re.compile(r"(です|ます|ました|でした|ません)。$")


def by_name(fmt, name):
    return next(f for f in fmt["fields"] if f["name"] == name)


def ja(draft, fmt, name):
    return draft[by_name(fmt, name)["id"]]["ja"]


def test_main_body_follows_past_style(fmt, draft):
    entry = draft[by_name(fmt, "報告内容")["id"]]
    assert entry["ja_status"] == "memo"
    lines = entry["ja"].split("\n")
    assert lines[0] == make_samples.OPENING
    assert lines[-1] == make_samples.CLOSING
    assert [l for l in lines if l.startswith("【")] == ["【背景】", "【状況】", "【対応】"]
    body = [l for l in lines[1:-1] if not l.startswith("【")]
    for l in body:
        assert l.startswith("・"), l
        if l != "・特になし":
            assert ENDING_RE.search(l), l
    assert "・9/26に移行リハーサル（第2回）を実施しました。" in lines
    assert "・レセプトの点検で3件のエラーが発生しました。" in lines


def test_labeled_memo_lines_go_to_their_fields(fmt, draft):
    assert ja(draft, fmt, "課題") == "・点数マスタの差分が未確認です。"
    assert ja(draft, fmt, "今後の予定").split("\n") == [
        "・10/3までにエラーの原因を調査します。", "・10/10に本番移行の稼働判定会議を行います。"]


def test_fixed_fields(fmt, draft):
    assert ja(draft, fmt, "件名") == TOPIC
    assert ja(draft, fmt, "報告日") == "2026/09/27", "過去事例の書式 yyyy/MM/dd"
    assert ja(draft, fmt, "工程") == PHASE
    assert ja(draft, fmt, "進捗") == STATUS
    assert ja(draft, fmt, "報告者") == "山田 太郎", "設定が空なら過去事例で最も多い値"
    assert ja(draft, fmt, "No") == "1\n2"
    assert ja(draft, fmt, "アクション") == "エラー原因の調査\n点数マスタの差分確認"
    assert ja(draft, fmt, "担当") == "山田\n佐藤"
    assert ja(draft, fmt, "期限") == "10/3\n10/5"
    assert ja(draft, fmt, "状態") == "対応中\n未着手"
    assert by_name(fmt, "上長コメント")["id"] not in draft
    assert by_name(fmt, "Summary (EN)")["id"] not in draft, "英語版の欄は日本語欄の英訳を使う"


def test_settings_override_reporter_and_date_format(fmt, imported, refs):
    s = dict(storage.DEFAULT_SETTINGS, reporter="高橋 一郎", date_format="yyyy年M月d日")
    r = generator_ja.generate(fmt, TOPIC, PHASE, STATUS, MEMO, refs, imported["cases"], s, TODAY)
    assert r[by_name(fmt, "報告者")["id"]]["ja"] == "高橋 一郎"
    assert r[by_name(fmt, "報告日")["id"]]["ja"] == "2026年9月27日"


def test_reuse_when_memo_missing(fmt, imported, refs):
    memo = "9/26に移行リハーサル（第2回）を実施した"
    r = generator_ja.generate(fmt, TOPIC, PHASE, STATUS, memo, refs, imported["cases"],
                              dict(storage.DEFAULT_SETTINGS), TODAY)
    issue = r[by_name(fmt, "課題")["id"]]
    assert issue["status"] == "reuse" and issue["review"] and "要確認" in issue["note"]
    assert issue["ja"]
    off = generator_ja.generate(fmt, TOPIC, PHASE, STATUS, memo, refs, imported["cases"],
                                dict(storage.DEFAULT_SETTINGS, reuse=False), TODAY)
    assert off[by_name(fmt, "課題")["id"]]["status"] == "empty"
    assert tu.replace_dates("報告日 2026/08/03 と 2026年8月3日", TODAY) == "報告日 2026/09/27 と 2026年9月27日"


@pytest.mark.parametrize("src, expected", [
    ("会議を行う", "会議を行います"), ("手順書を更新する", "手順書を更新します"),
    ("原因を調べる", "原因を調べます"), ("資料を送る", "資料を送ります"),
    ("結果を報告書に書いた", "結果を報告書に書きました"), ("原因を調べた", "原因を調べました"),
    ("対応できない", "対応できません"), ("確認している", "確認しています"), ("問題ない", "問題ありません"),
])
def test_desu_conversion(src, expected):
    assert tu.convert_ending(src, "desu") == expected


def test_cell_values_ja_and_en_columns(fmt, draft):
    values = report.cell_values(fmt, draft)
    main = draft[by_name(fmt, "報告内容")["id"]]
    assert values[(SHEET, "A8")] == main["ja"]
    assert values[(SHEET, "B30")] == main["en"], "Summary (EN) には英語"
    assert values[(SHEET, "B5")] == TOPIC
    assert values[(SHEET, "B24")] == "エラー原因の調査" and values[(SHEET, "B26")] == ""
    assert (SHEET, "B34") not in values, "上長コメントは記入しない"
    en_only = report.cell_values(fmt, draft, english_only=True)
    assert en_only[(SHEET, "A8")] == main["en"]
    assert en_only[(SHEET, "B5")] == draft[by_name(fmt, "件名")["id"]]["en"]


def test_export_changes_only_field_cells(fmt, draft, template, tmp_path):
    out = tmp_path / "report.xlsx"
    values = report.cell_values(fmt, draft)
    xw.write_cells(template, out, values)
    ws = analyzer.load_book(out)[SHEET]
    for (_, ref), v in values.items():
        assert (ws[ref].value or "") == v, ref
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(out) as b:
        assert [n for n in a.namelist() if n.startswith("xl/media/")] == [n for n in b.namelist() if n.startswith("xl/media/")]
        part = xw.sheet_parts(a)[SHEET]
        assert a.read(part).count(b"<mergeCell ") == b.read(part).count(b"<mergeCell ")


def test_clipboard_texts(fmt, draft):
    """日本語用と英語用のコピー文字列（受け入れテスト 7）。"""
    text_ja = report.clipboard_text(fmt, draft, "ja")
    blocks = text_ja.split("\n\n")
    assert blocks[0] == "【報告日】2026/09/27"
    assert f"【件名】{TOPIC}" in blocks
    assert any(b.startswith("【報告内容】\nお疲れ様です。") for b in blocks)
    assert "【No／アクション／担当／期限／状態】\n1｜エラー原因の調査｜山田｜10/3｜対応中\n2｜点数マスタの差分確認｜佐藤｜10/5｜未着手" in blocks
    assert "上長コメント" not in text_ja and "Summary" not in text_ja

    text_en = report.clipboard_text(fmt, draft, "en")
    en_blocks = text_en.split("\n\n")
    assert en_blocks[0] == "[Date] Sep 27, 2026"
    assert "[Phase] Migration rehearsal" in en_blocks and "[Status] Delayed" in en_blocks
    assert any(b.startswith("[Summary (EN)]\nPlease find my report below.") for b in en_blocks)
    assert any(b.startswith("[No. / Action / Owner / Due / Status]\n1 | ") for b in en_blocks)
    for name in ("Subject", "Issues", "Next Steps"):
        assert any(b.startswith(f"[{name}]") for b in en_blocks), name

    main = by_name(fmt, "報告内容")
    assert report.field_clip(fmt, main, draft[main["id"]], "ja") == draft[main["id"]]["ja"]
    assert report.field_clip(fmt, main, draft[main["id"]], "en") == draft[main["id"]]["en"]


def test_field_at_maps_english_cell_to_its_japanese_field(fmt):
    assert report.field_at(fmt, SHEET, "B30")["name"] == "報告内容"
    assert report.field_at(fmt, SHEET, "A8")["name"] == "報告内容"
    assert report.field_at(fmt, SHEET, "B25")["name"] == "アクション"
    assert report.field_at(fmt, SHEET, "H1") is None


def test_filename():
    assert report.make_filename("{date}_{phase}_{topic}", TODAY, "移行リハーサル", 'a/b:c?"d') == \
        "20260927_移行リハーサル_a_b_c_d"
    assert report.make_filename("{date}_{phase}_{topic}", TODAY, "", "件名") == "20260927_件名"
