"""過去事例の取り込み・個人情報・類似検索のテスト（受け入れテスト 3・4）。"""
from __future__ import annotations

import pytest

from app import cases, privacy
from tests import make_samples


def by_name(fmt, name):
    return next(f for f in fmt["fields"] if f["name"] == name)


def case_by_subject(fmt, items, subject):
    return next(c for c in items if cases.subject_of(fmt, c) == subject)


def test_import_skips_lock_and_xls(imported):
    assert len(imported["cases"]) == 5, imported["log"]
    assert any(".xls" in l and "スキップ" in l for l in imported["log"])
    assert not any("~$" in c["source"] for c in imported["cases"])


@pytest.mark.parametrize("sample", make_samples.CASES, ids=[c["file"] for c in make_samples.CASES])
def test_values_read_by_label_even_if_shifted(fmt, imported, sample):
    case = case_by_subject(fmt, imported["cases"], sample["subject"])
    v = case["values"]
    assert case["warnings"] == []
    assert v[by_name(fmt, "報告日")["id"]] == sample["date"].strftime("%Y/%m/%d")
    assert v[by_name(fmt, "工程")["id"]] == sample["phase"]
    assert v[by_name(fmt, "進捗")["id"]] == sample["status"]
    assert v[by_name(fmt, "報告者")["id"]] == sample["reporter"]
    assert v[by_name(fmt, "報告内容")["id"]] == "\n".join(sample["body"])
    assert v[by_name(fmt, "今後の予定")["id"]] == "\n".join(sample["plan"])
    assert v[by_name(fmt, "Summary (EN)")["id"]] == "\n".join(sample["summary"])
    assert [a for a in v[by_name(fmt, "アクション")["id"]] if a] == [row[1] for row in sample["table"]]


def test_shifted_case_is_the_old_version(fmt, imported):
    shifted = next(c for c in make_samples.CASES if c.get("shift"))
    case = case_by_subject(fmt, imported["cases"], shifted["subject"])
    assert case["values"][by_name(fmt, "工程")["id"]] == "移行リハーサル"


def test_dummy_patient_id_detected_and_masked(fmt, imported):
    assert len(imported["pii"]) == 1
    name, findings = imported["pii"][0]
    assert name.endswith("総合テスト.xlsx")
    assert [(f.kind, f.value, f.field) for f in findings] == [("患者ID", make_samples.DUMMY_PATIENT_ID, "課題")]
    case = case_by_subject(fmt, imported["cases"], "医事会計システム更新 総合テストの結果報告")
    issue = case["values"][by_name(fmt, "課題")["id"]]
    assert make_samples.DUMMY_PATIENT_ID not in issue and "患者ID ●●●●●●●" in issue
    assert case["masked"] is True


def test_import_can_be_cancelled_on_pii(fmt, samples):
    items: list[dict] = []
    log = cases.import_paths(fmt, items, [samples["reports_dir"]], on_pii=lambda path, findings: "skip")
    assert len(items) == 4
    assert any(l.startswith("中止:") and "総合テスト" in l for l in log)


@pytest.mark.parametrize("text, kinds", [
    ("患者ID 1234567 のデータ", ["患者ID"]),
    ("カルテ番号：00123456", ["患者ID"]),
    ("氏名：山田太郎", ["氏名"]),
    ("鈴木 一郎様より連絡", ["氏名"]),
    ("生年月日：昭和45年3月2日", ["生年月日"]),
    ("連絡先 03-1234-5678 まで", ["電話番号"]),
    ("携帯 09012345678", ["電話番号"]),
    ("2026/09/27 に 9/10 の件を担当：山田が対応。患者様への影響なし。ID 12", []),
    ("電子カルテ連携の仕様確認と検査仕様の同様の変更、各様式の多様な対応", []),
])
def test_privacy_patterns(text, kinds):
    assert [f.kind for f in privacy.find(text)] == kinds
    masked = privacy.mask(text)
    assert len(masked) == len(text)
    if kinds:
        assert privacy.MASK in masked


def test_translation_memory_from_english_column(imported):
    tm = cases.translation_memory(imported["cases"])
    assert ("以上、ご報告いたします。", "That concludes my report.") in tm
    assert ("【状況】", "[Status]") in tm
    assert any(ja.startswith("・データ移行は予定の時間内") for ja, _ in tm)


def test_similar_case_ranked_first(fmt, imported):
    index = cases.Index(fmt, imported["cases"])
    query = "移行リハーサル（第2回）の結果報告\nデータ移行を実施し、切り戻し手順を確認した"
    hits = index.search(query, phase="移行リハーサル")
    assert cases.subject_of(fmt, hits[0][1]) == "医事会計システム更新 移行リハーサル（第1回）の実施報告"
    assert hits[0][0] > hits[1][0]
    hits2 = index.search("総合テストの結果\nテストケースの合格率とレセプトの不具合", phase="総合テスト")
    assert cases.subject_of(fmt, hits2[0][1]) == "医事会計システム更新 総合テストの結果報告"


def test_phase_bonus(fmt, imported):
    index = cases.Index(fmt, imported["cases"])
    q = "医事会計システム更新 進捗報告"
    plain = {cases.phase_of(fmt, c): s for s, c in index.search(q)}
    boosted = {cases.phase_of(fmt, c): s for s, c in index.search(q, phase="開発")}
    assert boosted["開発"] == pytest.approx(plain["開発"] + cases.PHASE_BONUS, abs=1e-3)
    assert boosted["要件定義"] == plain["要件定義"]
