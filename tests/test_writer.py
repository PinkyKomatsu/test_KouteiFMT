"""Excel 書き込みエンジン（xlsx_writer）のテスト。受け入れテスト 6（Excel 出力）を含む。"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile

import pytest
from lxml import etree

from app import analyzer
from app import xlsx_writer as xw
from tests.conftest import SHEET

M = "{%s}" % xw.NS_MAIN

VALUES = {
    (SHEET, "B3"): "2026/09/27",
    (SHEET, "D3"): "移行リハーサル",
    (SHEET, "B5"): "件名のテスト",
    (SHEET, "A8"): "1行目\n2行目\n3行目",
    (SHEET, "B24"): "アクション1",
    (SHEET, "B30"): "Summary line 1\nSummary line 2",
}


def c14n(el) -> bytes:
    return etree.tostring(el, method="c14n")


def shared_strings(zf) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = xw.parse_xml(zf.read("xl/sharedStrings.xml"))
    return ["".join(si.itertext()) for si in root.iter(M + "si")]


def cell_map(zf, part) -> dict[str, tuple[str, str]]:
    """セル番地 -> (値, スタイル番号)。共有文字列・インライン文字列・数値をすべて文字列にそろえる。"""
    sst = shared_strings(zf)
    root = xw.parse_xml(zf.read(part))
    out = {}
    for c in root.iter(M + "c"):
        t = c.get("t")
        if t == "s":
            v = sst[int(c.findtext(M + "v"))]
        elif t == "inlineStr":
            v = "".join(c.find(M + "is").itertext())
        else:
            f = c.findtext(M + "f")
            v = ("=" + f) if f is not None else (c.findtext(M + "v") or "")
        out[c.get("r")] = (v, c.get("s", "0"))
    return out


@pytest.fixture(scope="module")
def output(template, tmp_path_factory):
    out = tmp_path_factory.mktemp("writer") / "out.xlsx"
    xw.write_cells(template, out, VALUES)
    return out


def test_parts_and_images_unchanged(template, output):
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(output) as b:
        assert a.testzip() is None and b.testzip() is None
        assert a.namelist() == b.namelist(), "パーツの構成と順序が同じ"
        sheet_part = xw.sheet_parts(a)[SHEET]
        changed = {sheet_part, "xl/styles.xml"}
        for n in a.namelist():
            if n not in changed:
                assert a.read(n) == b.read(n), f"{n} が変わっている"
        media = [n for n in a.namelist() if n.startswith("xl/media/")]
        assert media, "ロゴ画像がある"
        assert all(a.read(n) == b.read(n) for n in media)


def test_merges_print_settings_unchanged(template, output):
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(output) as b:
        part = xw.sheet_parts(a)[SHEET]
        sa, sb = xw.parse_xml(a.read(part)), xw.parse_xml(b.read(part))
    assert len(sa.findall(f".//{M}mergeCell")) == len(sb.findall(f".//{M}mergeCell")) > 0
    for root in (sa, sb):
        root.remove(root.find(M + "sheetData"))
    assert c14n(sa) == c14n(sb), "列幅・結合・印刷設定・図形の参照は完全に同じ"


def test_only_specified_cells_changed(template, output):
    """原本とセル単位で差分を取り、指定したセルだけが変わっていること。"""
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(output) as b:
        part = xw.sheet_parts(a)[SHEET]
        ca, cb = cell_map(a, part), cell_map(b, part)
        xfs_a = xw.parse_xml(a.read("xl/styles.xml")).find(M + "cellXfs").findall(M + "xf")
        xfs_b = xw.parse_xml(b.read("xl/styles.xml")).find(M + "cellXfs").findall(M + "xf")
    written = {ref for (_, ref) in VALUES}
    assert set(ca) <= set(cb)
    changed = {ref for ref in cb if ca.get(ref, ("", "0"))[0] != cb[ref][0]}
    assert changed == written
    for (_, ref), v in VALUES.items():
        assert cb[ref][0] == v
    # スタイル：書いていないセルは同じ番号。書いたセルは罫線・フォント・塗り・表示形式が同じ
    for ref, (_, s) in ca.items():
        if ref not in written:
            assert cb[ref][1] == s, ref
        else:
            xa, xb = xfs_a[int(s)], xfs_b[int(cb[ref][1])]
            for attr in ("borderId", "fontId", "fillId", "numFmtId"):
                assert xa.get(attr) == xb.get(attr), f"{ref} {attr}"


def test_wrap_style_cloned_not_modified(template, output):
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(output) as b:
        st_a, st_b = xw.parse_xml(a.read("xl/styles.xml")), xw.parse_xml(b.read("xl/styles.xml"))
    xfs_a = st_a.find(M + "cellXfs").findall(M + "xf")
    xfs_b = st_b.find(M + "cellXfs").findall(M + "xf")
    assert len(xfs_b) > len(xfs_a)
    for x, y in zip(xfs_a, xfs_b):
        assert c14n(x) == c14n(y), "既存のスタイルは変えない"
    for x in xfs_b[len(xfs_a):]:
        assert x.find(M + "alignment").get("wrapText") == "1"
    for tag in ("fonts", "fills", "borders", "cellStyleXfs", "cellStyles"):
        assert c14n(st_a.find(M + tag)) == c14n(st_b.find(M + tag)), tag


def test_values_readable(output):
    ws = analyzer.load_book(output)[SHEET]
    assert ws["A8"].value == "1行目\n2行目\n3行目"
    assert ws["B5"].value == "件名のテスト"
    assert "A8:F15" in [str(r) for r in ws.merged_cells.ranges]


def test_formula_overwrite_removes_calcchain(template, tmp_path):
    out = tmp_path / "formula.xlsx"
    xw.write_cells(template, out, {(SHEET, "A42"): "上書き"})
    with zipfile.ZipFile(template) as a:
        assert "xl/calcChain.xml" in a.namelist()
    with zipfile.ZipFile(out) as b:
        assert "xl/calcChain.xml" not in b.namelist()
        assert b"calcChain" not in b.read("[Content_Types].xml")
        assert b"calcChain" not in b.read("xl/_rels/workbook.xml.rels")
    assert analyzer.load_book(out)[SHEET]["A42"].value == "上書き"


def test_rows_and_cells_inserted_in_order(template, tmp_path):
    out = tmp_path / "insert.xlsx"
    xw.write_cells(template, out, {(SHEET, "H60"): "行を追加", (SHEET, "H3"): "列を追加", (SHEET, "A2"): "空行に追加"})
    with zipfile.ZipFile(out) as b:
        root = xw.parse_xml(b.read(xw.sheet_parts(b)[SHEET]))
    rows = [int(r.get("r")) for r in root.iter(M + "row")]
    assert rows == sorted(rows)
    for row in root.iter(M + "row"):
        cols = [xw.split_ref(c.get("r"))[0] for c in row.iter(M + "c")]
        assert cols == sorted(cols)
        assert row.get("spans") is None
    ws = analyzer.load_book(out)[SHEET]
    assert (ws["H60"].value, ws["H3"].value, ws["A2"].value) == ("行を追加", "列を追加", "空行に追加")


def test_xls_rejected(samples, tmp_path):
    xls = next(samples["reports_dir"].rglob("*.xls"))
    with pytest.raises(xw.XlsNotSupported, match=r"\.xlsx として保存し直してください"):
        xw.write_cells(xls, tmp_path / "x.xlsx", {})


def test_excel_opens_without_repair(template, output, tmp_path):
    """Excel で開いて修復メッセージが出ないこと（Excel がない環境ではスキップ）。"""
    if sys.platform != "win32":
        pytest.skip("Windows 以外")
    probe = subprocess.run(["powershell", "-NoProfile", "-Command",
                            "if ([type]::GetTypeFromProgID('Excel.Application')) { 'yes' } else { 'no' }"],
                           capture_output=True, text=True)
    if probe.stdout.strip() != "yes":
        pytest.skip("Excel がインストールされていません")
    formula = tmp_path / "formula_excel.xlsx"
    xw.write_cells(template, formula, {(SHEET, "A42"): "上書き"})
    script = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$out = @()
try {
  foreach ($p in $args) {
    $wb = $xl.Workbooks.Open($p, 0, $true)
    $ws = $wb.Worksheets.Item(1)
    $out += [pscustomobject]@{ caption = $wb.Windows.Item(1).Caption; name = $wb.Name;
                               b5 = $ws.Range('B5').Text; shapes = $ws.Shapes.Count; merged = $ws.Range('A8').MergeArea.Address(0,0) }
    $wb.Close($false)
  }
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
$out | ConvertTo-Json -Compress
"""
    ps = tmp_path / "excel_check.ps1"
    ps.write_text(script, encoding="utf-8-sig")
    proc = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps),
                           str(output), str(formula)], capture_output=True, timeout=180)
    stdout = proc.stdout.decode("utf-8", "replace")
    assert proc.returncode == 0, stdout + proc.stderr.decode("cp932", "replace")
    data = json.loads(stdout.strip().splitlines()[-1])
    data = data if isinstance(data, list) else [data]
    for d in data:
        assert not any(w in f"{d['caption']} {d['name']}" for w in ("修復", "Repaired")), d
        assert d["shapes"] == 2, "ロゴ画像と図形が残っている"
        assert d["merged"] == "A8:F15"
    assert data[0]["b5"] == "件名のテスト"
