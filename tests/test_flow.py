"""一連の動作テスト（受け入れ基準）。

    python -m unittest tests.test_flow -v

1. 記入欄の検出（上長コメントは記入しない、タイトル下の空行は欄にならない）
2. 行がずれた過去報告書もラベル基準で読める
3. 類似検索で内容の近い事例が 1 位になる
4. 生成結果が過去事例の見出し・箇条書き・文末・結びの定型文にそろう
5. 出力 Excel の結合セル数・画像・書式が原本と同じ（calcChain・行／セル挿入も確認）
6. Excel で開いても修復メッセージが出ない（Excel がある環境のみ。COM 経由）
7. オフライン（通信系モジュールを使っていない／exe を data ごとコピーして動く）
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app import analyzer, cases, generator, storage  # noqa: E402
from app import xlsx_writer as xw  # noqa: E402
from tests import make_samples  # noqa: E402

TODAY = date(2026, 9, 26)
TOPIC = "メールサーバー障害の件"
MEMO = """9/25 夕方にメールサーバーで障害が発生した
原因はディスク容量の不足
不要なログを削除して復旧済み
課題：容量の監視設定が不十分
【今後の対応】
容量監視のアラートを追加する
10/3までに運用手順書を更新する
監視設定の見直し｜山田｜10/3｜対応中
手順書の更新｜佐藤｜10/3｜未着手"""

SHEET = make_samples.SHEET
M = "{%s}" % xw.NS_MAIN
ENDING_RE = re.compile(r"(です|ます|ました|でした|ません)。$")


def c14n(el) -> bytes:
    return etree.tostring(el, method="c14n")


class FlowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="report_test_"))
        cls.samples = make_samples.make_all(cls.tmp / "samples")
        storage.set_data_dir(cls.tmp / "data")
        cls.fmt = analyzer.analyze(cls.samples["format"], "週次報告")
        storage.save_format(cls.fmt, template_src=cls.samples["format"])
        cls.cases: list[dict] = []
        cls.import_log = cases.import_paths(cls.fmt, cls.cases, [cls.samples["reports_dir"]])
        storage.save_cases(cls.fmt["name"], cls.cases)
        cls.template = storage.template_path(cls.fmt)
        cls._results = None

    @classmethod
    def tearDownClass(cls):
        if not os.environ.get("KEEP_TEST_FILES"):
            shutil.rmtree(cls.tmp, ignore_errors=True)
        else:
            print(f"\nテストファイル: {cls.tmp}")

    # ---- helpers ----
    def field(self, name) -> dict:
        return next(f for f in self.fmt["fields"] if f["name"] == name)

    def case_by_subject(self, subject) -> dict:
        return next(c for c in self.cases if cases.subject_of(self.fmt, c) == subject)

    def refs(self, topic=TOPIC, memo=MEMO):
        hits = cases.Index(self.fmt, self.cases).search(topic + "\n" + memo)
        return [c for s, c in hits[:3] if s > 0]

    def results(self):
        if FlowTest._results is None:
            FlowTest._results = generator.generate(self.fmt, TOPIC, MEMO, self.refs(), self.cases,
                                                   dict(storage.DEFAULT_SETTINGS), TODAY)
        return FlowTest._results

    def export(self, name="out.xlsx", values=None) -> Path:
        out = self.tmp / name
        if values is None:
            texts = {k: v["text"] for k, v in self.results().items()}
            values = generator.to_cell_values(self.fmt, texts)
        xw.write_cells(self.template, out, values)
        return out

    # ---- 1 ----
    def test_1_fields_detected(self):
        expect = {
            "報告日": ("date", "right", ["B3"]),
            "報告者": ("reporter", "right", ["D3"]),
            "部署": ("dept", "right", ["F3"]),
            "件名": ("subject", "right", ["B4"]),
            "報告内容": ("body", "down", ["A7"]),
            "課題": ("body", "right", ["B15"]),
            "今後の対応": ("body", "right", ["B18"]),
            "No": ("list", "table", [f"A{r}" for r in range(23, 28)]),
            "アクション": ("list", "table", [f"B{r}" for r in range(23, 28)]),
            "担当": ("list", "table", [f"D{r}" for r in range(23, 28)]),
            "期限": ("list", "table", [f"E{r}" for r in range(23, 28)]),
            "状態": ("list", "table", [f"F{r}" for r in range(23, 28)]),
            "上長コメント": ("skip", "right", ["B29"]),
        }
        got = {f["name"]: (f["kind"], f["direction"], f["cells"]) for f in self.fmt["fields"]}
        self.assertEqual(got, expect)
        self.assertFalse(self.field("上長コメント")["enabled"], "上長コメントは既定で使用オフ")
        self.assertTrue(all(f["enabled"] for f in self.fmt["fields"] if f["name"] != "上長コメント"))
        tables = {f["table_id"] for f in self.fmt["fields"] if f["direction"] == "table"}
        self.assertEqual(len(tables), 1)
        # タイトル（A1）の下の空行（2 行目）は欄にならない
        self.assertFalse(any(f["label_cell"] == "A1" for f in self.fmt["fields"]))
        self.assertFalse(any(xw.split_ref(r)[1] == 2 for f in self.fmt["fields"] for r in f["cells"]))

    # ---- 2 ----
    def test_2_cases_imported_and_shifted_readable(self):
        self.assertEqual(len(self.cases), 5, self.import_log)
        self.assertTrue(any(".xls" in l for l in self.import_log), "xls は案内付きでスキップ")
        self.assertFalse(any(c["source"].split(os.sep)[-1].startswith("~$") for c in self.cases))
        for sample in make_samples.CASES:
            case = self.case_by_subject(sample["subject"])
            v = case["values"]
            self.assertEqual(case["warnings"], [], sample["file"])
            self.assertEqual(v[self.field("報告日")["id"]], sample["date"].strftime("%Y/%m/%d"))
            self.assertEqual(v[self.field("報告者")["id"]], sample["reporter"])
            self.assertEqual(v[self.field("報告内容")["id"]], "\n".join(sample["body"]))
            self.assertEqual(v[self.field("課題")["id"]], "\n".join(sample["issue"]))
            self.assertEqual(v[self.field("今後の対応")["id"]], "\n".join(sample["plan"]))
            actions = [a for a in v[self.field("アクション")["id"]] if a]
            self.assertEqual(actions, [row[1] for row in sample["table"]])
        shifted = self.case_by_subject("基幹システム移行の進捗報告")
        self.assertEqual(shifted["values"][self.field("報告者")["id"]], "佐藤 花子")

    # ---- 3 ----
    def test_3_similar_case_ranked_first(self):
        hits = cases.Index(self.fmt, self.cases).search(TOPIC + "\n" + MEMO)
        self.assertEqual(cases.subject_of(self.fmt, hits[0][1]), "ファイルサーバー障害の発生と復旧について")
        self.assertGreater(hits[0][0], hits[1][0])
        hits2 = cases.Index(self.fmt, self.cases).search("情報セキュリティ研修の追加実施\n新入社員向けに研修を行った")
        self.assertEqual(cases.subject_of(self.fmt, hits2[0][1]), "新人向けセキュリティ研修の実施報告")

    # ---- 4 ----
    def test_4_generated_text_follows_past_style(self):
        r = self.results()
        get = lambda name: r[self.field(name)["id"]]

        main = get("報告内容")
        self.assertEqual(main["status"], "memo")
        lines = main["text"].split("\n")
        self.assertEqual(lines[0], make_samples.OPENING, main["text"])
        self.assertEqual(lines[-1], make_samples.CLOSING, main["text"])
        headings = [l for l in lines if l.startswith("【")]
        self.assertEqual(headings, ["【背景】", "【状況】", "【対応】"], main["text"])
        body = [l for l in lines[1:-1] if not l.startswith("【")]
        self.assertTrue(body)
        for l in body:
            self.assertTrue(l.startswith("・"), l)
            if l != "・特になし":
                self.assertRegex(l, ENDING_RE)
        self.assertIn("・9/25 夕方にメールサーバーで障害が発生しました。", lines)
        self.assertIn("・原因はディスク容量の不足です。", lines)

        issue = get("課題")
        self.assertEqual(issue["text"], "・容量の監視設定が不十分です。")
        plan = get("今後の対応")["text"].split("\n")
        self.assertEqual(plan, ["・容量監視のアラートを追加します。", "・10/3までに運用手順書を更新します。"])

        self.assertEqual(get("件名")["text"], TOPIC)
        self.assertEqual(get("報告日")["text"], "2026/09/26")
        self.assertEqual(get("報告日")["status"], "auto")
        self.assertEqual(get("報告者")["text"], "山田 太郎", "未設定なら過去事例で最も多い値")
        self.assertEqual(get("部署")["text"], "情報システム部")

        self.assertEqual(get("No")["text"], "1\n2")
        self.assertEqual(get("アクション")["text"], "監視設定の見直し\n手順書の更新")
        self.assertEqual(get("担当")["text"], "山田\n佐藤")
        self.assertEqual(get("期限")["text"], "10/3\n10/3")
        self.assertEqual(get("状態")["text"], "対応中\n未着手")
        self.assertNotIn(self.field("上長コメント")["id"], r)
        for v in r.values():
            self.assertIn(v["status"], ("memo", "reuse", "auto", "empty"))
            self.assertIn("note", v)

    def test_4b_reuse_when_memo_missing(self):
        memo = "9/25 夕方にメールサーバーで障害が発生した"
        r = generator.generate(self.fmt, TOPIC, memo, self.refs(memo=memo), self.cases,
                               dict(storage.DEFAULT_SETTINGS), TODAY)
        issue = r[self.field("課題")["id"]]
        self.assertEqual(issue["status"], "reuse")
        self.assertIn("要確認", issue["note"])
        self.assertTrue(issue["text"])
        off = dict(storage.DEFAULT_SETTINGS, reuse=False)
        r2 = generator.generate(self.fmt, TOPIC, memo, self.refs(memo=memo), self.cases, off, TODAY)
        self.assertEqual(r2[self.field("課題")["id"]]["status"], "empty")
        self.assertEqual(tu_replace("報告日 2026/08/03 と 2026年8月3日"), "報告日 2026/09/26 と 2026年9月26日")

    # ---- 5 ----
    def test_5_output_keeps_format(self):
        out = self.export()
        with zipfile.ZipFile(self.template) as a, zipfile.ZipFile(out) as b:
            self.assertEqual(a.testzip(), None)
            self.assertEqual(b.testzip(), None)
            names_a, names_b = a.namelist(), b.namelist()
            self.assertEqual(names_a, names_b, "パーツ構成・順序が同じ")
            sheet_part = xw.sheet_parts(a)[SHEET]
            changed = {sheet_part, "xl/styles.xml"}
            for n in names_a:
                if n not in changed:
                    self.assertEqual(a.read(n), b.read(n), f"{n} が変わっている")
            media = [n for n in names_a if n.startswith("xl/media/")]
            self.assertTrue(media, "ロゴ画像がある")

            sa, sb = xw.parse_xml(a.read(sheet_part)), xw.parse_xml(b.read(sheet_part))
            self.assertEqual(len(sa.findall(f".//{M}mergeCell")), len(sb.findall(f".//{M}mergeCell")))
            # sheetData 以外（列幅・結合・印刷設定・図形の参照など）は完全に同じ
            for root in (sa, sb):
                root.remove(root.find(M + "sheetData"))
            self.assertEqual(c14n(sa), c14n(sb))

            # 既存のスタイルは変わらず、折り返しの複製が末尾に追加されているだけ
            st_a, st_b = xw.parse_xml(a.read("xl/styles.xml")), xw.parse_xml(b.read("xl/styles.xml"))
            xfs_a = st_a.find(M + "cellXfs").findall(M + "xf")
            xfs_b = st_b.find(M + "cellXfs").findall(M + "xf")
            self.assertGreater(len(xfs_b), len(xfs_a), "折り返しスタイルが追加される")
            for x, y in zip(xfs_a, xfs_b):
                self.assertEqual(c14n(x), c14n(y))
            for x in xfs_b[len(xfs_a):]:
                self.assertEqual(x.find(M + "alignment").get("wrapText"), "1")
            for tag in ("fonts", "fills", "borders", "numFmts", "cellStyleXfs", "cellStyles", "dxfs"):
                ea, eb = st_a.find(M + tag), st_b.find(M + tag)
                self.assertEqual(ea is None, eb is None)
                if ea is not None:
                    self.assertEqual(c14n(ea), c14n(eb), tag)

            # 書いたセルのスタイルは原本と同じ罫線・フォント（折り返しだけ追加）
            def cells(root):
                return {c.get("r"): c for c in root.iter(M + "c")}
            ca = cells(xw.parse_xml(a.read(sheet_part)))
            cb = cells(xw.parse_xml(b.read(sheet_part)))
            for ref, c in ca.items():
                self.assertIn(ref, cb)
                s_a, s_b = int(c.get("s", 0)), int(cb[ref].get("s", 0))
                for attr in ("borderId", "fontId", "fillId", "numFmtId"):
                    self.assertEqual(xfs_a[s_a].get(attr), xfs_b[s_b].get(attr), f"{ref} {attr}")

        wb = analyzer.load_book(out)
        ws = wb[SHEET]
        texts = {k: v["text"] for k, v in self.results().items()}
        self.assertEqual(ws["B4"].value, TOPIC)
        self.assertEqual(ws["A7"].value, texts[self.field("報告内容")["id"]])
        self.assertEqual(ws["B3"].value, "2026/09/26")
        self.assertEqual(ws["A23"].value, "1")
        self.assertEqual(ws["B23"].value, "監視設定の見直し")
        self.assertEqual(ws["B24"].value, "手順書の更新")
        self.assertIsNone(ws["B25"].value)
        self.assertIsNone(ws["B29"].value, "上長コメントは記入しない")
        self.assertIn("A7:F14", [str(r) for r in ws.merged_cells.ranges])

    def test_5b_formula_overwrite_removes_calcchain(self):
        out = self.tmp / "formula.xlsx"
        xw.write_cells(self.template, out, {(SHEET, "A33"): "上書き"})
        with zipfile.ZipFile(self.template) as a:
            self.assertIn("xl/calcChain.xml", a.namelist())
        with zipfile.ZipFile(out) as b:
            self.assertNotIn("xl/calcChain.xml", b.namelist())
            self.assertNotIn(b"calcChain", b.read("[Content_Types].xml"))
            self.assertNotIn(b"calcChain", b.read("xl/_rels/workbook.xml.rels"))
        self.assertEqual(analyzer.load_book(out)[SHEET]["A33"].value, "上書き")

    def test_5c_insert_rows_and_cells_in_order(self):
        out = self.tmp / "insert.xlsx"
        xw.write_cells(self.template, out, {(SHEET, "H40"): "行を追加", (SHEET, "H3"): "列を追加",
                                            (SHEET, "A2"): "空行に追加"})
        with zipfile.ZipFile(out) as b:
            root = xw.parse_xml(b.read(xw.sheet_parts(b)[SHEET]))
        rows = [int(r.get("r")) for r in root.iter(M + "row")]
        self.assertEqual(rows, sorted(rows))
        for row in root.iter(M + "row"):
            cols = [xw.split_ref(c.get("r"))[0] for c in row.iter(M + "c")]
            self.assertEqual(cols, sorted(cols))
            self.assertIsNone(row.get("spans"))
        ws = analyzer.load_book(out)[SHEET]
        self.assertEqual((ws["H40"].value, ws["H3"].value, ws["A2"].value), ("行を追加", "列を追加", "空行に追加"))

    def test_5d_xls_rejected(self):
        xls = next(self.samples["reports_dir"].rglob("*.xls"))
        with self.assertRaises(xw.XlsNotSupported):
            xw.write_cells(xls, self.tmp / "x.xlsx", {})
        with self.assertRaises(xw.XlsNotSupported):
            analyzer.analyze(xls)

    # ---- 6 ----
    def test_6_excel_opens_without_repair(self):
        if sys.platform != "win32":
            self.skipTest("Windows 以外")
        probe = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "if ([type]::GetTypeFromProgID('Excel.Application')) { 'yes' } else { 'no' }"],
            capture_output=True, text=True)
        if probe.stdout.strip() != "yes":
            self.skipTest("Excel がインストールされていません")
        files = [self.export("excel_check.xlsx"), self.tmp / "formula.xlsx"]
        if not files[1].exists():
            xw.write_cells(self.template, files[1], {(SHEET, "A33"): "上書き"})
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
    $out += [pscustomobject]@{ path = $p; caption = $wb.Windows.Item(1).Caption; name = $wb.Name;
                               b4 = $ws.Range('B4').Text; shapes = $ws.Shapes.Count; merged = $ws.Range('A7').MergeArea.Address(0,0) }
    $wb.Close($false)
  }
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
$out | ConvertTo-Json -Compress
"""
        ps = self.tmp / "excel_check.ps1"
        ps.write_text(script, encoding="utf-8-sig")
        proc = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps)]
                              + [str(f) for f in files], capture_output=True, timeout=180)
        stdout = proc.stdout.decode("utf-8", "replace")
        self.assertEqual(proc.returncode, 0, stdout + proc.stderr.decode("cp932", "replace"))
        data = json.loads(stdout.strip().splitlines()[-1])
        data = data if isinstance(data, list) else [data]
        for d in data:
            text = f"{d['caption']} {d['name']}"
            self.assertNotRegex(text, r"修復|Repaired", d)
            self.assertEqual(d["shapes"], 2, "ロゴ画像と図形が残っている")
            self.assertEqual(d["merged"], "A7:F14")
        self.assertEqual(data[0]["b4"], TOPIC)

    # ---- 7 ----
    def test_7_offline(self):
        banned = re.compile(r"^\s*(import|from)\s+(socket|urllib|http|requests|ftplib|smtplib|ssl|asyncio)\b", re.M)
        for py in list((ROOT / "app").glob("*.py")) + [ROOT / "main.py"]:
            self.assertIsNone(banned.search(py.read_text(encoding="utf-8")), f"{py.name} が通信系モジュールを使っている")

        exe = ROOT / "dist" / "ReportAssistant.exe"
        if not exe.exists():
            self.skipTest("exe 未ビルド（build.bat を実行すると確認できます）")
        portable = self.tmp / "別のPC"
        portable.mkdir()
        shutil.copy2(exe, portable / exe.name)
        shutil.copytree(storage.data_dir(), portable / "data")
        env = {k: v for k, v in os.environ.items() if k != "REPORT_APP_DATA"}
        out_dir = self.tmp / "selftest_out"
        proc = subprocess.run([str(portable / exe.name), "--selftest", str(out_dir)],
                              cwd=str(self.tmp), env=env, timeout=300)
        result = json.loads((out_dir / "selftest.json").read_text(encoding="utf-8"))
        self.assertTrue(result["ok"], result)
        self.assertEqual(proc.returncode, 0)
        self.assertEqual(Path(result["data_dir"]).resolve(), (portable / "data").resolve())
        self.assertTrue(Path(result["output"]).exists())


def tu_replace(text: str) -> str:
    from app import textutil
    return textutil.replace_dates(text, TODAY)


if __name__ == "__main__":
    unittest.main(verbosity=2)
