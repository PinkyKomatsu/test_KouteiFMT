"""テスト用の架空フォーマットと過去報告書 5 件を作る。

    python tests/make_samples.py [出力フォルダ]     （既定: tests/samples）

フォーマット：タイトル、報告日／報告者／部署、件名、【報告内容】の大枠、課題、今後の対応、
No／アクション／担当／期限／状態の表、上長コメント、ロゴ画像、図形、数式（calcChain 付き）、印刷設定。
過去報告書のうち 1 件（基幹システム移行）は、先頭に 1 行多い旧版（行が 1 行ずれている）。
"""
from __future__ import annotations

import struct
import sys
import zipfile
import zlib
from datetime import datetime
from pathlib import Path

import openpyxl
from lxml import etree
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

SHEET = "報告書"
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GRAY = PatternFill("solid", fgColor="EEEEEE")
TOP_LEFT = Alignment(horizontal="left", vertical="top")  # 折り返しなし（書き込み時に複製されるか確認する）

OPENING = "お疲れ様です。"
CLOSING = "以上、ご報告いたします。"

CASES = [
    {
        "file": "2026-08-03_ファイルサーバー障害.xlsx",
        "date": datetime(2026, 8, 3), "reporter": "山田 太郎", "dept": "情報システム部",
        "subject": "ファイルサーバー障害の発生と復旧について",
        "body": [OPENING,
                 "【背景】", "・8/1 午前にファイルサーバーでディスク障害が発生しました。", "・共有フォルダが約2時間利用できませんでした。",
                 "【状況】", "・原因はディスクの経年劣化と判明しました。", "・予備ディスクに交換し、サーバーは復旧しました。",
                 "【対応】", "・全サーバーのディスク状態を点検します。",
                 CLOSING],
        "issue": ["・障害の検知が遅れました。", "・サーバー監視の仕組みが不十分です。"],
        "plan": ["・ディスク監視を導入する予定です。", "・点検結果を来週報告します。"],
        "table": [("1", "ディスク点検", "山田", "8/10", "対応中"), ("2", "監視ツール選定", "佐藤", "8/20", "未着手")],
    },
    {
        "file": "2026-08-10_セキュリティ研修.xlsx",
        "date": datetime(2026, 8, 10), "reporter": "山田 太郎", "dept": "情報システム部",
        "subject": "新人向けセキュリティ研修の実施報告",
        "body": [OPENING,
                 "【背景】", "・新入社員5名を対象に情報セキュリティ研修を企画しました。",
                 "【状況】", "・8/5 に研修を実施しました。", "・理解度テストの平均点は85点でした。",
                 "【対応】", "・フォローアップ研修を計画します。",
                 CLOSING],
        "issue": ["・パスワード管理の理解が不十分です。"],
        "plan": ["・来月に追加研修を実施する予定です。"],
        "table": [("1", "追加研修の準備", "鈴木", "9/1", "未着手")],
    },
    {
        "file": "2026-08-17_A社問い合わせ.xlsx",
        "date": datetime(2026, 8, 17), "reporter": "山田 太郎", "dept": "情報システム部",
        "subject": "顧客A社からの問い合わせ対応状況",
        "body": [OPENING,
                 "【背景】", "・A社から請求書の記載内容について問い合わせがありました。",
                 "【状況】", "・経理部と確認し、記載誤りであることを確認しました。", "・訂正版の請求書を送付しました。",
                 "【対応】", "・請求書のチェック手順を見直します。",
                 CLOSING],
        "issue": ["・請求書の二重チェックができていません。"],
        "plan": ["・チェックリストを作成する予定です。"],
        "table": [("1", "チェックリスト作成", "田中", "8/31", "対応中")],
    },
    {
        "file": "2026-08-24_基幹システム移行_旧版.xlsx",
        "shift": 1,
        "date": datetime(2026, 8, 24), "reporter": "佐藤 花子", "dept": "情報システム部",
        "subject": "基幹システム移行の進捗報告",
        "body": [OPENING,
                 "【背景】", "・基幹システムを新基盤へ移行するプロジェクトを進めています。",
                 "【状況】", "・データ移行のリハーサルが完了しました。", "・性能試験で一部の画面に遅延が見つかりました。",
                 "【対応】", "・遅延の原因を調査します。",
                 CLOSING],
        "issue": ["・性能試験の結果が目標に届いていません。"],
        "plan": ["・9/15 に本番移行を予定しています。"],
        "table": [("1", "遅延の原因調査", "佐藤", "9/5", "対応中"), ("2", "本番移行判定", "山田", "9/10", "未着手")],
    },
    {
        "file": "2026-08-31_複合機コスト削減.xlsx",
        "date": datetime(2026, 8, 31), "reporter": "山田 太郎", "dept": "情報システム部",
        "subject": "複合機リース契約見直しによるコスト削減",
        "body": [OPENING,
                 "【背景】", "・複合機のリース契約が更新時期を迎えました。",
                 "【状況】", "・3社から見積もりを取得し、比較しました。", "・年間約40万円の削減が見込めます。",
                 "【対応】", "・契約先の変更手続きを進めます。",
                 CLOSING],
        "issue": ["・切り替え時の業務影響を確認できていません。"],
        "plan": ["・10月から新契約に切り替える予定です。"],
        "table": [("1", "契約手続き", "高橋", "9/20", "対応中")],
    },
]


# ---------------------------------------------------------------------------
# シートの組み立て
# ---------------------------------------------------------------------------

def _label(ws, rng: str, text: str):
    first = rng.split(":")[0]
    if ":" in rng:
        ws.merge_cells(rng)
    _box(ws, rng)
    c = ws[first]
    c.value = text
    c.fill = GRAY
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")


def _box(ws, rng: str):
    if ":" in rng and rng not in [str(r) for r in ws.merged_cells.ranges]:
        ws.merge_cells(rng)
    cells = ws[rng]
    if not isinstance(cells, tuple):
        cells = ((cells,),)
    for row in cells:
        for c in (row if isinstance(row, tuple) else (row,)):
            c.border = BOX


def build_sheet(ws, values: dict | None = None, off: int = 0):
    v = values or {}

    def R(r: int) -> int:
        return r + off

    for col in "ABCDEF":
        ws.column_dimensions[col].width = 14
    if off:
        ws["A1"] = "社外秘"
    # タイトル（下の行は余白：罫線なし）
    ws.merge_cells(f"A{R(1)}:F{R(1)}")
    ws[f"A{R(1)}"] = "業務報告書"
    ws[f"A{R(1)}"].font = Font(bold=True, size=16)
    ws[f"A{R(1)}"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[R(1)].height = 30

    # 報告日／報告者／部署
    _label(ws, f"A{R(3)}", "報告日")
    _box(ws, f"B{R(3)}")
    _label(ws, f"C{R(3)}", "報告者")
    _box(ws, f"D{R(3)}")
    _label(ws, f"E{R(3)}", "部署")
    _box(ws, f"F{R(3)}")
    # 件名
    _label(ws, f"A{R(4)}", "件名")
    _box(ws, f"B{R(4)}:F{R(4)}")
    # 報告内容
    _label(ws, f"A{R(6)}:F{R(6)}", "【報告内容】")
    _box(ws, f"A{R(7)}:F{R(14)}")
    ws[f"A{R(7)}"].alignment = TOP_LEFT
    # 課題・今後の対応
    _label(ws, f"A{R(15)}:A{R(17)}", "課題")
    _box(ws, f"B{R(15)}:F{R(17)}")
    ws[f"B{R(15)}"].alignment = TOP_LEFT
    _label(ws, f"A{R(18)}:A{R(20)}", "今後の対応")
    _box(ws, f"B{R(18)}:F{R(20)}")
    ws[f"B{R(18)}"].alignment = TOP_LEFT
    # 表
    _label(ws, f"A{R(22)}", "No")
    _label(ws, f"B{R(22)}:C{R(22)}", "アクション")
    _label(ws, f"D{R(22)}", "担当")
    _label(ws, f"E{R(22)}", "期限")
    _label(ws, f"F{R(22)}", "状態")
    for r in range(R(23), R(28)):
        _box(ws, f"A{r}")
        _box(ws, f"B{r}:C{r}")
        _box(ws, f"D{r}")
        _box(ws, f"E{r}")
        _box(ws, f"F{r}")
    # 上長コメント
    _label(ws, f"A{R(29)}:A{R(31)}", "上長コメント")
    _box(ws, f"B{R(29)}:F{R(31)}")
    # 数式（表の件数。余白に置く）
    ws[f"A{R(33)}"] = f"=COUNTA(B{R(23)}:B{R(27)})"

    # 印刷設定
    ws.print_area = f"A1:F{R(31)}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    if v:
        ws[f"B{R(3)}"] = v["date"]
        ws[f"B{R(3)}"].number_format = "yyyy/mm/dd"
        ws[f"D{R(3)}"] = v["reporter"]
        ws[f"F{R(3)}"] = v["dept"]
        ws[f"B{R(4)}"] = v["subject"]
        ws[f"A{R(7)}"] = "\n".join(v["body"])
        ws[f"B{R(15)}"] = "\n".join(v["issue"])
        ws[f"B{R(18)}"] = "\n".join(v["plan"])
        for i, row in enumerate(v["table"]):
            r = R(23) + i
            for col, val in zip(("A", "B", "D", "E", "F"), row):
                ws[f"{col}{r}"] = val


# ---------------------------------------------------------------------------
# ロゴ画像・図形・calcChain を ZIP に直接追加する（openpyxl の画像追加は Pillow が必要なため）
# ---------------------------------------------------------------------------

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

DRAWING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:oneCellAnchor>
<xdr:from><xdr:col>5</xdr:col><xdr:colOff>76200</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from>
<xdr:ext cx="838200" cy="304800"/>
<xdr:pic>
<xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="838200" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:oneCellAnchor>
<xdr:twoCellAnchor>
<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>31</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>33</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Stamp"/><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></xdr:spPr>
</xdr:sp>
<xdr:clientData/>
</xdr:twoCellAnchor>
</xdr:wsDr>
"""

DRAWING_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>
"""

CALC_CHAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A33" i="1"/></calcChain>
"""

# sheet XML で drawing 要素より前に来る要素
_BEFORE_DRAWING = ("printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
                   "customProperties", "cellWatches", "ignoredErrors", "smartTags")


def png_bytes(w: int, h: int, rgb=(37, 99, 235)) -> bytes:
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _xml(data: bytes):
    return etree.fromstring(data)


def _dump(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def add_drawing_and_calcchain(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        order = [i.filename for i in z.infolist()]
        parts = {n: z.read(n) for n in order}

    sheet = "xl/worksheets/sheet1.xml"
    parts["xl/media/image1.png"] = png_bytes(110, 40)
    parts["xl/drawings/drawing1.xml"] = DRAWING_XML.encode("utf-8")
    parts["xl/drawings/_rels/drawing1.xml.rels"] = DRAWING_RELS.encode("utf-8")

    rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
    rels = _xml(parts[rels_name]) if rels_name in parts else etree.Element("{%s}Relationships" % NS_PKG, nsmap={None: NS_PKG})
    etree.SubElement(rels, "{%s}Relationship" % NS_PKG, Id="rIdLogo1",
                     Type=NS_R + "/drawing", Target="../drawings/drawing1.xml")
    parts[rels_name] = _dump(rels)

    root = _xml(parts[sheet])
    drawing = etree.Element("{%s}drawing" % NS_MAIN, nsmap={"r": NS_R})
    drawing.set("{%s}id" % NS_R, "rIdLogo1")
    last = None
    for child in root:
        if etree.QName(child).localname in _BEFORE_DRAWING:
            last = child
    if last is not None:
        last.addnext(drawing)
    else:
        root.append(drawing)
    parts[sheet] = _dump(root)

    parts["xl/calcChain.xml"] = CALC_CHAIN.encode("utf-8")
    wb_rels = _xml(parts["xl/_rels/workbook.xml.rels"])
    etree.SubElement(wb_rels, "{%s}Relationship" % NS_PKG, Id="rIdCalc1", Type=NS_R + "/calcChain", Target="calcChain.xml")
    parts["xl/_rels/workbook.xml.rels"] = _dump(wb_rels)

    ct = _xml(parts["[Content_Types].xml"])
    if not any(d.get("Extension") == "png" for d in ct.iter("{%s}Default" % NS_CT)):
        first_override = ct.find("{%s}Override" % NS_CT)
        png = etree.Element("{%s}Default" % NS_CT, Extension="png", ContentType="image/png")
        if first_override is not None:
            first_override.addprevious(png)
        else:
            ct.append(png)
    etree.SubElement(ct, "{%s}Override" % NS_CT, PartName="/xl/drawings/drawing1.xml",
                     ContentType="application/vnd.openxmlformats-officedocument.drawing+xml")
    etree.SubElement(ct, "{%s}Override" % NS_CT, PartName="/xl/calcChain.xml",
                     ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
    parts["[Content_Types].xml"] = _dump(ct)

    for name in parts:
        if name not in order:
            order.append(name)
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            z.writestr(name, parts[name])
    tmp.replace(path)


# ---------------------------------------------------------------------------

def _new_book(values=None, off=0):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    build_sheet(ws, values, off)
    return wb


def make_all(out_dir) -> dict:
    out = Path(out_dir)
    reports = out / "reports"
    reports.mkdir(parents=True, exist_ok=True)

    fmt_path = out / "format.xlsx"
    _new_book().save(fmt_path)
    add_drawing_and_calcchain(fmt_path)

    paths = []
    for case in CASES:
        p = reports / case["file"]
        _new_book(case, case.get("shift", 0)).save(p)
        paths.append(p)

    # 取り込みで無視されるべきファイル
    sub = reports / "old"
    sub.mkdir(exist_ok=True)
    (reports / "~$2026-08-03_ファイルサーバー障害.xlsx").write_bytes(b"lock file")
    (sub / "旧形式の報告書.xls").write_bytes(b"\xd0\xcf\x11\xe0dummy")
    return {"format": fmt_path, "reports_dir": reports, "reports": paths, "cases": CASES}


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "samples"
    info = make_all(target)
    print(f"フォーマット: {info['format']}")
    for p in info["reports"]:
        print(f"過去報告書: {p}")
