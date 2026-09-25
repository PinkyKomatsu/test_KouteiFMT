"""書式を壊さない Excel 書き込み。

openpyxl で保存すると画像・図形などが失われるため、xlsx/xlsm を ZIP として開き、
対象シートの XML だけを lxml で編集する（ElementTree は名前空間の接頭辞を書き換えて
Excel の修復エラーになるため使わない）。それ以外のパーツは中身をそのまま複製する。

    write_cells(template, out, {("報告書", "B3"): "2026/09/26", ...})
"""
from __future__ import annotations

import copy
import os
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_XML = "http://www.w3.org/XML/1998/namespace"

M = "{%s}" % NS_MAIN
REL = "{%s}" % NS_PKG_REL
CT = "{%s}" % NS_CT

REL_OFFICE_DOC = NS_R + "/officeDocument"
REL_STYLES = NS_R + "/styles"
REL_CALCCHAIN = NS_R + "/calcChain"

SUPPORTED_EXT = (".xlsx", ".xlsm")


class XlsNotSupported(Exception):
    """旧形式 .xls が渡された。"""


class WriteError(Exception):
    """書き込みできないファイル・指定。"""


def check_supported(path) -> None:
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".xls":
        raise XlsNotSupported(
            f"「{p.name}」は旧形式（.xls）のため読み込めません。\n"
            "Excel で開き、「名前を付けて保存」で Excel ブック（*.xlsx）として保存し直してください。"
        )
    if ext not in SUPPORTED_EXT:
        raise WriteError(f"「{p.name}」は対応していない形式です（.xlsx / .xlsm のみ）。")
    if not zipfile.is_zipfile(p):
        raise WriteError(f"「{p.name}」を Excel ファイルとして開けません。")


# ---------------------------------------------------------------------------
# セル番地
# ---------------------------------------------------------------------------

_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def col_to_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref: str) -> tuple[int, int]:
    """'B3' -> (列番号 2, 行番号 3)"""
    m = _REF_RE.match(ref.strip())
    if not m:
        raise WriteError(f"セル番地が正しくありません: {ref}")
    return col_to_index(m.group(1)), int(m.group(2))


def make_ref(col: int, row: int) -> str:
    return f"{index_to_col(col)}{row}"


# XML 1.0 で使えない制御文字
_ILLEGAL_XML_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    return _ILLEGAL_XML_RE.sub("", text)


# ---------------------------------------------------------------------------
# XML
# ---------------------------------------------------------------------------

_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


def parse_xml(data: bytes):
    return etree.fromstring(data, _PARSER)


def serialize_xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_path(part: str) -> str:
    d, name = posixpath.split(part)
    return posixpath.join(d, "_rels", name + ".rels")


def _resolve(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def workbook_part(zf: zipfile.ZipFile) -> str:
    try:
        rels = parse_xml(zf.read("_rels/.rels"))
        for r in rels.iter(REL + "Relationship"):
            if r.get("Type") == REL_OFFICE_DOC:
                return _resolve("", r.get("Target"))
    except KeyError:
        pass
    return "xl/workbook.xml"


def sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    """シート名 -> シート XML のパーツ名"""
    wb_part = workbook_part(zf)
    wb = parse_xml(zf.read(wb_part))
    rels = parse_xml(zf.read(_rels_path(wb_part)))
    targets = {r.get("Id"): r.get("Target") for r in rels.iter(REL + "Relationship")}
    out = {}
    for s in wb.iter(M + "sheet"):
        target = targets.get(s.get("{%s}id" % NS_R))
        if target:
            out[s.get("name")] = _resolve(wb_part, target)
    return out


def _find_rel_target(zf, wb_part, rel_type) -> tuple[str | None, str | None]:
    """ワークブックの rels から指定タイプのパーツ名と rels のパーツ名を返す。"""
    rels_part = _rels_path(wb_part)
    rels = parse_xml(zf.read(rels_part))
    for r in rels.iter(REL + "Relationship"):
        if r.get("Type") == rel_type:
            return _resolve(wb_part, r.get("Target")), rels_part
    return None, rels_part


# ---------------------------------------------------------------------------
# スタイル（折り返しスタイルの複製）
# ---------------------------------------------------------------------------

class _Styles:
    def __init__(self, data: bytes | None):
        self.root = parse_xml(data) if data else None
        self.cache: dict[int, int] = {}
        self.changed = False

    def ensure_wrap(self, s: str | None) -> str | None:
        """スタイル s に折り返しがなければ、折り返し付きの複製を作ってその番号を返す。"""
        if self.root is None:
            return s
        cell_xfs = self.root.find(M + "cellXfs")
        if cell_xfs is None:
            return s
        xfs = cell_xfs.findall(M + "xf")
        idx = int(s) if s not in (None, "") else 0
        if idx >= len(xfs):
            return s
        align = xfs[idx].find(M + "alignment")
        if align is not None and align.get("wrapText") in ("1", "true"):
            return s
        if idx not in self.cache:
            new_xf = copy.deepcopy(xfs[idx])
            new_align = new_xf.find(M + "alignment")
            if new_align is None:
                new_align = etree.Element(M + "alignment")
                new_xf.insert(0, new_align)  # alignment は xf の先頭の子要素
            new_align.set("wrapText", "1")
            new_xf.set("applyAlignment", "1")
            cell_xfs.append(new_xf)
            count = len(cell_xfs.findall(M + "xf"))
            cell_xfs.set("count", str(count))
            self.cache[idx] = count - 1
            self.changed = True
        return str(self.cache[idx])


# ---------------------------------------------------------------------------
# シート XML の編集
# ---------------------------------------------------------------------------

def _row_num(row, fallback: int) -> int:
    r = row.get("r")
    return int(r) if r else fallback


def _get_row(sheet_data, rnum: int):
    prev = 0
    for row in sheet_data.findall(M + "row"):
        n = _row_num(row, prev + 1)
        if row.get("r") is None:
            row.set("r", str(n))
        if n == rnum:
            return row
        if n > rnum:
            new = etree.Element(M + "row", r=str(rnum))
            row.addprevious(new)
            return new
        prev = n
    return etree.SubElement(sheet_data, M + "row", r=str(rnum))


def _get_cell(row, col: int, rnum: int):
    prev = 0
    for c in row.findall(M + "c"):
        ref = c.get("r")
        n = split_ref(ref)[0] if ref else prev + 1
        if ref is None:
            c.set("r", make_ref(n, rnum))
        if n == col:
            return c
        if n > col:
            new = etree.Element(M + "c", r=make_ref(col, rnum))
            c.addprevious(new)
            return new
        prev = n
    # extLst があればその前に入れる
    ext = row.find(M + "extLst")
    new = etree.Element(M + "c", r=make_ref(col, rnum))
    if ext is not None:
        ext.addprevious(new)
    else:
        row.append(new)
    return new


def _unshare_formula(sheet_data, si: str) -> None:
    """共有数式の親セルを上書きする場合、同じ si を参照するセルの数式を外して値だけ残す。"""
    for f in sheet_data.iter(M + "f"):
        if f.get("t") == "shared" and f.get("si") == si:
            f.getparent().remove(f)


def _set_cell(sheet_data, ref: str, value, styles: _Styles) -> tuple[bool, str | None]:
    """1 セルに書き込む。戻り値は (数式を上書きしたか, 警告)。"""
    col, rnum = split_ref(ref)
    row = _get_row(sheet_data, rnum)
    row.attrib.pop("spans", None)
    c = _get_cell(row, col, rnum)

    warning = None
    f = c.find(M + "f")
    had_formula = f is not None
    if had_formula and f.get("t") == "shared" and f.get("ref"):
        _unshare_formula(sheet_data, f.get("si"))
        warning = f"{ref} の共有数式を上書きしたため、同じ数式を共有していたセルは値のみになりました。"

    for child in list(c):
        if child.tag != M + "extLst":
            c.remove(child)
    for attr in ("t", "cm", "vm"):
        c.attrib.pop(attr, None)

    text = clean_text(value)
    if text == "":
        return had_formula, warning

    if "\n" in text:
        new_s = styles.ensure_wrap(c.get("s"))
        if new_s is not None and new_s != c.get("s"):
            c.set("s", new_s)
    c.set("t", "inlineStr")
    is_el = etree.Element(M + "is")
    t_el = etree.SubElement(is_el, M + "t")
    t_el.text = text
    t_el.set("{%s}space" % NS_XML, "preserve")
    ext = c.find(M + "extLst")
    if ext is not None:
        ext.addprevious(is_el)
    else:
        c.append(is_el)
    return had_formula, warning


# ---------------------------------------------------------------------------
# 書き込み本体
# ---------------------------------------------------------------------------

def _drop_calc_chain(zf, wb_part, modified: dict, drop: set) -> None:
    calc_part, rels_part = _find_rel_target(zf, wb_part, REL_CALCCHAIN)
    if not calc_part and "xl/calcChain.xml" in zf.namelist():
        calc_part = "xl/calcChain.xml"
    if not calc_part:
        return
    drop.add(calc_part)

    rels = parse_xml(modified.get(rels_part) or zf.read(rels_part))
    for r in list(rels.iter(REL + "Relationship")):
        if r.get("Type") == REL_CALCCHAIN:
            r.getparent().remove(r)
    modified[rels_part] = serialize_xml(rels)

    ct = parse_xml(modified.get("[Content_Types].xml") or zf.read("[Content_Types].xml"))
    for o in list(ct.iter(CT + "Override")):
        if o.get("PartName", "").lstrip("/") == calc_part:
            o.getparent().remove(o)
    modified["[Content_Types].xml"] = serialize_xml(ct)


def write_cells(template, out, values: dict) -> list[str]:
    """template を複製して out に保存し、values のセルに書き込む。

    values: {(シート名, "B3"): 値}。値が "" / None ならセルの値を消す（書式は残す）。
    戻り値は警告メッセージのリスト。
    """
    template, out = Path(template), Path(out)
    check_supported(template)
    if out.exists() and out.resolve() == template.resolve():
        raise WriteError("出力先がフォーマットの原本と同じです。別の名前を指定してください。")

    by_sheet: dict[str, dict[str, object]] = {}
    for (sheet, ref), v in values.items():
        split_ref(ref)  # 番地チェック
        by_sheet.setdefault(sheet, {})[ref.replace("$", "").upper()] = v

    warnings: list[str] = []
    with zipfile.ZipFile(template) as zin:
        wb_part = workbook_part(zin)
        parts = sheet_parts(zin)
        styles_part, _ = _find_rel_target(zin, wb_part, REL_STYLES)
        styles = _Styles(zin.read(styles_part) if styles_part and styles_part in zin.namelist() else None)

        modified: dict[str, bytes] = {}
        drop: set[str] = set()
        formula_hit = False

        for sheet, cells in by_sheet.items():
            part = parts.get(sheet)
            if not part:
                raise WriteError(f"シート「{sheet}」がフォーマットに見つかりません。")
            root = parse_xml(zin.read(part))
            sheet_data = root.find(M + "sheetData")
            if sheet_data is None:
                raise WriteError(f"シート「{sheet}」の構造を読み取れません。")
            for ref in sorted(cells, key=lambda r: (split_ref(r)[1], split_ref(r)[0])):
                hit, warn = _set_cell(sheet_data, ref, cells[ref], styles)
                formula_hit |= hit
                if warn:
                    warnings.append(warn)
            modified[part] = serialize_xml(root)

        if styles.changed:
            modified[styles_part] = serialize_xml(styles.root)
        if formula_hit:
            _drop_calc_chain(zin, wb_part, modified, drop)

        out.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix="~report_", suffix=out.suffix, dir=out.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in zin.infolist():
                    if info.filename in drop:
                        continue
                    data = modified.get(info.filename)
                    if data is None:
                        data = zin.read(info.filename)
                    zout.writestr(info, data, compress_type=info.compress_type)
            os.replace(tmp, out)
        except BaseException:
            try:
                os.unlink(tmp)
            except OSError:
                pass
            raise
    return warnings
