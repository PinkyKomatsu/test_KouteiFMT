"""エントリポイント。

    ReportAssistant.exe                GUI を起動
    ReportAssistant.exe --selftest D   GUI を出さずに一連の処理を行い、結果を D\\selftest.json に書く
                                       （exe を別の PC にコピーしたときの動作確認用）
"""
from __future__ import annotations

import json
import sys
import traceback
from datetime import date
from pathlib import Path


def enable_dpi_awareness() -> None:
    """Windows の高 DPI 表示でぼやけないようにする（Tk の作成前に呼ぶ）。"""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # System DPI aware
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def selftest(out_dir: str) -> int:
    from app import cases, generator, storage, xlsx_writer

    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    result = {"ok": False, "data_dir": None}
    try:
        result["data_dir"] = str(storage.data_dir())
        names = storage.list_formats()
        if not names:
            raise RuntimeError("フォーマットが登録されていません")
        fmt = storage.load_format(names[0])
        all_cases = storage.load_cases(fmt["name"])
        topic = "セルフテスト：サーバー障害の件"
        memo = "サーバーで障害が発生した\n課題：監視が不十分\n確認作業｜担当者｜10/1｜対応中"
        hits = cases.Index(fmt, all_cases).search(topic + "\n" + memo, top=3)
        refs = [c for s, c in hits if s > 0]
        results = generator.generate(fmt, topic, memo, refs, all_cases, storage.load_settings(), date.today())
        values = generator.to_cell_values(fmt, {k: v["text"] for k, v in results.items()})
        template = storage.template_path(fmt)
        dest = out / ("selftest" + template.suffix)
        xlsx_writer.write_cells(template, dest, values)
        result.update(ok=True, format=fmt["name"], cases=len(all_cases), hits=len(refs), output=str(dest))
    except Exception as e:
        result["error"] = f"{e}\n{traceback.format_exc()}"
    with open(out / "selftest.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return 0 if result["ok"] else 1


def main(argv: list[str]) -> int:
    if "--selftest" in argv:
        i = argv.index("--selftest")
        return selftest(argv[i + 1] if i + 1 < len(argv) else ".")
    enable_dpi_awareness()
    from app.gui import run
    run()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
