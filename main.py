"""エントリポイント。

    MedAcctReport.exe                               画面を起動
    MedAcctReport.exe --selftest 出力先 --format フォーマット.xlsx --reports 過去報告書のフォルダ
        画面を出さずに「フォーマット取り込み → 事例取り込み → 生成 → プレビュー → コピー → 出力」を行い、
        結果を 出力先\\selftest.json に書く（通信は socket ごと禁止した状態で行う）
    MedAcctReport.exe --gui-selftest 出力先 --format … --reports …
        実際の画面を自動で操作して同じ流れを確かめ、各画面の画像を 出力先 に保存する
"""
from __future__ import annotations

import argparse
import json
import logging
import logging.handlers
import os
import socket
import sys
import time
import traceback
from datetime import date
from pathlib import Path

log = logging.getLogger("medacct")


def setup_logging() -> Path:
    from app import storage
    path = storage.log_dir() / "app.log"
    handler = logging.handlers.RotatingFileHandler(path, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    return path


def block_network() -> None:
    """自己診断中は通信を禁止する（実行時の通信ゼロの確認用）。"""
    def deny(*_a, **_k):
        raise RuntimeError("通信は禁止されています（オフライン確認）")
    socket.socket.connect = deny
    socket.socket.connect_ex = deny
    socket.create_connection = deny
    socket.getaddrinfo = deny


SAMPLE_TOPIC = "移行リハーサル（第2回）の結果報告"
SAMPLE_PHASE = "移行リハーサル"
SAMPLE_STATUS = "遅延"
SAMPLE_MEMO = """9/26に移行リハーサル（第2回）を実施した
データ移行は計画の95%まで完了した
レセプトの点検で3件のエラーが発生した
課題：点数マスタの差分が未確認
【今後の予定】
10/3までにエラーの原因を調査する
10/10に本番移行の稼働判定会議を行う
エラー原因の調査｜山田｜10/3｜対応中"""


def selftest(out_dir: Path, format_path: str, reports: str) -> dict:
    """画面なしで一連の処理を行う。"""
    from app import analyzer, cases, glossary, report, storage, translator_en, xlsx_writer
    block_network()
    out_dir.mkdir(parents=True, exist_ok=True)
    result: dict = {"ok": False, "steps": {}}
    t0 = time.perf_counter()
    try:
        result["data_dir"] = str(storage.data_dir())
        result["model_dir"] = str(storage.model_dir(storage.load_settings()))
        fmt = analyzer.analyze(format_path, "自己診断")
        storage.save_format(fmt, template_src=format_path)
        result["steps"]["format"] = {"fields": len(fmt["fields"]), "review": [f["name"] for f in analyzer.review_fields(fmt)]}
        items: list[dict] = []
        log_lines = cases.import_paths(fmt, items, [reports], on_pii=lambda p, f: "mask")
        storage.save_cases(fmt["name"], items)
        result["steps"]["cases"] = {"count": len(items), "log": log_lines}
        t = translator_en.Translator(storage.model_dir(storage.load_settings()), glossary.load(),
                                     cases.translation_memory(items))
        hits = cases.Index(fmt, items).search(SAMPLE_TOPIC + "\n" + SAMPLE_MEMO, phase=SAMPLE_PHASE)
        refs = [c for s, c in hits[:3] if s > 0]
        g0 = time.perf_counter()
        draft = report.build_draft(fmt, SAMPLE_TOPIC, SAMPLE_PHASE, SAMPLE_STATUS, SAMPLE_MEMO, refs, items,
                                   storage.load_settings(), t, date.today())
        result["steps"]["generate"] = {"seconds": round(time.perf_counter() - g0, 2), "model": t.available,
                                       "model_error": t.error,
                                       "top_case": cases.subject_of(fmt, hits[0][1]) if hits else None}
        preview = report.preview_values(fmt, draft)
        result["steps"]["preview"] = {"cells": sum(len(v) for v in preview.values())}
        clip_ja, clip_en = report.clipboard_text(fmt, draft, "ja"), report.clipboard_text(fmt, draft, "en")
        (out_dir / "copy_ja.txt").write_text(clip_ja, encoding="utf-8")
        (out_dir / "copy_en.txt").write_text(clip_en, encoding="utf-8")
        result["steps"]["copy"] = {"ja_chars": len(clip_ja), "en_chars": len(clip_en)}
        template = storage.template_path(fmt)
        out = out_dir / ("selftest" + template.suffix)
        xlsx_writer.write_cells(template, out, report.cell_values(fmt, draft))
        out_en = out_dir / ("selftest_EN" + template.suffix)
        xlsx_writer.write_cells(template, out_en, report.cell_values(fmt, draft, english_only=True))
        result["steps"]["export"] = {"output": str(out), "output_en": str(out_en)}
        result["ok"] = True
    except Exception as e:
        result["error"] = f"{e}\n{traceback.format_exc()}"
    result["seconds"] = round(time.perf_counter() - t0, 2)
    (out_dir / "selftest.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def gui_selftest(app, out_dir: Path, format_path: str, reports: str) -> dict:
    """実際の画面を自動で操作する（ダイアログは使わない経路で呼ぶ）。"""
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import QApplication
    from app import report
    from app.ui.main_window import MainWindow
    block_network()
    out_dir.mkdir(parents=True, exist_ok=True)
    result: dict = {"ok": False, "steps": {}}
    win = MainWindow()
    win.resize(1600, 960)
    win.show()

    def pump(ms=300):
        end = time.perf_counter() + ms / 1000
        while time.perf_counter() < end:
            QApplication.processEvents()
            time.sleep(0.02)

    def shot(name):
        pump()
        win.grab().save(str(out_dir / name))

    try:
        # フォーマット取り込み → 保存
        win.tabs.setCurrentWidget(win.format_page)
        assert win.format_page.import_file(format_path, "自己診断"), "フォーマットを解析できません"
        shot("1_format.png")
        assert win.format_page.save(confirm=False), "保存できません"
        result["steps"]["format"] = len(win.state.fmt["fields"])
        # 事例取り込み
        win.tabs.setCurrentWidget(win.cases_page)
        win.cases_page.import_paths([reports], policy="mask")
        shot("2_cases.png")
        result["steps"]["cases"] = len(win.state.cases)
        # 生成
        page = win.compose_page
        win.tabs.setCurrentWidget(page)
        page.set_inputs(SAMPLE_TOPIC, SAMPLE_PHASE, SAMPLE_STATUS, SAMPLE_MEMO)
        page.search()
        g0 = time.perf_counter()
        page.generate(sync=True)
        result["steps"]["generate_seconds"] = round(time.perf_counter() - g0, 2)
        result["steps"]["model"] = win.state.translator().available
        shot("3_compose.png")
        # プレビュー（記入予定の値が表示されているか）
        view = page.preview.views[win.state.fmt["sheets"][0]]
        subject = next(f for f in win.state.fmt["fields"] if f["kind"] == "subject")
        shown = view.item_at_ref(subject["cells"][0]).text()
        result["steps"]["preview_subject"] = shown
        assert shown == SAMPLE_TOPIC, f"プレビューの件名が違います: {shown}"
        page.en_preview.setChecked(True)
        shot("4_preview_en.png")
        page.en_preview.setChecked(False)
        # コピー
        ja = page.copy_all("ja")
        assert QGuiApplication.clipboard().text() == ja
        en = page.copy_all("en")
        assert QGuiApplication.clipboard().text() == en
        (out_dir / "gui_copy_ja.txt").write_text(ja, encoding="utf-8")
        (out_dir / "gui_copy_en.txt").write_text(en, encoding="utf-8")
        result["steps"]["copy"] = {"ja": len(ja), "en": len(en)}
        # 出力
        out = page.export(path=str(out_dir / "gui_output.xlsx"), open_after=False, register=True)
        out_en = page.export(english_only=True, path=str(out_dir / "gui_output_EN.xlsx"), open_after=False)
        result["steps"]["export"] = [str(out), str(out_en)]
        result["steps"]["cases_after_register"] = len(win.state.cases)
        win.tabs.setCurrentWidget(win.settings_page)
        shot("5_settings.png")
        result["ok"] = True
    except Exception as e:
        result["error"] = f"{e}\n{traceback.format_exc()}"
    (out_dir / "gui_selftest.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    win.format_page.dirty = False
    win.close()
    return result


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--selftest", metavar="DIR")
    ap.add_argument("--gui-selftest", metavar="DIR")
    ap.add_argument("--format")
    ap.add_argument("--reports")
    args, _ = ap.parse_known_args(argv)

    diag_dir = args.selftest or args.gui_selftest
    if diag_dir:
        # 自己診断のデータは出力先の中に作る（配布フォルダの data/ を汚さない）
        from app import storage
        storage.set_data_dir(Path(diag_dir) / "data")
    log_path = setup_logging()
    log.info("起動 %s", " ".join(argv))

    if args.selftest:
        r = selftest(Path(args.selftest), args.format, args.reports)
        return 0 if r["ok"] else 1

    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QFontDatabase, QGuiApplication
    from PySide6.QtWidgets import QApplication, QMessageBox

    # 高 DPI：拡大率を丸めずに使う（Qt 6 は Windows の Per-Monitor DPI に対応済み）
    QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv[:1])
    for family in ("Yu Gothic UI", "Meiryo UI"):
        if family in QFontDatabase.families():
            app.setFont(QFont(family, 9))
            break

    def excepthook(exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        log.error("予期しないエラー\n%s", text)
        QMessageBox.critical(None, "エラー", f"予期しないエラーが発生しました。\n{exc}\n\n詳細は {log_path} に記録しました。")

    sys.excepthook = excepthook

    if args.gui_selftest:
        r = gui_selftest(app, Path(args.gui_selftest), args.format, args.reports)
        return 0 if r["ok"] else 1

    from app.ui.main_window import MainWindow
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
