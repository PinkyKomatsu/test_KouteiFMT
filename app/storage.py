"""JSON 保存。

保存先は exe（開発時はプロジェクト）と同じ場所の data/。
書き込めない場合は %LOCALAPPDATA%\\ReportAssistant\\data を使う。

    data/
      settings.json
      formats/<名前>/format.json      フォーマット定義
      formats/<名前>/template.xlsx    取り込んだ原本のコピー（.xlsm ならそのまま .xlsm）
      formats/<名前>/cases.json       過去事例
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

APP_NAME = "ReportAssistant"

DEFAULT_SETTINGS = {
    "reporter": "",
    "department": "",
    "date_format": "auto",
    "reuse": True,               # メモのない欄に過去事例の文面を流用する
    "highlight": True,           # 流用・自動入力の欄を色分け表示する
    "register_output": True,     # 出力した報告書を過去事例に登録する（既定値）
    "open_after_export": True,   # 出力後に Excel で開く（既定値）
    "output_dir": "",
    "filename_pattern": "{date}_{topic}",
    "last_format": "",
}

_data_dir: Path | None = None


def base_dir() -> Path:
    """exe 実行時は exe のフォルダ、開発時はプロジェクトのルート。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _writable(d: Path) -> bool:
    try:
        d.mkdir(parents=True, exist_ok=True)
        probe = d / f".write_test_{os.getpid()}"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def data_dir() -> Path:
    global _data_dir
    if _data_dir is None:
        candidates = []
        if os.environ.get("REPORT_APP_DATA"):
            candidates.append(Path(os.environ["REPORT_APP_DATA"]))
        candidates.append(base_dir() / "data")
        local = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
        candidates.append(Path(local) / APP_NAME / "data")
        for c in candidates:
            if _writable(c):
                _data_dir = c
                break
        else:
            raise RuntimeError("データの保存先に書き込めません: " + ", ".join(map(str, candidates)))
    return _data_dir


def set_data_dir(path) -> Path:
    """保存先を明示的に指定する（テスト用）。"""
    global _data_dir
    _data_dir = Path(path)
    _data_dir.mkdir(parents=True, exist_ok=True)
    return _data_dir


# ---------------------------------------------------------------------------
# JSON 入出力（一時ファイルに書いてから置き換える）
# ---------------------------------------------------------------------------

def load_json(path: Path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default
    except (OSError, json.JSONDecodeError):
        # 壊れたファイルは退避して既定値で続行する
        try:
            shutil.copy2(path, str(path) + ".broken")
        except OSError:
            pass
        return default


def save_json(path: Path, obj) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=path.name, suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    s = dict(DEFAULT_SETTINGS)
    s.update(load_json(data_dir() / "settings.json", {}))
    return s


def save_settings(settings: dict) -> None:
    save_json(data_dir() / "settings.json", settings)


# ---------------------------------------------------------------------------
# フォーマット・過去事例
# ---------------------------------------------------------------------------

_FORBIDDEN_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def safe_name(name: str) -> str:
    """ファイル名・フォルダ名に使えない文字を「_」に置き換える。"""
    s = _FORBIDDEN_RE.sub("_", (name or "").strip()).strip(" .")
    return s or "無題"


def formats_dir() -> Path:
    return data_dir() / "formats"


def format_dir(name: str) -> Path:
    return formats_dir() / safe_name(name)


def list_formats() -> list[str]:
    root = formats_dir()
    if not root.exists():
        return []
    names = []
    for d in sorted(root.iterdir()):
        if (d / "format.json").exists():
            fmt = load_json(d / "format.json", {})
            names.append(fmt.get("name") or d.name)
    return names


def load_format(name: str) -> dict | None:
    return load_json(format_dir(name) / "format.json", None)


def save_format(fmt: dict, template_src=None) -> dict:
    """定義を保存する。template_src を渡すと原本を template.xlsx（.xlsm）としてコピーする。"""
    d = format_dir(fmt["name"])
    d.mkdir(parents=True, exist_ok=True)
    if template_src:
        ext = Path(template_src).suffix.lower()
        dest = d / f"template{ext}"
        if Path(template_src).resolve() != dest.resolve():
            shutil.copy2(template_src, dest)
        fmt["template"] = dest.name
    save_json(d / "format.json", fmt)
    return fmt


def template_path(fmt: dict) -> Path:
    return format_dir(fmt["name"]) / fmt.get("template", "template.xlsx")


def delete_format(name: str) -> None:
    shutil.rmtree(format_dir(name), ignore_errors=True)


def load_cases(name: str) -> list[dict]:
    return load_json(format_dir(name) / "cases.json", [])


def save_cases(name: str, cases: list[dict]) -> None:
    save_json(format_dir(name) / "cases.json", cases)
