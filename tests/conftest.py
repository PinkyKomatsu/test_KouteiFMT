"""テスト共通の準備。

- テスト中は socket による通信を禁止する（実行時の通信ゼロの確認）
- 架空のフォーマットと過去報告書をセッションごとに 1 回だけ作る
"""
from __future__ import annotations

import socket
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tests import make_samples  # noqa: E402

SHEET = make_samples.SHEET


class NetworkBlocked(RuntimeError):
    pass


def _deny(*args, **kwargs):
    raise NetworkBlocked("テスト中の通信は禁止されています")


@pytest.fixture(autouse=True, scope="session")
def block_network():
    """socket の接続・名前解決をすべて失敗させる。"""
    saved = {
        "connect": socket.socket.connect,
        "connect_ex": socket.socket.connect_ex,
        "create_connection": socket.create_connection,
        "getaddrinfo": socket.getaddrinfo,
    }
    socket.socket.connect = _deny
    socket.socket.connect_ex = _deny
    socket.create_connection = _deny
    socket.getaddrinfo = _deny
    yield
    socket.socket.connect = saved["connect"]
    socket.socket.connect_ex = saved["connect_ex"]
    socket.create_connection = saved["create_connection"]
    socket.getaddrinfo = saved["getaddrinfo"]


@pytest.fixture(scope="session")
def samples(tmp_path_factory):
    return make_samples.make_all(tmp_path_factory.mktemp("samples"))


@pytest.fixture(scope="session")
def data_dir(tmp_path_factory):
    from app import storage
    return storage.set_data_dir(tmp_path_factory.mktemp("data"))


@pytest.fixture(scope="session")
def fmt(samples, data_dir):
    """解析して保存したフォーマット定義。"""
    from app import analyzer, storage
    definition = analyzer.analyze(samples["format"], "工程報告")
    storage.save_format(definition, template_src=samples["format"])
    return definition


@pytest.fixture(scope="session")
def template(fmt):
    from app import storage
    return storage.template_path(fmt)


@pytest.fixture(scope="session")
def imported(fmt, samples):
    """過去報告書を取り込んだ結果（個人情報は伏せ字）。{"cases", "log", "pii"}"""
    from app import cases, storage
    found = []

    def on_pii(path, findings):
        found.append((path.name, findings))
        return "mask"

    items: list[dict] = []
    log = cases.import_paths(fmt, items, [samples["reports_dir"]], on_pii=on_pii)
    storage.save_cases(fmt["name"], items)
    return {"cases": items, "log": log, "pii": found}


def model_available() -> bool:
    from app import storage
    return (storage.default_model_dir() / "model.bin").exists()


TODAY = date(2026, 9, 27)
TOPIC = "移行リハーサル（第2回）の結果報告"
PHASE = "移行リハーサル"
STATUS = "遅延"
MEMO = """9/26に移行リハーサル（第2回）を実施した
データ移行は計画の95%まで完了した
レセプトの点検で3件のエラーが発生した
課題：点数マスタの差分が未確認
【今後の予定】
10/3までにエラーの原因を調査する
10/10に本番移行の稼働判定会議を行う
エラー原因の調査｜山田｜10/3｜対応中
点数マスタの差分確認｜佐藤｜10/5｜未着手"""


@pytest.fixture(scope="session")
def translator(imported):
    from app import cases, glossary, storage, translator_en
    return translator_en.Translator(storage.default_model_dir(), glossary.load(),
                                    cases.translation_memory(imported["cases"]))


@pytest.fixture(scope="session")
def refs(fmt, imported):
    from app import cases
    hits = cases.Index(fmt, imported["cases"]).search(TOPIC + "\n" + MEMO, phase=PHASE)
    return [c for s, c in hits[:3] if s > 0]


@pytest.fixture(scope="session")
def draft(fmt, imported, refs, translator):
    from app import report, storage
    return report.build_draft(fmt, TOPIC, PHASE, STATUS, MEMO, refs, imported["cases"],
                              dict(storage.DEFAULT_SETTINGS), translator, TODAY)
