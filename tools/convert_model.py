"""翻訳モデルの準備（開発 PC で 1 回だけ実行する）。

Helsinki-NLP/opus-mt-ja-en をダウンロードし、CTranslate2 の int8 形式に変換して
models/ja-en に保存します。実行時のアプリは torch / transformers を使いません。

    uv venv --python 3.11 .venv-dev
    uv pip install --python .venv-dev\\Scripts\\python.exe -r requirements-dev.txt
    .venv-dev\\Scripts\\python tools\\convert_model.py

保存されるもの:
    models/ja-en/model.bin           CTranslate2 形式のモデル（int8）
    models/ja-en/source.spm          日本語側の sentencepiece モデル
    models/ja-en/target.spm          英語側の sentencepiece モデル
    models/ja-en/shared_vocabulary.* 語彙
    models/ja-en/model_info.json     変換元と日時
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_MODEL = "Helsinki-NLP/opus-mt-ja-en"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--model", default=DEFAULT_MODEL, help="Hugging Face のモデル名")
    ap.add_argument("--output", default=str(ROOT / "models" / "ja-en"), help="保存先フォルダ")
    ap.add_argument("--quantization", default="int8", help="量子化（既定 int8）")
    ap.add_argument("--force", action="store_true", help="保存先があっても上書きする")
    args = ap.parse_args(argv)

    try:
        from ctranslate2.converters import TransformersConverter
        from huggingface_hub import snapshot_download
    except ImportError as e:
        print(f"必要なライブラリがありません: {e}\nrequirements-dev.txt をインストールしてください。", file=sys.stderr)
        return 1

    out = Path(args.output)
    if out.exists() and not args.force:
        print(f"{out} はすでにあります。作り直すときは --force を付けてください。")
        return 0

    print(f"ダウンロード中: {args.model}")
    src = snapshot_download(args.model, allow_patterns=[
        "*.json", "*.bin", "*.safetensors", "*.spm", "vocab*", "*.txt", "*.model"])

    print(f"変換中（{args.quantization}）: {out}")
    tmp = out.with_name(out.name + ".tmp")
    if tmp.exists():
        shutil.rmtree(tmp)
    converter = TransformersConverter(src, copy_files=["source.spm", "target.spm"])
    converter.convert(str(tmp), quantization=args.quantization, force=True)
    (tmp / "model_info.json").write_text(json.dumps({
        "source_model": args.model,
        "quantization": args.quantization,
        "converted_at": datetime.now().isoformat(timespec="seconds"),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    if out.exists():
        shutil.rmtree(out)
    tmp.rename(out)

    size = sum(f.stat().st_size for f in out.iterdir()) / 1e6
    print(f"完了しました: {out}（{size:.0f} MB）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
