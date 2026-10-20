# -*- coding: utf-8 -*-
"""
build_deck.py - デッキフォルダ内の slideNN_*.py を順に実行して 1 つの pptx を組み立てる

使い方:
    python build_deck.py <deck_dir> [--force]

<deck_dir>/
    deck.json            {"name": "出力ファイル名(拡張子なし)", "footer": "右下表記"}（任意）
    python/slide01_*.py  各ファイルに build(prs) を定義。ファイル名順に 1 枚ずつ追加する
    <name>.pptx          生成物（毎回作り直す）

安全策:
  * 前回ビルド後に pptx が手で編集されていたら上書きせず停止する（--force で上書き）
  * 1 枚が失敗しても残りはビルドし、失敗スライドはエラー表示のスライドに置き換える
"""
import importlib.util
import json
import os
import sys
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import pptx_helpers as H  # noqa: E402


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    if not args:
        print(__doc__)
        return 2
    deck_dir = os.path.abspath(args[0])
    cfg = {}
    cfg_path = os.path.join(deck_dir, "deck.json")
    if os.path.exists(cfg_path):
        with open(cfg_path, encoding="utf-8") as f:
            cfg = json.load(f)
    name = cfg.get("name") or os.path.basename(deck_dir)
    out = os.path.join(deck_dir, name + ".pptx")
    stamp = os.path.join(deck_dir, ".build_stamp")

    if os.path.exists(out) and not force:
        last = float(open(stamp).read()) if os.path.exists(stamp) else 0
        if os.path.getmtime(out) > last + 2:
            print(f"停止: {out} は前回のビルド後に変更されています（手修正の可能性）。\n"
                  f"上書きしてよい場合のみ --force を付けて再実行してください。")
            return 3

    py_dir = os.path.join(deck_dir, "python")
    files = sorted(f for f in os.listdir(py_dir) if f.startswith("slide") and f.endswith(".py"))
    prs = H.new_deck(footer=cfg.get("footer", ""))
    failed = []
    for fn in files:
        before = len(prs.slides)
        try:
            spec = importlib.util.spec_from_file_location(fn[:-3], os.path.join(py_dir, fn))
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            mod.build(prs)
        except Exception as e:  # 1 枚の失敗で全体を止めない
            # 途中まで追加されたスライドを取り除き、エラースライドに置き換える
            sldIdLst = prs.slides._sldIdLst
            while len(prs.slides) > before:
                rid = sldIdLst[-1].rId
                prs.part.drop_rel(rid)
                sldIdLst.remove(sldIdLst[-1])
            s = H.content_slide(prs, f"BUILD ERROR: {fn}")
            H._textbox(s, *H.BODY, str(e)[:400], size=12, color=H.COLOR_ACCENT, check=False)
            failed.append((fn, e, traceback.format_exc(limit=3)))
        print(f"  {fn}: slide {before + 1}-{len(prs.slides)}" + ("  ← 失敗" if failed and failed[-1][0] == fn else ""))

    errors, warns = H.check_presentation(prs)
    tmp = out + ".tmp.pptx"
    prs.save(tmp)
    try:
        os.replace(tmp, out)
    except PermissionError:
        print(f"保存できません: {out} が PowerPoint で開かれている可能性があります。閉じてから再実行してください。"
              f"\n（ビルド結果は {tmp} に残しています）")
        return 4
    with open(stamp, "w") as f:
        f.write(str(os.path.getmtime(out)))

    print(f"\n出力: {out}  ({len(prs.slides)} 枚)")
    for fn, e, tb in failed:
        print(f"\n[ERROR] {fn}: {e}\n{tb}")
    for m in errors:
        print(f"[ERROR] {m}")
    for m in warns:
        print(f"[WARN]  {m}")
    return 1 if failed or errors else 0


if __name__ == "__main__":
    sys.exit(main())
