# -*- coding: utf-8 -*-
"""
list_parts.py - 登録済み部品の一覧を reference/parts.md に書き出す（部品を追加したら再実行）

使い方:
    python list_parts.py            # reference/parts.md を更新
    python list_parts.py --count    # 件数だけ表示
"""
import glob
import inspect
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import pptx_helpers as H  # noqa: E402

ORDER = ["スライド枠", "レイアウト", "テキスト", "カード", "表", "フロー", "構造", "図解", "数値", "グラフ",
         "スケジュール", "画像", "アイコン", "線・矢印", "補助", "会議・報告", "リスク・課題", "費用・効果",
         "要件・仕様", "目標管理", "顧客・利用者", "意思決定", "フレームワーク"]
ROOT = os.path.join(HERE, "..")


def main():
    reg = H.REGISTRY
    if "--count" in sys.argv:
        print(len(reg))
        return
    ex = {f: open(f, encoding="utf-8").read()
          for f in glob.glob(os.path.join(ROOT, "examples", "*", "python", "*.py"))}

    def examples(name):
        pat = re.compile(r"\b" + re.escape(name) + r"\(")
        hits = [os.path.relpath(f, ROOT).replace(os.sep, "/") for f, t in ex.items() if pat.search(t)]
        return "<br>".join(sorted(hits)[:2]) or "（他の部品の内部で使用）"

    cats = {}
    for name, meta in reg.items():
        cats.setdefault(meta["category"], []).append((name, meta))
    lines = ["# 部品一覧（自動生成: scripts/list_parts.py）", "",
             f"全 {len(reg)} 部品。`from pptx_helpers import *` で使える。単位は pt、area は (x, y, w, h)。",
             "見た目は `examples/parts/slides/`（カタログ画像）、使い方の実例は「実例」列のスクリプトを読む。", ""]
    for c in sorted(cats, key=lambda k: ORDER.index(k) if k in ORDER else 99):
        lines += [f"## {c}（{len(cats[c])}）", "", "| 関数 | 用途 | 引数 | 実例 |", "|---|---|---|---|"]
        for name, meta in cats[c]:
            sig = str(inspect.signature(meta["fn"])).replace("|", "\\|")
            doc = (inspect.getdoc(meta["fn"]) or "").split("\n")
            extra = " ".join(d.strip() for d in doc if d.strip() and d.strip() != meta["desc"])
            desc = meta["desc"] + (f"<br>{extra}" if extra else "")
            lines.append(f"| `{name}` | {desc} | `{sig}` | {examples(name)} |")
        lines.append("")
    out = os.path.join(ROOT, "reference", "parts.md")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"{len(reg)} 部品 → {os.path.abspath(out)}")
    for c in sorted(cats, key=lambda k: ORDER.index(k) if k in ORDER else 99):
        print(f"  {c}: {len(cats[c])}")


if __name__ == "__main__":
    main()
