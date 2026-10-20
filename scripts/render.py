# -*- coding: utf-8 -*-
"""
render.py - PowerPoint 本体（COM）で pptx を PNG 画像に書き出す

使い方:
    python render.py <pptx> [<out_dir>] [--slides 3,5,7] [--width 1600]

  * out_dir 省略時は pptx と同じ場所の slides/ に slide_NN.png を出力
  * PowerPoint で既に開いているファイルやウィンドウには触れない（読み取り専用で別に開き、自分が開いたものだけ閉じる）
"""
import os
import sys


def main():
    argv = sys.argv[1:]
    opts = {}
    pos = []
    i = 0
    while i < len(argv):
        if argv[i].startswith("--"):
            opts[argv[i][2:]] = argv[i + 1]
            i += 2
        else:
            pos.append(argv[i])
            i += 1
    if not pos:
        print(__doc__)
        return 2
    pptx = os.path.abspath(pos[0])
    out_dir = os.path.abspath(pos[1]) if len(pos) > 1 else os.path.join(os.path.dirname(pptx), "slides")
    os.makedirs(out_dir, exist_ok=True)
    width = int(opts.get("width", 1600))
    only = {int(s) for s in opts["slides"].split(",")} if "slides" in opts else None

    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    was_open = app.Presentations.Count
    pres = app.Presentations.Open(pptx, True, False, False)   # ReadOnly, Untitled=False, WithWindow=False
    try:
        n = pres.Slides.Count
        height = int(width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth)
        done = []
        for k in range(1, n + 1):
            if only and k not in only:
                continue
            path = os.path.join(out_dir, f"slide_{k:02d}.png")
            pres.Slides(k).Export(path, "PNG", width, height)
            done.append(path)
    finally:
        pres.Close()
        if was_open == 0 and app.Presentations.Count == 0:
            app.Quit()
    print(f"{len(done)} 枚を出力: {out_dir}")
    for p in done:
        print("  " + p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
