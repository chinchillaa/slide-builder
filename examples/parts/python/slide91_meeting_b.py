from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "会議・報告部品②", breadcrumb="部品カタログ：会議・報告",
                      lead="weekly_report / notice_box")
    top, bot = split_rows(BODY, [1.6, 1], 14)
    weekly_report(s, top, ("注意", "テストの遅れを来週中に回復する見込み"), ["結合テストを70%完了", "移行リハーサル実施"],
                  ["結合テスト完了", "ユーザー教育の準備"], ["不具合の修正に時間がかかっている"], period="2026/10/13〜10/17")
    a, b = split_cols(bot, [1, 1], 16)
    notice_box(s, a, "お知らせ", ["11月から新しい申請画面に切り替わります"], kind="info")
    notice_box(s, b, "注意事項", ["旧画面からの申請は10月末で終了します"], kind="warn")
