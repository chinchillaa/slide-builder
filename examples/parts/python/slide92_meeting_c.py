from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "会議・報告部品③", breadcrumb="部品カタログ：会議・報告",
                      lead="meeting_bodies / revision_history")
    top, bot = split_rows(BODY, [1, 1], 16)
    meeting_bodies(s, top, [("ステアリングコミッティ", "月1回", "役員、PM", "重要事項の意思決定"),
                            ("定例会", "週1回", "PM、各チームリーダー", "進捗と課題の共有"),
                            ("分科会", "随時", "担当者", "個別テーマの検討")])
    revision_history(s, bot, [("1.0", "2026/10/01", "初版作成", "佐藤"), ("1.1", "2026/10/15", "スケジュールを更新", "鈴木"),
                              ("1.2", "2026/10/20", "体制図を追加", "高橋")])
