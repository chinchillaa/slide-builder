from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "リスク・課題部品②", breadcrumb="部品カタログ：リスク・課題",
                      lead="risk_table / issue_tracker")
    top, bot = split_rows(BODY, [1, 1], 16)
    risk_table(s, top, [("ベンダーの納期遅延", "高", "中", "週次で進捗を確認する", "PMO"),
                        ("仕様変更の多発", "中", "高", "変更管理の手続きを定める", "PM"),
                        ("利用者の反発", "低", "中", "早期に説明会を開く", "業務チーム")])
    issue_tracker(s, bot, [("帳票レイアウトの確定", "高", "鈴木", "10/24", "進行中"), ("権限設計", "中", "高橋", "10/31", "未着手"),
                           ("テストデータの準備", "高", "田中", "10/20", "遅延")])
