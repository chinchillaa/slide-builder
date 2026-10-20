from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "リスク・課題部品①", breadcrumb="部品カタログ：リスク・課題",
                      lead="risk_matrix（影響度×発生可能性）")
    risk_matrix(s, BODY, [("要員の不足", 3, 2), ("ベンダーの納期遅延", 3, 3), ("仕様変更の多発", 2, 3),
                          ("データ移行の失敗", 3, 1), ("利用者の反発", 1, 2)])
