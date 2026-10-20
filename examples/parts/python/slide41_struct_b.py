from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品②", breadcrumb="部品カタログ：構造",
                      lead="org_chart（体制図）")
    org_chart(s, BODY, ("プロジェクトオーナー\n取締役 山田", [
        ("PMO\n情報システム部 佐藤", [("業務チーム\n経理部 鈴木", ["経理担当", "営業担当"]),
                                     ("システムチーム\n情シス 高橋", ["アプリ担当", "インフラ担当"]),
                                     ("移行チーム\n情シス 田中", [])])]))
