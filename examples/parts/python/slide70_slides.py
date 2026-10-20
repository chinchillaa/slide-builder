from pptx_helpers import *

def build(prs):
    agenda_slide(prs, ["本検討の概要", "現状分析", "今後の進め方", "質疑応答"], current=1)
    closing_slide(prs, contact="お問い合わせ：情報システム部（内線 1234）")
