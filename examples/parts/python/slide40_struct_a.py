from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品①", breadcrumb="部品カタログ：構造",
                      lead="logic_tree（3階層まで）")
    logic_tree(s, BODY, ("利益の減少", [("売上の減少", ["客数の減少", "客単価の低下"]),
                                         ("費用の増加", ["原材料費の高騰", "人件費の増加", "物流費の増加"])]))
