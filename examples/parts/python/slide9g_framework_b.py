from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "フレームワーク部品②", breadcrumb="部品カタログ：フレームワーク",
                      lead="five_forces / business_model_canvas")
    l, r = split_cols(BODY, [1, 1.3], 20)
    five_forces(s, l, ["価格競争が激しい"], {"新規参入の脅威": ["IT企業の参入"], "代替品の脅威": ["無料アプリ"],
                                          "売り手の交渉力": ["部品の寡占"], "買い手の交渉力": ["大口顧客が強い"]})
    business_model_canvas(s, r, {"KP": ["販売代理店"], "KA": ["開発", "保守"], "KR": ["技術者"], "VP": ["手間なく申請できる"],
                                 "CR": ["専任担当"], "CH": ["Web", "代理店"], "CS": ["中小企業"], "C$": ["人件費", "サーバ費"],
                                 "R$": ["月額利用料"]})
