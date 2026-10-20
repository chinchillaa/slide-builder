from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "仕様書型の本文", breadcrumb="部品カタログ：テキスト",
                      lead="spec_text（①見出し＋段落、i. ii. と イ) ロ) の階層、__下線__で強調） / code_block")
    l, r = split_cols(BODY, [1.6, 1], 20)
    spec_text(s, l, [("ID管理機能", "利用者ごとに一意のIDを発行し、__複数のアプリで同じIDを利用できる__ことが必要となる。"),
                     ("同意管理機能", ["第三者へ提供する際は、本人の同意を取得すること", (1, "同意の取得方法"), (2, "画面で取得"),
                                       (2, "書面で取得"), (1, "同意の撤回")])])
    code_block(s, r, '{\n  "blood_glucose": {\n    "value": 95,\n    "unit": "mg/dL"\n  },\n  "measured_at": "2026-10-20"\n}', title="データ例（JSON）")
