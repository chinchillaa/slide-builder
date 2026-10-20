from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "顧客・利用者部品①", breadcrumb="部品カタログ：顧客・利用者",
                      lead="customer_journey（行動・接点・課題＋感情の曲線）")
    customer_journey(s, BODY, ["認知", "検討", "申込", "利用", "継続"],
                     [("行動", ["広告を見る", "比較サイトで調べる", "Webで申し込む", "サービスを使う", "更新する"]),
                      ("接点", ["SNS", "Webサイト", "申込フォーム", "アプリ", "メール"]),
                      ("課題", ["知られていない", "違いが分かりにくい", "入力項目が多い", "使い方が難しい", "更新を忘れる"])],
                     emotions=[0, 1, -2, 1, 2])
