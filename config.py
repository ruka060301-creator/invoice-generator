"""
請求書自動作成アプリ - 設定モジュール

このファイルには「変わりうる値」だけを集めています。
テンプレートのレイアウトが変わったときは、原則ここだけを直せば済むようにしています。
"""
import datetime
# ---------------------------------------------------------------------------
# CSV(ねっぱん予約リスト)の設定
# ---------------------------------------------------------------------------

# ねっぱんCSVの文字コード。Windowsの日本語CSVは基本これ。
CSV_ENCODING = "cp932"

# CSVの日付形式。例: 2026/9/27(ゼロ埋めなし)
CSV_DATE_FORMAT = "%Y/%m/%d"

# CSVから読み込む列名。ここに書いた列だけを読み込む。
# → 氏名・電話番号・住所・メールアドレスなどの個人情報は最初から読み込まない設計。
COL_CHECKIN = "チェックイン日"
COL_CHECKOUT = "チェックアウト日"
COL_NIGHTS = "泊数"
COL_SITE = "予約サイト名称"
COL_ROOM_TYPE = "部屋タイプ名称"
COL_AMOUNT = "料金合計額"
COL_RESERVATION_NO = "予約番号"  # 重複除去の判定キー
COL_ORDER_DATE = "申込日"        # 除外判定と重複時の優先順位に使う
COL_NIGHT_INDEX = "泊目"  # 重複行を除くためのフィルタ用(貼り付けはしない)

# 実際に読み込む列の一覧
USE_COLUMNS = [
    COL_NIGHT_INDEX,
    COL_RESERVATION_NO,
    COL_ORDER_DATE,
    COL_CHECKIN,
    COL_CHECKOUT,
    COL_NIGHTS,
    COL_SITE,
    COL_ROOM_TYPE,
    COL_AMOUNT,
]

# ---------------------------------------------------------------------------
# 貼り付けシートの設定
# ---------------------------------------------------------------------------

# シート名は施設ごとに違うため、
# 完全一致ではなく「この文字列で終わるシート」を探す方式にする。
PASTE_SHEET_SUFFIX = "貼り付けシート"

# データを書き始める行(1行目はヘッダー)
PASTE_START_ROW = 2

# CSVの列 → 貼り付けシートの列(番号)の対応
# 1=A, 2=B, ... 5=E, 6=F, 8=H, 9=I, 10=J, 26=Z
PASTE_COLUMN_MAP = {
    COL_CHECKIN: 5,     # E列
    COL_CHECKOUT: 6,    # F列
    COL_NIGHTS: 8,      # H列
    COL_SITE: 9,        # I列
    COL_ROOM_TYPE: 10,  # J列
    COL_AMOUNT: 26,     # Z列
}

# ---------------------------------------------------------------------------
# 請求書シートの設定
# ---------------------------------------------------------------------------

INVOICE_SHEET_NAME = "請求書"

# 明細行の開始行。55行目が貼り付けシートの2行目に対応する。
DETAIL_START_ROW = 55

# 集計式 SUM(H55:H179) の終端。ここが明細行の上限になる。
DETAIL_END_ROW = 179

# 1回で処理できる最大件数
MAX_RECORDS = DETAIL_END_ROW - DETAIL_START_ROW + 1  # = 125

# 明細行の列(番号)
DETAIL_COL_CHECKIN = 3    # C列
DETAIL_COL_CHECKOUT = 4   # D列
DETAIL_COL_NIGHTS = 5     # E列
DETAIL_COL_SITE = 6       # F列
DETAIL_COL_ROOM_TYPE = 7  # G列
DETAIL_COL_AMOUNT = 8     # H列

# 毎月書き換えが必要なセル
CELL_COMMISSION_LABEL = "A17"  
CELL_INVOICE_NUMBER = "F7"     
CELL_ISSUE_DATE = "H7"         # 発行日(翌月1日)
CELL_DUE_DATE = "H9"           # 支払期限(翌月末日)

# ---------------------------------------------------------------------------
# 請求ルール
# ---------------------------------------------------------------------------

# この日より前に申し込まれた予約は請求対象外とする
MIN_ORDER_DATE = datetime.date(2026, 3, 1)
# Agoda経由の予約は手数料が引かれた金額で入ってくるため、割り戻して総額にする
ADJUST_SITE_NAME = "Agoda"
ADJUST_DIVISOR = 0.88

# 送客手数料の料率(請求書シートの数式内で使用)
COMMISSION_RATE = 0.05

# 出力ファイル名のパターン
OUTPUT_FILENAME_PATTERN = "{year}{month:02d}請求書_{hotel}様.xlsx"