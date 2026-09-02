"""
請求書自動作成アプリ - CSV読み込みモジュール

ねっぱんからエクスポートしたCSVを読み込み、請求対象の行だけに絞り込む。
Excelの構造には一切依存しない(= 純粋なデータ処理だけを担当する)。
"""

import pandas as pd

import config


class CsvLoadError(Exception):
    """CSVの読み込み・整形に失敗したときに投げる例外。

    GUI側でこの例外を捕まえて、そのままメッセージダイアログに出せるようにしている。
    """


def load_reservations(csv_path, year, month):
    """ねっぱんCSVを読み込み、指定年月にチェックアウトした予約だけを返す。

    Args:
        csv_path: ねっぱんCSVのパス
        year:  対象年(例: 2026)
        month: 対象月(例: 8)

    Returns:
        pandas.DataFrame: 貼り付けに必要な6列のみ。チェックアウト日の昇順。

    Raises:
        CsvLoadError: 読み込み・整形に失敗した場合
    """
    df = _read_csv(csv_path)
    df = _drop_duplicate_nights(df)
    df = _parse_order_date(df)
    df = _filter_by_order_date(df)
    df = _drop_duplicate_reservations(df)
    df = _filter_by_month(df, year, month)
    df = _sort_and_select(df)
    return df


def _read_csv(csv_path):
    """CSVを読み込む。必要な列だけを読み込むので個人情報はメモリに乗らない。"""
    try:
        df = pd.read_csv(
            csv_path,
            encoding=config.CSV_ENCODING,
            usecols=config.USE_COLUMNS,
        )
    except UnicodeDecodeError:
        raise CsvLoadError(
            f"文字コードを {config.CSV_ENCODING} として読めませんでした。\n"
            "ねっぱんからエクスポートしたCSVかどうか確認してください。"
        )
    except ValueError as e:
        # usecols に指定した列がCSVに存在しない場合はここに来る
        raise CsvLoadError(
            "CSVに必要な列が見つかりませんでした。\n"
            f"詳細: {e}"
        )
    except FileNotFoundError:
        raise CsvLoadError(f"CSVファイルが見つかりません:\n{csv_path}")

    if df.empty:
        raise CsvLoadError("CSVにデータが1件も入っていません。")

    return df


def _drop_duplicate_nights(df):
    """「1泊 = 1行」形式のCSVを「1予約 = 1行」に正規化する。

    ねっぱんのエクスポート形式によっては、8泊の予約が8行に分かれて出力される。
    料金合計額は全行に同じ総額が入っているため、そのまま集計すると売上が
    泊数倍に膨らむ。泊目 == 1 の行だけ残すことでこれを防ぐ。

    もともと「1予約 = 1行」のCSVなら全行が泊目 == 1 なので、この処理をしても
    件数は変わらない(= どちらの形式でも安全に動く)。
    """
    return df[df[config.COL_NIGHT_INDEX] == 1].copy()

def _parse_order_date(df):
    """申込日を日付型に変換する。"""
    order_date = pd.to_datetime(
        df[config.COL_ORDER_DATE],
        format=config.CSV_DATE_FORMAT,
        errors="coerce",
    )

    if order_date.isna().any():
        bad_count = int(order_date.isna().sum())
        raise CsvLoadError(
            f"申込日を日付として読めない行が {bad_count} 件ありました。\n"
            f"想定している形式: {config.CSV_DATE_FORMAT}(例: 2026/9/27)"
        )

    df = df.copy()
    df[config.COL_ORDER_DATE] = order_date
    return df


def _filter_by_order_date(df):
    """申込日が基準日より前の予約を除外する。

    契約開始前に申し込まれた予約は請求対象外とするため。
    """
    threshold = pd.Timestamp(config.MIN_ORDER_DATE)
    filtered = df[df[config.COL_ORDER_DATE] >= threshold].copy()

    if filtered.empty:
        raise CsvLoadError(
            f"申込日が {config.MIN_ORDER_DATE} 以降の予約が1件もありませんでした。"
        )
    return filtered


def _drop_duplicate_reservations(df):
    """予約番号が重複している場合、申込日が最新の行だけを残す。

    同じ予約が「予約」と「変更」で複数行に出てくることがあるため、
    最も新しい申込日の行(=変更後の内容)を採用する。
    """
    df = df.sort_values(config.COL_ORDER_DATE)
    return df.drop_duplicates(
        subset=config.COL_RESERVATION_NO,
        keep="last",
    ).copy()


def _filter_by_month(df, year, month):
    """チェックアウト日が指定年月の行だけを残す。"""
    checkout = pd.to_datetime(
        df[config.COL_CHECKOUT],
        format=config.CSV_DATE_FORMAT,
        errors="coerce",  # 変換できない値は NaT にする
    )

    if checkout.isna().any():
        bad_count = int(checkout.isna().sum())
        raise CsvLoadError(
            f"チェックアウト日を日付として読めない行が {bad_count} 件ありました。\n"
            f"想定している形式: {config.CSV_DATE_FORMAT}(例: 2026/9/27)"
        )

    # 年と月が両方一致する行だけを残す
    mask = (checkout.dt.year == year) & (checkout.dt.month == month)
    filtered = df[mask].copy()

    if filtered.empty:
        raise CsvLoadError(
            f"{year}年{month}月にチェックアウトした予約が1件もありませんでした。\n"
            "対象年月、またはCSVの出力期間を確認してください。"
        )

    # 後続処理で日付として扱えるよう、変換済みの値で上書きしておく
    filtered[config.COL_CHECKOUT] = checkout[mask]
    filtered[config.COL_CHECKIN] = pd.to_datetime(
        filtered[config.COL_CHECKIN],
        format=config.CSV_DATE_FORMAT,
        errors="coerce",
    )

    return filtered


def _sort_and_select(df):
    """チェックアウト日順に並べ、貼り付けに使う6列だけを残す。"""
    df = df.sort_values(config.COL_CHECKOUT)
    columns = list(config.PASTE_COLUMN_MAP.keys())
    return df[columns].reset_index(drop=True)