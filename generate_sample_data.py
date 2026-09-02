"""
動作確認用のサンプルデータ生成スクリプト。

ねっぱんCSVと同じ列構成の、完全に架空のデータを生成する。
実際の予約データは一切使用していない。
"""

import argparse
import os
import random
from datetime import date, timedelta

import pandas as pd

import config

OUTPUT_DIR = "sample_data"
OUTPUT_NAME = "sample_reservations.csv"

# 架空の予約サイト・部屋タイプ・プラン
SITES = ["楽天トラベル", "じゃらんnet", "一休.com", "Expedia",
         "Booking.com", "Agoda"]
ROOM_TYPES = ["スタンダードツイン", "デラックスダブル", "和室10畳",
              "ファミリールーム", "コンドミニアム2BR"]
PLANS = ["素泊まりプラン", "朝食付きプラン", "連泊割プラン", "早期予約プラン"]

# 架空の宿泊者名(実在の人物とは無関係)
SAMPLE_NAMES = ["山田 太郎", "佐藤 花子", "鈴木 一郎", "田中 二郎",
                "高橋 三郎", "伊藤 四季", "渡辺 五月", "中村 六花"]


def generate(count, year, month, seed=0):
    """指定年月にチェックアウトする架空の予約データを生成する。"""
    random.seed(seed)  # 毎回同じデータが出るよう固定

    month_start = date(year, month, 1)
    next_month = (month_start + timedelta(days=32)).replace(day=1)
    month_end = next_month - timedelta(days=1)
    days_in_month = (month_end - month_start).days + 1

    rows = []
    for i in range(count):
        nights = random.choice([1, 1, 1, 2, 2, 3, 5])
        checkout = month_start + timedelta(days=random.randrange(days_in_month))
        checkin = checkout - timedelta(days=nights)

       
        # 一部が MIN_ORDER_DATE より前になるよう、申込日を広めに散らす
        order_date = checkin - timedelta(days=random.randint(1, 300))

        site = random.choice(SITES)
        adults = random.randint(1, 4)
        unit_price = random.randrange(6000, 20001, 500)
        total = unit_price * adults * nights

        # 1予約が泊数ぶんの行に分かれる形式を再現する
        for night_index in range(1, nights + 1):
            rows.append({
                "予約ID": 100000 + i,
                "予約区分": "予約",
                "予約番号": f"SAMPLE-{i:05d}",
                "泊目": night_index,
                "チェックイン日": _fmt(checkin),
                "チェックアウト日": _fmt(checkout),
                "申込日": _fmt(order_date),
                "泊数": nights,
                "予約サイト名称": site,
                "部屋タイプ名称": random.choice(ROOM_TYPES),
                "商品プラン名称": random.choice(PLANS),
                "室数": 1,
                "宿泊者氏名": random.choice(SAMPLE_NAMES),
                "大人人数計": adults,
                "子供人数計": 0,
                "幼児人数計": 0,
                "料金合計額": total,
                "大人単価": unit_price,
                "決済方法": random.choice(["事前カード決済", "現地決済"]),
            })

    return pd.DataFrame(rows)


def _fmt(d):
    """ねっぱんCSVと同じ日付形式にする(ゼロ埋めなし)。"""
    return f"{d.year}/{d.month}/{d.day}"


def main():
    parser = argparse.ArgumentParser(description="サンプルデータを生成する")
    parser.add_argument("--count", type=int, default=100, help="予約件数")
    parser.add_argument("--year", type=int, default=2026, help="対象年")
    parser.add_argument("--month", type=int, default=9, help="対象月")
    args = parser.parse_args()

    df = generate(args.count, args.year, args.month)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_DIR, OUTPUT_NAME)
    df.to_csv(path, index=False, encoding=config.CSV_ENCODING)

    print(f"生成しました: {path}")
    print(f"  予約件数: {args.count} 件 / 行数: {len(df)} 行")


if __name__ == "__main__":
    main()