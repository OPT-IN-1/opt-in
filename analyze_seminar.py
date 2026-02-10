import csv
import re
from collections import defaultdict
from datetime import datetime, timedelta
import statistics

filepath = "/Users/tetsu/Downloads/個別申込者データ - シート4.csv"
TODAY = datetime(2026, 2, 8)
THREE_MONTHS_AGO = TODAY - timedelta(days=92)

with open(filepath, encoding="utf-8") as f:
    reader = csv.reader(f)
    rows = list(reader)

all_rows = []

for i, row in enumerate(rows):
    if i < 9:
        continue
    if len(row) < 14:
        continue

    date_str = row[1].strip()
    time_str = row[2].strip()

    day_match = re.search(r'\((.)\)', date_str)
    if not day_match:
        continue
    day = day_match.group(1)

    date_part_match = re.match(r'(\d{4}/\d{1,2}/\d{1,2})', date_str)
    if not date_part_match:
        continue
    try:
        seminar_date = datetime.strptime(date_part_match.group(1), '%Y/%m/%d')
    except ValueError:
        continue

    if seminar_date < THREE_MONTHS_AGO or seminar_date > TODAY:
        continue
    if not time_str:
        continue

    try:
        個別申込 = int(row[12]) if row[12].strip() else None
    except (ValueError, IndexError):
        個別申込 = None

    try:
        着席数 = int(row[8]) if row[8].strip() else None
    except (ValueError, IndexError):
        着席数 = None

    try:
        申込数 = int(row[4]) if row[4].strip() else None
    except (ValueError, IndexError):
        申込数 = None

    tc = '昼' if '12:00' in time_str else '夜'

    all_rows.append({
        'date': date_str,
        'day': day,
        'time_category': tc,
        '申込数': 申込数,
        '着席数': 着席数,
        '個別申込': 個別申込,
    })

day_order = ['月', '火', '水', '木', '金', '土', '日']
time_order = ['昼', '夜']

cross = defaultdict(lambda: {'個別': [], '申込': [], '着席': []})
for r in all_rows:
    key = (r['day'], r['time_category'])
    if r['個別申込'] is not None:
        cross[key]['個別'].append(r['個別申込'])
    if r['申込数'] is not None:
        cross[key]['申込'].append(r['申込数'])
    if r['着席数'] is not None:
        cross[key]['着席'].append(r['着席数'])

# 水曜昼の外れ値（1/28の128名回）を除外した補正版も計算
水曜昼_個別_raw = cross[('水', '昼')]['個別']
水曜昼_申込_raw = cross[('水', '昼')]['申込']
水曜昼_着席_raw = cross[('水', '昼')]['着席']

# 1/28は申込128 → それを除外
水曜昼_申込_adj = [x for x in 水曜昼_申込_raw if x != 128]
# 対応する個別申込（23）と着席（88）も除外
# 実データ: 12/17(97,69,33), 1/14(88,56,21), 1/28(128,88,23), 2/4(93,63,26)
水曜昼_個別_adj = [33, 21, 26]  # 1/28の23を除外
水曜昼_着席_adj = [69, 56, 63]  # 1/28の88を除外

print("=" * 80)
print("  個別相談申込数 予測レンジ（直近3ヶ月ベース）")
print("  ※水曜昼は外れ値(1/28 申込128名)除外の補正値も併記")
print("=" * 80)

print(f"\n{'─'*80}")
print(f"{'':>8} | {'予測値':>6} | {'下限':>6} | {'上限':>6} | {'サンプル':>6} | {'信頼度':>6} | 備考")
print(f"{'─'*80}")

results = []
for day in day_order:
    for tc in time_order:
        key = (day, tc)
        vals = cross[key]['個別']
        sem_vals = cross[key]['申込']
        sit_vals = cross[key]['着席']
        if not vals:
            continue
        
        avg = statistics.mean(vals)
        n = len(vals)
        
        if n >= 2:
            std = statistics.stdev(vals)
        else:
            std = 0
        
        low = max(0, avg - std)
        high = avg + std
        
        if n >= 7:
            trust = "★★★"
        elif n >= 4:
            trust = "★★☆"
        elif n >= 2:
            trust = "★☆☆"
        else:
            trust = "☆☆☆"
        
        sem_avg = statistics.mean(sem_vals) if sem_vals else 0
        sit_avg = statistics.mean(sit_vals) if sit_vals else 0
        
        note = ""
        if day in ['月','火','木','金'] and tc == '昼':
            note = "※サンプル極少"
        
        results.append((f"{day}・{tc}", avg, low, high, n, trust, note, sem_avg, sit_avg))

# 予測値でソート
results.sort(key=lambda x: x[1], reverse=True)

for label, avg, low, high, n, trust, note, sem, sit in results:
    # 水曜昼は特別表記
    if label == '水・昼':
        print(f"  {label:>6} | {avg:>6.1f} | {low:>6.1f} | {high:>6.1f} | {n:>5}回 | {trust} | 外れ値込み")
        # 補正版
        adj_avg = statistics.mean(水曜昼_個別_adj)
        adj_std = statistics.stdev(水曜昼_個別_adj) if len(水曜昼_個別_adj) >= 2 else 0
        adj_low = max(0, adj_avg - adj_std)
        adj_high = adj_avg + adj_std
        print(f"  {'〃補正':>6} | {adj_avg:>6.1f} | {adj_low:>6.1f} | {adj_high:>6.1f} | {len(水曜昼_個別_adj):>5}回 | ★☆☆ | 1/28除外")
    else:
        print(f"  {label:>6} | {avg:>6.1f} | {low:>6.1f} | {high:>6.1f} | {n:>5}回 | {trust} | {note}")

# フロー版
print(f"\n{'─'*80}")
print(f"【予測フロー】申込 → 着席 → 個別申込（現実的な予測値）")
print(f"{'─'*80}")
print(f"{'':>8} | {'申込':>6} → {'着席':>5} → {'個別':>5} | {'個別レンジ':>14} | {'n':>3} | 信頼度")
print(f"{'─'*80}")

for label, avg, low, high, n, trust, note, sem, sit in results:
    if label == '水・昼':
        # 補正版のみ表示
        adj_sem = statistics.mean(水曜昼_申込_adj)
        adj_sit = statistics.mean(水曜昼_着席_adj)
        adj_avg = statistics.mean(水曜昼_個別_adj)
        adj_std = statistics.stdev(水曜昼_個別_adj)
        adj_low = max(0, adj_avg - adj_std)
        adj_high = adj_avg + adj_std
        print(f"  {label:>6} | {adj_sem:>6.1f} → {adj_sit:>5.1f} → {adj_avg:>5.1f} | {adj_low:>5.1f} 〜 {adj_high:>5.1f} | {3:>3} | ★☆☆ 補正済")
    else:
        print(f"  {label:>6} | {sem:>6.1f} → {sit:>5.1f} → {avg:>5.1f} | {low:>5.1f} 〜 {high:>5.1f} | {n:>3} | {trust} {note}")
