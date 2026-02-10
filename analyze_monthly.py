import pandas as pd
import warnings
warnings.filterwarnings('ignore')

df = pd.read_csv('/Users/tetsu/Downloads/個別申込者データ - シート3.csv', encoding='utf-8')

# カラム定義
age_col = '年齢'
income_col = '現在の年収を教えてください'
credit_col = 'クレジットカードはお持ちですか？'
willingness_col = '勉強会でお話ししたスクールを\n受講してみたいですか？ \n（回答内容が個別相談会に影響を及ぼすことはありません。安心して正直にご回答ください。）'
job_col = '職業を教えてください'
result_col = '結果'
execution_col = '実施可否'

# フラグ作成
df['成約フラグ'] = df[result_col].isin(['成約', 'GH成約（クロスセル/99万）']).astype(int)
df['実施フラグ'] = df[execution_col].isin(['実施済み', '実施', '再アポ実施済み']).astype(int)

# 月カラム
df['申込月'] = pd.to_datetime(df['申込日時']).dt.to_period('M').astype(str)
months = sorted(df['申込月'].unique())

# ============================================
# 0. 月別サマリー
# ============================================
print("=" * 100)
print("■ 月別 全体サマリー")
print("=" * 100)

header = f"{'月':>8s} | {'申込数':>6s} | {'実施数':>6s} | {'実施率':>7s} | {'成約数':>6s} | {'対申込成約率':>10s} | {'対実施成約率':>10s}"
print(header)
print("-" * 100)

for m in months:
    sub = df[df['申込月'] == m]
    n = len(sub)
    ex = sub['実施フラグ'].sum()
    cv = sub['成約フラグ'].sum()
    ex_rate = ex / n * 100 if n > 0 else 0
    cv_app = cv / n * 100 if n > 0 else 0
    cv_ex = cv / ex * 100 if ex > 0 else 0
    print(f"  {m:>8s} | {n:6d} | {ex:6d} | {ex_rate:6.1f}% | {cv:6d} | {cv_app:8.1f}% | {cv_ex:8.1f}%")

# 合計
n = len(df); ex = df['実施フラグ'].sum(); cv = df['成約フラグ'].sum()
print(f"  {'合計':>7s} | {n:6d} | {ex:6d} | {ex/n*100:6.1f}% | {cv:6d} | {cv/n*100:8.1f}% | {cv/ex*100:8.1f}%")


def monthly_cross(df, attr_col, attr_label, months):
    """月別 × 属性 × 成約率を出力"""
    valid = df.dropna(subset=[attr_col])
    categories = sorted(valid[attr_col].unique(), key=lambda x: str(x))
    
    print(f"\n\n{'=' * 100}")
    print(f"■ 月別 × {attr_label} × 成約率")
    print(f"{'=' * 100}")
    
    # --- 月別 × 属性: 申込数 ---
    print(f"\n[申込数]")
    # ヘッダー
    cats_short = [str(c)[:12] for c in categories]
    header = f"{'月':>8s} |" + "|".join([f" {c:>12s}" for c in cats_short]) + f" | {'合計':>6s}"
    print(header)
    print("-" * len(header))
    
    for m in months:
        sub = valid[valid['申込月'] == m]
        vals = []
        for cat in categories:
            cnt = len(sub[sub[attr_col] == cat])
            vals.append(cnt)
        total = sum(vals)
        row = f"  {m:>8s} |" + "|".join([f" {v:12d}" for v in vals]) + f" | {total:6d}"
        print(row)
    
    # 合計行
    vals = [len(valid[valid[attr_col] == cat]) for cat in categories]
    total = sum(vals)
    row = f"  {'合計':>7s} |" + "|".join([f" {v:12d}" for v in vals]) + f" | {total:6d}"
    print(row)
    
    # --- 月別 × 属性: 成約数 ---
    print(f"\n[成約数]")
    print(header)
    print("-" * len(header))
    
    for m in months:
        sub = valid[valid['申込月'] == m]
        vals = []
        for cat in categories:
            cnt = sub[sub[attr_col] == cat]['成約フラグ'].sum()
            vals.append(int(cnt))
        total = sum(vals)
        row = f"  {m:>8s} |" + "|".join([f" {v:12d}" for v in vals]) + f" | {total:6d}"
        print(row)
    
    vals = [int(valid[valid[attr_col] == cat]['成約フラグ'].sum()) for cat in categories]
    total = sum(vals)
    row = f"  {'合計':>7s} |" + "|".join([f" {v:12d}" for v in vals]) + f" | {total:6d}"
    print(row)
    
    # --- 月別 × 属性: 成約率(%) ---
    print(f"\n[成約率 %]")
    cats_short_pct = [str(c)[:12] for c in categories]
    header_pct = f"{'月':>8s} |" + "|".join([f" {c:>12s}" for c in cats_short_pct]) + f" | {'合計':>6s}"
    print(header_pct)
    print("-" * len(header_pct))
    
    for m in months:
        sub = valid[valid['申込月'] == m]
        vals = []
        for cat in categories:
            cat_sub = sub[sub[attr_col] == cat]
            n = len(cat_sub)
            cv = cat_sub['成約フラグ'].sum()
            rate = cv / n * 100 if n > 0 else float('nan')
            vals.append(rate)
        total_n = len(sub)
        total_cv = sub['成約フラグ'].sum()
        total_rate = total_cv / total_n * 100 if total_n > 0 else 0
        
        row = f"  {m:>8s} |" + "|".join([f" {v:11.1f}%" if not pd.isna(v) else f" {'-':>12s}" for v in vals]) + f" | {total_rate:5.1f}%"
        print(row)
    
    # 合計行
    vals = []
    for cat in categories:
        cat_sub = valid[valid[attr_col] == cat]
        n = len(cat_sub)
        cv = cat_sub['成約フラグ'].sum()
        rate = cv / n * 100 if n > 0 else float('nan')
        vals.append(rate)
    total_rate = valid['成約フラグ'].sum() / len(valid) * 100
    row = f"  {'合計':>7s} |" + "|".join([f" {v:11.1f}%" if not pd.isna(v) else f" {'-':>12s}" for v in vals]) + f" | {total_rate:5.1f}%"
    print(row)
    
    # --- 月ごとの属性構成比 ---
    print(f"\n[月別 {attr_label} 構成比 %]")
    print(header_pct)
    print("-" * len(header_pct))
    
    for m in months:
        sub = valid[valid['申込月'] == m]
        total_n = len(sub)
        vals = []
        for cat in categories:
            cnt = len(sub[sub[attr_col] == cat])
            pct = cnt / total_n * 100 if total_n > 0 else 0
            vals.append(pct)
        row = f"  {m:>8s} |" + "|".join([f" {v:11.1f}%" for v in vals]) + f" | {100.0:5.1f}%"
        print(row)


# 各属性で月別クロス
monthly_cross(df, age_col, '年齢', months)
monthly_cross(df, income_col, '年収', months)
monthly_cross(df, job_col, '職業', months)
monthly_cross(df, credit_col, 'クレカ有無', months)
monthly_cross(df, willingness_col, '入会意欲', months)

print("\n\n" + "=" * 100)
print("分析完了")
