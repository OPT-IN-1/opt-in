import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# CSVファイル読み込み
df = pd.read_csv('/Users/tetsu/Downloads/個別申込者データ - シート3.csv', encoding='utf-8')

print("=" * 80)
print("【データ概要】")
print(f"総レコード数: {len(df)}")
print(f"\nカラム一覧:")
for i, col in enumerate(df.columns):
    print(f"  {i}: {col}")

# 対象カラムを特定
age_col = '年齢'
income_col = '現在の年収を教えてください'
credit_col = 'クレジットカードはお持ちですか？'
willingness_col = '勉強会でお話ししたスクールを\n受講してみたいですか？ \n（回答内容が個別相談会に影響を及ぼすことはありません。安心して正直にご回答ください。）'

print("\n" + "=" * 80)
print("【各カラムの値確認】")
for col_name, col in [("年齢", age_col), ("年収", income_col), ("クレカ有無", credit_col), ("入会意欲", willingness_col)]:
    print(f"\n--- {col_name} ---")
    vc = df[col].value_counts(dropna=False)
    for val, cnt in vc.items():
        print(f"  {val}: {cnt}件 ({cnt/len(df)*100:.1f}%)")

print("\n" + "=" * 80)
print("=" * 80)
print("\n【1. 個々のデータ分類】")

# ====== 年齢 ======
print("\n" + "-" * 60)
print("■ 年齢分布")
print("-" * 60)
age_vc = df[age_col].value_counts(dropna=False).sort_index()
total = len(df)
for val, cnt in age_vc.items():
    bar = "█" * int(cnt / total * 100)
    print(f"  {str(val):30s} | {cnt:4d}件 ({cnt/total*100:5.1f}%) {bar}")

# ====== 年収 ======
print("\n" + "-" * 60)
print("■ 年収分布")
print("-" * 60)
income_vc = df[income_col].value_counts(dropna=False)
for val, cnt in income_vc.items():
    bar = "█" * int(cnt / total * 100)
    print(f"  {str(val):30s} | {cnt:4d}件 ({cnt/total*100:5.1f}%) {bar}")

# ====== クレカ有無 ======
print("\n" + "-" * 60)
print("■ クレジットカード有無")
print("-" * 60)
credit_vc = df[credit_col].value_counts(dropna=False)
for val, cnt in credit_vc.items():
    bar = "█" * int(cnt / total * 100)
    print(f"  {str(val):30s} | {cnt:4d}件 ({cnt/total*100:5.1f}%) {bar}")

# ====== 入会意欲 ======
print("\n" + "-" * 60)
print("■ 入会意欲")
print("-" * 60)
will_vc = df[willingness_col].value_counts(dropna=False)
for val, cnt in will_vc.items():
    bar = "█" * int(cnt / total * 100)
    print(f"  {str(val):30s} | {cnt:4d}件 ({cnt/total*100:5.1f}%) {bar}")


print("\n" + "=" * 80)
print("=" * 80)
print("\n【2. クロス分析】")

# ====== 年齢 × 年収 ======
print("\n" + "-" * 60)
print("■ クロス分析①: 年齢 × 年収")
print("-" * 60)
ct1 = pd.crosstab(df[age_col], df[income_col], margins=True, margins_name='合計')
print(ct1.to_string())
print("\n[行割合 %]")
ct1_pct = pd.crosstab(df[age_col], df[income_col], normalize='index') * 100
print(ct1_pct.round(1).to_string())

# ====== 年齢 × クレカ有無 ======
print("\n" + "-" * 60)
print("■ クロス分析②: 年齢 × クレカ有無")
print("-" * 60)
ct2 = pd.crosstab(df[age_col], df[credit_col], margins=True, margins_name='合計')
print(ct2.to_string())
print("\n[行割合 %]")
ct2_pct = pd.crosstab(df[age_col], df[credit_col], normalize='index') * 100
print(ct2_pct.round(1).to_string())

# ====== 年齢 × 入会意欲 ======
print("\n" + "-" * 60)
print("■ クロス分析③: 年齢 × 入会意欲")
print("-" * 60)
ct3 = pd.crosstab(df[age_col], df[willingness_col], margins=True, margins_name='合計')
print(ct3.to_string())
print("\n[行割合 %]")
ct3_pct = pd.crosstab(df[age_col], df[willingness_col], normalize='index') * 100
print(ct3_pct.round(1).to_string())

# ====== 年収 × クレカ有無 ======
print("\n" + "-" * 60)
print("■ クロス分析④: 年収 × クレカ有無")
print("-" * 60)
ct4 = pd.crosstab(df[income_col], df[credit_col], margins=True, margins_name='合計')
print(ct4.to_string())
print("\n[行割合 %]")
ct4_pct = pd.crosstab(df[income_col], df[credit_col], normalize='index') * 100
print(ct4_pct.round(1).to_string())

# ====== 年収 × 入会意欲 ======
print("\n" + "-" * 60)
print("■ クロス分析⑤: 年収 × 入会意欲")
print("-" * 60)
ct5 = pd.crosstab(df[income_col], df[willingness_col], margins=True, margins_name='合計')
print(ct5.to_string())
print("\n[行割合 %]")
ct5_pct = pd.crosstab(df[income_col], df[willingness_col], normalize='index') * 100
print(ct5_pct.round(1).to_string())

# ====== クレカ有無 × 入会意欲 ======
print("\n" + "-" * 60)
print("■ クロス分析⑥: クレカ有無 × 入会意欲")
print("-" * 60)
ct6 = pd.crosstab(df[credit_col], df[willingness_col], margins=True, margins_name='合計')
print(ct6.to_string())
print("\n[行割合 %]")
ct6_pct = pd.crosstab(df[credit_col], df[willingness_col], normalize='index') * 100
print(ct6_pct.round(1).to_string())

print("\n" + "=" * 80)
print("分析完了")
