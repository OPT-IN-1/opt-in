import pandas as pd
import warnings
warnings.filterwarnings('ignore')

df = pd.read_csv('/Users/tetsu/Downloads/個別申込者データ - シート3.csv', encoding='utf-8')

# カラム定義
age_col = '年齢'
income_col = '現在の年収を教えてください'
credit_col = 'クレジットカードはお持ちですか？'
willingness_col = '勉強会でお話ししたスクールを\n受講してみたいですか？ \n（回答内容が個別相談会に影響を及ぼすことはありません。安心して正直にご回答ください。）'
result_col = '結果'
execution_col = '実施可否'

# 成約フラグ作成（成約 + GH成約をまとめる）
df['成約フラグ'] = df[result_col].isin(['成約', 'GH成約（クロスセル/99万）']).astype(int)

# 実施済みフラグ（実施済み or 実施）
df['実施フラグ'] = df[execution_col].isin(['実施済み', '実施', '再アポ実施済み']).astype(int)

total = len(df)
total_closed = df['成約フラグ'].sum()
total_executed = df['実施フラグ'].sum()

print("=" * 90)
print("【成約データ概要】")
print("=" * 90)
print(f"総申込数:           {total}件")
print(f"個別相談実施数:       {total_executed}件（実施率: {total_executed/total*100:.1f}%）")
print(f"成約数:             {total_closed}件（成約 {(df[result_col]=='成約').sum()} + GH成約 {(df[result_col]=='GH成約（クロスセル/99万）').sum()}）")
print(f"対申込成約率:         {total_closed/total*100:.1f}%")
print(f"対実施成約率:         {total_closed/total_executed*100:.1f}%")

print(f"\n--- 結果の内訳 ---")
for val, cnt in df[result_col].value_counts(dropna=False).items():
    print(f"  {str(val):35s} {cnt:4d}件 ({cnt/total*100:5.1f}%)")


def cross_with_conversion(df, col, col_label):
    """属性 × 成約率のクロス分析"""
    # NaN除外
    valid = df.dropna(subset=[col])
    
    grouped = valid.groupby(col).agg(
        申込数=('成約フラグ', 'count'),
        実施数=('実施フラグ', 'sum'),
        成約数=('成約フラグ', 'sum'),
    ).reset_index()
    
    grouped['実施率'] = (grouped['実施数'] / grouped['申込数'] * 100).round(1)
    grouped['対申込成約率'] = (grouped['成約数'] / grouped['申込数'] * 100).round(1)
    grouped['対実施成約率'] = grouped.apply(
        lambda r: round(r['成約数'] / r['実施数'] * 100, 1) if r['実施数'] > 0 else 0.0, axis=1
    )
    
    # 合計行
    total_row = pd.DataFrame([{
        col: '【合計】',
        '申込数': grouped['申込数'].sum(),
        '実施数': grouped['実施数'].sum(),
        '成約数': grouped['成約数'].sum(),
        '実施率': round(grouped['実施数'].sum() / grouped['申込数'].sum() * 100, 1),
        '対申込成約率': round(grouped['成約数'].sum() / grouped['申込数'].sum() * 100, 1),
        '対実施成約率': round(grouped['成約数'].sum() / grouped['実施数'].sum() * 100, 1),
    }])
    grouped = pd.concat([grouped, total_row], ignore_index=True)
    
    grouped['申込数'] = grouped['申込数'].astype(int)
    grouped['実施数'] = grouped['実施数'].astype(int)
    grouped['成約数'] = grouped['成約数'].astype(int)
    
    print(f"\n{'=' * 90}")
    print(f"■ {col_label} × 成約率")
    print(f"{'=' * 90}")
    
    # テーブル表示
    header = f"{'':30s} | {'申込数':>6s} | {'実施数':>6s} | {'実施率':>7s} | {'成約数':>6s} | {'対申込成約率':>10s} | {'対実施成約率':>10s}"
    print(header)
    print("-" * 100)
    
    for _, row in grouped.iterrows():
        label = str(row[col])[:28]
        is_total = label == '【合計】'
        prefix = ">>> " if is_total else "    "
        print(f"{prefix}{label:26s} | {int(row['申込数']):6d} | {int(row['実施数']):6d} | {row['実施率']:6.1f}% | {int(row['成約数']):6d} | {row['対申込成約率']:8.1f}% | {row['対実施成約率']:8.1f}%")
    
    return grouped


# ============================
# 各属性 × 成約率
# ============================
print("\n\n" + "#" * 90)
print("# 各属性 × 成約率 クロス分析")
print("#" * 90)

g_age = cross_with_conversion(df, age_col, '年齢')
g_income = cross_with_conversion(df, income_col, '年収')
g_credit = cross_with_conversion(df, credit_col, 'クレカ有無')
g_will = cross_with_conversion(df, willingness_col, '入会意欲')


# ============================
# 3次元クロス分析（属性×属性×成約率）
# ============================
def cross_3way(df, row_col, col_col, row_label, col_label):
    """2つの属性でクロスし、各セルの成約率を表示"""
    valid = df.dropna(subset=[row_col, col_col])
    
    # 申込数
    ct_total = pd.crosstab(valid[row_col], valid[col_col], margins=True, margins_name='合計')
    
    # 成約数
    ct_conv = valid.groupby([row_col, col_col])['成約フラグ'].sum().unstack(fill_value=0)
    ct_conv['合計'] = ct_conv.sum(axis=1)
    ct_conv.loc['合計'] = ct_conv.sum()
    
    # 成約率
    ct_rate = (ct_conv / ct_total * 100).round(1)
    
    print(f"\n{'=' * 90}")
    print(f"■ {row_label} × {col_label} × 成約率")
    print(f"{'=' * 90}")
    
    print(f"\n[申込数]")
    print(ct_total.to_string())
    
    print(f"\n[成約数]")
    print(ct_conv.to_string())
    
    print(f"\n[成約率 %]")
    print(ct_rate.to_string())


print("\n\n" + "#" * 90)
print("# 属性 × 属性 × 成約率 クロス分析")
print("#" * 90)

cross_3way(df, age_col, income_col, '年齢', '年収')
cross_3way(df, age_col, credit_col, '年齢', 'クレカ有無')
cross_3way(df, age_col, willingness_col, '年齢', '入会意欲')
cross_3way(df, income_col, credit_col, '年収', 'クレカ有無')
cross_3way(df, income_col, willingness_col, '年収', '入会意欲')
cross_3way(df, credit_col, willingness_col, 'クレカ有無', '入会意欲')


# ============================
# 契約金額の分析
# ============================
print(f"\n\n{'=' * 90}")
print("■ 補足: 成約者の契約プラン内訳")
print(f"{'=' * 90}")
closed = df[df['成約フラグ'] == 1]
print(f"\n成約者数: {len(closed)}件")
print(f"\n--- 契約プラン ---")
print(closed['契約プラン名'].value_counts(dropna=False).to_string())
print(f"\n--- 契約金額 ---")
print(closed['契約金額(売上)'].value_counts(dropna=False).to_string())

print("\n\n" + "=" * 90)
print("分析完了")
