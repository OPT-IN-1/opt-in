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
front_col = 'フロント\n登録経路'

df['成約フラグ'] = df[result_col].isin(['成約', 'GH成約（クロスセル/99万）']).astype(int)
df['実施フラグ'] = df[execution_col].isin(['実施済み', '実施', '再アポ実施済み']).astype(int)

# --- フロント流入経路を大分類に集約 ---
def categorize_route(val):
    if pd.isna(val):
        return '不明'
    val = str(val)
    if val == '不明':
        return '不明'
    if val.startswith('さきAI_YTチャンネルプロフ'):
        return 'さきAI_YTプロフ'
    if val.startswith('さきAI_YTQR'):
        return 'さきAI_YT(QR)'
    if val.startswith('さきAI_YT'):
        return 'さきAI_YT(動画)'
    if val.startswith('さきAI業務効率化'):
        return 'さきAI業務効率化'
    if val.startswith('たくむAIインスタ'):
        return 'たくむAIインスタ'
    if val.startswith('たくむAI業務効率化'):
        return 'たくむAI業務効率化'
    if val.startswith('たくむビジ系インスタ'):
        return 'たくむビジ系インスタ'
    if val.startswith('たくむビジ系ハウス'):
        return 'たくむビジ系ハウス'
    if val.startswith('たくむYT'):
        return 'たくむYT'
    if val.startswith('みさをインスタ'):
        return 'みさをインスタ'
    if val.startswith('みさをハウス'):
        return 'みさをハウス'
    if val.startswith('えむ'):
        return 'えむ'
    if val.startswith('lp01_Meta'):
        return 'Meta広告(LP01)'
    if val.startswith('lp02_Meta'):
        return 'Meta広告(LP02)'
    if val.startswith('ビジたくインスタ'):
        return 'たくむビジ系インスタ'
    return 'その他'

df['流入経路大分類'] = df[front_col].apply(categorize_route)
route_col = '流入経路大分類'

# さらにチャネル大分類
def channel_category(val):
    if 'さきAI_YT' in val:
        return 'さきAI YouTube系'
    if 'さきAI業務効率化' in val:
        return 'さきAI その他'
    if 'たくむ' in val:
        return 'たくむ系'
    if 'みさを' in val:
        return 'みさを系'
    if 'えむ' in val:
        return 'えむ系'
    if 'Meta' in val:
        return 'Meta広告系'
    if val == '不明':
        return '不明'
    return 'その他'

df['チャネル大分類'] = df[route_col].apply(channel_category)

valid = df.dropna(subset=[front_col])

# ====================================
# 1. 流入経路大分類 サマリー
# ====================================
print("=" * 110)
print("■ フロント流入経路（大分類）× 成約率")
print("=" * 110)

routes = df[route_col].value_counts().index.tolist()

header = f"{'流入経路':20s} | {'件数':>5s} | {'実施数':>5s} | {'実施率':>6s} | {'成約数':>5s} | {'対申込成約率':>9s} | {'対実施成約率':>9s}"
print(header)
print("-" * 110)

rows = []
for r in routes:
    sub = df[df[route_col] == r]
    n = len(sub)
    ex = int(sub['実施フラグ'].sum())
    cv = int(sub['成約フラグ'].sum())
    rate_a = cv / n * 100 if n > 0 else 0
    rate_e = cv / ex * 100 if ex > 0 else 0
    ex_r = ex / n * 100 if n > 0 else 0
    rows.append((r, n, ex, ex_r, cv, rate_a, rate_e))

rows.sort(key=lambda x: x[4], reverse=True)  # 成約数降順

for r, n, ex, ex_r, cv, rate_a, rate_e in rows:
    print(f"  {r:18s} | {n:5d} | {ex:5d} | {ex_r:5.1f}% | {cv:5d} | {rate_a:7.1f}% | {rate_e:7.1f}%")

n_all = len(df); ex_all = int(df['実施フラグ'].sum()); cv_all = int(df['成約フラグ'].sum())
print(f"  {'【合計】':18s} | {n_all:5d} | {ex_all:5d} | {ex_all/n_all*100:5.1f}% | {cv_all:5d} | {cv_all/n_all*100:7.1f}% | {cv_all/ex_all*100:7.1f}%")

# ====================================
# 2. チャネル大分類 サマリー
# ====================================
print(f"\n\n{'=' * 110}")
print("■ チャネル大分類 × 成約率")
print("=" * 110)

channels = df['チャネル大分類'].value_counts().index.tolist()
header = f"{'チャネル':18s} | {'件数':>5s} | {'実施数':>5s} | {'実施率':>6s} | {'成約数':>5s} | {'対申込成約率':>9s} | {'対実施成約率':>9s}"
print(header)
print("-" * 100)

ch_rows = []
for ch in channels:
    sub = df[df['チャネル大分類'] == ch]
    n = len(sub); ex = int(sub['実施フラグ'].sum()); cv = int(sub['成約フラグ'].sum())
    rate_a = cv/n*100 if n>0 else 0; rate_e = cv/ex*100 if ex>0 else 0; ex_r = ex/n*100 if n>0 else 0
    ch_rows.append((ch, n, ex, ex_r, cv, rate_a, rate_e))

ch_rows.sort(key=lambda x: x[4], reverse=True)
for ch, n, ex, ex_r, cv, rate_a, rate_e in ch_rows:
    print(f"  {ch:16s} | {n:5d} | {ex:5d} | {ex_r:5.1f}% | {cv:5d} | {rate_a:7.1f}% | {rate_e:7.1f}%")


# ====================================
# 3. 流入経路 × 各属性 クロス分析
# ====================================
def route_attr_cross(df, route_col, attr_col, attr_label, min_n=5):
    """流入経路 × 属性の構成比と成約率"""
    valid = df.dropna(subset=[attr_col])
    categories = sorted(valid[attr_col].unique(), key=lambda x: str(x))
    routes = valid[route_col].value_counts()
    routes = routes[routes >= min_n].index.tolist()
    
    def shorten(c, ml=12):
        s = str(c); return s[:ml] if len(s)>ml else s
    cats_short = [shorten(c) for c in categories]
    
    # 構成比
    print(f"\n{'=' * 110}")
    print(f"■ 流入経路 × {attr_label}【構成比 %】（n≧{min_n}の経路のみ）")
    print(f"{'=' * 110}")
    
    hdr = f"{'流入経路':20s} | {'n':>4s} |" + "|".join([f" {c:>12s}" for c in cats_short])
    print(hdr)
    print("-" * len(hdr))
    
    for r in routes:
        sub = valid[valid[route_col] == r]
        n = len(sub)
        vals = [len(sub[sub[attr_col]==cat])/n*100 for cat in categories]
        row = f"  {r:18s} | {n:4d} |" + "|".join([f" {v:11.1f}%" for v in vals])
        print(row)
    
    # 成約率
    print(f"\n{'=' * 110}")
    print(f"■ 流入経路 × {attr_label}【成約率 %】（n≧{min_n}の経路のみ）")
    print(f"{'=' * 110}")
    
    hdr2 = f"{'流入経路':20s} | {'n':>4s} |" + "|".join([f" {c:>12s}" for c in cats_short]) + f" | {'全体':>6s}"
    print(hdr2)
    print("-" * len(hdr2))
    
    for r in routes:
        sub = valid[valid[route_col] == r]
        n = len(sub)
        vals = []
        for cat in categories:
            cs = sub[sub[attr_col]==cat]
            cn = len(cs)
            ccv = cs['成約フラグ'].sum()
            rate = ccv/cn*100 if cn>0 else float('nan')
            vals.append(rate)
        total_rate = sub['成約フラグ'].sum()/n*100
        row = f"  {r:18s} | {n:4d} |" + "|".join([f" {v:11.1f}%" if not pd.isna(v) else f" {'---':>12s}" for v in vals]) + f" | {total_rate:5.1f}%"
        print(row)


# 入会意欲の短縮
will_map = {
    '入会するか悩んでいる': '悩んでいる',
    '入会をあまり考えていない': 'あまり考えてない',
    '入会をほぼ決めている': 'ほぼ決めている',
    '入会を全く考えていない': '全く考えてない',
    '入会を前向きに検討している': '前向き検討',
    '入会を決めており今すぐ始めたい': '今すぐ始めたい',
}
df['入会意欲短'] = df[willingness_col].map(will_map)

route_attr_cross(df, route_col, age_col, '年齢', min_n=8)
route_attr_cross(df, route_col, income_col, '年収', min_n=8)
route_attr_cross(df, route_col, job_col, '職業', min_n=8)
route_attr_cross(df, route_col, credit_col, 'クレカ有無', min_n=8)
route_attr_cross(df, route_col, '入会意欲短', '入会意欲', min_n=8)

print("\n\n" + "=" * 110)
print("分析完了")
