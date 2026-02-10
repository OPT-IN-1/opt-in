// ============================================================================
// 個別申込者データ 自動分析スクリプト (Google Apps Script)
// 
// 使い方:
//   1. スプレッドシートで「拡張機能 > Apps Script」を開く
//   2. このコードを全てコピーしてエディタに貼り付け
//   3. CONFIG.SOURCE_SHEET_NAME を元データのシート名に変更
//   4. 「runAllAnalysis」を選択して▶実行
//   5. 毎日自動実行したい場合は setupDailyTrigger() を一度実行
// ============================================================================

// ===== 設定 =====
const CONFIG = {
  SOURCE_SHEET_NAME: 'シート3',  // ← 元データのシート名に変更してください
  
  // カラム検索キーワード（ヘッダー行のテキストに含まれる文字列で特定）
  COL_KEYWORDS: {
    age:         '年齢',
    income:      '現在の年収',
    job:         '職業を教えて',
    credit:      'クレジットカード',
    willingness: '受講してみたい',
    result:      '結果',
    execution:   '実施可否',
    appDate:     '申込日時',
    staff:       '担当者',
    frontRoute:  null, // 特殊処理（下記参照）
    seminarDate: 'セミナー参加日',
  },
  
  // 成約と判定する値
  CONVERSION_VALUES: ['成約', 'GH成約（クロスセル/99万）'],
  // 実施済みと判定する値
  EXECUTION_VALUES: ['実施済み', '実施', '再アポ実施済み'],
  
  // 出力シート名
  SHEETS: {
    attr:    '1_属性分布',
    conv:    '2_属性x成約率',
    monthly: '3_月別推移',
    seminar: '4_セミナー別',
    staff:   '5_担当者別',
    route:   '6_流入経路別',
  },
};


// ============================================================================
// ユーティリティ関数
// ============================================================================

/**
 * ヘッダー行からキーワードを含むカラムのインデックスを検索
 */
function findColIdx(headers, keyword) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').replace(/\n/g, '');
    if (h.indexOf(keyword) !== -1) return i;
  }
  return -1;
}

/**
 * フロント登録経路カラムを特定（「フロント」と「登録経路」を含み「集計シート」を含まない）
 */
function findFrontRouteIdx(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').replace(/\n/g, '');
    if (h.indexOf('フロント') !== -1 && h.indexOf('登録経路') !== -1 && h.indexOf('集計シート') === -1) {
      return i;
    }
  }
  return -1;
}

/**
 * 担当者カラムを特定（「個別相談」と「担当者」を含む）
 */
function findStaffIdx(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').replace(/\n/g, '');
    if (h.indexOf('個別相談') !== -1 && h.indexOf('担当者') !== -1) return i;
  }
  return -1;
}

/**
 * シートを取得（なければ作成）してクリア
 */
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  return sheet;
}

/**
 * 2D配列をシートに一括書き出し
 */
function writeBlock(sheet, data, startRow, startCol) {
  if (data.length === 0) return;
  const numRows = data.length;
  const numCols = Math.max(...data.map(r => r.length));
  // パディング
  const padded = data.map(r => {
    const row = r.slice();
    while (row.length < numCols) row.push('');
    return row;
  });
  sheet.getRange(startRow, startCol, numRows, numCols).setValues(padded);
}

/**
 * セルの値を文字列として取得（null/undefined/空を統一）
 */
function strVal(v) {
  if (v === null || v === undefined || v === '') return '';
  return String(v).trim();
}

/**
 * パーセント文字列（小数1桁）
 */
function pct(num, denom) {
  if (!denom || denom === 0) return '-';
  return (num / denom * 100).toFixed(1) + '%';
}

/**
 * 日付から "YYYY-MM" を取得
 */
function toYearMonth(val) {
  if (!val) return '';
  let d;
  if (val instanceof Date) {
    d = val;
  } else {
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  return y + '-' + m;
}

/**
 * 日付を "MM/DD(曜日) HH:MM" に整形
 */
function formatSeminarDate(val) {
  if (!val) return '';
  let d;
  if (val instanceof Date) {
    d = val;
  } else {
    d = new Date(val);
  }
  if (isNaN(d.getTime())) return '';
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  const hh = ('0' + d.getHours()).slice(-2);
  const mi = ('0' + d.getMinutes()).slice(-2);
  return mm + '/' + dd + '(' + days[d.getDay()] + ') ' + hh + ':' + mi;
}


// ============================================================================
// データ前処理
// ============================================================================

/**
 * フロント流入経路を大分類に変換
 */
function categorizeRoute(val) {
  const s = strVal(val);
  if (!s || s === '不明') return '不明';
  if (s.indexOf('さきAI_YTチャンネルプロフ') === 0) return 'さきAI_YTプロフ';
  if (s.indexOf('さきAI_YTQR') === 0) return 'さきAI_YT(QR)';
  if (s.indexOf('さきAI_YT') === 0) return 'さきAI_YT(動画)';
  if (s.indexOf('さきAI業務効率化') === 0) return 'さきAI業務効率化';
  if (s.indexOf('たくむAIインスタ') === 0) return 'たくむAIインスタ';
  if (s.indexOf('たくむAI業務効率化') === 0) return 'たくむAI業務効率化';
  if (s.indexOf('たくむビジ系インスタ') === 0 || s.indexOf('ビジたくインスタ') === 0) return 'たくむビジ系インスタ';
  if (s.indexOf('たくむビジ系ハウス') === 0) return 'たくむビジ系ハウス';
  if (s.indexOf('たくむYT') === 0) return 'たくむYT';
  if (s.indexOf('みさをインスタ') === 0) return 'みさをインスタ';
  if (s.indexOf('みさをハウス') === 0) return 'みさをハウス';
  if (s.indexOf('えむ') === 0) return 'えむ';
  if (s.indexOf('lp01_Meta') === 0) return 'Meta広告(LP01)';
  if (s.indexOf('lp02_Meta') === 0) return 'Meta広告(LP02)';
  return 'その他';
}

/**
 * 入会意欲を短縮名に変換
 */
function shortenWillingness(val) {
  const s = strVal(val);
  const map = {
    '入会するか悩んでいる': '悩んでいる',
    '入会をあまり考えていない': 'あまり考えてない',
    '入会をほぼ決めている': 'ほぼ決めている',
    '入会を全く考えていない': '全く考えてない',
    '入会を前向きに検討している': '前向き検討',
    '入会を決めており今すぐ始めたい': '今すぐ始めたい',
  };
  return map[s] || s;
}


// ============================================================================
// 集計エンジン
// ============================================================================

/**
 * カテゴリ別の件数カウント
 * @return {Object} { category: count, ... } （ソート済み配列も返す）
 */
function countBy(rows, colIdx) {
  const counts = {};
  rows.forEach(row => {
    const v = strVal(row[colIdx]);
    if (!v) return;
    counts[v] = (counts[v] || 0) + 1;
  });
  return counts;
}

/**
 * カテゴリ別の成約分析
 * @return {Array} [{ category, total, executed, converted, execRate, convAppRate, convExRate }, ...]
 */
function conversionBy(rows, colIdx, convIdx, execIdx) {
  const groups = {};
  rows.forEach(row => {
    const cat = strVal(row[colIdx]);
    if (!cat) return;
    if (!groups[cat]) groups[cat] = { total: 0, executed: 0, converted: 0 };
    groups[cat].total++;
    if (row[execIdx]) groups[cat].executed++;
    if (row[convIdx]) groups[cat].converted++;
  });
  
  const result = [];
  for (const cat in groups) {
    const g = groups[cat];
    result.push({
      category: cat,
      total: g.total,
      executed: g.executed,
      converted: g.converted,
    });
  }
  return result;
}

/**
 * 2次元クロス集計（行カテゴリ × 列カテゴリ）
 * @return { rowCats, colCats, counts[row][col], totals }
 */
function crossTab(rows, rowIdx, colIdx) {
  const data = {};
  const rowSet = {};
  const colSet = {};
  
  rows.forEach(row => {
    const r = strVal(row[rowIdx]);
    const c = strVal(row[colIdx]);
    if (!r || !c) return;
    rowSet[r] = true;
    colSet[c] = true;
    if (!data[r]) data[r] = {};
    data[r][c] = (data[r][c] || 0) + 1;
  });
  
  const rowCats = Object.keys(rowSet).sort();
  const colCats = Object.keys(colSet).sort();
  
  return { rowCats, colCats, data };
}

/**
 * 2次元クロス × 成約率
 * @return { rowCats, colCats, totalData[r][c], convData[r][c], rateData[r][c] }
 */
function crossConvRate(rows, rowIdx, colIdx, convFlagIdx) {
  const totalD = {};
  const convD = {};
  const rowSet = {};
  const colSet = {};
  
  rows.forEach(row => {
    const r = strVal(row[rowIdx]);
    const c = strVal(row[colIdx]);
    if (!r || !c) return;
    rowSet[r] = true;
    colSet[c] = true;
    if (!totalD[r]) totalD[r] = {};
    if (!convD[r]) convD[r] = {};
    totalD[r][c] = (totalD[r][c] || 0) + 1;
    convD[r][c] = (convD[r][c] || 0) + (row[convFlagIdx] ? 1 : 0);
  });
  
  const rowCats = Object.keys(rowSet).sort();
  const colCats = Object.keys(colSet).sort();
  
  return { rowCats, colCats, totalData: totalD, convData: convD };
}


// ============================================================================
// メインデータローダー
// ============================================================================

function loadAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  if (!srcSheet) throw new Error('シート「' + CONFIG.SOURCE_SHEET_NAME + '」が見つかりません');
  
  const allData = srcSheet.getDataRange().getValues();
  if (allData.length < 2) throw new Error('データが空です');
  
  const headers = allData[0];
  
  // カラムインデックス特定
  const cols = {
    age:         findColIdx(headers, CONFIG.COL_KEYWORDS.age),
    income:      findColIdx(headers, CONFIG.COL_KEYWORDS.income),
    job:         findColIdx(headers, CONFIG.COL_KEYWORDS.job),
    credit:      findColIdx(headers, CONFIG.COL_KEYWORDS.credit),
    willingness: findColIdx(headers, CONFIG.COL_KEYWORDS.willingness),
    result:      findColIdx(headers, CONFIG.COL_KEYWORDS.result),
    execution:   findColIdx(headers, CONFIG.COL_KEYWORDS.execution),
    appDate:     findColIdx(headers, CONFIG.COL_KEYWORDS.appDate),
    staff:       findStaffIdx(headers),
    frontRoute:  findFrontRouteIdx(headers),
    seminarDate: findColIdx(headers, CONFIG.COL_KEYWORDS.seminarDate),
  };
  
  // 検証
  for (const key in cols) {
    if (cols[key] === -1) {
      Logger.log('WARNING: カラム「' + key + '」が見つかりませんでした');
    }
  }
  
  // データ行（ヘッダー除く）に派生カラムを追加
  const CONV_FLAG = headers.length;     // 成約フラグ
  const EXEC_FLAG = headers.length + 1; // 実施フラグ
  const APP_MONTH = headers.length + 2; // 申込月
  const ROUTE_CAT = headers.length + 3; // 流入経路大分類
  const WILL_SHORT = headers.length + 4; // 入会意欲短縮
  const SEM_LABEL = headers.length + 5;  // セミナー日ラベル
  
  const rows = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i].slice();
    
    // 成約フラグ
    const resultVal = strVal(row[cols.result]);
    row[CONV_FLAG] = CONFIG.CONVERSION_VALUES.indexOf(resultVal) !== -1 ? 1 : 0;
    
    // 実施フラグ
    const execVal = strVal(row[cols.execution]);
    row[EXEC_FLAG] = CONFIG.EXECUTION_VALUES.indexOf(execVal) !== -1 ? 1 : 0;
    
    // 申込月
    row[APP_MONTH] = toYearMonth(row[cols.appDate]);
    
    // 流入経路大分類
    row[ROUTE_CAT] = cols.frontRoute !== -1 ? categorizeRoute(row[cols.frontRoute]) : '';
    
    // 入会意欲短縮
    row[WILL_SHORT] = cols.willingness !== -1 ? shortenWillingness(row[cols.willingness]) : '';
    
    // セミナー日ラベル
    row[SEM_LABEL] = cols.seminarDate !== -1 ? formatSeminarDate(row[cols.seminarDate]) : '';
    
    rows.push(row);
  }
  
  return {
    ss, headers, rows, cols,
    idx: { CONV_FLAG, EXEC_FLAG, APP_MONTH, ROUTE_CAT, WILL_SHORT, SEM_LABEL },
  };
}


// ============================================================================
// シート1: 属性分布
// ============================================================================

function writeSheet1_AttributeDistribution(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.attr);
  
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレジットカード有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  let currentRow = 1;
  const totalN = rows.length;
  
  // タイトル
  const output = [['個別申込者データ 属性分布', '', '', ''], ['総数: ' + totalN + '件', '', '', ''], []];
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    output.push(['■ ' + attr.label, '', '', '']);
    output.push([attr.label, '件数', '割合', '']);
    
    const counts = countBy(rows, attr.colIdx);
    // 件数降順ソート
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    sorted.forEach(([cat, cnt]) => {
      output.push([cat, cnt, pct(cnt, totalN), '']);
    });
    output.push([]); // 空行
  });
  
  writeBlock(sheet, output, 1, 1);
  
  // ヘッダー行を太字に
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
}


// ============================================================================
// シート2: 属性 × 成約率
// ============================================================================

function writeSheet2_AttributeConversion(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.conv);
  
  const output = [['属性 × 成約率 クロス分析', '', '', '', '', '', '', ''], []];
  
  // --- Part A: 各属性 × 成約率 ---
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレカ有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    output.push(['■ ' + attr.label + ' × 成約率', '', '', '', '', '', '', '']);
    output.push([attr.label, '申込数', '実施数', '実施率', '成約数', '対申込成約率', '対実施成約率', '']);
    
    const results = conversionBy(rows, attr.colIdx, idx.CONV_FLAG, idx.EXEC_FLAG);
    results.sort((a, b) => {
      const rateA = a.total > 0 ? a.converted / a.total : 0;
      const rateB = b.total > 0 ? b.converted / b.total : 0;
      return rateB - rateA;
    });
    
    let sumT = 0, sumE = 0, sumC = 0;
    results.forEach(r => {
      output.push([
        r.category, r.total, r.executed, pct(r.executed, r.total),
        r.converted, pct(r.converted, r.total), pct(r.converted, r.executed), ''
      ]);
      sumT += r.total; sumE += r.executed; sumC += r.converted;
    });
    output.push(['合計', sumT, sumE, pct(sumE, sumT), sumC, pct(sumC, sumT), pct(sumC, sumE), '']);
    output.push([]);
  });
  
  // --- Part B: クロス分析（6パターン）---
  const crossPairs = [
    { label: '年齢 × 年収', rowIdx: cols.age, colIdx: cols.income },
    { label: '年齢 × クレカ有無', rowIdx: cols.age, colIdx: cols.credit },
    { label: '年齢 × 入会意欲', rowIdx: cols.age, colIdx: idx.WILL_SHORT },
    { label: '年収 × クレカ有無', rowIdx: cols.income, colIdx: cols.credit },
    { label: '年収 × 入会意欲', rowIdx: cols.income, colIdx: idx.WILL_SHORT },
    { label: 'クレカ有無 × 入会意欲', rowIdx: cols.credit, colIdx: idx.WILL_SHORT },
  ];
  
  crossPairs.forEach(cp => {
    if (cp.rowIdx === -1 || cp.colIdx === -1) return;
    
    const cr = crossConvRate(rows, cp.rowIdx, cp.colIdx, idx.CONV_FLAG);
    
    // 申込数テーブル
    output.push(['■ ' + cp.label + '【申込数】']);
    const headerRow = [''].concat(cr.colCats).concat(['合計']);
    output.push(headerRow);
    
    cr.rowCats.forEach(rc => {
      const row = [rc];
      let rowTotal = 0;
      cr.colCats.forEach(cc => {
        const v = (cr.totalData[rc] && cr.totalData[rc][cc]) || 0;
        row.push(v);
        rowTotal += v;
      });
      row.push(rowTotal);
      output.push(row);
    });
    output.push([]);
    
    // 成約率テーブル
    output.push(['■ ' + cp.label + '【成約率 %】']);
    output.push([''].concat(cr.colCats).concat(['合計']));
    
    cr.rowCats.forEach(rc => {
      const row = [rc];
      let rowTotalN = 0, rowTotalC = 0;
      cr.colCats.forEach(cc => {
        const n = (cr.totalData[rc] && cr.totalData[rc][cc]) || 0;
        const c = (cr.convData[rc] && cr.convData[rc][cc]) || 0;
        row.push(n > 0 ? pct(c, n) : '-');
        rowTotalN += n;
        rowTotalC += c;
      });
      row.push(pct(rowTotalC, rowTotalN));
      output.push(row);
    });
    output.push([]);
  });
  
  writeBlock(sheet, output, 1, 1);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
}


// ============================================================================
// シート3: 月別推移
// ============================================================================

function writeSheet3_MonthlyTrend(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.monthly);
  
  const output = [['月別推移分析', '', '', '', '', '', '', ''], []];
  
  // 月のリストを取得
  const monthSet = {};
  rows.forEach(r => { const m = r[idx.APP_MONTH]; if (m) monthSet[m] = true; });
  const months = Object.keys(monthSet).sort();
  
  // --- サマリー ---
  output.push(['■ 月別サマリー']);
  output.push(['月', '申込数', '実施数', '実施率', '成約数', '対申込成約率', '対実施成約率']);
  
  months.forEach(m => {
    const sub = rows.filter(r => r[idx.APP_MONTH] === m);
    const n = sub.length;
    const ex = sub.filter(r => r[idx.EXEC_FLAG]).length;
    const cv = sub.filter(r => r[idx.CONV_FLAG]).length;
    output.push([m, n, ex, pct(ex, n), cv, pct(cv, n), pct(cv, ex)]);
  });
  
  // 合計
  const n = rows.length;
  const ex = rows.filter(r => r[idx.EXEC_FLAG]).length;
  const cv = rows.filter(r => r[idx.CONV_FLAG]).length;
  output.push(['合計', n, ex, pct(ex, n), cv, pct(cv, n), pct(cv, ex)]);
  output.push([]);
  
  // --- 月別 × 各属性の成約率 ---
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレカ有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    
    // カテゴリ一覧
    const catSet = {};
    rows.forEach(r => { const v = strVal(r[attr.colIdx]); if (v) catSet[v] = true; });
    const categories = Object.keys(catSet).sort();
    
    output.push(['■ 月別 × ' + attr.label + '【成約率 %】']);
    output.push(['月'].concat(categories).concat(['全体']));
    
    months.forEach(m => {
      const sub = rows.filter(r => r[idx.APP_MONTH] === m);
      const row = [m];
      categories.forEach(cat => {
        const catSub = sub.filter(r => strVal(r[attr.colIdx]) === cat);
        const cn = catSub.length;
        const ccv = catSub.filter(r => r[idx.CONV_FLAG]).length;
        row.push(cn > 0 ? pct(ccv, cn) : '-');
      });
      const totalN = sub.length;
      const totalCv = sub.filter(r => r[idx.CONV_FLAG]).length;
      row.push(pct(totalCv, totalN));
      output.push(row);
    });
    
    // 合計行
    const totalRow = ['合計'];
    categories.forEach(cat => {
      const catAll = rows.filter(r => strVal(r[attr.colIdx]) === cat);
      const cn = catAll.length;
      const ccv = catAll.filter(r => r[idx.CONV_FLAG]).length;
      totalRow.push(cn > 0 ? pct(ccv, cn) : '-');
    });
    totalRow.push(pct(cv, n));
    output.push(totalRow);
    output.push([]);
    
    // 構成比
    output.push(['■ 月別 × ' + attr.label + '【構成比 %】']);
    output.push(['月'].concat(categories));
    
    months.forEach(m => {
      const sub = rows.filter(r => r[idx.APP_MONTH] === m && strVal(r[attr.colIdx]) !== '');
      const totalN = sub.length;
      const row = [m];
      categories.forEach(cat => {
        const cnt = sub.filter(r => strVal(r[attr.colIdx]) === cat).length;
        row.push(totalN > 0 ? pct(cnt, totalN) : '-');
      });
      output.push(row);
    });
    output.push([]);
  });
  
  writeBlock(sheet, output, 1, 1);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
}


// ============================================================================
// シート4: セミナー別
// ============================================================================

function writeSheet4_SeminarAnalysis(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.seminar);
  
  const output = [['セミナー別分析', '', '', '', '', '', '', ''], []];
  
  // セミナー日のリスト
  const semRows = rows.filter(r => r[idx.SEM_LABEL] !== '');
  const semSet = {};
  semRows.forEach(r => {
    const key = r[idx.SEM_LABEL];
    if (!semSet[key]) semSet[key] = [];
    semSet[key].push(r);
  });
  
  // セミナー日をソート（元のセミナー日時でソート）
  const semKeys = Object.keys(semSet).sort((a, b) => {
    // MM/DD形式なので、年をまたぐ場合を考慮
    // 10,11,12月 → 2025年、01,02月 → 2026年として比較
    const monthA = parseInt(a.substring(0, 2));
    const monthB = parseInt(b.substring(0, 2));
    const yearA = monthA >= 10 ? 2025 : 2026;
    const yearB = monthB >= 10 ? 2025 : 2026;
    if (yearA !== yearB) return yearA - yearB;
    return a.localeCompare(b);
  });
  
  // サマリー
  output.push(['■ セミナー別サマリー']);
  output.push(['セミナー日', '申込数', '実施数', '実施率', '成約数', '対申込成約率', '対実施成約率']);
  
  semKeys.forEach(key => {
    const sub = semSet[key];
    const n = sub.length;
    const ex = sub.filter(r => r[idx.EXEC_FLAG]).length;
    const cv = sub.filter(r => r[idx.CONV_FLAG]).length;
    output.push([key, n, ex, pct(ex, n), cv, pct(cv, n), pct(cv, ex)]);
  });
  
  const semN = semRows.length;
  const semEx = semRows.filter(r => r[idx.EXEC_FLAG]).length;
  const semCv = semRows.filter(r => r[idx.CONV_FLAG]).length;
  output.push(['合計', semN, semEx, pct(semEx, semN), semCv, pct(semCv, semN), pct(semCv, semEx)]);
  output.push([]);
  
  // セミナー別 × 各属性（構成比 + 成約率）
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレカ有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    
    const catSet = {};
    semRows.forEach(r => { const v = strVal(r[attr.colIdx]); if (v) catSet[v] = true; });
    const categories = Object.keys(catSet).sort();
    
    // 構成比
    output.push(['■ セミナー別 × ' + attr.label + '【構成比 %】']);
    output.push(['セミナー日', 'n'].concat(categories));
    
    semKeys.forEach(key => {
      const sub = semSet[key].filter(r => strVal(r[attr.colIdx]) !== '');
      const n = sub.length;
      if (n === 0) return;
      const row = [key, n];
      categories.forEach(cat => {
        const cnt = sub.filter(r => strVal(r[attr.colIdx]) === cat).length;
        row.push(pct(cnt, n));
      });
      output.push(row);
    });
    output.push([]);
    
    // 成約率
    output.push(['■ セミナー別 × ' + attr.label + '【成約率 %】']);
    output.push(['セミナー日', 'n'].concat(categories).concat(['全体']));
    
    semKeys.forEach(key => {
      const sub = semSet[key].filter(r => strVal(r[attr.colIdx]) !== '');
      const n = sub.length;
      if (n === 0) return;
      const row = [key, n];
      categories.forEach(cat => {
        const catSub = sub.filter(r => strVal(r[attr.colIdx]) === cat);
        const cn = catSub.length;
        const ccv = catSub.filter(r => r[idx.CONV_FLAG]).length;
        row.push(cn > 0 ? pct(ccv, cn) : '-');
      });
      const totalCv = sub.filter(r => r[idx.CONV_FLAG]).length;
      row.push(pct(totalCv, n));
      output.push(row);
    });
    output.push([]);
  });
  
  writeBlock(sheet, output, 1, 1);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
}


// ============================================================================
// シート5: 担当者別
// ============================================================================

function writeSheet5_StaffAnalysis(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.staff);
  
  if (cols.staff === -1) {
    writeBlock(sheet, [['担当者カラムが見つかりませんでした']], 1, 1);
    return;
  }
  
  const output = [['担当者別分析', '', '', '', '', '', '', '', ''], []];
  
  // 担当者リスト
  const staffSet = {};
  rows.forEach(r => {
    const s = strVal(r[cols.staff]);
    if (s) {
      if (!staffSet[s]) staffSet[s] = [];
      staffSet[s].push(r);
    }
  });
  
  // 全体の成約率（比較用）
  const allN = rows.length;
  const allCv = rows.filter(r => r[idx.CONV_FLAG]).length;
  const allCvRate = allN > 0 ? allCv / allN * 100 : 0;
  
  // --- サマリー ---
  output.push(['■ 担当者別サマリー（成約率降順）']);
  output.push(['担当者', '担当数', '実施数', '実施率', '成約数', '対担当成約率', '対実施成約率', '全体成約率', '差分']);
  
  const staffEntries = [];
  for (const staff in staffSet) {
    const sub = staffSet[staff];
    const n = sub.length;
    const ex = sub.filter(r => r[idx.EXEC_FLAG]).length;
    const cv = sub.filter(r => r[idx.CONV_FLAG]).length;
    const rate = n > 0 ? cv / n * 100 : 0;
    staffEntries.push({ staff, n, ex, cv, rate });
  }
  staffEntries.sort((a, b) => b.rate - a.rate);
  
  staffEntries.forEach(e => {
    const diff = e.rate - allCvRate;
    const diffStr = (diff >= 0 ? '+' : '') + diff.toFixed(1) + 'pt';
    output.push([
      e.staff, e.n, e.ex, pct(e.ex, e.n),
      e.cv, pct(e.cv, e.n), pct(e.cv, e.ex),
      pct(allCv, allN), diffStr
    ]);
  });
  output.push([]);
  
  // --- 担当者ごとの属性別成約率（上位担当者のみ: 10件以上）---
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレカ有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  const mainStaff = staffEntries.filter(e => e.n >= 10);
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    
    // 全体のカテゴリ一覧
    const catSet = {};
    rows.forEach(r => { const v = strVal(r[attr.colIdx]); if (v) catSet[v] = true; });
    const categories = Object.keys(catSet).sort();
    
    output.push(['■ 担当者別 × ' + attr.label + '【成約率 %】（担当10件以上）']);
    output.push(['担当者', 'n'].concat(categories).concat(['全体']));
    
    mainStaff.forEach(e => {
      const sub = staffSet[e.staff].filter(r => strVal(r[attr.colIdx]) !== '');
      const n = sub.length;
      const row = [e.staff, n];
      
      categories.forEach(cat => {
        const catSub = sub.filter(r => strVal(r[attr.colIdx]) === cat);
        const cn = catSub.length;
        const ccv = catSub.filter(r => r[idx.CONV_FLAG]).length;
        row.push(cn > 0 ? pct(ccv, cn) : '-');
      });
      
      const totalCv = sub.filter(r => r[idx.CONV_FLAG]).length;
      row.push(pct(totalCv, n));
      output.push(row);
    });
    
    // 全体平均
    const allValid = rows.filter(r => strVal(r[attr.colIdx]) !== '');
    const avgRow = ['【全体平均】', allValid.length];
    categories.forEach(cat => {
      const catAll = allValid.filter(r => strVal(r[attr.colIdx]) === cat);
      const cn = catAll.length;
      const ccv = catAll.filter(r => r[idx.CONV_FLAG]).length;
      avgRow.push(cn > 0 ? pct(ccv, cn) : '-');
    });
    avgRow.push(pct(allValid.filter(r => r[idx.CONV_FLAG]).length, allValid.length));
    output.push(avgRow);
    output.push([]);
  });
  
  writeBlock(sheet, output, 1, 1);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
}


// ============================================================================
// シート6: 流入経路別
// ============================================================================

function writeSheet6_RouteAnalysis(ctx) {
  const { ss, rows, cols, idx } = ctx;
  const sheet = getOrCreateSheet(ss, CONFIG.SHEETS.route);
  
  const output = [['フロント流入経路別分析', '', '', '', '', '', '', ''], []];
  
  // --- 大分類サマリー ---
  output.push(['■ 流入経路（大分類）× 成約率']);
  output.push(['流入経路', '件数', '実施数', '実施率', '成約数', '対申込成約率', '対実施成約率']);
  
  const routeResults = conversionBy(rows, idx.ROUTE_CAT, idx.CONV_FLAG, idx.EXEC_FLAG);
  routeResults.sort((a, b) => b.converted - a.converted);
  
  let sumT = 0, sumE = 0, sumC = 0;
  routeResults.forEach(r => {
    output.push([
      r.category, r.total, r.executed, pct(r.executed, r.total),
      r.converted, pct(r.converted, r.total), pct(r.converted, r.executed)
    ]);
    sumT += r.total; sumE += r.executed; sumC += r.converted;
  });
  output.push(['合計', sumT, sumE, pct(sumE, sumT), sumC, pct(sumC, sumT), pct(sumC, sumE)]);
  output.push([]);
  
  // --- チャネル大分類 ---
  function channelCategory(routeCat) {
    if (routeCat.indexOf('さきAI_YT') !== -1) return 'さきAI YouTube系';
    if (routeCat.indexOf('さきAI業務効率化') !== -1) return 'さきAI その他';
    if (routeCat.indexOf('たくむ') !== -1) return 'たくむ系';
    if (routeCat.indexOf('みさを') !== -1) return 'みさを系';
    if (routeCat.indexOf('えむ') !== -1) return 'えむ系';
    if (routeCat.indexOf('Meta') !== -1) return 'Meta広告系';
    if (routeCat === '不明') return '不明';
    return 'その他';
  }
  
  // チャネルフラグを一時的に追加
  const CH_IDX = rows[0] ? rows[0].length : 0;
  rows.forEach(r => { r[CH_IDX] = channelCategory(r[idx.ROUTE_CAT]); });
  
  output.push(['■ チャネル大分類 × 成約率']);
  output.push(['チャネル', '件数', '実施数', '実施率', '成約数', '対申込成約率', '対実施成約率']);
  
  const chResults = conversionBy(rows, CH_IDX, idx.CONV_FLAG, idx.EXEC_FLAG);
  chResults.sort((a, b) => b.converted - a.converted);
  
  chResults.forEach(r => {
    output.push([
      r.category, r.total, r.executed, pct(r.executed, r.total),
      r.converted, pct(r.converted, r.total), pct(r.converted, r.executed)
    ]);
  });
  output.push([]);
  
  // --- 流入経路 × 各属性 ---
  const attrList = [
    { label: '年齢', colIdx: cols.age },
    { label: '年収', colIdx: cols.income },
    { label: '職業', colIdx: cols.job },
    { label: 'クレカ有無', colIdx: cols.credit },
    { label: '入会意欲', colIdx: idx.WILL_SHORT },
  ];
  
  // n >= 8 の経路のみ
  const mainRoutes = routeResults.filter(r => r.total >= 8).map(r => r.category);
  
  attrList.forEach(attr => {
    if (attr.colIdx === -1) return;
    
    const catSet = {};
    rows.forEach(r => { const v = strVal(r[attr.colIdx]); if (v) catSet[v] = true; });
    const categories = Object.keys(catSet).sort();
    
    // 構成比
    output.push(['■ 流入経路 × ' + attr.label + '【構成比 %】（n≧8）']);
    output.push(['流入経路', 'n'].concat(categories));
    
    mainRoutes.forEach(route => {
      const sub = rows.filter(r => r[idx.ROUTE_CAT] === route && strVal(r[attr.colIdx]) !== '');
      const n = sub.length;
      if (n === 0) return;
      const row = [route, n];
      categories.forEach(cat => {
        const cnt = sub.filter(r => strVal(r[attr.colIdx]) === cat).length;
        row.push(pct(cnt, n));
      });
      output.push(row);
    });
    output.push([]);
    
    // 成約率
    output.push(['■ 流入経路 × ' + attr.label + '【成約率 %】（n≧8）']);
    output.push(['流入経路', 'n'].concat(categories).concat(['全体']));
    
    mainRoutes.forEach(route => {
      const sub = rows.filter(r => r[idx.ROUTE_CAT] === route && strVal(r[attr.colIdx]) !== '');
      const n = sub.length;
      if (n === 0) return;
      const row = [route, n];
      categories.forEach(cat => {
        const catSub = sub.filter(r => strVal(r[attr.colIdx]) === cat);
        const cn = catSub.length;
        const ccv = catSub.filter(r => r[idx.CONV_FLAG]).length;
        row.push(cn > 0 ? pct(ccv, cn) : '-');
      });
      const totalCv = sub.filter(r => r[idx.CONV_FLAG]).length;
      row.push(pct(totalCv, n));
      output.push(row);
    });
    output.push([]);
  });
  
  writeBlock(sheet, output, 1, 1);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
}


// ============================================================================
// メイン実行関数
// ============================================================================

/**
 * 全分析を実行（メニューまたはトリガーから呼び出し）
 */
function runAllAnalysis() {
  const startTime = new Date();
  Logger.log('分析開始: ' + startTime.toLocaleString());
  
  // データ読み込み
  const ctx = loadAllData();
  Logger.log('データ読み込み完了: ' + ctx.rows.length + '行');
  
  // 各シート書き出し
  writeSheet1_AttributeDistribution(ctx);
  Logger.log('シート1_属性分布 完了');
  
  writeSheet2_AttributeConversion(ctx);
  Logger.log('シート2_属性x成約率 完了');
  
  writeSheet3_MonthlyTrend(ctx);
  Logger.log('シート3_月別推移 完了');
  
  writeSheet4_SeminarAnalysis(ctx);
  Logger.log('シート4_セミナー別 完了');
  
  writeSheet5_StaffAnalysis(ctx);
  Logger.log('シート5_担当者別 完了');
  
  writeSheet6_RouteAnalysis(ctx);
  Logger.log('シート6_流入経路別 完了');
  
  const endTime = new Date();
  const elapsed = (endTime - startTime) / 1000;
  Logger.log('全分析完了: ' + elapsed.toFixed(1) + '秒');
  
  // 完了通知（UIがある場合）
  try {
    SpreadsheetApp.getUi().alert('分析完了！（' + elapsed.toFixed(1) + '秒）\n\n6つの分析シートが更新されました。');
  } catch (e) {
    // トリガー実行時はUIがないためエラーを無視
  }
}


// ============================================================================
// トリガー設定
// ============================================================================

/**
 * 毎日午前9時に自動実行するトリガーを設定
 * ※ この関数を一度だけ手動実行してください
 */
function setupDailyTrigger() {
  // 既存のトリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAllAnalysis') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新規トリガー作成
  ScriptApp.newTrigger('runAllAnalysis')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  Logger.log('日次トリガーを設定しました（毎日午前9〜10時に実行）');
  
  try {
    SpreadsheetApp.getUi().alert('日次トリガーを設定しました！\n\n毎日午前9〜10時に自動で分析が更新されます。');
  } catch (e) {
    // UIがない場合
  }
}


// ============================================================================
// カスタムメニュー（スプレッドシートを開いた時に自動追加）
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 データ分析')
    .addItem('今すぐ分析を実行', 'runAllAnalysis')
    .addItem('毎日自動実行を設定', 'setupDailyTrigger')
    .addToUi();
}
