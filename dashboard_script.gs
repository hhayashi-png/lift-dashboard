// =====================================================
// 📊 LIFT 売上ダッシュボード 自動生成スクリプト v2.0
// 使い方:
//   1. Googleスプレッドシートを開く
//   2. 拡張機能 > Apps Script
//   3. このコードを全て貼り付けて保存
//   4. createDashboard() を選択して ▷ 実行
// =====================================================

// ========== 設定（必要に応じて変更） ==========
var CFG = {
  srcName:  '売上管理',          // 元データシート名（部分一致でOK）
  dashName: '📊ダッシュボード',   // ダッシュボードシート名
  maxScanRow: 10,               // ヘッダー検索上限行
  dataStartRow: 6,              // データ開始行（通常は変更不要）
};

// カラーパレット
var C = {
  // ベース
  pageBg:    '#F5F2FA',
  white:     '#FFFFFF',
  border:    '#DDD5EC',
  text:      '#1C1028',
  text2:     '#7A6880',
  // ヘッダー
  header:    '#1A1040',
  headerTxt: '#FFFFFF',
  // SV行
  svBg:      '#2D1B6E',
  svTxt:     '#FFFFFF',
  svSubBg:   '#3D2880',
  // 達成率カラー
  good:      '#E8F5E9', goodTxt: '#1B5E20', goodDark: '#388E3C',
  warn:      '#FFF8E1', warnTxt: '#E65100', warnDark: '#F57C00',
  bad:       '#FFEBEE', badTxt:  '#B71C1C', badDark:  '#D32F2F',
  // タイプ
  direct:    '#EEF2FF',
  franchise: '#FFF9F0',
  // アクセント
  gold:      '#F9A825',
  accent:    '#7B3FC4',
  accentLt:  '#EDE7FF',
};

// ========== ユーティリティ ==========

function yen(v) {
  if (!v || isNaN(v)) return '¥0';
  return '¥' + Math.round(Number(v)).toLocaleString();
}

function pct(v) {
  if (v === null || v === undefined || v === '') return '0%';
  var n = parseFloat(String(v).replace('%', ''));
  if (isNaN(n)) return '0%';
  if (n <= 1 && n > 0) n = Math.round(n * 100); // 0.5 → 50%
  else n = Math.round(n);
  return n + '%';
}

function toNum(v) {
  if (!v && v !== 0) return 0;
  return parseFloat(String(v).replace(/[¥,\s%]/g, '')) || 0;
}

function achNum(v) {
  var n = toNum(v);
  if (n > 0 && n <= 1) return Math.round(n * 100);
  return Math.round(n);
}

function miniBar(pctVal, width) {
  width = width || 12;
  var n = Math.min(Math.round(pctVal / 100 * width), width);
  return '█'.repeat(n) + '░'.repeat(width - n);
}

function achColor(achPct) {
  if (achPct >= 100) return { bg: C.good, txt: C.goodTxt, bar: C.goodDark };
  if (achPct >= 60)  return { bg: C.warn, txt: C.warnTxt, bar: C.warnDark };
  return               { bg: C.bad,  txt: C.badTxt,  bar: C.badDark  };
}

// ========== シート検索 ==========

function findSheet(ss, name) {
  var sheets = ss.getSheets();
  // 完全一致
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === name) return sheets[i];
  }
  // 部分一致
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(name) !== -1) return sheets[i];
  }
  return null;
}

// ========== カラム検出 ==========

function detectColumns(data) {
  var cols = {
    sv: 0, type: 1, store: 2,
    target: -1, actual: -1, ach: -1, landing: -1,
    headerRow: -1
  };

  for (var r = 0; r < Math.min(CFG.maxScanRow, data.length); r++) {
    var row = data[r];
    var found = false;
    for (var c = 0; c < row.length; c++) {
      var v = String(row[c] || '').trim();
      if (v === 'SV' || v === 'SV名') { cols.sv = c; found = true; }
      if (v === '直or加' || v === '直or加盟' || v === '種別') cols.type = c;
      if (v === '店舗名' || v === '店舗') { cols.store = c; found = true; }
      // 目標: 最初に見つかったものをターゲットとする
      if ((v === '目標' || v === '月目標') && cols.target === -1) cols.target = c;
      // 実績/進捗
      if ((v === '進捗' || v === '実績' || v === '売上') && cols.actual === -1) cols.actual = c;
      // 達成率
      if ((v === '達成率' || v === '達成%') && cols.ach === -1) cols.ach = c;
      // 着地見込
      if ((v === '着地' || v === '着地見込' || v === '着地予測') && cols.landing === -1) cols.landing = c;
    }
    if (found && cols.target > -1) {
      cols.headerRow = r;
      break;
    }
  }

  // フォールバック（ヘッダーが見つからない場合）
  if (cols.target === -1) cols.target = 3;
  if (cols.actual === -1) cols.actual = cols.target + 1;
  if (cols.ach    === -1) cols.ach    = cols.actual + 1;
  if (cols.landing === -1) cols.landing = cols.ach + 1;

  Logger.log('Detected cols: ' + JSON.stringify(cols));
  return cols;
}

// ========== メタデータ取得（経過日数など） ==========

function parseMeta(data) {
  var meta = {
    yyyymm: '',
    elapsedDays: 0,
    totalDays: 0,
    elapsedPct: 0,
  };

  for (var r = 0; r < Math.min(5, data.length); r++) {
    var row = data[r];
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      var vStr = String(v || '');
      // 202603 形式の年月
      if (/^2026\d{2}$/.test(vStr) || /^2025\d{2}$/.test(vStr)) {
        meta.yyyymm = vStr;
      }
      // 経過日数（整数 1-31）
      if (typeof v === 'number' && v >= 1 && v <= 31 && meta.elapsedDays === 0) {
        meta.elapsedDays = v;
      }
      // 月間日数（28-31）
      if (typeof v === 'number' && v >= 28 && v <= 31 && meta.totalDays === 0) {
        meta.totalDays = v;
      }
      // 経過率 % (0.5 = 50% or 57 = 57)
      if ((typeof v === 'number' && v > 0 && v <= 1) || vStr.match(/^\d+%$/)) {
        if (meta.elapsedPct === 0) {
          meta.elapsedPct = typeof v === 'number' ? Math.round(v * 100) : parseInt(vStr);
        }
      }
    }
  }

  // フォールバック
  if (!meta.yyyymm) {
    var now = new Date();
    meta.yyyymm = String(now.getFullYear()) + String(now.getMonth() + 1).padStart(2, '0');
  }
  if (meta.totalDays === 0) {
    var y = parseInt(meta.yyyymm.slice(0, 4));
    var m = parseInt(meta.yyyymm.slice(4, 6));
    meta.totalDays = new Date(y, m, 0).getDate();
  }
  if (meta.elapsedDays === 0) {
    meta.elapsedDays = new Date().getDate();
  }
  if (meta.elapsedPct === 0 && meta.totalDays > 0) {
    meta.elapsedPct = Math.round(meta.elapsedDays / meta.totalDays * 100);
  }

  Logger.log('Meta: ' + JSON.stringify(meta));
  return meta;
}

// ========== 店舗データ解析 ==========

function parseStores(data, cols) {
  var startRow = Math.max(cols.headerRow + 1, CFG.dataStartRow - 1); // 0-indexed
  var stores = [];
  var lastSV = '';

  for (var r = startRow; r < data.length; r++) {
    var row = data[r];
    var sv    = String(row[cols.sv]    || '').trim();
    var type  = String(row[cols.type]  || '').trim();
    var store = String(row[cols.store] || '').trim();

    // SVの引き継ぎ（空白行でも前のSVを使う）
    if (sv) lastSV = sv;
    else sv = lastSV;

    if (!store) continue;  // 店舗名が空の行はスキップ

    var target  = toNum(row[cols.target]);
    var actual  = toNum(row[cols.actual]);
    var achRaw  = row[cols.ach];
    var landing = cols.landing >= 0 ? toNum(row[cols.landing]) : 0;

    // 達成率計算
    var achPct;
    if (achRaw !== null && achRaw !== undefined && achRaw !== '') {
      achPct = achNum(achRaw);
    } else if (target > 0) {
      achPct = Math.round(actual / target * 100);
    } else {
      achPct = 0;
    }

    stores.push({
      sv:      sv,
      type:    type,
      store:   store,
      target:  target,
      actual:  actual,
      achPct:  achPct,
      landing: landing,
    });
  }

  return stores;
}

// ========== SV別グループ化 ==========

function groupBySV(stores) {
  var map = {};
  var order = [];

  stores.forEach(function(s) {
    if (!map[s.sv]) {
      map[s.sv] = { sv: s.sv, stores: [], totTarget: 0, totActual: 0 };
      order.push(s.sv);
    }
    map[s.sv].stores.push(s);
    map[s.sv].totTarget += s.target;
    map[s.sv].totActual += s.actual;
  });

  // 達成率計算
  order.forEach(function(sv) {
    var g = map[sv];
    g.achPct = g.totTarget > 0 ? Math.round(g.totActual / g.totTarget * 100) : 0;
    g.storeCount = g.stores.length;
  });

  return order.map(function(sv) { return map[sv]; });
}

// ========== セルスタイル ヘルパー ==========

function styleRange(sheet, r1, c1, r2, c2, opts) {
  var range = sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  if (opts.bg)     range.setBackground(opts.bg);
  if (opts.txt)    range.setFontColor(opts.txt);
  if (opts.bold)   range.setFontWeight('bold');
  if (opts.size)   range.setFontSize(opts.size);
  if (opts.align)  range.setHorizontalAlignment(opts.align);
  if (opts.valign) range.setVerticalAlignment(opts.valign);
  if (opts.wrap)   range.setWrap(opts.wrap);
  if (opts.italic) range.setFontStyle('italic');
  if (opts.borders) {
    var b = range.getBorder ? null : null;
    if (opts.borders === 'all') {
      range.setBorder(true, true, true, true, false, false,
        C.border, SpreadsheetApp.BorderStyle.SOLID);
    }
    if (opts.borders === 'bottom') {
      range.setBorder(false, false, true, false, false, false,
        C.border, SpreadsheetApp.BorderStyle.SOLID);
    }
    if (opts.borders === 'outer') {
      range.setBorder(true, true, true, true, false, false,
        C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  return range;
}

function setCell(sheet, r, c, val, opts) {
  var cell = sheet.getRange(r, c);
  cell.setValue(val);
  if (opts) styleRange(sheet, r, c, r, c, opts);
  return cell;
}

function mergeCells(sheet, r1, c1, r2, c2) {
  sheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1).merge();
}

// ========== ダッシュボード描画 ==========

// 列構成（1-indexed）
var COL = {
  pad:   1,   // A: 余白
  name:  2,   // B: SV/店舗名
  type:  3,   // C: 直or加
  tgt:   4,   // D: 目標
  act:   5,   // E: 実績
  ach:   6,   // F: 達成率
  bar:   7,   // G: 進捗バー
  land:  8,   // H: 着地見込
  stat:  9,   // I: 状態
  END:   9,
};

// --- タイトルセクション ---
function drawTitle(sheet, meta, r) {
  var yr = meta.yyyymm.slice(0, 4);
  var mo = meta.yyyymm.slice(4, 6);

  // 行1: タイトルバー
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue('📊  LIFT 売上ダッシュボード')
    .setBackground(C.header)
    .setFontColor(C.headerTxt)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 52);
  r++;

  // 行2: サブヘッダー（年月 + 経過情報）
  var subLabel = yr + '年' + parseInt(mo) + '月  ｜  '
    + meta.elapsedDays + '日経過 / ' + meta.totalDays + '日  ｜  '
    + '月次経過率 ' + meta.elapsedPct + '%';
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue(subLabel)
    .setBackground('#2D1B6E')
    .setFontColor('#CCB8FF')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 30);
  r++;

  return r;
}

// --- 全体サマリーセクション ---
function drawOverallSummary(sheet, stores, meta, r) {
  // セクションヘッダー
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue('  全体進捗サマリー')
    .setBackground(C.accentLt)
    .setFontColor(C.accent)
    .setFontSize(10)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 26);
  r++;

  // 合計計算
  var totTarget = 0, totActual = 0, totLanding = 0, achCount = 0;
  stores.forEach(function(s) {
    totTarget  += s.target;
    totActual  += s.actual;
    totLanding += s.landing;
    if (s.achPct >= 100) achCount++;
  });
  var totAchPct = totTarget > 0 ? Math.round(totActual / totTarget * 100) : 0;
  var achColors = achColor(totAchPct);

  // KPI列ヘッダー
  var headers = ['合計目標', '合計実績', '達成率', '着地見込', '目標達成店舗'];
  var vals    = [yen(totTarget), yen(totActual), totAchPct + '%', yen(totLanding), achCount + ' / ' + stores.length + '店舗'];

  // KPIカード行（ヘッダー）
  var kpiCols  = [[1,2], [3,4], [5,6], [7,8], [9,9]];  // [start, end] mergeの列範囲
  // 実際の列は COL.name〜COL.stat を5分割
  // 簡単に2列ずつに割り当て
  var starts = [2, 3, 4, 5, 6, 7, 8, 9];

  // KPIラベル行
  var kpiData = [
    { label: '合計目標',    val: yen(totTarget),   col: 2 },
    { label: '合計実績',    val: yen(totActual),   col: 3 },
    { label: '達成率',      val: totAchPct + '%',  col: 4 },
    { label: '着地見込',    val: yen(totLanding),  col: 5 },
    { label: '目標達成',    val: achCount + '/' + stores.length, col: 6 },
    { label: '対経過率',    val: (totAchPct > 0 && meta.elapsedPct > 0 ? (totAchPct >= meta.elapsedPct ? '▲ペース良好' : '▼ペース遅れ') : '--'), col: 7 },
  ];

  // KPIラベル行
  sheet.getRange(r, 1, 1, 1).setValue('').setBackground(C.pageBg);
  kpiData.forEach(function(kpi) {
    sheet.getRange(r, kpi.col).setValue(kpi.label)
      .setBackground('#EDE7FF')
      .setFontColor(C.accent)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  // 空きセル
  sheet.getRange(r, 8).setBackground('#EDE7FF');
  sheet.getRange(r, 9).setBackground('#EDE7FF');
  sheet.setRowHeight(r, 20);
  r++;

  // KPI値行
  sheet.getRange(r, 1, 1, 1).setValue('').setBackground(C.pageBg);
  kpiData.forEach(function(kpi) {
    var isAch = kpi.label === '達成率';
    var cellBg = isAch ? achColors.bg : C.white;
    var cellTxt = isAch ? achColors.txt : C.text;
    sheet.getRange(r, kpi.col).setValue(kpi.val)
      .setBackground(cellBg)
      .setFontColor(cellTxt)
      .setFontSize(13)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(false, false, true, false, false, false, achColors.bar, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange(r, 8).setBackground(C.white).setValue('').setBorder(false, false, true, false, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(r, 9).setBackground(C.white).setValue('').setBorder(false, false, true, false, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(r, 38);
  r++;

  // 進捗バー行
  var barFilled = Math.min(Math.round(totAchPct / 100 * 40), 40);
  var barStr = '█'.repeat(barFilled) + '░'.repeat(40 - barFilled)
    + '  ' + totAchPct + '%';
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue(barStr)
    .setBackground(C.white)
    .setFontColor(achColors.bar)
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Courier New');
  sheet.setRowHeight(r, 24);
  r++;

  // 経過率バー
  var ePct = meta.elapsedPct;
  var eBarFilled = Math.min(Math.round(ePct / 100 * 40), 40);
  var eBarStr = '▒'.repeat(eBarFilled) + '·'.repeat(40 - eBarFilled)
    + '  経過 ' + ePct + '%（' + meta.elapsedDays + '/' + meta.totalDays + '日）';
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue(eBarStr)
    .setBackground(C.white)
    .setFontColor(C.text2)
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily('Courier New');
  sheet.setRowHeight(r, 22);
  r++;

  // 外枠
  sheet.getRange(r - 5, 1, 5, COL.END)
    .setBorder(true, true, true, true, false, false, C.accent, SpreadsheetApp.BorderStyle.SOLID);

  // 区切り行
  sheet.getRange(r, 1, 1, COL.END).setBackground(C.pageBg);
  sheet.setRowHeight(r, 10);
  r++;

  return r;
}

// --- SVランキングセクション ---
function drawSVRanking(sheet, svGroups, r) {
  // ヘッダー
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue('  SV別ランキング')
    .setBackground(C.accentLt)
    .setFontColor(C.accent)
    .setFontSize(10)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 26);
  r++;

  // 表ヘッダー行
  var rankHeaders = ['#', 'SV名', '店舗数', '合計目標', '合計実績', '達成率', '進捗', '着地見込', '状態'];
  rankHeaders.forEach(function(h, i) {
    sheet.getRange(r, i + 1).setValue(h)
      .setBackground(C.header)
      .setFontColor(C.headerTxt)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeight(r, 24);
  r++;

  // SVを達成率順にソート
  var sorted = svGroups.slice().sort(function(a, b) { return b.achPct - a.achPct; });
  var rankMedals = ['🥇', '🥈', '🥉'];

  sorted.forEach(function(g, i) {
    var col = achColor(g.achPct);
    var rank = i < 3 ? rankMedals[i] : (i + 1);
    var landing = g.stores.reduce(function(s, st) { return s + st.landing; }, 0);
    var bar = miniBar(Math.min(g.achPct, 100), 8);
    var status = g.achPct >= 100 ? '🟢' : g.achPct >= 60 ? '🟡' : '🔴';

    var rowData = [rank, g.sv, g.storeCount + '店', yen(g.totTarget), yen(g.totActual), g.achPct + '%', bar, yen(landing), status];
    rowData.forEach(function(v, j) {
      var cell = sheet.getRange(r, j + 1);
      cell.setValue(v).setBackground(col.bg).setFontColor(col.txt)
        .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
      if (j === 1) cell.setHorizontalAlignment('left').setFontWeight('bold');
      if (j === 6) cell.setFontFamily('Courier New').setFontSize(9);
    });
    // 外枠
    sheet.getRange(r, 1, 1, COL.END)
      .setBorder(false, false, true, false, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // 外枠
  sheet.getRange(r - sorted.length - 1, 1, sorted.length + 1, COL.END)
    .setBorder(true, true, true, true, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 区切り
  sheet.getRange(r, 1, 1, COL.END).setBackground(C.pageBg);
  sheet.setRowHeight(r, 14);
  r++;

  return r;
}

// --- SV詳細セクション ---
function drawSVSection(sheet, group, meta, r) {
  var sv = group.sv;
  var stores = group.stores;
  var totTarget = group.totTarget;
  var totActual = group.totActual;
  var achPct    = group.achPct;
  var col       = achColor(achPct);
  var directCnt = stores.filter(function(s) { return s.type.indexOf('直') !== -1; }).length;
  var franchiseCnt = stores.length - directCnt;

  // SV名ヘッダー行
  var svLabel = sv + '  ｜  '
    + (directCnt > 0 ? '直営' + directCnt : '')
    + (directCnt > 0 && franchiseCnt > 0 ? '・' : '')
    + (franchiseCnt > 0 ? '加盟' + franchiseCnt : '')
    + '  計' + stores.length + '店舗';
  var svRight = '合計  ' + yen(totActual) + '  /  ' + yen(totTarget) + '  （' + achPct + '%）';

  sheet.getRange(r, COL.pad, 1, COL.name - COL.pad + 1).merge()
    .setValue('  ' + svLabel)
    .setBackground(C.svBg)
    .setFontColor(C.svTxt)
    .setFontSize(11)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.getRange(r, COL.type, 1, COL.END - COL.type + 1).merge()
    .setValue(svRight + '  ')
    .setBackground(C.svBg)
    .setFontColor('#C8B4FF')
    .setFontSize(10)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 30);
  r++;

  // 店舗ヘッダー行
  var colHdrs = ['', '店舗名', '種別', '目標', '実績', '達成率', '進捗バー', '着地見込', ''];
  colHdrs.forEach(function(h, i) {
    sheet.getRange(r, i + 1).setValue(h)
      .setBackground('#4A2D8E')
      .setFontColor('#CCB8FF')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeight(r, 20);
  r++;

  // 店舗行
  stores.forEach(function(s, i) {
    var isDirect   = s.type.indexOf('直') !== -1;
    var rowBg      = isDirect ? C.direct : C.franchise;
    var achCol     = achColor(s.achPct);
    var bar        = miniBar(Math.min(s.achPct, 100), 10);
    var statusIcon = s.achPct >= 100 ? '✅' : s.achPct >= 60 ? '⚠️' : '🔴';

    // 行番号（スキップ列）
    sheet.getRange(r, COL.pad).setValue(i + 1)
      .setBackground(rowBg).setFontColor(C.text2).setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 店舗名
    sheet.getRange(r, COL.name).setValue(s.store)
      .setBackground(rowBg).setFontColor(C.text)
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    // 直or加
    sheet.getRange(r, COL.type).setValue(isDirect ? '直営' : '加盟')
      .setBackground(isDirect ? '#D5CCFF' : '#FFE8CC')
      .setFontColor(isDirect ? '#3730A3' : '#92400E')
      .setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 目標
    sheet.getRange(r, COL.tgt).setValue(yen(s.target))
      .setBackground(rowBg).setFontColor(C.text2)
      .setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
    // 実績
    sheet.getRange(r, COL.act).setValue(yen(s.actual))
      .setBackground(rowBg).setFontColor(C.text)
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
    // 達成率
    sheet.getRange(r, COL.ach).setValue(s.achPct + '%')
      .setBackground(achCol.bg).setFontColor(achCol.txt)
      .setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 進捗バー
    sheet.getRange(r, COL.bar).setValue(bar)
      .setBackground(rowBg).setFontColor(achCol.bar)
      .setFontFamily('Courier New').setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // 着地見込
    sheet.getRange(r, COL.land).setValue(s.landing > 0 ? yen(s.landing) : '-')
      .setBackground(rowBg).setFontColor(C.text2)
      .setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
    // ステータスアイコン
    sheet.getRange(r, COL.stat).setValue(statusIcon)
      .setBackground(rowBg).setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    // 行区切り
    sheet.getRange(r, 1, 1, COL.END)
      .setBorder(false, false, true, false, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID_LIGHT);
    sheet.setRowHeight(r, 26);
    r++;
  });

  // 小計行
  var subLanding = stores.reduce(function(s, st) { return s + st.landing; }, 0);
  var subAchPct  = totTarget > 0 ? Math.round(totActual / totTarget * 100) : 0;
  var subBar     = miniBar(Math.min(subAchPct, 100), 10);
  var subCol     = achColor(subAchPct);

  sheet.getRange(r, COL.pad).setBackground('#EDE7FF').setValue('');
  sheet.getRange(r, COL.name).setValue('  小計').setBackground('#EDE7FF').setFontColor(C.accent).setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.getRange(r, COL.type).setBackground('#EDE7FF').setValue('');
  sheet.getRange(r, COL.tgt).setValue(yen(totTarget)).setBackground('#EDE7FF').setFontColor(C.text2).setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle').setFontWeight('bold');
  sheet.getRange(r, COL.act).setValue(yen(totActual)).setBackground('#EDE7FF').setFontColor(C.accent).setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle').setFontWeight('bold');
  sheet.getRange(r, COL.ach).setValue(subAchPct + '%').setBackground(subCol.bg).setFontColor(subCol.txt).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(r, COL.bar).setValue(subBar).setBackground('#EDE7FF').setFontColor(subCol.bar).setFontFamily('Courier New').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(r, COL.land).setValue(subLanding > 0 ? yen(subLanding) : '-').setBackground('#EDE7FF').setFontColor(C.text2).setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle').setFontWeight('bold');
  sheet.getRange(r, COL.stat).setBackground('#EDE7FF').setValue('');
  sheet.getRange(r, 1, 1, COL.END)
    .setBorder(true, true, true, true, false, false, C.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(r, 26);
  r++;

  // セクション全体の外枠
  sheet.getRange(r - stores.length - 2, 1, stores.length + 2, COL.END)
    .setBorder(true, true, true, true, false, false, C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 区切り行
  sheet.getRange(r, 1, 1, COL.END).setBackground(C.pageBg);
  sheet.setRowHeight(r, 12);
  r++;

  return r;
}

// --- フッター ---
function drawFooter(sheet, r) {
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sheet.getRange(r, 1, 1, COL.END).merge()
    .setValue('最終更新: ' + now + '  ｜  毎時自動更新  ｜  📊 LIFT Sales Dashboard')
    .setBackground(C.header)
    .setFontColor('#6B5C8A')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(r, 24);
}

// --- 列幅 ---
function setColWidths(sheet) {
  sheet.setColumnWidth(COL.pad,  30);   // A
  sheet.setColumnWidth(COL.name, 195);  // B
  sheet.setColumnWidth(COL.type, 50);   // C
  sheet.setColumnWidth(COL.tgt,  115);  // D
  sheet.setColumnWidth(COL.act,  115);  // E
  sheet.setColumnWidth(COL.ach,  68);   // F
  sheet.setColumnWidth(COL.bar,  105);  // G
  sheet.setColumnWidth(COL.land, 115);  // H
  sheet.setColumnWidth(COL.stat, 40);   // I
}

// ========== 自動更新トリガー ==========

function setupTrigger_() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'createDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 1時間おきに自動更新
  ScriptApp.newTrigger('createDashboard')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('毎時自動更新トリガーを設定しました');
}

// ========== シートのカラム設定（GID確認用）==========

function checkSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = ss.getSheets().map(function(s) {
    return s.getName() + ' (GID: ' + s.getSheetId() + ')';
  }).join('\n');
  SpreadsheetApp.getUi().alert('シート一覧:\n' + info);
}
