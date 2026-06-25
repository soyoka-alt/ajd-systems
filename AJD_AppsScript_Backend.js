// ============================================================
// ALL JAPAN DMC — 報価システム バックエンド
// Google Apps Script — ウェブアプリとしてデプロイ
//
// デプロイ手順：
// 1. Google Sheetsで新規スプレッドシートを作成
// 2. メニュー → 拡張機能 → Apps Script
// 3. このコードを全部貼り付けて保存（Ctrl+S）
// 4. デプロイ → 新しいデプロイ → 種類：ウェブアプリ
//    実行ユーザー：自分 / アクセス：自分と同じドメインの全員
// 5. デプロイURLをHTMLの設定ページに貼り付ける
// ============================================================

const SHEET_QUOTES = '报价记录';
const SHEET_ITIN   = '行程记录';
const SHEET_STATS  = '统计汇总';

const QUOTE_HEADERS = [
  'ID','保存时间','操作人','客户名称','旅行类型','目的地',
  '出发日期','天数','人数','方案名','备注',
  '货币','净成本','报价总价','人均报价','Markup%','预计利润',
  '状态','费用明细JSON','行程JSON','完整数据JSON'
];

// ── CORS headers helper ──
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function makeResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── POST handler ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;
    if      (action === 'saveQuote')    result = saveQuote(data);
    else if (action === 'getQuotes')    result = getQuotes(data);
    else if (action === 'updateStatus') result = updateStatus(data);
    else if (action === 'deleteQuote')  result = deleteQuote(data);
    else if (action === 'getStats')     result = getStats();
    else result = { success: false, error: 'Unknown action: ' + action };
    return makeResponse(result);
  } catch(err) {
    return makeResponse({ success: false, error: err.message });
  }
}

// ── GET handler (for testConn and getQuotes) ──
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'ping';
    const callback = e.parameter && e.parameter.callback;
    let result;
    if      (action === 'getQuotes') result = getQuotes(e.parameter);
    else if (action === 'getStats')  result = getStats();
    else if (action === 'ping')      result = { success: true, message: 'AJD OK' };
    else result = { success: false, error: 'Unknown action' };
    // Support JSONP for file:// origins (Chrome CORS bypass)
    if(callback){
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return makeResponse(result);
  } catch(err) {
    const callback = e.parameter && e.parameter.callback;
    const errObj = { success: false, error: err.message };
    if(callback){
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(errObj) + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return makeResponse(errObj);
  }
}

// ── シート初期化 ──
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let qSheet = ss.getSheetByName(SHEET_QUOTES);
  if (!qSheet) {
    qSheet = ss.insertSheet(SHEET_QUOTES);
    const hdr = qSheet.getRange(1, 1, 1, QUOTE_HEADERS.length);
    hdr.setValues([QUOTE_HEADERS]);
    hdr.setBackground('#1a2744').setFontColor('#ffffff').setFontWeight('bold');
    qSheet.setFrozenRows(1);
    [180,140,120,180,120,160,100,60,60,120,160,60,80,100,80,70,80,70,60,60,60].forEach((w,i) => {
      qSheet.setColumnWidth(i+1, w);
    });
  }

  let iSheet = ss.getSheetByName(SHEET_ITIN);
  if (!iSheet) {
    iSheet = ss.insertSheet(SHEET_ITIN);
    const ITIN_H = ['ID','报价ID','客户名称','方案名','Day','日期','标题','时间','类型','内容','备注'];
    iSheet.getRange(1,1,1,ITIN_H.length).setValues([ITIN_H])
      .setBackground('#1a2744').setFontColor('#ffffff').setFontWeight('bold');
    iSheet.setFrozenRows(1);
  }

  let stSheet = ss.getSheetByName(SHEET_STATS);
  if (!stSheet) stSheet = ss.insertSheet(SHEET_STATS);

  return { qSheet, iSheet, stSheet };
}

// ── 報価保存 ──
function saveQuote(data) {
  const { qSheet, iSheet } = initSheets();
  const opt = data.opt || {};
  const operator = data.operator || '不明';
  const now = new Date();
  const pax = opt.pax || 1;
  const mk  = opt.markup || 0;

  let net = 0;
  (opt.cats || []).forEach(cat => {
    (cat.items || []).forEach(item => {
      net += cat.pp ? (item.c||0) * (item.q||1) * pax : (item.c||0) * (item.q||1);
    });
  });
  const quote  = net * (1 + mk / 100);
  const profit = quote - net;

  const quoteId = data.quoteId || Utilities.getUuid().slice(0,8).toUpperCase();
  const existing = findRowById(qSheet, quoteId);

  const row = [
    quoteId, now, operator,
    opt.clientName || '', opt.tourType || '', opt.dest || '',
    opt.startDate || '', opt.days || 0, opt.pax || 0,
    opt.name || 'Option A', opt.optNote || '',
    data.currency || 'JPY',
    Math.round(net), Math.round(quote), Math.round(quote / pax),
    mk, Math.round(profit),
    data.status || '草稿',
    JSON.stringify(opt.cats || []),
    JSON.stringify(opt.itin || []),
    JSON.stringify(data.allOpts || [opt])
  ];

  if (existing > 0) {
    qSheet.getRange(existing, 1, 1, row.length).setValues([row]);
    colorRow(qSheet, existing, data.status || '草稿');
  } else {
    qSheet.appendRow(row);
    colorRow(qSheet, qSheet.getLastRow(), data.status || '草稿');
  }

  writeItinRows(iSheet, quoteId, opt);
  updateStats(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STATS));

  return { success: true, id: quoteId };
}

// ── 報価一覧取得 ──
function getQuotes(params) {
  const { qSheet } = initSheets();
  const data = qSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, quotes: [] };

  const headers = data[0];
  let quotes = data.slice(1)
    .filter(r => r[0])
    .map(row => {
      const q = {};
      headers.forEach((h, i) => { q[h] = row[i]; });
      if (q['保存时间'] instanceof Date) {
        q['保存时间'] = Utilities.formatDate(q['保存时间'], Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
      }
      return q;
    });

  if (params && params.operator) quotes = quotes.filter(q => q['操作人'] === params.operator);
  if (params && params.status)   quotes = quotes.filter(q => q['状态'] === params.status);

  return { success: true, quotes: quotes.reverse() };
}

// ── ステータス更新 ──
function updateStatus(data) {
  const { qSheet } = initSheets();
  const rowIdx = findRowById(qSheet, data.id);
  if (rowIdx < 0) return { success: false, error: 'Not found' };
  const statusCol = QUOTE_HEADERS.indexOf('状态') + 1;
  qSheet.getRange(rowIdx, statusCol).setValue(data.status);
  colorRow(qSheet, rowIdx, data.status);
  return { success: true };
}

// ── 削除 ──
function deleteQuote(data) {
  const { qSheet } = initSheets();
  const rowIdx = findRowById(qSheet, data.id);
  if (rowIdx < 0) return { success: false, error: 'Not found' };
  qSheet.deleteRow(rowIdx);
  return { success: true };
}

// ── 統計 ──
function getStats() {
  const { qSheet } = initSheets();
  const data = qSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, stats: { total: 0 } };

  const rows = data.slice(1).filter(r => r[0]);
  const byStatus = {}, byOperator = {}, byType = {};
  let totalQ = 0, totalN = 0;

  rows.forEach(r => {
    const status = r[17] || '草稿';
    const op     = r[2]  || '不明';
    const type   = r[4]  || '不明';
    byStatus[status]   = (byStatus[status]   || 0) + 1;
    byOperator[op]     = (byOperator[op]     || 0) + 1;
    byType[type]       = (byType[type]       || 0) + 1;
    totalQ += (r[13] || 0);
    totalN += (r[12] || 0);
  });

  return { success: true, stats: {
    total: rows.length, byStatus, byOperator, byType,
    totalQuoteAmt: Math.round(totalQ),
    totalNetAmt:   Math.round(totalN),
    totalProfit:   Math.round(totalQ - totalN)
  }};
}

// ── ヘルパー ──
function findRowById(sheet, id) {
  if (!id) return -1;
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function colorRow(sheet, rowIdx, status) {
  const colors = { '草稿':'#ffffff','已发送':'#fff9e6','成单':'#e8f5ee','取消':'#fde8e8','跟进中':'#e8f0f8' };
  sheet.getRange(rowIdx, 1, 1, QUOTE_HEADERS.length).setBackground(colors[status] || '#ffffff');
}

function writeItinRows(iSheet, quoteId, opt) {
  const vals = iSheet.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][1]) === String(quoteId)) iSheet.deleteRow(i + 1);
  }
  (opt.itin || opt.itinerary || []).forEach((day, di) => {
    (day.slots || []).filter(s => s.n || s.name).forEach(sl => {
      iSheet.appendRow([
        Utilities.getUuid().slice(0,8), quoteId,
        opt.clientName || '', opt.name || '',
        'DAY ' + (di+1), day.date || '', day.title || '',
        sl.t || sl.time || '', sl.type || '',
        sl.n || sl.name || '', sl.note || ''
      ]);
    });
  });
}

function updateStats(stSheet) {
  if (!stSheet) return;
  const s = getStats().stats;
  stSheet.clearContents();
  stSheet.getRange('A1').setValue('ALL JAPAN DMC — 報価統計').setFontSize(14).setFontWeight('bold').setFontColor('#1a2744');
  stSheet.getRange('A2').setValue('更新：' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'));
  stSheet.getRange('A4:B4').setValues([['指標','数値']]).setFontWeight('bold');
  stSheet.getRange('A5:B8').setValues([
    ['報価総数', s.total || 0],
    ['報価総金額', s.totalQuoteAmt || 0],
    ['純コスト合計', s.totalNetAmt || 0],
    ['予想利益合計', s.totalProfit || 0]
  ]);
  let r = 10;
  stSheet.getRange('A'+r).setValue('ステータス別').setFontWeight('bold');
  Object.entries(s.byStatus || {}).forEach(([k,v]) => stSheet.getRange('A'+(++r)+':B'+r).setValues([[k,v]]));
  r++;
  stSheet.getRange('A'+r).setValue('担当者別').setFontWeight('bold');
  Object.entries(s.byOperator || {}).forEach(([k,v]) => stSheet.getRange('A'+(++r)+':B'+r).setValues([[k,v]]));
}
