/**
 * 紅茶媽媽三重神農店：訂單後端（Google Apps Script）
 * 網站送出訂單 → 寫進這份 Google 試算表的「訂單」分頁 → 推播到店長的 LINE
 * 部署成「網頁應用程式」後，把 …/exec 網址填到網站 data.js 的 orderEndpoint
 *
 * 要設定的值放在「專案設定 → 指令碼屬性」，不要寫在程式裡：
 *   LINE_TOKEN   LINE Messaging API 的 Channel access token（長期）
 *   NOTIFY_TO    要收通知的 LINE userId（U 開頭）或群組 groupId（C 開頭），可用逗號隔開放多個
 *   BIND_CODE    （選填）想把通知改送到群組時用的暗號，例如 shennong-2026
 *   （SECRET 會自動產生，用來保護「接單」連結，不用自己填）
 *
 * 接單流程：LINE 通知下面有「接單／無法接單」按鈕 → 店員按了 → 試算表「處理狀態」寫入「已接單」
 * → 客人的訂單頁面自動顯示「已接單」。直接在試算表「處理狀態」欄打字也會顯示給客人看。
 */

// 這個 Apps Script 部署後的網址（和網站 data.js 的 orderEndpoint 一樣）
var WEB_URL = 'https://script.google.com/macros/s/AKfycbzk2CTeRzCQkpwKTYddIx7swi7cGdcXuW6FDVyTdTAJFWwbbKmS1uHj9znbSvXQ5G1c/exec';
var TZ = 'Asia/Taipei';
var MIN_LEAD = 15;   // 後端最後一道關：取餐時間至少要比現在晚幾分鐘（網站本身是 40 分鐘）

var SHEET_NAME = '訂單';
var HEADERS = ['訂單編號', '下單時間', '取餐方式', '日期', '時間', '姓名', '電話', '公司', '地址', '距離',
               '品項', '杯數', '金額', '備註', 'LINE 通知', '處理狀態', '接單時間', '查詢碼'];
var COL_STATUS = 16, COL_TIME = 17, COL_TOKEN = 18;

function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.a === 'status') return json_(orderStatus(q.no, q.t, q.tel));
  if (q.a === 'staff') return staffPage_(q.no, q.k);
  return json_({ ok: true, service: '神農店訂單' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 網站按「送出訂單」時呼叫 */
function submitOrder(o) {
  if (!o || o.website) return { ok: false, error: '訂單格式不正確' }; // 機器人填了隱藏欄位
  var err = checkOrder_(o) || checkTime_(o);
  if (err) return { ok: false, error: err };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var tz = 'Asia/Taipei';
    var now = new Date();
    var prefix = Utilities.formatDate(now, tz, 'MMdd') + '-';
    var todayCount = sheet.getLastRow() < 2 ? 0 : sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .filter(function (r) { return String(r[0]).indexOf(prefix) === 0; }).length;
    var orderNo = prefix + ('00' + (todayCount + 1)).slice(-3);

    var cups = o.items.reduce(function (s, l) { return s + l.qty; }, 0);
    var itemsTxt = o.items.map(function (l) { return l.name + ' ×' + l.qty + '（' + l.opts + '）單價 $' + l.unit; }).join('\n');
    var token = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    var notified = pushLine_(formatOrder_(o, orderNo), acceptButtons_(o, orderNo));

    sheet.appendRow([orderNo, now, o.mode, o.day, o.time, o.name, "'" + o.tel, o.company, o.addr, o.dist,
                     itemsTxt, cups, o.total, (o.discount ? (o.promo || '優惠') + ' −' + o.discount + '　' : '') + (o.code ? '碼:' + o.code + '　' : '') + (Number(o.bags) > 0 ? '塑膠袋×' + Number(o.bags) + '　' : '') + (o.note || ''), notified ? '已通知' : '未通知', '', '', token]);
    return { ok: true, orderNo: orderNo, notified: notified, token: token };
  } finally {
    lock.releaseLock();
  }
}

function checkOrder_(o) {
  if (!o.items || !o.items.length) return '沒有品項';
  if (o.items.length > 60) return '品項太多，請直接打電話訂購';
  if (!o.name || !o.tel) return '請填姓名和電話';
  if (String(o.tel).length > 60 || String(o.tel).replace(/\D/g, '').length < 8) return '電話格式不對';
  if (o.mode !== '自取' && !o.addr) return '外送要填地址';
  var sum = o.items.reduce(function (s, l) {
    return s + Math.max(0, Number(l.unit) || 0) * Math.max(0, Math.min(99, Number(l.qty) || 0));
  }, 0);
  var disc = Number(o.discount) || 0;
  var bags = Math.max(0, Math.min(99, Math.floor(Number(o.bags) || 0)));
  if (Number(o.bagFee || 0) !== bags) return '塑膠袋金額有誤，請重新整理網頁再下單';
  if (sum !== Number(o.subtotal) || disc < 0 || disc > Math.floor(sum / 2) || sum - disc + bags !== Number(o.total))
    return '金額計算有誤，請重新整理網頁再下單';
  return '';
}

/** 擋掉已經過了的時段（例如客人開著舊網頁，過了截止時間才按送出） */
function checkTime_(o) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.day || '') || !/^\d{2}:\d{2}$/.test(o.time || '')) return '請選取餐日期和時間';
  var earliest = Utilities.formatDate(new Date(Date.now() + MIN_LEAD * 60000), TZ, 'yyyy-MM-dd HH:mm');
  if (o.day + ' ' + o.time < earliest) return '這個時段已經來不及了，請重新整理網頁再選時間';
  return '';
}

function formatOrder_(o, no) {
  var L = ['🧋 新訂單 ' + no, o.mode + '｜' + o.dayLabel + ' ' + o.time];
  if (o.company) L.push('公司：' + o.company);
  L.push('姓名：' + o.name, '電話：' + o.tel);
  if (o.addr) L.push('地址：' + o.addr + '（' + o.dist + '）');
  L.push('────────');
  o.items.forEach(function (l) { L.push(l.name + ' ×' + l.qty + '\n　' + l.opts + '\n　單價 $' + l.unit + ' × ' + l.qty + ' = $' + l.unit * l.qty); });
  if (Number(o.bags) > 0) L.push('塑膠袋 ×' + Number(o.bags) + '　$' + Number(o.bagFee));
  L.push('────────');
  if (o.discount) {
    L.push('飲料小計 $' + o.subtotal, '優惠：' + (o.promo || '優惠') + '　−$' + o.discount);
    (Array.isArray(o.promoDetail) ? o.promoDetail : []).slice(0, 30).forEach(function (d) { L.push('　' + String(d).slice(0, 80)); });
  }
  if (o.code) L.push('優惠碼：' + o.code);
  L.push('共 ' + o.items.reduce(function (s, l) { return s + l.qty; }, 0) + ' 杯　合計 $' + o.total);
  if (o.note) L.push('備註：' + o.note);
  return L.join('\n').slice(0, 4900);
}

/** LINE 訊息下面的「接單／無法接單」按鈕 */
function acceptButtons_(o, no) {
  var url = WEB_URL + '?a=staff&no=' + encodeURIComponent(no) + '&k=' + staffKey_(no);
  return { type: 'template', altText: '訂單 ' + no + '：點開來接單',
    template: { type: 'buttons', text: ('訂單 ' + no + '｜' + o.mode + ' ' + o.dayLabel + ' ' + o.time).slice(0, 160),
      actions: [{ type: 'uri', label: '✅ 接單／無法接單', uri: url }] } };
}

function staffKey_(no) {
  var p = PropertiesService.getScriptProperties(), secret = p.getProperty('SECRET');
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', secret); }
  return Utilities.computeHmacSha256Signature(String(no), secret)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('').slice(0, 20);
}

/** 用訂單編號找試算表的那一列（從最新的往回找） */
function findRow_(sheet, no) {
  var last = sheet.getLastRow();
  if (!no || last < 2) return 0;
  var start = Math.max(2, last - 499);
  var ids = sheet.getRange(start, 1, last - start + 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(no)) return start + i;
  return 0;
}

/** 客人查訂單狀態：網站每隔一陣子來問一次 */
/**
 * 客人查訂單：訂單編號、電話二擇一
 * - 只給訂單編號 → 回傳那一張
 * - 只給電話（完整號碼）→ 回傳這支電話最近 5 張訂單
 * 只回傳狀態和取餐時間，不回傳姓名、地址
 */
function orderStatus(no, token, tel) {
  no = String(no || '').trim();
  var sheet = getSheet_(), notFound = { ok: false, error: '查不到訂單，請確認訂單編號或電話' };
  var pick = function (v) {
    return { orderNo: v[0], status: v[COL_STATUS - 1].trim(), at: v[COL_TIME - 1], mode: v[2], day: v[3], time: v[4] };
  };
  if (no) {
    var row = findRow_(sheet, no);
    if (!row) return notFound;
    var one = pick(sheet.getRange(row, 1, 1, COL_TOKEN).getDisplayValues()[0]);
    return { ok: true, orderNo: one.orderNo, status: one.status, at: one.at, mode: one.mode, day: one.day, time: one.time, orders: [one] };
  }
  var digits = String(tel || '').split(/#|分機|轉|ext/i)[0].replace(/\D/g, '');
  if (digits.length < 8) return { ok: false, error: '請輸入訂單編號，或下單時留的完整電話' };
  var last = sheet.getLastRow();
  if (last < 2) return notFound;
  var first = Math.max(2, last - 499);
  var rows = sheet.getRange(first, 1, last - first + 1, COL_TOKEN).getDisplayValues(), found = [];
  for (var i = rows.length - 1; i >= 0 && found.length < 5; i--) {
    var phones = String(rows[i][6]).split('／').map(function (x) { return x.split(/#|分機|轉|ext/i)[0].replace(/\D/g, ''); });
    if (phones.indexOf(digits) >= 0) found.push(pick(rows[i]));
  }
  return found.length ? { ok: true, orders: found } : notFound;
}

/** 店員在 LINE 點按鈕後看到的頁面 */
function staffPage_(no, key) {
  var sheet = getSheet_(), row = findRow_(sheet, no);
  if (!row || key !== staffKey_(no)) return HtmlService.createHtmlOutput('<p style="font:20px sans-serif;padding:24px">連結不正確或訂單不存在。</p>');
  var v = sheet.getRange(row, 1, 1, COL_TOKEN).getDisplayValues()[0];
  var esc = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var when = v[3], tm = v[4];
  var html = '<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
    'body{font:17px/1.6 -apple-system,"PingFang TC","Noto Sans TC",sans-serif;margin:0;padding:20px;color:#222}' +
    'h1{font-size:26px;margin:0 0 4px}pre{white-space:pre-wrap;background:#f5f2ea;padding:12px;border-radius:8px;font:inherit;font-size:15px}' +
    'button{display:block;width:100%;font-size:20px;padding:16px;border:0;border-radius:10px;margin-top:12px;color:#fff}' +
    '#ok{background:#2f6b3a}#no{background:#a33}#st{font-size:20px;font-weight:700;margin:12px 0}</style>' +
    '<h1>訂單 ' + esc(no) + '</h1><div>' + esc(v[2]) + '｜' + esc(when) + ' ' + esc(tm) + '｜' + esc(v[5]) + ' ' + esc(v[6]) + '</div>' +
    (v[8] ? '<div>地址：' + esc(v[8]) + '</div>' : '') +
    '<pre>' + esc(v[10]) + '\n共 ' + esc(v[11]) + ' 杯　$' + esc(v[12]) + (v[13] ? '\n備註：' + esc(v[13]) : '') + '</pre>' +
    '<div id="st">' + (v[COL_STATUS - 1] ? '目前狀態：' + esc(v[COL_STATUS - 1]) : '還沒有人接單') + '</div>' +
    '<button id="ok" onclick="go(\'已接單\')">✅ 接單</button>' +
    '<button id="no" onclick="if(confirm(\'確定無法接這張單？請記得打電話告訴客人。\'))go(\'無法接單\')">❌ 無法接單（會請客人來電）</button>' +
    '<script>function go(s){document.getElementById("st").textContent="處理中…";' +
    'google.script.run.withSuccessHandler(function(r){document.getElementById("st").textContent=r}).staffSet(' +
    JSON.stringify(String(no)) + ',' + JSON.stringify(String(key)) + ',s)}</script>';
  return HtmlService.createHtmlOutput(html).setTitle('接單 ' + no)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 店員頁面按下按鈕時呼叫 */
function staffSet(no, key, status) {
  if (key !== staffKey_(no) || ['已接單', '無法接單'].indexOf(status) < 0) return '連結不正確';
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(), row = findRow_(sheet, no);
    if (!row) return '找不到訂單';
    sheet.getRange(row, COL_STATUS, 1, 2).setValues([[status, "'" + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm')]]);
    return '✔ 已設定為「' + status + '」，客人在網站上會看到。' + (status === '無法接單' ? '請打電話通知客人。' : '');
  } finally {
    lock.releaseLock();
  }
}

/** 推播到 LINE；沒設定或失敗回傳 false（訂單照樣存進試算表） */
function pushLine_(text, extra) {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty('LINE_TOKEN');
  var to = (p.getProperty('NOTIFY_TO') || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (!token || !to.length) return false;
  var ok = false;
  to.forEach(function (id) {
    try {
      var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ to: id, messages: [{ type: 'text', text: text }].concat(extra ? [extra] : []) })
      });
      if (res.getResponseCode() === 200) ok = true;
      else console.error('LINE push ' + res.getResponseCode() + ' ' + res.getContentText());
    } catch (e) { console.error(e); }
  });
  return ok;
}

/**
 * 所有 POST 都進這裡：
 * - 網站送出的訂單（type: 'order'）
 * - （選填）LINE Webhook：把官方帳號拉進店裡的 LINE 群組後，在群組裡打「綁定 暗號」，
 *   通知就會改送到這個群組，店員都看得到。
 */
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '格式不正確' }); }
  if (body.type === 'order') {
    try { return json_(submitOrder(body)); }
    catch (err) { console.error(err); return json_({ ok: false, error: '系統忙碌中' }); }
  }
  try {
    var p = PropertiesService.getScriptProperties();
    var code = p.getProperty('BIND_CODE');
    (body.events || []).forEach(function (ev) {
      if (ev.type !== 'message' || ev.message.type !== 'text' || !code) return;
      if (ev.message.text.trim() !== '綁定 ' + code) return;
      var id = ev.source.groupId || ev.source.roomId || ev.source.userId;
      p.setProperty('NOTIFY_TO', id);
      reply_(ev.replyToken, '綁定完成，之後的線上訂單會傳到這裡。');
    });
  } catch (err) { console.error(err); }
  return ContentService.createTextOutput('OK');
}

function reply_(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + PropertiesService.getScriptProperties().getProperty('LINE_TOKEN') },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] })
  });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');  // 舊表補上新欄位名稱
  }
  return sh;
}

/** 設定完在編輯器裡按「執行」跑這個，店長的 LINE 應該會收到一則測試訊息 */
function testLine() {
  var ok = pushLine_('✅ 線上訂購通知測試：設定成功');
  Logger.log(ok ? '成功，請看 LINE' : '失敗，請檢查 LINE_TOKEN 和 NOTIFY_TO');
}
