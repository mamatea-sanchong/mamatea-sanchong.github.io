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
 * 接單流程：LINE 通知卡片上有「接單／無法接單」按鈕 → 店員按了 → 試算表「處理狀態」寫入「已接單」
 * → 客人的訂單頁面自動顯示「已接單」。飲料做好按「完成」，客人會看到「可以來取餐／已出發」。
 * 「處理狀態」欄是下拉選單（已接單／無法接單／已完成／已取消），直接在試算表選也會顯示給客人看。
 *
 * 第一次貼上這版：編輯器上方函式選 setup → 執行（授權、加下拉選單和顏色、裝提醒排程），再部署新版本。
 * 排程每 5 分鐘跑一次：
 *   - 訂單 10 分鐘沒人接、取餐時間在 90 分鐘內 → LINE 再提醒一次（每張只提醒一次）
 *   - 每天開店前 30 分鐘 → LINE 傳今天的訂單摘要和「今日訂單」頁面連結
 */

// 這個 Apps Script 部署後的網址（和網站 data.js 的 orderEndpoint 一樣）
var WEB_URL = 'https://script.google.com/macros/s/AKfycbzk2CTeRzCQkpwKTYddIx7swi7cGdcXuW6FDVyTdTAJFWwbbKmS1uHj9znbSvXQ5G1c/exec';
var TZ = 'Asia/Taipei';
var MIN_LEAD = 15;   // 後端最後一道關：取餐時間至少要比現在晚幾分鐘（網站本身是 40 分鐘）

var SHEET_NAME = '訂單';
var HEADERS = ['訂單編號', '下單時間', '取餐方式', '日期', '時間', '姓名', '電話', '公司', '地址', '距離',
               '品項', '杯數', '金額', '備註', 'LINE 通知', '處理狀態', '接單時間', '查詢碼', '優惠', '原因'];
var COL_NOTE = 14, COL_NOTIFY = 15, COL_STATUS = 16, COL_TIME = 17, COL_TOKEN = 18, COL_PROMO = 19, COL_REASON = 20;
var NCOL = HEADERS.length;
var STATUSES = ['已接單', '無法接單', '已完成', '已取消'];
var REASONS = ['品項售完', '這個時段太滿', '超出外送範圍', '今天臨時公休'];
var WD = ['日', '一', '二', '三', '四', '五', '六'];
// 營業時間（和網站 data.js 一樣）：用來決定開店前什麼時候傳今日摘要
var HOURS = { 1: '11:00', 2: '11:00', 3: '11:00', 4: '11:00', 5: '11:00', 6: '12:00' };
var REMIND_AFTER = 10, REMIND_WITHIN = 90, SUMMARY_BEFORE = 30;  // 分鐘

function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.a === 'status') return json_(orderStatus(q.no, q.t, q.tel));
  if (q.a === 'staff') return staffPage_(q.no, q.k);
  if (q.a === 'set') return quickSet_(q.no, q.k, q.s);
  if (q.a === 'today') return todayPage_(q.d, q.k);
  if (q.a === 'cancel') return json_(customerCancel(q.no, q.t));
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
    // 用今天最大的號碼 +1（刪過列也不會重號）
    var maxNo = 0, lastRow = sheet.getLastRow();
    if (lastRow >= 2) sheet.getRange(Math.max(2, lastRow - 499), 1, lastRow - Math.max(2, lastRow - 499) + 1, 1).getValues()
      .forEach(function (r) { var x = String(r[0]); if (x.indexOf(prefix) === 0) maxNo = Math.max(maxNo, Number(x.slice(prefix.length)) || 0); });
    var orderNo = prefix + ('00' + (maxNo + 1)).slice(-3);

    var cups = o.items.reduce(function (s, l) { return s + l.qty; }, 0);
    var itemsTxt = o.items.map(function (l) { return l.name + ' ×' + l.qty + '（' + l.opts + '）單價 $' + l.unit; }).join('\n');
    var token = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    var notified = pushLine_([orderCard_(o, orderNo)]);
    var promo = (o.discount ? (o.promo || '優惠') + ' −' + o.discount + '　' : '') + (o.code ? '碼:' + o.code + '　' : '') + (Number(o.bags) > 0 ? '塑膠袋×' + Number(o.bags) : '');

    sheet.appendRow([orderNo, now, o.mode, o.day, o.time, o.name, "'" + o.tel, o.company, o.addr, o.dist,
                     itemsTxt, cups, o.total, o.note || '', notified ? '已通知' : '未通知', '', '', token, promo.trim(), '']);
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

/** 取餐日：今天／明天／⚠️ 預訂 11/22（日），店員一眼分得出來 */
function dayWord_(day) {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var tmr = Utilities.formatDate(new Date(Date.now() + 864e5), TZ, 'yyyy-MM-dd');
  var d = String(day || '').replace(/\//g, '-');
  if (d === today) return '今天';
  var m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var md = m ? (+m[2]) + '/' + (+m[3]) + '（' + WD[new Date(+m[1], +m[2] - 1, +m[3]).getDay()] + '）' : d;
  return d === tmr ? '明天 ' + md : '⚠️ 預訂 ' + md;
}
function telDigits_(tel) { return String(tel || '').split(/／|#|分機|轉|ext/i)[0].replace(/[^\d]/g, ''); }

/** LINE 通知：一張卡片，最上面是取餐日和時間，下面是品項，按鈕在同一張 */
function orderCard_(o, no) {
  var when = dayWord_(o.day) + ' ' + o.time + ' ' + o.mode;
  var cups = o.items.reduce(function (s, l) { return s + l.qty; }, 0);
  var T = function (text, opt) { var x = { type: 'text', text: String(text).slice(0, 300) || ' ', wrap: true }; for (var k in opt) x[k] = opt[k]; return x; };
  var body = [];
  o.items.slice(0, 30).forEach(function (l) {
    body.push({ type: 'box', layout: 'horizontal', margin: 'md', contents: [
      { type: 'box', layout: 'vertical', flex: 5, contents: [T(l.name + ' ×' + l.qty, { weight: 'bold' }), T(l.opts, { size: 'sm', color: '#666666' })] },
      T('$' + l.unit * l.qty, { flex: 2, align: 'end' })] });
  });
  if (Number(o.bags) > 0) body.push(T('塑膠袋 ×' + Number(o.bags) + '　$' + Number(o.bagFee), { size: 'sm', margin: 'md' }));
  body.push({ type: 'separator', margin: 'lg' });
  if (o.discount) body.push(T('優惠：' + (o.promo || '優惠') + '　−$' + o.discount, { size: 'sm', color: '#3e5b45', margin: 'md' }));
  if (o.code) body.push(T('優惠碼：' + o.code, { size: 'sm', color: '#3e5b45' }));
  body.push(T('共 ' + cups + ' 杯　合計 $' + o.total, { weight: 'bold', size: 'lg', margin: 'md' }));
  if (o.note) body.push(T('📝 客人備註：' + o.note, { color: '#a8322a', weight: 'bold', margin: 'md' }));
  if (o.addr) body.push(T('📍 ' + (o.company ? o.company + '　' : '') + o.addr + '（' + o.dist + '）', { size: 'sm', margin: 'md' }));
  var staffUrl = WEB_URL + '?a=staff&no=' + encodeURIComponent(no) + '&k=' + staffKey_(no);
  // 店員只用 LINE：接單、做好了都是按一下就完成；無法接單或要看細節才進接單頁
  var setUrl = function (st) { return WEB_URL + '?a=set&no=' + encodeURIComponent(no) + '&k=' + staffKey_(no) + '&s=' + encodeURIComponent(st); };
  var footer = [
    { type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
      { type: 'button', style: 'primary', color: '#3e5b45', action: { type: 'uri', label: '✅ 接單', uri: setUrl('已接單') } },
      { type: 'button', style: 'primary', color: '#3f5b74', action: { type: 'uri', label: '🧋 做好了', uri: setUrl('已完成') } }] },
    { type: 'button', style: 'secondary', height: 'sm', action: { type: 'uri', label: '❌ 無法接單／看明細', uri: staffUrl } }];
  var td = telDigits_(o.tel);
  if (td) footer.push({ type: 'button', style: 'secondary', action: { type: 'uri', label: '📞 打給 ' + String(o.name).slice(0, 8), uri: 'tel:' + td } });
  return { type: 'flex', altText: ('🧋 ' + when + '｜' + cups + ' 杯 $' + o.total + '｜' + no).slice(0, 400),
    contents: { type: 'bubble',
      header: { type: 'box', layout: 'vertical', backgroundColor: dayWord_(o.day) === '今天' ? '#e3e9e1' : '#f1e7da', contents: [
        T(when, { weight: 'bold', size: 'xl' }), T('訂單 ' + no + '｜' + o.name + ' ' + o.tel, { size: 'sm', color: '#555555' })] },
      body: { type: 'box', layout: 'vertical', contents: body },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: footer } } };
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

/**
 * 客人查訂單：訂單編號、電話二擇一
 * - 只給訂單編號 → 回傳那一張
 * - 只給電話（完整號碼）→ 回傳這支電話最近 5 張訂單
 * 回傳狀態、取餐時間和品項明細，不回傳姓名、電話、地址、備註
 */
function orderStatus(no, token, tel) {
  no = String(no || '').trim();
  var sheet = getSheet_(), notFound = { ok: false, error: '查不到訂單，請確認訂單編號或電話' };
  var pick = function (v) {
    // 明細：品項、杯數、金額。姓名、電話、地址、備註不回傳（只用訂單編號也查得到，所以不放個人資料）
    return { orderNo: v[0], status: String(v[COL_STATUS - 1]).trim(), at: v[COL_TIME - 1], mode: v[2], day: v[3], time: v[4],
             dist: v[9], items: v[10], cups: v[11], total: v[12], reason: v[COL_REASON - 1] || '' };
  };
  if (no) {
    var row = findRow_(sheet, no);
    if (!row) return notFound;
    var one = pick(sheet.getRange(row, 1, 1, NCOL).getDisplayValues()[0]);
    var res = { ok: true, orders: [one] };
    for (var k in one) res[k] = one[k];
    return res;
  }
  var digits = String(tel || '').split(/#|分機|轉|ext/i)[0].replace(/\D/g, '');
  if (digits.length < 8) return { ok: false, error: '請輸入訂單編號，或下單時留的完整電話' };
  var last = sheet.getLastRow();
  if (last < 2) return notFound;
  var first = Math.max(2, last - 499);
  var rows = sheet.getRange(first, 1, last - first + 1, NCOL).getDisplayValues(), found = [];
  for (var i = rows.length - 1; i >= 0 && found.length < 5; i--) {
    var phones = String(rows[i][6]).split('／').map(function (x) { return x.split(/#|分機|轉|ext/i)[0].replace(/\D/g, ''); });
    if (phones.indexOf(digits) >= 0) found.push(pick(rows[i]));
  }
  return found.length ? { ok: true, orders: found } : notFound;
}

/** 客人在店家接單前自己取消（要有下單時拿到的查詢碼） */
function customerCancel(no, token) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(), row = findRow_(sheet, no);
    if (!row) return { ok: false, error: '查不到訂單' };
    var v = sheet.getRange(row, 1, 1, NCOL).getDisplayValues()[0];
    if (!token || String(token) !== v[COL_TOKEN - 1]) return { ok: false, error: '只能在下單的手機上取消' };
    var st = String(v[COL_STATUS - 1]).trim();
    if (st === '已取消') return { ok: true, status: st };
    if (st) return { ok: false, error: '店家已經處理這張單，要取消請打電話' };
    sheet.getRange(row, COL_STATUS, 1, 2).setValues([['已取消', "'" + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm')]]);
    pushLine_('🚫 客人取消了訂單 ' + no + '\n' + dayWord_(v[3]) + ' ' + v[4] + ' ' + v[2] + '｜' + v[5] + '\n不用做這張了。');
    return { ok: true, status: '已取消' };
  } finally {
    lock.releaseLock();
  }
}

var esc_ = function (x) { return String(x == null ? '' : x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
var STAFF_CSS = 'body{font:17px/1.6 -apple-system,"PingFang TC","Noto Sans TC",sans-serif;margin:0;padding:20px;color:#222;background:#fff}' +
  'h1{font-size:26px;margin:0 0 4px}.when{font-size:24px;font-weight:700;margin:4px 0}.pre{white-space:pre-wrap;background:#f5f2ea;padding:12px;border-radius:8px;font-size:15px;margin:10px 0}' +
  '.note{background:#f3dfdc;color:#7a231c;font-weight:700;padding:10px 12px;border-radius:8px;margin:8px 0}a{color:#3f5b74}' +
  'button{display:block;width:100%;font:inherit;font-size:20px;padding:15px;border:0;border-radius:10px;margin-top:10px;color:#fff;cursor:pointer}' +
  '.ok{background:#2f6b3a}.no{background:#a33}.done{background:#3f5b74}.ghost{background:#eee;color:#333;font-size:16px;padding:10px}' +
  '.st{font-size:20px;font-weight:700;margin:14px 0 4px}.rs button{background:#fff;color:#a33;border:1px solid #a33;font-size:17px;padding:11px}' +
  '.card{border:1px solid #ddd;border-left:6px solid #bbb;border-radius:8px;padding:10px 12px;margin:12px 0}.card.s0{border-left-color:#c0392b;background:#fdf3f2}' +
  '.card.s1{border-left-color:#2f6b3a}.card.s2{border-left-color:#999;opacity:.6}.card .t{font-size:20px;font-weight:700}.small{font-size:14px;color:#666}' +
  '.row{display:flex;gap:8px}.row button{font-size:16px;padding:10px;margin-top:8px}.tabs a{display:inline-block;margin-right:14px}table{border-collapse:collapse;width:100%}td{padding:4px 0;border-bottom:1px solid #eee}td:last-child{text-align:right;font-weight:700}';

// 注意：Apps Script 的網頁是包在 iframe 裡，頁面裡的連結都要開在最外層（<base target="_top">），
// 不然在 LINE 和手機瀏覽器會出現 ERR_BLOCKED_BY_RESPONSE。
/** 店員在 LINE 點按鈕後看到的頁面 */
function staffPage_(no, key) {
  var sheet = getSheet_(), row = findRow_(sheet, no);
  if (!row || key !== staffKey_(no)) return HtmlService.createHtmlOutput('<p style="font:20px sans-serif;padding:24px">連結不正確或訂單不存在。</p>');
  var v = sheet.getRange(row, 1, 1, NCOL).getDisplayValues()[0];
  var st = String(v[COL_STATUS - 1]).trim(), td = telDigits_(v[6]);
  var html = '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><base target="_top"><style>' + STAFF_CSS + '</style>' +
    '<h1>訂單 ' + esc_(no) + '</h1>' +
    '<div class="when">' + esc_(dayWord_(v[3])) + ' ' + esc_(v[4]) + ' ' + esc_(v[2]) + '</div>' +
    '<div>' + esc_(v[5]) + (v[7] ? '（' + esc_(v[7]) + '）' : '') + '　' + (td ? '<a href="tel:' + td + '">📞 ' + esc_(v[6]) + '</a>' : esc_(v[6])) + '</div>' +
    (v[8] ? '<div>📍 <a target="_blank" href="https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(v[8]) + '">' + esc_(v[8]) + '</a>（' + esc_(v[9]) + '）</div>' : '') +
    '<div class="pre">' + esc_(v[10]) + '\n共 ' + esc_(v[11]) + ' 杯　$' + esc_(v[12]) + (v[COL_PROMO - 1] ? '\n' + esc_(v[COL_PROMO - 1]) : '') + '</div>' +
    (v[COL_NOTE - 1] ? '<div class="note">📝 客人備註：' + esc_(v[COL_NOTE - 1]) + '</div>' : '') +
    '<div class="st" id="st">' + (st ? '目前狀態：' + esc_(st) + (v[COL_TIME - 1] ? '（' + esc_(v[COL_TIME - 1]) + '）' : '') + (v[COL_REASON - 1] ? '　' + esc_(v[COL_REASON - 1]) : '') : '還沒有人接單') + '</div>' +
    // 已經處理過的單：按鈕先收起來，避免兩個店員重複按
    // 已接單的單：直接顯示「做好了」，其他按鈕收在「要修改狀態」裡
    (st === '已接單' ? '<button class="done" onclick="this.remove();go(\'已完成\')">🧋 做好了（' + (v[2] === '自取' ? '通知客人可以來取' : '已出發外送') + '）</button>' : '') +
    (st ? '<button class="ghost" id="chg" onclick="this.remove();document.getElementById(\'acts\').style.display=\'block\'">要修改狀態</button>' : '') +
    '<div id="acts" style="display:' + (st ? 'none' : 'block') + '">' +
    (st === '已接單' ? '' : '<button class="ok" onclick="go(\'已接單\')">✅ 接單</button>') +
    '<button class="no" onclick="document.getElementById(\'rs\').style.display=\'block\';this.style.display=\'none\'">❌ 無法接單</button>' +
    '<div class="rs" id="rs" style="display:none"><p class="small">選原因，客人在網站上會看到：</p>' +
    REASONS.map(function (r) { return '<button onclick="go(\'無法接單\',\'' + r + '\')">' + r + '</button>'; }).join('') +
    '<button onclick="var r=prompt(\'其他原因（客人會看到）\');if(r)go(\'無法接單\',r)">其他原因…</button></div></div>' +
    '<p class="small" style="margin-top:24px"><a href="' + WEB_URL + '?a=today&k=' + staffKey_('today') + '">📋 看今天全部訂單</a></p>' +
    '<script>function go(s,r){document.getElementById("st").textContent="處理中…";document.getElementById("acts").style.display="none";' +
    'google.script.run.withSuccessHandler(function(m){document.getElementById("st").textContent=m}).staffSet(' +
    JSON.stringify(String(no)) + ',' + JSON.stringify(String(key)) + ',s,r||"")}</script>';
  return HtmlService.createHtmlOutput(html).setTitle('接單 ' + no)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** LINE 卡片上的「接單」「做好了」：按一下就設定好，畫面只顯示結果 */
function quickSet_(no, key, status) {
  var page = function (icon, title, sub) {
    return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><base target="_top">' +
      '<div style="font:20px/1.6 -apple-system,\'PingFang TC\',\'Noto Sans TC\',sans-serif;padding:48px 24px;text-align:center;color:#222">' +
      '<div style="font-size:64px">' + icon + '</div><div style="font-size:26px;font-weight:700;margin:8px 0">' + esc_(title) + '</div>' +
      '<div style="color:#555">' + sub + '</div><p style="color:#888;font-size:16px;margin-top:32px">可以關掉這頁，回到 LINE。</p></div>')
      .setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  };
  if (key !== staffKey_(no) || ['已接單', '已完成'].indexOf(status) < 0) return page('⚠️', '連結不正確', '');
  var sheet = getSheet_(), row = findRow_(sheet, no);
  if (!row) return page('⚠️', '找不到訂單 ' + no, '');
  var v = sheet.getRange(row, 1, 1, NCOL).getDisplayValues()[0], cur = String(v[COL_STATUS - 1]).trim();
  var when = esc_(dayWord_(v[3]) + ' ' + v[4] + ' ' + v[2] + '｜' + v[5]);
  if (cur === '已取消') return page('🚫', '客人已經取消這張單', when + '<br>不用做了。');
  if (cur === '無法接單') return page('❌', '這張單已設為無法接單', when + '<br>要改的話，請按卡片上的「無法接單／看明細」。');
  if (cur === status || (cur === '已完成' && status === '已接單')) return page('✔', no + ' 已經是「' + cur + '」', when + '<br>' + esc_(v[COL_TIME - 1] || ''));
  var msg = staffSet(no, key, status, '');
  var ok = msg.indexOf('✔') === 0;
  return page(ok ? (status === '已完成' ? '🧋' : '✅') : '⚠️', ok ? no + (status === '已完成' ? ' 做好了' : ' 已接單') : msg,
    when + (ok ? '<br>' + (status === '已完成' ? (v[2] === '自取' ? '客人會看到「可以來取囉」。' : '客人會看到「已出發外送」。') : '客人會看到「店家已接單」。飲料做好時，按同一張卡片的「🧋 做好了」。') : ''));
}

/** 店員頁面按下按鈕時呼叫（接單頁和今日訂單頁共用） */
function staffSet(no, key, status, reason) {
  if ((key !== staffKey_(no) && key !== staffKey_('today')) || STATUSES.indexOf(status) < 0 || status === '已取消') return '連結不正確';
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(), row = findRow_(sheet, no);
    if (!row) return '找不到訂單';
    if (String(sheet.getRange(row, COL_STATUS).getDisplayValue()).trim() === '已取消') return '客人已經取消這張單了，不用做。';
    sheet.getRange(row, COL_STATUS, 1, 2).setValues([[status, "'" + Utilities.formatDate(new Date(), TZ, 'M/d HH:mm')]]);
    sheet.getRange(row, COL_REASON).setValue(status === '無法接單' ? String(reason || '').slice(0, 60) : '');
    return '✔ 已設定為「' + status + '」，客人在網站上會看到。';
  } finally {
    lock.releaseLock();
  }
}

/** 今日訂單：按取餐時間排，可以直接接單／完成；最上面是備料彙總 */
function todayPage_(d, key) {
  if (key !== staffKey_('today')) return HtmlService.createHtmlOutput('<p style="font:20px sans-serif;padding:24px">連結不正確。</p>');
  var day = /^\d{4}-\d{2}-\d{2}$/.test(d || '') ? d : Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var list = ordersOn_(day);
  var live = list.filter(function (v) { return ['無法接單', '已取消'].indexOf(String(v[COL_STATUS - 1]).trim()) < 0; });
  var prep = {}, cupsAll = 0, sumAll = 0;
  live.forEach(function (v) {
    cupsAll += Number(v[11]) || 0; sumAll += Number(v[12]) || 0;
    String(v[10]).split('\n').forEach(function (ln) { var m = ln.match(/^(.+?) ×(\d+)/); if (m) prep[m[1]] = (prep[m[1]] || 0) + Number(m[2]); });
  });
  var link = function (k, label) { var x = Utilities.formatDate(new Date(Date.now() + k * 864e5), TZ, 'yyyy-MM-dd'); return '<a href="' + WEB_URL + '?a=today&d=' + x + '&k=' + key + '">' + (x === day ? '<b>' + label + '</b>' : label) + '</a>'; };
  var html = '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><base target="_top"><style>' + STAFF_CSS + '</style>' +
    '<div class="tabs">' + link(0, '今天') + link(1, '明天') + link(2, '後天') + '</div>' +
    '<h1>' + esc_(dayWord_(day).replace('⚠️ 預訂 ', '')) + ' 訂單</h1><div class="small">' + live.length + ' 張・' + cupsAll + ' 杯・$' + sumAll + '（不含無法接單、已取消）</div>' +
    (cupsAll ? '<h3>備料</h3><table>' + Object.keys(prep).sort(function (a, b) { return prep[b] - prep[a]; }).map(function (k) { return '<tr><td>' + esc_(k) + '</td><td>' + prep[k] + '</td></tr>'; }).join('') + '</table>' : '') +
    '<h3>訂單（按取餐時間）</h3>' + (list.length ? '' : '<p class="small">還沒有訂單。</p>') +
    list.map(function (v) {
      var st = String(v[COL_STATUS - 1]).trim(), cls = !st ? 's0' : st === '已接單' ? 's1' : 's2', no = v[0], td = telDigits_(v[6]);
      return '<div class="card ' + cls + '" id="c' + esc_(no) + '"><div class="t">' + esc_(v[4]) + ' ' + esc_(v[2]) + '　<span class="small">' + esc_(no) + '</span></div>' +
        '<div>' + esc_(v[5]) + '　' + (td ? '<a href="tel:' + td + '">' + esc_(v[6]) + '</a>' : esc_(v[6])) + (v[8] ? '<br>📍 ' + esc_(v[8]) : '') + '</div>' +
        '<div class="pre">' + esc_(v[10]) + '\n共 ' + esc_(v[11]) + ' 杯　$' + esc_(v[12]) + '</div>' +
        (v[COL_NOTE - 1] ? '<div class="note">📝 ' + esc_(v[COL_NOTE - 1]) + '</div>' : '') +
        '<div class="small" id="s' + esc_(no) + '">' + (st ? esc_(st) + ' ' + esc_(v[COL_TIME - 1]) + (v[COL_REASON - 1] ? '　' + esc_(v[COL_REASON - 1]) : '') : '⚠️ 還沒接單') + '</div>' +
        (st === '已取消' || st === '無法接單' || st === '已完成' ? '' : '<div class="row">' +
          (st ? '' : '<button class="ok" onclick="go(\'' + esc_(no) + '\',\'已接單\')">✅ 接單</button>') +
          '<button class="done" onclick="go(\'' + esc_(no) + '\',\'已完成\')">🧋 做好了</button></div>') + '</div>';
    }).join('') +
    '<p class="small">無法接單請從 LINE 通知的按鈕進去選原因。頁面不會自動更新，按<a href="' + WEB_URL + '?a=today&d=' + day + '&k=' + key + '">重新整理</a>。</p>' +
    '<script>function go(n,s){var e=document.getElementById("s"+n);e.textContent="處理中…";' +
    'google.script.run.withSuccessHandler(function(m){e.textContent=m}).staffSet(n,' + JSON.stringify(key) + ',s,"")}</script>';
  return HtmlService.createHtmlOutput(html).setTitle('今日訂單').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 某一天的訂單列（按取餐時間排序） */
function ordersOn_(day) {
  var sheet = getSheet_(), last = sheet.getLastRow();
  if (last < 2) return [];
  var first = Math.max(2, last - 499);
  return sheet.getRange(first, 1, last - first + 1, NCOL).getDisplayValues()
    .filter(function (v) { return String(v[3]).replace(/\//g, '-') === day; })
    .sort(function (a, b) { return a[4] < b[4] ? -1 : a[4] > b[4] ? 1 : 0; });
}

/** 排程（每 5 分鐘）：漏單提醒、開店前今日摘要 */
function tick() {
  var now = new Date(), nowMin = now.getHours() * 60 + now.getMinutes();
  var tz = function (f, d) { return Utilities.formatDate(d || now, TZ, f); };
  var hm = tz('HH:mm'), today = tz('yyyy-MM-dd'), p = PropertiesService.getScriptProperties();
  var toMin = function (t) { var a = String(t).split(':'); return +a[0] * 60 + +a[1]; };
  nowMin = toMin(hm);
  // 1) 今日摘要：開店前 SUMMARY_BEFORE 分鐘之後的第一次排程
  var dow = Number(tz('u')) % 7, open = HOURS[dow];
  if (open && nowMin >= toMin(open) - SUMMARY_BEFORE && nowMin < toMin(open) + 60 && p.getProperty('SUM') !== today) {
    p.setProperty('SUM', today);
    var list = ordersOn_(today).filter(function (v) { return ['無法接單', '已取消'].indexOf(String(v[COL_STATUS - 1]).trim()) < 0; });
    if (list.length) {
      var cups = list.reduce(function (s, v) { return s + (Number(v[11]) || 0); }, 0);
      var waiting = list.filter(function (v) { return !String(v[COL_STATUS - 1]).trim(); }).length;
      pushLine_('☀️ 今天的線上訂單：' + list.length + ' 張，共 ' + cups + ' 杯，最早 ' + list[0][4] + '。' +
        (waiting ? '\n⚠️ 還有 ' + waiting + ' 張沒接單。' : '') + '\n' +
        list.map(function (v) { return v[4] + ' ' + v[2] + '　' + v[5] + '　' + v[11] + ' 杯'; }).join('\n') +
        '\n\n📋 今日訂單：' + WEB_URL + '?a=today&k=' + staffKey_('today'));
    }
  }
  // 2) 漏單提醒：下單超過 REMIND_AFTER 分鐘還沒人接，而且取餐時間在 REMIND_WITHIN 分鐘內（每張只提醒一次）
  var sheet = getSheet_(), last = sheet.getLastRow();
  if (last < 2) return;
  var first = Math.max(2, last - 99), rng = sheet.getRange(first, 1, last - first + 1, NCOL);
  var rows = rng.getDisplayValues(), raw = rng.getValues();
  var reminded = (p.getProperty('REMINDED') || '').split(',');
  rows.forEach(function (v, i) {
    var no = String(v[0]), st = String(v[COL_STATUS - 1]).trim();
    if (!no || st || reminded.indexOf(no) >= 0) return;
    var day = String(v[3]).replace(/\//g, '-'), time = String(v[4]).slice(0, 5);
    var placed = raw[i][1] instanceof Date ? raw[i][1].getTime() : 0;
    if (day !== today || !placed || Date.now() - placed < REMIND_AFTER * 60000) return;
    var left = toMin(time) - nowMin;
    if (left > REMIND_WITHIN || left < -30) return;
    reminded.push(no);
    pushLine_('⏰ 訂單 ' + no + ' 還沒人接單！\n今天 ' + time + ' ' + v[2] + '｜' + v[5] + '（還有 ' + Math.max(0, left) + ' 分鐘）\n' +
      '👉 ' + WEB_URL + '?a=staff&no=' + encodeURIComponent(no) + '&k=' + staffKey_(no));
  });
  p.setProperty('REMINDED', reminded.filter(String).slice(-60).join(','));
}

/**
 * 第一次貼上這版時，在編輯器選 setup 按執行一次：
 * 補欄位名稱、「處理狀態」改下拉選單並加顏色、裝每 5 分鐘的排程
 */
function setup() {
  var sh = getSheet_(), rows = Math.max(1000, sh.getMaxRows()) - 1;
  sh.getRange(1, 1, 1, NCOL).setValues([HEADERS]).setFontWeight('bold');
  var stRange = sh.getRange(2, COL_STATUS, rows, 1);
  stRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  // 顏色：還沒接單＝淡紅（有訂單編號的列才上色），已接單＝綠，已完成／已取消／無法接單＝灰
  var rng = sh.getRange(2, 1, rows, NCOL);
  var rules = sh.getConditionalFormatRules().filter(function (r) { return String(r.getRanges()[0].getA1Notation()).indexOf('A2') !== 0; });
  var col = String.fromCharCode(64 + COL_STATUS);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$' + col + '2="")').setBackground('#fbe3e0').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + col + '2="已接單"').setBackground('#e3efe3').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($' + col + '2="已完成",$' + col + '2="已取消",$' + col + '2="無法接單")').setFontColor('#999999').setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
  Logger.log('完成：處理狀態下拉選單、顏色、每 5 分鐘排程都設定好了。今日訂單頁：' + WEB_URL + '?a=today&k=' + staffKey_('today'));
}

/** 推播到 LINE；text 可以是文字，或 LINE 訊息物件的陣列。沒設定或失敗回傳 false（訂單照樣存進試算表） */
function pushLine_(text) {
  var p = PropertiesService.getScriptProperties();
  var token = p.getProperty('LINE_TOKEN');
  var to = (p.getProperty('NOTIFY_TO') || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (!token || !to.length) return false;
  var messages = Array.isArray(text) ? text : [{ type: 'text', text: String(text).slice(0, 4900) }];
  var ok = false;
  to.forEach(function (id) {
    try {
      var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ to: id, messages: messages })
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
  } else if (sh.getRange(1, NCOL).getValue() !== HEADERS[NCOL - 1]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');  // 舊表補上新欄位名稱
  }
  return sh;
}

/** 設定完在編輯器裡按「執行」跑這個，店長的 LINE 應該會收到一則測試訊息 */
function testLine() {
  var ok = pushLine_('✅ 線上訂購通知測試：設定成功');
  Logger.log(ok ? '成功，請看 LINE' : '失敗，請檢查 LINE_TOKEN 和 NOTIFY_TO');
}
