/**
 * 豆豆貼 線上點餐 API（Google Apps Script + Google Sheets）
 *
 * 工作表：
 *   商品      — 菜單資料，網站從這裡讀取品項與價格
 *   訂單      — 每筆訂單一列（狀態可在表內用下拉選單更新）
 *   訂單明細  — 每個品項一列
 *
 * 端點（部署為「網頁應用程式」後的網址）：
 *   GET  ?action=products                 取得上架中的商品
 *   GET  ?action=order&id=訂單編號&phone=電話   查詢訂單狀態
 *   POST {action:'createOrder', ...}      建立訂單（價格一律以「商品」表為準）
 *
 * 安裝步驟請見同資料夾的 README.md。
 */

const CONFIG = {
  SHEET_PRODUCTS: '商品',
  SHEET_ORDERS: '訂單',
  SHEET_ITEMS: '訂單明細',
  TIMEZONE: 'Asia/Taipei',
  ORDER_PREFIX: 'DD',
  METHODS: ['內用', '外帶', '宅配'],
  STATUSES: ['新訂單', '製作中', '可取餐', '已完成', '已取消'],
  MAX_QTY_PER_ITEM: 99,
  MAX_LINES: 30,
  RATE_LIMIT_SECONDS: 30,      // 同一支電話多久內不能重複下單
  PRODUCTS_CACHE_SECONDS: 60,  // 商品清單快取秒數（改表後最多延遲這麼久；編輯商品表會自動清快取）
  NOTIFY_EMAIL: ''             // 填入店家 Email 可收到新訂單通知；留空則不寄
};

const PRODUCT_HEADERS = ['id', '分類', '名稱', '說明', '價格', '標籤', '圖片', '上架', '排序'];
const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', '地址', '品項摘要', '件數', '總額', '備註'];
const ITEM_HEADERS = ['訂單編號', '商品ID', '品名', '單價', '數量', '小計'];

// 初始菜單（與目前網站相同）。「圖片」可填 dumpling / gyoza / soup / side，或完整圖片網址。
const SEED_PRODUCTS = [
  ['p1', '招牌鍋貼', '黃金鮮肉鍋貼', '台灣豬肉與鮮甜高麗菜，金黃酥脆。', 85, '人氣 NO.1', 'dumpling', true, 1],
  ['p2', '招牌鍋貼', '香辣韭菜鍋貼', '韭菜香氣濃郁，微辣更開胃。', 90, '微辣', 'gyoza', true, 2],
  ['p3', '招牌鍋貼', '玉米鮮肉鍋貼', '香甜玉米與鮮嫩肉餡的溫柔搭配。', 90, '爆汁', 'dumpling', true, 3],
  ['p4', '招牌鍋貼', '鮮蝦手作鍋貼', '彈牙鮮蝦，層次豐富的海味。', 115, '主廚推薦', 'gyoza', true, 4],
  ['p5', '精選水餃', '高麗菜鮮肉水餃', '清甜爽口，皮薄餡飽滿。', 80, '經典', 'gyoza', true, 5],
  ['p6', '精選水餃', '韭菜鮮肉水餃', '鮮香多汁，家常好滋味。', 85, '人氣', 'dumpling', true, 6],
  ['p7', '湯品／飲料', '暖心酸辣湯', '酸香微辣，搭配鍋貼剛剛好。', 45, '推薦搭配', 'soup', true, 7],
  ['p8', '湯品／飲料', '古早味紅茶', '茶香清爽，解膩回甘。', 30, '冰飲', 'soup', true, 8],
  ['p9', '經典小菜', '爽脆涼拌小黃瓜', '清脆酸甜，開胃首選。', 45, '清爽', 'side', true, 9],
  ['p10', '經典小菜', '香滷豆干', '滷香入味，越嚼越香。', 40, '經典', 'side', true, 10],
  ['gift', '冷凍禮盒', '招牌招財鍋貼禮盒 50 入', '適合全家分享的冷凍鍋貼組。', 599, '冷凍宅配', 'dumpling', true, 11]
];

/* ------------------------------------------------------------------ 選單與初始化 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('豆豆貼')
    .addItem('初始化工作表', 'setup')
    .addItem('清除商品快取', 'clearProductCache')
    .addToUi();
}

/** 建立三張工作表、表頭、範例商品、上架勾選框與訂單狀態下拉選單。可重複執行，不會覆蓋既有資料。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  const products = ensureSheet_(ss, CONFIG.SHEET_PRODUCTS, PRODUCT_HEADERS);
  if (products.getLastRow() < 2) {
    products.getRange(2, 1, SEED_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(SEED_PRODUCTS);
  }
  const rows = Math.max(products.getLastRow() - 1, 1) + 50;
  products.getRange(2, 8, rows, 1).insertCheckboxes();
  products.getRange(2, 5, rows, 1).setNumberFormat('0');

  const orders = ensureSheet_(ss, CONFIG.SHEET_ORDERS, ORDER_HEADERS);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build();
  orders.getRange(2, 3, orders.getMaxRows() - 1, 1).setDataValidation(statusRule);
  orders.getRange(2, 2, orders.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  ensureSheet_(ss, CONFIG.SHEET_ITEMS, ITEM_HEADERS);

  const blank = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  clearProductCache();
  ss.toast('工作表已就緒', '豆豆貼', 5);
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#38271f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 編輯「商品」表時清掉快取，網站下次讀取就會拿到新資料。 */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === CONFIG.SHEET_PRODUCTS) clearProductCache();
}

function clearProductCache() {
  CacheService.getScriptCache().remove('products');
}

/* ------------------------------------------------------------------ HTTP 入口 */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    switch (p.action || 'products') {
      case 'products': return json_({ ok: true, products: getProducts_() });
      case 'order': return json_(lookupOrder_(p.id, p.phone));
      case 'ping': return json_({ ok: true, time: new Date().toISOString() });
      default: return json_({ ok: false, error: '未知的操作' });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: '系統忙碌中，請稍後再試' });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '資料格式錯誤' });
  }
  try {
    if (body.action === 'createOrder') return json_(createOrder_(body));
    return json_({ ok: false, error: '未知的操作' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : '系統忙碌中，請稍後再試' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ 商品 */

function getProducts_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('products');
  if (hit) return JSON.parse(hit);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_PRODUCTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
  const list = rows
    .filter(r => String(r[0]).trim() && r[7] === true && Number(r[4]) >= 0 && String(r[4]) !== '')
    .map(r => ({
      id: String(r[0]).trim(),
      cat: String(r[1]).trim(),
      name: String(r[2]).trim(),
      desc: String(r[3]).trim(),
      price: Math.round(Number(r[4])),
      tag: String(r[5]).trim(),
      img: String(r[6]).trim(),
      sort: Number(r[8]) || 9999
    }))
    .sort((a, b) => a.sort - b.sort)
    .map(({ sort, ...p }) => p);

  cache.put('products', JSON.stringify(list), CONFIG.PRODUCTS_CACHE_SECONDS);
  return list;
}

/* ------------------------------------------------------------------ 訂單 */

function createOrder_(body) {
  // 機器人陷阱欄位：真人看不到，有填就當作沒送出
  if (body.website) return { ok: false, error: '送出失敗，請重新整理後再試' };

  const method = String(body.method || '');
  const name = clean_(body.name, 50);
  const phone = clean_(body.phone, 20);
  const address = clean_(body.address, 200);
  const note = clean_(body.note, 200);

  if (CONFIG.METHODS.indexOf(method) < 0) return fail_('請選擇取餐方式');
  if (!name) return fail_('請填寫聯絡姓名');
  if (!/^[0-9+() -]{8,20}$/.test(phone) || phone.replace(/\D/g, '').length < 8) return fail_('電話格式不正確');
  if (method === '宅配' && address.length < 5) return fail_('請填寫完整宅配地址');

  // 合併重複品項，並以「商品」表的價格重新計算（不信任前端傳來的價格）
  const items = Array.isArray(body.items) ? body.items : [];
  if (!items.length) return fail_('購物車是空的');
  if (items.length > CONFIG.MAX_LINES) return fail_('品項太多，請分批下單');
  const qtyById = {};
  for (const it of items) {
    const id = String((it && it.id) || '');
    const qty = Number(it && it.qty);
    if (!id || !Number.isInteger(qty) || qty < 1 || qty > CONFIG.MAX_QTY_PER_ITEM) return fail_('商品數量不正確');
    qtyById[id] = Math.min(CONFIG.MAX_QTY_PER_ITEM, (qtyById[id] || 0) + qty);
  }
  const catalog = {};
  getProducts_().forEach(p => { catalog[p.id] = p; });
  const lines = [];
  for (const id in qtyById) {
    const p = catalog[id];
    if (!p) return fail_('有商品已下架，請重新整理頁面後再試');
    lines.push({ id: id, name: p.name, price: p.price, qty: qtyById[id], subtotal: p.price * qtyById[id] });
  }
  const total = lines.reduce((s, l) => s + l.subtotal, 0);
  const count = lines.reduce((s, l) => s + l.qty, 0);

  // 防止連續重複送出
  const cache = CacheService.getScriptCache();
  const rlKey = 'rl_' + phone.replace(/\D/g, '');
  if (cache.get(rlKey)) return fail_('剛剛已收到您的訂單，請稍候再送出');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_('目前訂單較多，請稍後再試');
  let orderId;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    orderId = nextOrderId_();
    const now = new Date();
    const summary = lines.map(l => l.name + ' x' + l.qty).join('、');
    ss.getSheetByName(CONFIG.SHEET_ORDERS).appendRow([
      orderId, now, CONFIG.STATUSES[0], method, name, "'" + phone, address, summary, count, total, note
    ]);
    const itemSheet = ss.getSheetByName(CONFIG.SHEET_ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, lines.length, ITEM_HEADERS.length)
      .setValues(lines.map(l => [orderId, l.id, l.name, l.price, l.qty, l.subtotal]));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put(rlKey, '1', CONFIG.RATE_LIMIT_SECONDS);

  notify_(orderId, method, name, phone, address, note, lines, total);
  return { ok: true, orderId: orderId, method: method, count: count, total: total };
}

/** 訂單編號：DD + 日期 + 當日流水號，例如 DD20261116-0007。須在鎖內呼叫。 */
function nextOrderId_() {
  const date = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + date;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return CONFIG.ORDER_PREFIX + date + '-' + ('000' + n).slice(-4);
}

/** 查詢訂單：需同時提供訂單編號與下單電話，避免他人亂猜編號看到資料。 */
function lookupOrder_(id, phone) {
  id = String(id || '').trim();
  const digits = String(phone || '').replace(/\D/g, '');
  if (!id || digits.length < 8) return { ok: false, error: '請提供訂單編號與電話' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ORDERS);
  const found = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  if (found) {
    const r = sh.getRange(found.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0];
    if (String(r[5]).replace(/\D/g, '') === digits) {
      return { ok: true, orderId: r[0], status: r[2], method: r[3], summary: r[7], count: r[8], total: r[9],
        createdAt: Utilities.formatDate(new Date(r[1]), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm') };
    }
  }
  return { ok: false, error: '查無此訂單' };
}

function notify_(orderId, method, name, phone, address, note, lines, total) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  try {
    const body = [
      '訂單編號：' + orderId, '取餐方式：' + method, '姓名：' + name, '電話：' + phone,
      address ? '地址：' + address : '', note ? '備註：' + note : '', '',
      lines.map(l => '・' + l.name + ' x' + l.qty + '　NT$ ' + l.subtotal).join('\n'), '',
      '總額：NT$ ' + total
    ].join('\n');
    MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '【豆豆貼】新訂單 ' + orderId, body);
  } catch (err) {
    console.error('通知信寄送失敗', err);
  }
}

/* ------------------------------------------------------------------ 工具 */

/** 去頭尾空白、限制長度，並擋掉會被試算表當成公式的開頭字元。 */
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s) && !/^\+?[0-9() -]+$/.test(s)) s = "'" + s;
  return s;
}

function fail_(msg) {
  return { ok: false, error: msg };
}

/** 在編輯器中手動執行，可快速測試建單流程（會真的寫入一筆測試訂單）。 */
function testCreateOrder() {
  const res = createOrder_({ method: '外帶', name: '測試', phone: '0912345678', items: [{ id: 'p1', qty: 2 }, { id: 'p7', qty: 1 }] });
  console.log(JSON.stringify(res));
}
