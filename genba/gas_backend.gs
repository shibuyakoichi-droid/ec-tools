// ============================================================
// mofmofu 工程管理 — Google Apps Script バックエンド (v4)
// v4: 商品4種対応・平本靴下追加・洗濯工程追加・ログAPI追加
// ============================================================

const SS_ID = '1l9QyWxdYcyTTqR7ZMBs-VQS0rj8c2qUp9q399jmy-30';

const LINE_TOKEN = 'D+dfSKKq9S3ZXx2xUDAE9r1b8Syt1tgRyGzddtMKHuhMJTDV1EDg+TsZNUfmW+XyszQpXshH9n9xBZ3XA5naGpcWT4a/xjl+bNtPxha2HSFbORUNrbZzYMJ/2Tl382QkDejinEoA8R0wGbVoeAdCzgdB04t89/1o/w1cDnyilFU=';
const GROUP_ID   = 'Cba1029b36b497b2840668bb08b7e7405';
const KG_PER_PCS = 0.030;

// ============================================================
// HTTP ハンドラ
// ============================================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'getMasters') return respond(getMasters());
    if (action === 'getLogs')    return respond(getLogsData(Number(e.parameter.limit) || 50));
    return respond({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.events)              return handleLineWebhook(payload);
    if (payload.action === 'recordEntry') return respond(recordEntry(payload));
    return respond({ error: 'Unknown action: ' + payload.action });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

function handleLineWebhook(payload) {
  payload.events.forEach(function(event) {
    if (event.source && event.source.type === 'group') {
      const gid = event.source.groupId;
      Logger.log('GROUP_ID: ' + gid);
      const ss = SpreadsheetApp.openById(SS_ID);
      let sheet = ss.getSheetByName('_GroupIdLog');
      if (!sheet) sheet = ss.insertSheet('_GroupIdLog');
      sheet.appendRow([new Date(), gid, event.source.userId || '']);
    }
  });
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// getMasters — 商品・SKU・工程・在庫スナップショットを返す
// ============================================================

function getMasters() {
  const ss = SpreadsheetApp.openById(SS_ID);

  const products = sheetToObjects(ss.getSheetByName('Products'));
  const skus     = sheetToObjects(ss.getSheetByName('SKUs'));
  const stages   = sheetToObjects(ss.getSheetByName('Stages'));
  const snapshot = buildSnapshot(ss);

  stages.forEach(function(s) {
    s.order    = Number(s.order);
    s.extInput = (s.extInput === true || s.extInput === 'TRUE');
    s.next     = s.next ? String(s.next).split(',').map(function(x) { return x.trim(); }).filter(Boolean) : [];
  });

  skus.forEach(function(s) {
    s.isSet  = (s.isSet === true || s.isSet === 'TRUE');
    s.shared = (s.shared === true || s.shared === 'TRUE');
    s.components = s.components
      ? String(s.components).split(',').map(function(c) { return c.trim(); }).filter(Boolean)
      : [];
  });

  return { products, skus, stages, snapshot };
}

// ============================================================
// getLogs — 操作ログを返す
// ============================================================

function getLogsData(limit) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Log');
  if (!sheet || sheet.getLastRow() < 2) return { logs: [] };

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1).reverse().slice(0, limit);

  const logs = data.map(function(r) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = r[i]; });
    return obj;
  });

  return { logs: logs };
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function(r) { return r[0] !== '' && r[0] !== null; })
    .map(function(r) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = r[i]; });
      return obj;
    });
}

// ============================================================
// buildSnapshot — 在庫・工程合計・差分を生成
// ============================================================

function buildSnapshot(ss) {
  const snapSheet = ss.getSheetByName('Snapshot');
  const snapshot  = {};
  if (snapSheet && snapSheet.getLastRow() > 1) {
    snapSheet.getDataRange().getValues().slice(1).forEach(function(row) {
      const skuCode = row[0], stage = row[1], qty = row[2];
      if (!skuCode) return;
      if (!snapshot[skuCode]) snapshot[skuCode] = {};
      snapshot[skuCode][stage] = Number(qty) || 0;
    });
  }

  // v4: 平本靴下さん・洗濯を工程合計に含める
  const factoryStages  = ['神木さん', '西尾さん', '平本靴下さん'];
  const processStages  = ['エルアイシー(未縫製)', '池本さん', '刑務所', 'エルアイシー(加工前)', '洗濯', '内職', '実在庫'];

  Object.keys(snapshot).forEach(function(sku) {
    const s = snapshot[sku];
    const factoryKg  = factoryStages.reduce(function(sum, f) { return sum + (s[f] || 0); }, 0);
    const factoryPcs = Math.round(factoryKg / KG_PER_PCS);
    const stagePcs   = processStages.reduce(function(sum, st) { return sum + (s[st] || 0); }, 0);
    const processTotal = factoryPcs + stagePcs;
    const backlog  = s['backlog'] || 0;
    const target   = s['target']  || 0;
    const shipped  = s['出荷数']  || 0;
    const physStock = s['実在庫'] || 0;

    s.factoryKg    = factoryKg;
    s.factoryPcs   = factoryPcs;
    s.kamikiPcs    = factoryPcs;
    s.processTotal = processTotal;
    s.backlog      = backlog;
    s.target       = target;
    s.shipped      = shipped;
    s.physStock    = physStock;
    s.produced     = physStock + shipped;
    s.diff         = backlog - processTotal;
    s.progress     = target > 0 ? Math.round((physStock + shipped) / target * 100) : null;
  });

  return snapshot;
}

// ============================================================
// recordEntry — ログ記録 + スナップショット更新
// ============================================================

function recordEntry(payload) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    appendLog(ss, payload);

    if (payload.mode === '出荷数更新') {
      recordShipment(ss, payload.skuCode, Number(payload.qty));
    } else if (payload.mode === '生産目標更新') {
      setTarget(ss, payload.skuCode, Number(payload.qty));
    } else if (payload.mode === '在庫修正') {
      setStageValue(ss, payload.skuCode, payload.dest, Number(payload.qty));
    } else {
      updateSnapshot(ss, payload);
    }

    const comps = getSetComponents(ss, payload.skuCode);
    if (comps.length > 0) {
      comps.forEach(function(c) { checkAndAlert(c, c); });
    } else {
      checkAndAlert(payload.skuCode, payload.skuName);
    }

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function getSetComponents(ss, skuCode) {
  const skus = sheetToObjects(ss.getSheetByName('SKUs'));
  const sku  = skus.find(function(s) { return s.code === skuCode; });
  if (!sku || !(sku.isSet === true || sku.isSet === 'TRUE') || !sku.components) return [];
  return String(sku.components).split(',').map(function(c) { return c.trim(); }).filter(Boolean);
}

function checkAndAlert(skuCode, skuName) {
  if (!LINE_TOKEN || !GROUP_ID) return;
  const ss   = SpreadsheetApp.openById(SS_ID);
  const snap = buildSnapshot(ss);
  if (snap[skuCode] && snap[skuCode].diff > 0) {
    const d = snap[skuCode];
    sendLineAlert(
      '⚠️ 不足アラート\n' + skuName + '\n' +
      '工程合計: ' + d.processTotal + '個 / 出荷残: ' + d.backlog + '個\n' +
      '不足: ' + d.diff + '個\n追加発注を検討してください'
    );
  }
}

function sendLineAlert(text) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({ to: GROUP_ID, messages: [{ type: 'text', text: text }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('LINE送信エラー: ' + e);
  }
}

function sendDailyReminder() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M月d日');
  sendLineAlert('📦 ' + today + ' 入力リマインド\n本日の進捗を入力してください。\n▼ https://liff.line.me/2009935318-eq7U4Fuc');
}

function appendLog(ss, p) {
  ss.getSheetByName('Log').appendRow([
    p.timestamp, p.user, p.userId, p.mode, p.product,
    p.skuCode, p.skuName, p.source || '', p.dest || '',
    Number(p.qty), p.unit, p.memo || ''
  ]);
}

// ============================================================
// Snapshot 更新
// ============================================================

function updateSnapshot(ss, p) {
  const sheet   = ss.getSheetByName('Snapshot');
  const qty     = Number(p.qty);
  const factories = ['神木さん', '西尾さん', '平本靴下さん'];

  if (p.mode === '外部入荷') {
    addToSnapshot(sheet, p.skuCode, p.dest, qty);
  } else if (p.mode === '工程移動') {
    if (factories.indexOf(p.source) !== -1) {
      const kgDelta = Math.round(qty * KG_PER_PCS * 1000) / 1000;
      addToSnapshot(sheet, p.skuCode, p.source, -kgDelta);
      addToSnapshot(sheet, p.skuCode, p.dest,    qty);
    } else {
      addToSnapshot(sheet, p.skuCode, p.source, -qty);
      addToSnapshot(sheet, p.skuCode, p.dest,    qty);
    }
  }
}

function setTarget(ss, skuCode, qty) {
  const sheet = ss.getSheetByName('Snapshot');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === skuCode && data[i][1] === 'target') {
      sheet.getRange(i + 1, 3).setValue(qty); return;
    }
  }
  sheet.appendRow([skuCode, 'target', qty]);
}

function setStageValue(ss, skuCode, stage, qty) {
  const sheet = ss.getSheetByName('Snapshot');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === skuCode && data[i][1] === stage) {
      sheet.getRange(i + 1, 3).setValue(qty); return;
    }
  }
  sheet.appendRow([skuCode, stage, qty]);
}

function recordShipment(ss, skuCode, qty) {
  const components = getSetComponents(ss, skuCode);
  const sheet = ss.getSheetByName('Snapshot');
  if (components.length > 0) {
    components.forEach(function(c) {
      addToSnapshot(sheet, c, '実在庫', -qty);
      addToSnapshot(sheet, c, '出荷数',  qty);
    });
    return;
  }
  addToSnapshot(sheet, skuCode, '実在庫', -qty);
  addToSnapshot(sheet, skuCode, '出荷数',  qty);
}

function addToSnapshot(sheet, skuCode, stage, delta) {
  if (!stage || stage === '—') return;
  if (stage === '出荷') {
    if (delta <= 0) return;
    stage = '出荷数';
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === skuCode && data[i][1] === stage) {
      const newVal = Math.max(0, Number(data[i][2]) + delta);
      sheet.getRange(i + 1, 3).setValue(newVal); return;
    }
  }
  if (delta > 0) sheet.appendRow([skuCode, stage, delta]);
}

// ============================================================
// setup2026 — v4構成でProducts/SKUs/Stagesシートを再構築
// GASエディタから手動で1回実行してください
// ============================================================

function setup2026() {
  setupProducts2026();
  setupSKUs2026();
  setupStages2026();
  Logger.log('✅ setup2026 完了');
}

function setupProducts2026() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('Products');
  if (!sheet) sheet = ss.insertSheet('Products');
  sheet.clearContents();
  sheet.appendRow(['code', 'name', 'type']);
  [
    ['babyleg-tubu',  'つぶつぶベビーレッグ', 'tubu'],
    ['babyleg-kusumi','くすみベビーレッグ',   'kusumi'],
    ['babyleg-linen', 'リネンベビーレッグ',   'linen'],
    ['babyleg-usude', '薄手ベビーレッグ',     'usude'],
  ].forEach(function(r) { sheet.appendRow(r); });
  sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#f3f0e8');
  Logger.log('✅ Products2026更新完了');
}

function setupSKUs2026() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('SKUs');
  if (!sheet) sheet = ss.insertSheet('SKUs');
  sheet.clearContents();
  sheet.appendRow(['product', 'code', 'name', 'isSet', 'components', 'shared']);

  const rows = [
    // ── 共有SKU（つぶつぶ・くすみ両方に表示、在庫は1つ）──
    ['babyleg-tubu', 'BL-KN',  'キナリ',      false, '', true],
    ['babyleg-tubu', 'BL-MOM', '杢オートミール', false, '', true],
    // ── つぶつぶ専用 ──
    ['babyleg-tubu', 'TBL-KNT', 'キナリツブ',   false, '', false],
    ['babyleg-tubu', 'TBL-SRT', 'シロツブ',     false, '', false],
    ['babyleg-tubu', 'TBL-CNP', 'カラーネップ', false, '', false],
    // ── くすみ専用 ──
    ['babyleg-kusumi', 'KBL-DP', 'ダスティピンク',   false, '', false],
    ['babyleg-kusumi', 'KBL-GR', 'グレー',           false, '', false],
    ['babyleg-kusumi', 'KBL-OL', 'オリーブ',         false, '', false],
    ['babyleg-kusumi', 'KBL-CG', 'チャコールグレー', false, '', false],
    ['babyleg-kusumi', 'KBL-MC', '杢チャコール',     false, '', false],
    ['babyleg-kusumi', 'KBL-MG', '杢グレー',         false, '', false],
    // ── リネン ──
    ['babyleg-linen', 'LBL-IV', '【リネン】くすみアイボリー', false, '', false],
    ['babyleg-linen', 'LBL-MP', '【リネン】ミルキーピンク',  false, '', false],
    ['babyleg-linen', 'LBL-GR', '【リネン】グレー',          false, '', false],
    ['babyleg-linen', 'LBL-MB', '【リネン】ミルクベージュ',  false, '', false],
    ['babyleg-linen', 'LBL-MT', '【リネン】ミント',          false, '', false],
    // ── 薄手 ──
    ['babyleg-usude', 'UBL-KN',  '【薄手】キナリ',         false, '', false],
    ['babyleg-usude', 'UBL-OAT', '【薄手】オートミール',   false, '', false],
    ['babyleg-usude', 'UBL-PP',  '【薄手】ベビーピンク',   false, '', false],
    ['babyleg-usude', 'UBL-BB',  '【薄手】ベビーブルー',   false, '', false],
    ['babyleg-usude', 'UBL-SG',  '【薄手】シルバーグレー', false, '', false],
    ['babyleg-usude', 'UBL-CG',  '【薄手】チャコールグレー', false, '', false],
    ['babyleg-usude', 'UBL-WH',  '【薄手】ホワイト',       false, '', false],
  ];
  rows.forEach(function(r) { sheet.appendRow(r); });
  sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f0e8');
  Logger.log('✅ SKUs2026更新完了: ' + rows.length + '件');
}

function setupStages2026() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('Stages');
  if (!sheet) sheet = ss.insertSheet('Stages');
  sheet.clearContents();
  sheet.appendRow(['code', 'name', 'order', 'unit', 'extInput', 'next']);
  [
    ['S01',  '神木さん',             1,  'kg', true,  'エルアイシー(未縫製)'],
    ['S01B', '西尾さん',             1,  'kg', true,  'エルアイシー(未縫製)'],
    ['S01C', '平本靴下さん',         1,  'kg', true,  'エルアイシー(未縫製)'],
    ['S02',  'エルアイシー(未縫製)', 2,  '個', false, '池本さん,刑務所,出荷'],
    ['S03',  '池本さん',             3,  '個', false, 'エルアイシー(加工前),出荷'],
    ['S04',  '刑務所',               3,  '個', false, 'エルアイシー(加工前),出荷'],
    ['S05',  'エルアイシー(加工前)', 4,  '個', false, '洗濯,内職,出荷'],
    ['S05B', '洗濯',                 5,  '個', false, '内職,出荷'],
    ['S06',  '内職',                 6,  '個', false, '実在庫,出荷'],
    ['S07',  '実在庫',               7,  '個', false, '出荷'],
  ].forEach(function(r) { sheet.appendRow(r); });
  sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f0e8');
  Logger.log('✅ Stages2026更新完了');
}

// ============================================================
// resetSnapshot2026 — 初期データ投入（手動1回実行）
// ============================================================

function resetSnapshot2026() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Snapshot');
  const last  = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 3).clearContent();
  // 既存データをBL-KN / BL-MOM / TBL-* / KBL-* に移行してください
  Logger.log('Snapshot クリア完了。既存在庫データを新SKUコードで再登録してください。');
}
