// ============================================================
// mofmofu 工程管理 — Google Apps Script バックエンド (v5)
// v5: 工場×商品種別の完全分離管理（12商品エントリ）
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
    if (payload.events)                   return handleLineWebhook(payload);
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
// getMasters
// ============================================================

function getMasters() {
  const ss       = SpreadsheetApp.openById(SS_ID);
  const products = sheetToObjects(ss.getSheetByName('Products'));
  const skus     = sheetToObjects(ss.getSheetByName('SKUs'));
  const stages   = sheetToObjects(ss.getSheetByName('Stages'));
  const snapshot = buildSnapshot(ss);

  stages.forEach(function(s) {
    s.order    = Number(s.order);
    s.extInput = (s.extInput === true || s.extInput === 'TRUE');
    s.next     = s.next ? String(s.next).split(',').map(function(x){ return x.trim(); }).filter(Boolean) : [];
  });
  skus.forEach(function(s) {
    s.isSet  = (s.isSet  === true || s.isSet  === 'TRUE');
    s.shared = (s.shared === true || s.shared === 'TRUE');
    s.components = s.components
      ? String(s.components).split(',').map(function(c){ return c.trim(); }).filter(Boolean)
      : [];
  });

  return { products, skus, stages, snapshot };
}

// ============================================================
// getLogs
// ============================================================

function getLogsData(limit) {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Log');
  if (!sheet || sheet.getLastRow() < 2) return { logs: [] };
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const logs    = rows.slice(1).reverse().slice(0, limit).map(function(r) {
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
// buildSnapshot
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

  const factoryStages = ['神木さん', '西尾さん', '平本靴下さん'];
  const processStages = ['エルアイシー(未縫製)', '池本さん', '刑務所', 'エルアイシー(加工前)', '洗濯', '内職', '実在庫'];

  Object.keys(snapshot).forEach(function(sku) {
    const s           = snapshot[sku];
    const factoryKg   = factoryStages.reduce(function(sum, f) { return sum + (s[f] || 0); }, 0);
    const factoryPcs  = Math.round(factoryKg / KG_PER_PCS);
    const stagePcs    = processStages.reduce(function(sum, st) { return sum + (s[st] || 0); }, 0);
    const processTotal = factoryPcs + stagePcs;
    const backlog     = s['backlog'] || 0;
    const target      = s['target']  || 0;
    const shipped     = s['出荷数']  || 0;
    const physStock   = s['実在庫']  || 0;

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
// recordEntry
// ============================================================

function recordEntry(payload) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    appendLog(ss, payload);
    if      (payload.mode === '出荷数更新')   recordShipment(ss, payload.skuCode, Number(payload.qty));
    else if (payload.mode === '生産目標更新') setTarget(ss, payload.skuCode, Number(payload.qty));
    else if (payload.mode === '在庫修正')     setStageValue(ss, payload.skuCode, payload.dest, Number(payload.qty));
    else                                       updateSnapshot(ss, payload);
    const comps = getSetComponents(ss, payload.skuCode);
    if (comps.length > 0) comps.forEach(function(c) { checkAndAlert(c, c); });
    else                   checkAndAlert(payload.skuCode, payload.skuName);
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
    sendLineAlert('⚠️ 不足アラート\n' + skuName + '\n工程合計: ' + d.processTotal + '個 / 出荷残: ' + d.backlog + '個\n不足: ' + d.diff + '個\n追加発注を検討してください');
  }
}

function sendLineAlert(text) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({ to: GROUP_ID, messages: [{ type: 'text', text: text }] }),
      muteHttpExceptions: true
    });
  } catch (e) { Logger.log('LINE送信エラー: ' + e); }
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
  const sheet     = ss.getSheetByName('Snapshot');
  const qty       = Number(p.qty);
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
    if (data[i][0] === skuCode && data[i][1] === 'target') { sheet.getRange(i+1,3).setValue(qty); return; }
  }
  sheet.appendRow([skuCode, 'target', qty]);
}

function setStageValue(ss, skuCode, stage, qty) {
  const sheet = ss.getSheetByName('Snapshot');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === skuCode && data[i][1] === stage) { sheet.getRange(i+1,3).setValue(qty); return; }
  }
  sheet.appendRow([skuCode, stage, qty]);
}

function recordShipment(ss, skuCode, qty) {
  const components = getSetComponents(ss, skuCode);
  const sheet = ss.getSheetByName('Snapshot');
  if (components.length > 0) {
    components.forEach(function(c) { addToSnapshot(sheet,c,'実在庫',-qty); addToSnapshot(sheet,c,'出荷数',qty); });
    return;
  }
  addToSnapshot(sheet, skuCode, '実在庫', -qty);
  addToSnapshot(sheet, skuCode, '出荷数',  qty);
}

function addToSnapshot(sheet, skuCode, stage, delta) {
  if (!stage || stage === '—') return;
  if (stage === '出荷') { if (delta <= 0) return; stage = '出荷数'; }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === skuCode && data[i][1] === stage) {
      sheet.getRange(i+1,3).setValue(Math.max(0, Number(data[i][2]) + delta)); return;
    }
  }
  if (delta > 0) sheet.appendRow([skuCode, stage, delta]);
}

// ============================================================
// setup2026v5 — Products/SKUs/Stagesシートを再構築
// GASエディタから手動で1回実行
// ============================================================

function setup2026v5() {
  setupProducts2026v5();
  setupSKUs2026v5();
  setupStages2026();
  Logger.log('✅ setup2026v5 完了');
}

function setupProducts2026v5() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('Products');
  if (!sheet) sheet = ss.insertSheet('Products');
  sheet.clearContents();
  sheet.appendRow(['code', 'name', 'type', 'factory', 'factoryStage']);

  const factories = [
    ['kamiki',   '神木',     '神木さん'],
    ['nishio',   '西尾',     '西尾さん'],
    ['hiramoto', '平本靴下', '平本靴下さん'],
  ];
  const types = [
    ['tubu',   'つぶつぶベビーレッグ'],
    ['kusumi', 'くすみベビーレッグ'],
    ['linen',  'リネンベビーレッグ'],
    ['usude',  '薄手ベビーレッグ'],
  ];

  types.forEach(function(t) {
    factories.forEach(function(f) {
      sheet.appendRow([
        'babyleg-' + t[0] + '-' + f[0],
        t[1] + '（' + f[1] + '）',
        t[0], f[0], f[2]
      ]);
    });
  });
  sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f3f0e8');
  Logger.log('✅ Products v5更新完了 (12件)');
}

function setupSKUs2026v5() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('SKUs');
  if (!sheet) sheet = ss.insertSheet('SKUs');
  sheet.clearContents();
  sheet.appendRow(['product', 'code', 'name', 'isSet', 'components', 'shared']);

  const factories = [
    { code:'kamiki',   suffix:'KMK' },
    { code:'nishio',   suffix:'WSO' },
    { code:'hiramoto', suffix:'HRT' },
  ];

  // ベースSKU定義
  const sharedBase = [['BL-KN','キナリ'],['BL-MOM','杢オートミール']];
  const tubuBase   = [['TBL-KNT','キナリツブ'],['TBL-SRT','シロツブ'],['TBL-CNP','カラーネップ']];
  const kusumiBase = [['KBL-DP','ダスティピンク'],['KBL-GR','グレー'],['KBL-OL','オリーブ'],['KBL-CG','チャコールグレー'],['KBL-MC','杢チャコール'],['KBL-MG','杢グレー']];
  const linenBase  = [['LBL-IV','【リネン】くすみアイボリー'],['LBL-MP','【リネン】ミルキーピンク'],['LBL-GR','【リネン】グレー'],['LBL-MB','【リネン】ミルクベージュ'],['LBL-MT','【リネン】ミント']];
  const usudeBase  = [['UBL-KN','【薄手】キナリ'],['UBL-OAT','【薄手】オートミール'],['UBL-PP','【薄手】ベビーピンク'],['UBL-BB','【薄手】ベビーブルー'],['UBL-SG','【薄手】シルバーグレー'],['UBL-CG','【薄手】チャコールグレー'],['UBL-WH','【薄手】ホワイト']];

  factories.forEach(function(f) {
    const s = '-' + f.suffix;
    const pf = 'babyleg-';
    // つぶつぶ（共有 + 専用）
    sharedBase.forEach(function(b) { sheet.appendRow([pf+'tubu-'+f.code,  b[0]+s, b[1], false, '', true]); });
    tubuBase.forEach(  function(b) { sheet.appendRow([pf+'tubu-'+f.code,  b[0]+s, b[1], false, '', false]); });
    // くすみ（共有 + 専用）
    sharedBase.forEach(function(b) { sheet.appendRow([pf+'kusumi-'+f.code, b[0]+s, b[1], false, '', true]); });
    kusumiBase.forEach(function(b) { sheet.appendRow([pf+'kusumi-'+f.code, b[0]+s, b[1], false, '', false]); });
    // リネン
    linenBase.forEach( function(b) { sheet.appendRow([pf+'linen-'+f.code,  b[0]+s, b[1], false, '', false]); });
    // 薄手
    usudeBase.forEach( function(b) { sheet.appendRow([pf+'usude-'+f.code,  b[0]+s, b[1], false, '', false]); });
  });

  sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f0e8');
  const count = sheet.getLastRow() - 1;
  Logger.log('✅ SKUs v5更新完了: ' + count + '件');
}

function setupStages2026() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('Stages');
  if (!sheet) sheet = ss.insertSheet('Stages');
  sheet.clearContents();
  sheet.appendRow(['code','name','order','unit','extInput','next']);
  [
    ['S01',  '神木さん',             1,'kg',true,  'エルアイシー(未縫製)'],
    ['S01B', '西尾さん',             1,'kg',true,  'エルアイシー(未縫製)'],
    ['S01C', '平本靴下さん',         1,'kg',true,  'エルアイシー(未縫製)'],
    ['S02',  'エルアイシー(未縫製)',2,'個',false,'池本さん,刑務所,出荷'],
    ['S03',  '池本さん',             3,'個',false,'エルアイシー(加工前),出荷'],
    ['S04',  '刑務所',               3,'個',false,'エルアイシー(加工前),出荷'],
    ['S05',  'エルアイシー(加工前)',4,'個',false,'洗濯,内職,出荷'],
    ['S05B', '洗濯',                 5,'個',false,'内職,出荷'],
    ['S06',  '内職',                 6,'個',false,'実在庫,出荷'],
    ['S07',  '実在庫',               7,'個',false,'出荷'],
  ].forEach(function(r) { sheet.appendRow(r); });
  sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#f3f0e8');
  Logger.log('✅ Stages更新完了');
}

// ============================================================
// addFactorySuffix2026 — 既存Snapshotコードに -KMK を付与
// （既存データはすべて神木さん分なので一括で -KMK を追加）
// GASエディタから手動で1回実行
// ============================================================

function addFactorySuffix2026() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Snapshot');
  if (!sheet) { Logger.log('❌ Snapshotシートが見つかりません'); return; }

  const data    = sheet.getDataRange().getValues();
  const SUFFIXES = ['-KMK', '-WSO', '-HRT'];
  let changed   = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][0] || '');
    if (!code) continue;
    // すでにサフィックスがついている行はスキップ
    if (SUFFIXES.some(function(s) { return code.endsWith(s); })) { skipped++; continue; }
    // 特殊キー（backlog/target/出荷数）はスキップ
    if (['backlog','target','出荷数'].indexOf(data[i][1]) !== -1) {
      // コードにサフィックスを付ける（targetやbacklogも工場別に管理）
    }
    sheet.getRange(i + 1, 1).setValue(code + '-KMK');
    changed++;
  }

  Logger.log('✅ factory suffix追加完了: ' + changed + '件変更 / ' + skipped + '件スキップ（既存サフィックス）');
}

// ============================================================
// migrateSnapshot2026 — 旧コード → 新コード（前回分）
// ============================================================

function migrateSnapshot2026() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Snapshot');
  if (!sheet) { Logger.log('❌ Snapshotシートが見つかりません'); return; }
  const codeMap = {
    'TBL-KN':'BL-KN','TBL-MOM':'BL-MOM',
    'TBL-CG':'KBL-CG','TBL-GR':'KBL-GR','TBL-DP':'KBL-DP',
    'TBL-OL':'KBL-OL','TBL-MG':'KBL-MG','TBL-MC':'KBL-MC',
  };
  const data = sheet.getDataRange().getValues();
  let changed = 0, deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const code = String(data[i][0] || '');
    if (!code) continue;
    if (code.startsWith('TBL-WSO-')) { sheet.deleteRow(i+1); deleted++; continue; }
    if (codeMap[code]) { sheet.getRange(i+1,1).setValue(codeMap[code]); changed++; }
  }
  Logger.log('✅ 移行完了: ' + changed + '件書き換え / ' + deleted + '件削除');
}
