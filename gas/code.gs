// =============================================
// 設定値はすべてスクリプトプロパティから取得する
// Apps Script エディタ → プロジェクトの設定 → スクリプトプロパティ で登録
//
//   CLOUD_RUN_EXTRACT_URL   : Cloud Run /extract URL
//   CLOUD_RUN_SHARED_SECRET : Cloud Run 呼び出し用シークレット（任意）
//   SPREADSHEET_ID  : スプレッドシートID
// =============================================

var CATEGORIES = [
  'タクシー代',
  '新幹線',
  '交通費（電車）',
  '飲食',
  '駐車／ガソリン',
  'スーパー（社内飲み買い出し）',
  '雑費（消耗品・備品）',
  '諸会費（交流会費）'
];

var MASTER_SHEET_NAME = '全データ';
var MAX_IMAGE_BYTES = 4 * 1024 * 1024; // 4MB

// ===== ヘルパー：スクリプトプロパティ取得 =====
function getConfig(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('スクリプトプロパティ「' + key + '」が未設定です。プロジェクトの設定で登録してください。');
  }
  return value;
}

// --- ファイル名からMIMEタイプを推測 ---
function guessMimeType(fileName) {
  var ext = (fileName || '').toLowerCase().split('.').pop();
  var map = {
    'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'png': 'image/png', 'gif': 'image/gif',
    'webp': 'image/webp', 'bmp': 'image/bmp',
    'tiff': 'image/tiff', 'tif': 'image/tiff',
    'heic': 'image/heic', 'heif': 'image/heif',
    'avif': 'image/avif', 'pdf': 'application/pdf'
  };
  return map[ext] || 'image/jpeg';
}

function pad2(num) {
  return ('0' + num).slice(-2);
}

function isValidYmd(year, month, day) {
  var y = Number(year);
  var m = Number(month);
  var d = Number(day);
  if (!isFinite(y) || !isFinite(m) || !isFinite(d)) return false;
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  var dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === (m - 1) && dt.getDate() === d;
}

// AIが返した日付を YYYY-MM-DD に正規化する。
// 年がない場合は captureYear を使う。
function normalizeReceiptDate(rawDate, captureYear) {
  if (!rawDate) return '';
  var src = String(rawDate).trim();
  if (!src) return '';
  var fallbackYear = Number(captureYear) || new Date().getFullYear();

  // YYYY-MM-DD / YYYY/MM/DD
  var ymd = src.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (ymd) {
    var y = Number(ymd[1]);
    var m = Number(ymd[2]);
    var d = Number(ymd[3]);
    if (!isValidYmd(y, m, d)) return '';
    return y + '-' + pad2(m) + '-' + pad2(d);
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // MM/DD or M/D は「月/日」として扱う
  var md = src.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (md) {
    var mm = Number(md[1]);
    var dd = Number(md[2]);
    if (!isValidYmd(fallbackYear, mm, dd)) return '';
    if (new Date(fallbackYear, mm - 1, dd) > today && isValidYmd(fallbackYear - 1, mm, dd)) {
      return (fallbackYear - 1) + '-' + pad2(mm) + '-' + pad2(dd);
    }
    return fallbackYear + '-' + pad2(mm) + '-' + pad2(dd);
  }

  // M月D日
  var jp = src.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (jp) {
    var jm = Number(jp[1]);
    var jd = Number(jp[2]);
    if (!isValidYmd(fallbackYear, jm, jd)) return '';
    if (new Date(fallbackYear, jm - 1, jd) > today && isValidYmd(fallbackYear - 1, jm, jd)) {
      return (fallbackYear - 1) + '-' + pad2(jm) + '-' + pad2(jd);
    }
    return fallbackYear + '-' + pad2(jm) + '-' + pad2(jd);
  }

  return '';
}

// ===== メイン処理 =====
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'read') {
      return handleRead(body);
    } else if (action === 'save') {
      return handleSave(body);
    } else {
      return jsonResponse({ success: false, error: '不明なアクション: ' + action });
    }
  } catch (err) {
    Logger.log('doPost エラー: ' + err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}

// --- 読み取り処理 ---
function handleRead(body) {
  if (!body.image) {
    return jsonResponse({ success: false, error: 'image フィールドがありません' });
  }
  var base64Data = body.image.replace(/^data:[^;]+;base64,/, '');
  var fileName = body.fileName || 'receipt.jpg';
  var mimeType = body.mimeType || guessMimeType(fileName);
  var captureYear = Number(body.captureYear) || new Date().getFullYear();
  if (mimeType === 'application/octet-stream') {
    mimeType = guessMimeType(fileName);
  }

  var byteSize = base64Data.length * 3 / 4;
  if (byteSize > MAX_IMAGE_BYTES) {
    return jsonResponse({
      success: false,
      error: 'ファイルサイズが大きすぎます（' + Math.round(byteSize / 1024 / 1024) + 'MB）。4MB以下にしてください。'
    });
  }

  var aiResult = callCloudRunExtractor(base64Data, mimeType, captureYear);
  var entries = aiResult.entries || [aiResult];
  for (var i = 0; i < entries.length; i++) {
    entries[i].date = normalizeReceiptDate(entries[i].date, captureYear);
  }

  return jsonResponse({
    success: true,
    entries: entries
  });
}

// --- 保存処理（マスターシートに一括保存） ---
function handleSave(body) {
  var entries = body.entries;

  if (!entries || entries.length === 0) {
    return jsonResponse({ success: false, error: '保存するデータがありません' });
  }

  var spreadsheetId = getConfig('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(spreadsheetId);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var master = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!master) {
      master = ss.insertSheet(MASTER_SHEET_NAME, 0);
      master.appendRow(['登録日時', '日付', '費目', '金額']);
      master.getRange('1:1').setFontWeight('bold');
    }

    var now = new Date();

    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      master.appendRow([
        now,
        entry.date,
        entry.category,
        Number(entry.amount)
      ]);
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({ success: true, saved: entries.length });
}

// --- Cloud Run に読み取りを委譲 ---
function callCloudRunExtractor(base64Data, mimeType, captureYear) {
  var endpoint = getConfig('CLOUD_RUN_EXTRACT_URL');
  var sharedSecret = PropertiesService.getScriptProperties().getProperty('CLOUD_RUN_SHARED_SECRET');

  var payload = {
    imageBase64: base64Data,
    mimeType: mimeType || 'image/jpeg',
    captureYear: captureYear || new Date().getFullYear()
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: sharedSecret ? { 'x-shared-secret': sharedSecret } : {}
  };

  var response = UrlFetchApp.fetch(endpoint, options);
  var responseCode = response.getResponseCode();
  var text = response.getContentText();
  var json = JSON.parse(text);

  if (responseCode >= 300 || !json.success) {
    throw new Error('Cloud Runエラー: ' + (json.error || text));
  }

  var result = { entries: json.entries || [] };
  for (var i = 0; i < result.entries.length; i++) {
    var entry = result.entries[i];
    if (CATEGORIES.indexOf(entry.category) === -1) {
      entry.category = '';
    }
    var amt = String(entry.amount || '').replace(/[,\s円¥\\-]/g, '');
    var num = parseInt(amt, 10);
    entry.amount = (num > 0 && isFinite(num)) ? String(num) : '';
  }
  return result;
}

// --- JSON レスポンスを返す ---
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 初期セットアップ用 =====
function setupCategorySheets() {
  var spreadsheetId = getConfig('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(spreadsheetId);

  var master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!master) {
    master = ss.insertSheet(MASTER_SHEET_NAME, 0);
    master.appendRow(['登録日時', '日付', '費目', '金額']);
    master.getRange('1:1').setFontWeight('bold');
  }

  for (var i = 0; i < CATEGORIES.length; i++) {
    var name = CATEGORIES[i];
    var sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
    }

    var formula = '=QUERY(\'' + MASTER_SHEET_NAME + '\'!A:D, "SELECT * WHERE C = \'' + name + '\' ORDER BY A DESC", 1)';
    sheet.getRange('A1').setFormula(formula);
    sheet.getRange('A1').setNote('この表は「' + MASTER_SHEET_NAME + '」シートから自動取得しています。編集は「' + MASTER_SHEET_NAME + '」シートで行ってください。');

    sheet.getRange('G1').setValue('月別集計');
    sheet.getRange('G2').setFormula(
      "=IFERROR(QUERY({ARRAYFORMULA(TEXT(B2:B,\"yyyy-mm\")), D2:D}, " +
      "\"select Col1, sum(Col2) where Col1 is not null and Col2 > 0 group by Col1 order by Col1 desc " +
      "label Col1 '月', sum(Col2) '合計金額'\", 0), {\"月\",\"合計金額\"})"
    );
    sheet.getRange('H:H').setNumberFormat('#,##0');
  }

  Logger.log('セットアップ完了: マスターシート + ' + CATEGORIES.length + '個の費目シートを作成しました');
}

// ===== テスト用 =====
function testConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('設定済みプロパティ: ' + Object.keys(props).join(', '));
  Logger.log('CLOUD_RUN_EXTRACT_URL: ' + (props['CLOUD_RUN_EXTRACT_URL'] ? '設定済み' : '未設定'));
  Logger.log('CLOUD_RUN_SHARED_SECRET: ' + (props['CLOUD_RUN_SHARED_SECRET'] ? '設定済み' : '未設定（認証なし）'));
  Logger.log('SPREADSHEET_ID: ' + (props['SPREADSHEET_ID'] ? '設定済み' : '未設定'));
}
