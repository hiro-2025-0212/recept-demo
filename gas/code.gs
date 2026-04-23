// =============================================
// 設定値はすべてスクリプトプロパティから取得する
// Apps Script エディタ → プロジェクトの設定 → スクリプトプロパティ で登録
//
//   GCP_PROJECT_ID  : Google Cloud プロジェクトID
//   VERTEX_LOCATION : Vertex AI のリージョン（例: asia-northeast1）
//   VERTEX_MODEL    : モデル名（例: gemini-2.5-flash）
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

// --- 読み取り処理 ---
function handleRead(body) {
  if (!body.image) {
    return jsonResponse({ success: false, error: 'image フィールドがありません' });
  }
  var base64Data = body.image.replace(/^data:[^;]+;base64,/, '');
  var fileName = body.fileName || 'receipt.jpg';
  var mimeType = body.mimeType || guessMimeType(fileName);
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

  var aiResult = callGemini(base64Data, mimeType);

  return jsonResponse({
    success: true,
    entries: aiResult.entries || [aiResult]
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

// --- Gemini API で領収書を読み取る ---
function callGemini(base64Data, mimeType) {
  var projectId = getConfig('GCP_PROJECT_ID');
  var location = PropertiesService.getScriptProperties().getProperty('VERTEX_LOCATION') || 'asia-northeast1';
  var model = PropertiesService.getScriptProperties().getProperty('VERTEX_MODEL') || 'gemini-2.5-flash';
  var endpoint = 'https://' + location + '-aiplatform.googleapis.com/v1/projects/' + projectId
    + '/locations/' + location + '/publishers/google/models/' + model + ':generateContent';
  var token = ScriptApp.getOAuthToken();

  var prompt = 'この画像またはPDFを分析して、経費情報をJSON形式で返してください。\n'
    + 'JSONのみを返し、他のテキストは含めないでください。\n\n'
    + '## 画像の種類を判別してください\n\n'
    + '### 通常の領収書・レシートの場合\n'
    + '1件分の情報を返してください:\n'
    + '{"entries":[{"date":"YYYY-MM-DD","amount":"数値のみ","category":"費目"}]}\n\n'
    + '### ICカード（PASMO/Suica等）の利用履歴の場合\n'
    + 'マーカーやペンで色付け・印をつけた行だけを読み取り、複数件返してください。\n'
    + '色付けされていない行は無視してください。\n'
    + '{"entries":[{"date":"YYYY-MM-DD","amount":"数値のみ","category":"交通費（電車）"},{"date":"YYYY-MM-DD","amount":"数値のみ","category":"交通費（電車）"}]}\n\n'
    + '## 費目の選択肢（この中から選ぶ）:\n'
    + CATEGORIES.join('\n')
    + '\n\n'
    + '## ルール:\n'
    + '- 金額は支払った合計金額を正の整数で返す（カンマ・円・マイナス記号は不要）\n'
    + '- 負の数や0は返さない。読み取れない場合は空文字にする\n'
    + '- ICカード履歴の場合も運賃の金額を正の整数で返す\n'
    + '- 日付はYYYY-MM-DD形式\n'
    + '- ICカード履歴の場合、費目は「交通費（電車）」にする\n'
    + '- 読み取れない項目は空文字にする\n'
    + '- 必ず{"entries":[...]}の形式で返す';

  var payload = {
    contents: [{
      role: 'user',
      parts: [
        { text: prompt },
        {
          inlineData: {
            mimeType: mimeType || 'image/jpeg',
            data: base64Data
          }
        }
      ]
    }]
  };

  var options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token
    },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var MAX_GEMINI_RETRIES = 3;
  var json = null;

  for (var retry = 0; retry <= MAX_GEMINI_RETRIES; retry++) {
    var response = UrlFetchApp.fetch(endpoint, options);
    var rawText = response.getContentText();
    var httpCode = response.getResponseCode();
    Logger.log('Gemini HTTP ' + httpCode + ' (試行' + (retry + 1) + ')');

    json = JSON.parse(rawText);

    if (json.error) {
      var errCode = json.error.code || 0;
      var errMsg = json.error.message || '';
      Logger.log('Gemini API エラー: ' + errCode + ' ' + errMsg);

      // 429=レート制限, 503=サーバー過負荷 → リトライ
      if ((errCode === 429 || errCode === 503) && retry < MAX_GEMINI_RETRIES) {
        var waitSec = Math.min(10 + retry * 10, 30); // 10秒, 20秒, 30秒
        Logger.log('レート制限: ' + waitSec + '秒待機して再試行...');
        Utilities.sleep(waitSec * 1000);
        continue;
      }
      throw new Error('Vertex AI APIエラー: ' + errMsg);
    }

    break; // 成功
  }

  if (!json.candidates || json.candidates.length === 0) {
    Logger.log('Gemini: candidates が空です');
    throw new Error('Gemini が応答を返しませんでした');
  }

  // gemini-2.5-flash は思考パートを含む場合があるため、
  // 全パートからテキストを結合して JSON を探す
  var parts = json.candidates[0].content.parts;
  var allText = '';
  for (var p = 0; p < parts.length; p++) {
    if (parts[p].text && !parts[p].thought) {
      allText += parts[p].text;
    }
  }

  // 思考パート以外にテキストがない場合、全パートから探す
  if (!allText) {
    for (var p2 = 0; p2 < parts.length; p2++) {
      if (parts[p2].text) {
        allText += parts[p2].text;
      }
    }
  }

  Logger.log('Gemini text: ' + allText.substring(0, 500));

  var jsonMatch = allText.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    Logger.log('JSON抽出失敗。テキスト全文: ' + allText);
    throw new Error('Gemini応答からJSONを抽出できませんでした');
  }

  try {
    var result = JSON.parse(jsonMatch[0]);

    if (!result.entries || !Array.isArray(result.entries)) {
      result = { entries: [result] };
    }

    for (var i = 0; i < result.entries.length; i++) {
      var entry = result.entries[i];
      if (CATEGORIES.indexOf(entry.category) === -1) {
        entry.category = '';
      }
      var amt = String(entry.amount || '').replace(/[,\s円¥\\-]/g, '');
      var num = parseInt(amt, 10);
      entry.amount = (num > 0 && isFinite(num)) ? String(num) : '';
    }

    Logger.log('Gemini 解析成功: ' + result.entries.length + '件');
    return result;
  } catch (parseErr) {
    Logger.log('JSONパース失敗: ' + jsonMatch[0]);
    throw new Error('Gemini応答のJSONパースに失敗: ' + parseErr.message);
  }
}

// --- JSON レスポンスを返す ---
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 初期セットアップ用 =====
// スプレッドシートに費目シート（自動フィルタ）を一括作成する
// Apps Script エディタから手動で1回だけ実行してください
function setupCategorySheets() {
  var spreadsheetId = getConfig('SPREADSHEET_ID');
  var ss = SpreadsheetApp.openById(spreadsheetId);

  var master = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!master) {
    master = ss.insertSheet(MASTER_SHEET_NAME, 0);
    master.appendRow(['登録日時', '日付', '費目', '金額', '画像URL']);
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
  }

  Logger.log('セットアップ完了: マスターシート + ' + CATEGORIES.length + '個の費目シートを作成しました');
}

// ===== テスト用 =====
function testConfig() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('設定済みプロパティ: ' + Object.keys(props).join(', '));
  Logger.log('GCP_PROJECT_ID: ' + (props['GCP_PROJECT_ID'] ? '設定済み' : '未設定'));
  Logger.log('VERTEX_LOCATION: ' + (props['VERTEX_LOCATION'] ? props['VERTEX_LOCATION'] : '未設定（asia-northeast1を利用）'));
  Logger.log('VERTEX_MODEL: ' + (props['VERTEX_MODEL'] ? props['VERTEX_MODEL'] : '未設定（gemini-2.5-flashを利用）'));
  Logger.log('SPREADSHEET_ID: ' + (props['SPREADSHEET_ID'] ? '設定済み' : '未設定'));
}

