/**
 * ZOOM ROOM自動作成スクリプト
 * 
 * スプレッドシートから未作成のZOOM ROOM情報を読み取り、
 * ZOOM APIを使用して自動的にROOMを作成します。
 * 
 * 対象: starカラムが空白（★が付いていない）のレコード
 * 
 * ミーティングタイプ: 固定時間なしの定期ミーティング
 * - 1つのROOM IDと参加URLを継続して使い続けることができます
 * - 同じリンクを何度でも使用可能
 * 
 */

/* ============================================================
 * グローバル定数
 * ============================================================ */
const SPREADSHEET_ID = 'ここに合宿用スプレッドシートIDを入力してください';
const SHEET_ID = 9999999999; // シートID（gidの数値部分）を入力してください

// カラム名定義（ヘッダー名と完全一致させる）
const COLUMN_NAMES = {
  USER: 'user',
  TITLE: 'title',
  HOSTKEY: 'hostkey',
  STAR: 'star',
  MEETING_ID: 'meeting_id',
  JOIN_URL: 'join_url', 
  STATUS: 'status',
  MESSAGE: 'message'
};

// ZOOM API設定（Properties Serviceから取得）
// Server-to-Server OAuth2認証を使用
const ZOOM_ACCOUNT_ID = PropertiesService.getScriptProperties().getProperty('ZOOM_ACCOUNT_ID') || '';
const ZOOM_CLIENT_ID = PropertiesService.getScriptProperties().getProperty('ZOOM_CLIENT_ID') || '';
const ZOOM_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('ZOOM_CLIENT_SECRET') || '';
const ZOOM_USER_EMAIL = PropertiesService.getScriptProperties().getProperty('ZOOM_USER_EMAIL') || '合宿用メルアド';

/* ============================================================
 * 初回セットアップ関数
 * ============================================================ */

/**
 * 初回セットアップ関数
 * 
 * ZOOM API認証情報（Server-to-Server OAuth2）をProperties Serviceに保存します
 * この関数は初回のみ手動で実行してください
 * 
 */
function setupConfiguration() {
  // ⚠️ 重要: 設定後はこの関数内の認証情報を削除してください
  
  // Server-to-Server OAuth2設定
  // ZOOM App Marketplace にてマイアプリを1つ作成してください。https://marketplace.zoom.us/
  const config = {
    'ZOOM_ACCOUNT_ID': 'アプリのアカウントID',
    'ZOOM_CLIENT_ID': 'アプリのクライアントID',
    'ZOOM_CLIENT_SECRET': 'アプリのクライアントシークレットID',
    'ZOOM_USER_EMAIL': '合宿用メルアド'
  };
  
  // Properties Serviceに保存
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties(config);
  
  Logger.log('設定が保存されました');
  Logger.log('⚠️ 注意: 設定後はコード内の認証情報を削除してください');
}

/* ============================================================
 * データ取得関数
 * ============================================================ */

/**
 * シートIDからシートオブジェクトを取得する関数
 * @param {number} sheetId - シートID
 * @returns {Sheet} シートオブジェクト
 */
function getSheetById(sheetId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const sheetIds = sheets.map((sheet) => sheet.getSheetId());
  const sheetIndex = sheetIds.findIndex((id) => id === sheetId);
  
  if (sheetIndex >= 0) {
    return sheets[sheetIndex];
  } else {
    throw new Error(`シートID ${sheetId} は存在しません。`);
  }
}

/**
 * シートIDを渡すと、すべてのRecordsをオブジェクトレコーズで取得するメソッド
 * @param {number} sheetId - シートID
 * @returns {Array} objArray - オブジェクト配列
 */
function getDataSheetRecords(sheetId) {
  const sheet = getSheetById(sheetId);
  
  const [header, ...records] = sheet.getDataRange().getValues();
  
  const objectRecords = records.map(record => {
    const obj = {};
    header.forEach((element, index) => obj[element] = record[index]);
    return obj;
  });
  
  return objectRecords;
}

/**
 * objectRecoresからstarのついていないRecordだけを取得するメソッド
 * @param {Array} objectRecords - オブジェクト配列
 * @returns {Array} フィルタリングされたオブジェクト配列
 */
function getRecordWithoutStar(objectRecords) {
  const records = objectRecords.filter(record => record[COLUMN_NAMES.STAR] !== '★');
  return records;
}

/* ============================================================
 * データ更新関数
 * ============================================================ */

/**
 * 受け取ったオブジェクトレコーズをシートに上書きする
 * @param {number} sheetId - シートID
 * @param {Array} objectRecords - オブジェクト配列
 * @returns {string} 完了メッセージ
 */
function setAllRecords_(sheetId, objectRecords) {
  const sheet = getSheetById(sheetId);
  
  // ヘッダー行を取得して順序を確定
  const [header] = sheet.getDataRange().getValues();
  
  // ヘッダーの順序に従って値を並べ替え
  const records = objectRecords.map(record => {
    return header.map(headerName => record[headerName] !== undefined ? record[headerName] : '');
  });
  
  sheet.getRange(2, 1, records.length, records[0].length).setValues(records);
  
  return 'シートに書き込み完了しました';
}

/* ============================================================
 * ZOOM API関数
 * ============================================================ */

/**
 * Server-to-Server OAuth2用のアクセストークンを取得
 * @return {string} アクセストークン
 */
function getZoomAccessToken() {
  try {
    // Basic認証用のBase64エンコード
    const credentials = Utilities.base64Encode(`${ZOOM_CLIENT_ID}:${ZOOM_CLIENT_SECRET}`);
    
    // トークン取得用のリクエスト
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Basic ${credentials}`,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: `grant_type=account_credentials&account_id=${ZOOM_ACCOUNT_ID}`
    };
    
    const url = 'https://zoom.us/oauth/token';
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const responseData = JSON.parse(response.getContentText());
      return responseData.access_token;
    } else {
      Logger.log(`アクセストークン取得エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
      return null;
    }
    
  } catch (error) {
    Logger.log(`getZoomAccessToken エラー: ${error.message}`);
    return null;
  }
}

/**
 * ZOOM API認証トークンを取得（Server-to-Server OAuth2）
 * @return {string} アクセストークン
 */
function getAuthToken() {
  return getZoomAccessToken();
}

/**
 * ZOOM APIを使用して永続的なROOMを作成（固定時間なしの定期ミーティング）
 * 1つのROOM IDと参加URLを継続して使い続けることができます
 * @param {string} topic - ミーティングのトピック（タイトル）
 * @param {string} password - ミーティングのパスワード
 * @return {Object} ミーティングデータ（id, join_url等）
 */
function createZoomMeeting(topic, password) {
  try {
    // 認証トークンを取得
    const token = getAuthToken();
    
    if (!token) {
      Logger.log('認証トークンの取得に失敗しました');
      return null;
    }
    
    // ミーティング設定（type: 8 = 固定時間なしの定期ミーティング）
    const meetingData = {
      topic: topic,
      type: 8, // 固定時間なしの定期ミーティング（継続利用可能）
      password: password,
      recurrence: {
        type: 1 // 1 = 日次（固定時刻なしの定期ミーティングで推奨）
        // end_after_number_of_occurrences, end_date_time, end_after_date, no_end をすべて省略
        // ZOOM APIの制約により、UI上は「1回」と表示される可能性があるが、実質的に常設ROOMとして機能
      },
      settings: {
        join_before_host: true, // オプション任意の時刻に参加することを参加者に許可する
        waiting_room: false, // false = 待機室を無効化（ホストがいなくても参加可能にする）
        host_video: true,
        participant_video: true,
        mute_upon_entry: false,
        audio: 'both', // 電話とコンピューター音声の両方
        auto_recording: 'none' // 録画を自動で開始しない
      }
    };
    
    // ZOOM APIにリクエスト
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(meetingData)
    };
    
    const url = `https://api.zoom.us/v2/users/${ZOOM_USER_EMAIL}/meetings`;
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 201) {
      const responseData = JSON.parse(response.getContentText());
      
      return {
        id: responseData.id,
        join_url: responseData.join_url,
        start_url: responseData.start_url,
        password: responseData.password
      };
    } else {
      Logger.log(`ZOOM API エラー: ${response.getResponseCode()} - ${response.getContentText()}`);
      return null;
    }
    
  } catch (error) {
    Logger.log(`createZoomMeeting エラー: ${error.message}`);
    return null;
  }
}

/* ============================================================
 * 招待メッセージ生成関数
 * ============================================================ */

/**
 * ZOOMミーティング招待メッセージを生成
 * @param {string} title - ミーティングタイトル
 * @param {string} meetingId - ミーティングID
 * @param {string} joinUrl - 参加URL
 * @param {string} hostkey - ホストキー
 * @returns {string} 招待メッセージ
 */
function generateInvitationMessage(title, meetingId, joinUrl, hostkey) {
  const message = `実行委員さんがあなたをスケジュール済みの Zoom ミーティングに招待しています。

トピック: ${title}

参加 Zoom ミーティング

${joinUrl}

ホストキー: ${hostkey}`;
  
  return message;
}

/**
 * 招待メッセージをダイアログで表示
 * @param {string} message - 表示するメッセージ
 */
function showInvitationDialog(message) {
  // GASのHTMLダイアログを表示（コピペ可能なテキストエリア付き）
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Segoe UI', Arial, sans-serif;
          padding: 20px;
          max-width: 600px;
        }
        textarea {
          width: 100%;
          height: 200px;
          padding: 10px;
          border: 1px solid #ddd;
          border-radius: 4px;
          font-size: 13px;
          font-family: 'Courier New', monospace;
          resize: vertical;
        }
        button {
          background-color: #4CAF50;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
          margin-top: 10px;
        }
        button:hover {
          background-color: #45a049;
        }
        .copy-message {
          color: #4CAF50;
          font-size: 12px;
          margin-top: 5px;
        }
      </style>
    </head>
    <body>
      <h3>🎉 ZOOM ROOM作成完了</h3>
      <p>以下の招待メッセージをコピーしてください：</p>
      <textarea id="invitationText" readonly>${message.replace(/`/g, '\\`')}</textarea>
      <button onclick="copyToClipboard()">📋 クリップボードにコピー</button>
      <div id="copyMessage" class="copy-message"></div>
      <script>
        function copyToClipboard() {
          const textarea = document.getElementById('invitationText');
          textarea.focus();
          textarea.select();
          try {
            document.execCommand('copy');
            document.getElementById('copyMessage').textContent = '✓ コピーしました！';
          } catch(err) {
            document.getElementById('copyMessage').textContent = '※ コピーに失敗しました。手動でコピーしてください。';
          }
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(650)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ZOOM招待メッセージ');
}

/* ============================================================
 * メイン処理関数
 * ============================================================ */

/**
 * メイン実行関数
 * starのついていないレコードを抽出してZOOM ROOMを作成
 * @returns {Object} 実行結果
 */
function createZoomRooms() {
  try {
    Logger.log('ZOOM ROOM作成処理を開始します');
    
    // すべてのrecordsを取得
    const objectRecords = getDataSheetRecords(SHEET_ID);
    const withoutStarRecords = getRecordWithoutStar(objectRecords);
    
    Logger.log(`${withoutStarRecords.length}件の未作成ROOMが見つかりました`);
    
    let createdCount = 0;
    let failedCount = 0;
    const errors = [];
    const createdRooms = []; // 作成されたROOMの情報を保存
    
    // 各レコードを処理
    withoutStarRecords.forEach((record, index) => {
      try {
        const title = record[COLUMN_NAMES.TITLE];
        const password = record[COLUMN_NAMES.HOSTKEY]; // hostkeyカラムからパスワードを取得
        
        // 必須項目のチェック
        if (!title || !password) {
          Logger.log(`必須項目が不足: ${title || 'タイトルなし'}`);
          failedCount++;
          errors.push({ record: index + 1, error: '必須項目が不足' });
          return;
        }
        
        // ZOOM ROOMを作成（固定時間なしの定期ミーティング）
        Logger.log(`ROOM作成中: ${title}`);
        const meetingData = createZoomMeeting(title, password);
        
        if (meetingData) {
          // 招待メッセージを生成
          const invitationMessage = generateInvitationMessage(
            title,
            meetingData.id,
            meetingData.join_url,
            password
          );
          
          // レコードに情報を追加
          record[COLUMN_NAMES.MEETING_ID] = meetingData.id;
          record[COLUMN_NAMES.JOIN_URL] = meetingData.join_url;
          record[COLUMN_NAMES.STAR] = '★';
          record[COLUMN_NAMES.STATUS] = '作成済み';
          record[COLUMN_NAMES.MESSAGE] = invitationMessage;
          
          // 作成されたROOMの情報を保存
          createdRooms.push({
            title: title,
            meetingId: meetingData.id,
            joinUrl: meetingData.join_url,
            password: password
          });
          
          Logger.log(`ROOM作成成功: ${title} (ID: ${meetingData.id})`);
          createdCount++;
        } else {
          record[COLUMN_NAMES.STATUS] = '作成失敗';
          Logger.log(`ROOM作成失敗: ${title}`);
          failedCount++;
          errors.push({ record: index + 1, error: 'ZOOM API呼び出し失敗' });
        }
        
      } catch (error) {
        record[COLUMN_NAMES.STATUS] = 'エラー: ' + error.message;
        Logger.log(`エラー: ${record[COLUMN_NAMES.TITLE]} - ${error.message}`);
        failedCount++;
        errors.push({ record: index + 1, error: error.message });
      }
    });
    
    // すべてのrecordsをシートに書き込み
    setAllRecords_(SHEET_ID, objectRecords);
    
    Logger.log(`処理完了: ${createdCount}件作成, ${failedCount}件失敗`);
    
    // 作成されたROOMが1件以上ある場合、招待メッセージを表示
    if (createdCount > 0 && createdRooms.length > 0) {
      // 複数のROOMがある場合は、1つのメッセージにまとめる
      let allMessages = '';
      createdRooms.forEach((room, index) => {
        allMessages += generateInvitationMessage(
          room.title,
          room.meetingId,
          room.joinUrl,
          room.password
        );
        if (index < createdRooms.length - 1) {
          allMessages += '\n\n---\n\n'; // 区切り線
        }
      });
      
      showInvitationDialog(allMessages);
    }
    
    return {
      success: true,
      total: withoutStarRecords.length,
      created: createdCount,
      failed: failedCount,
      errors: errors
    };
    
  } catch (error) {
    Logger.log(`致命的なエラー: ${error.message}`);
    return {
      success: false,
      total: 0,
      created: 0,
      failed: 0,
      errors: [{ error: error.message }]
    };
  }
}

/* ============================================================
 * テスト関数
 * ============================================================ */

/**
 * テスト実行関数（1行だけ処理）
 */
function testCreateZoomRoom() {
  try {
    Logger.log('テスト実行を開始します');
    
    const objectRecords = getDataSheetRecords(SHEET_ID);
    const withoutStarRecords = getRecordWithoutStar(objectRecords);
    
    if (withoutStarRecords.length === 0) {
      Logger.log('処理対象のレコードがありません');
      return;
    }
    
    // 最初の1件だけ処理
    const testRecord = withoutStarRecords[0];
    const title = testRecord[COLUMN_NAMES.TITLE];
    const password = testRecord[COLUMN_NAMES.HOSTKEY]; // hostkeyカラムからパスワードを取得
    
    Logger.log(`テスト: ${title} - ${password} (固定時間なしの定期ミーティング)`);
    const meetingData = createZoomMeeting(title, password);
    
    if (meetingData) {
      Logger.log(`成功: ID=${meetingData.id}, URL=${meetingData.join_url}`);
    } else {
      Logger.log('失敗');
    }
    
  } catch (error) {
    Logger.log(`テストエラー: ${error.message}`);
  }
}

/**
 * 設定確認関数
 */
function checkConfiguration() {
  Logger.log('設定確認:');
  Logger.log(`Spreadsheet ID: ${SPREADSHEET_ID}`);
  Logger.log(`Sheet ID: ${SHEET_ID}`);
  Logger.log('認証方式: Server-to-Server OAuth2');
  Logger.log(`Account ID: ${ZOOM_ACCOUNT_ID ? '設定済み' : '設定されていません'}`);
  Logger.log(`Client ID: ${ZOOM_CLIENT_ID ? ZOOM_CLIENT_ID.substring(0, 10) + '...' : '設定されていません'}`);
  Logger.log(`Client Secret: ${ZOOM_CLIENT_SECRET ? '設定済み' : '設定されていません'}`);
  Logger.log(`ZOOM User Email: ${ZOOM_USER_EMAIL}`);
}
