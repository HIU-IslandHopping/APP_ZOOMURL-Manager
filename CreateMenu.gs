/**
 * スプレッドシートを開いたときに実行される関数
 * ZOOMカスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // ZOOMメニューを作成
  ui.createMenu('ZOOM')
    .addItem('🎯 ROOMを作成', 'onCreateZoomRooms')
    .addSeparator()
    .addItem('⚙️ 設定確認', 'onCheckConfiguration')
    .addToUi();
  
  Logger.log('ZOOMカスタムメニューを作成しました');
}

/**
 * カスタムメニューから「ROOMを作成」が選択されたときの処理
 */
function onCreateZoomRooms() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // 確認ダイアログ
    const response = ui.alert(
      'ZOOM ROOM作成',
      '未作成のZOOM ROOMを一括作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      Logger.log('ユーザーがキャンセルしました');
      return;
    }
    
    // ROOM作成処理を実行
    Logger.log('ZOOM ROOM作成を開始します');
    const result = createZoomRooms();
    
    // 結果を表示
    if (result.success) {
      ui.alert(
        '✅ 処理完了',
        `${result.created}件のROOMを作成しました。\n\n失敗: ${result.failed}件\n詳細はログを確認してください。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ エラー',
        'ROOMの作成中にエラーが発生しました。\n詳細はログを確認してください。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log(`onCreateZoomRooms エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

/**
 * カスタムメニューから「設定確認」が選択されたときの処理
 */
function onCheckConfiguration() {
  try {
    checkConfiguration();
    SpreadsheetApp.getUi().alert(
      '設定確認',
      '設定情報をログに出力しました。\nApps Scriptの実行ログを確認してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log(`onCheckConfiguration エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}
