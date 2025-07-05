/**
 * スプレッドシートが開かれたときにカスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カオナビ連携')
    .addItem('メンバー情報を取得', 'getMemberInfo')
    .addItem('シート情報を取得', 'getSheetInfo')
    .addItem('カスタムシートを作成', 'createCustomSheet')
    .addSeparator()
    .addItem('認証情報を更新', 'updateAuthInfo')
    .addToUi();
}

/**
 * メンバー情報を取得する
 */
function getMemberInfo() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理中です...', 'メンバー情報取得', 5);
    
    // 認証情報を確認
    if (!AuthService.checkAuth()) {
      AuthService.showAuthDialog();
      return;
    }
    
    // 出力方法を選択
    const outputType = UIService.showOutputTypeDialog();
    if (!outputType) return;
    
    // メンバー情報を取得
    const memberData = KaonaviService.getMemberInfo();
    
    // シートに出力
    SheetService.outputMemberData(memberData, outputType);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('メンバー情報の取得が完了しました', '完了', 3);
    
  } catch (error) {
    console.error('メンバー情報取得エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'メンバー情報の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * シート情報を取得する
 */
function getSheetInfo() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理中です...', 'シート情報取得', 5);
    
    // 認証情報を確認
    if (!AuthService.checkAuth()) {
      AuthService.showAuthDialog();
      return;
    }
    
    // シートを選択
    const selectedSheet = UIService.showSheetSelectionDialog();
    if (!selectedSheet) return;
    
    // 出力方法を選択
    const outputType = UIService.showOutputTypeDialog();
    if (!outputType) return;
    
    // シート情報を取得
    const sheetData = KaonaviService.getSheetInfo(selectedSheet.id);
    
    // シートに出力
    SheetService.outputSheetData(sheetData, selectedSheet, outputType);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('シート情報の取得が完了しました', '完了', 3);
    
  } catch (error) {
    console.error('シート情報取得エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'シート情報の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * カスタムシートを作成する
 */
function createCustomSheet() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('処理中です...', 'カスタムシート作成', 5);
    
    // 認証情報を確認
    if (!AuthService.checkAuth()) {
      AuthService.showAuthDialog();
      return;
    }
    
    // 定義シートの確認
    const definitionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('kaonavi_custom_def');
    if (!definitionSheet) {
      SpreadsheetApp.getUi().alert('エラー', 'カスタムシート定義用シート "kaonavi_custom_def" が見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 定義を読み込み
    const definitions = CustomSheetService.parseDefinitions(definitionSheet);
    
    // データを取得・結合
    const customData = CustomSheetService.createCustomData(definitions);
    
    // カスタムシートを作成
    SheetService.outputCustomData(customData, definitions);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('カスタムシートの作成が完了しました', '完了', 3);
    
  } catch (error) {
    console.error('カスタムシート作成エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'カスタムシートの作成に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 認証情報を更新する
 */
function updateAuthInfo() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('確認', '現在の認証情報を削除して新しい認証情報を設定しますか？', ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      AuthService.clearAuth();
      AuthService.showAuthDialog();
    }
  } catch (error) {
    console.error('認証情報更新エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '認証情報の更新に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}