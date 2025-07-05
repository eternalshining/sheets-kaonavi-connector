/**
 * UI関連のサービス
 */
const UIService = {
  
  /**
   * 出力方法選択ダイアログを表示する
   * @returns {string|null} 選択された出力方法（'new' または 'overwrite'）
   */
  showOutputTypeDialog() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '出力方法の選択',
      '新規シートを作成しますか？\n\n「はい」: 新規シートに出力\n「いいえ」: 現在のシートに上書き\n「キャンセル」: 処理を中止',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      return 'new'; // 新規シートに作成
    } else if (response === ui.Button.NO) {
      return 'overwrite'; // 現在のシートに上書き
    } else {
      return null; // キャンセル
    }
  },
  
  /**
   * シート選択ダイアログを表示する
   * @returns {Object|null} 選択されたシート情報
   */
  showSheetSelectionDialog() {
    try {
      // シートレイアウト一覧を取得
      const layouts = KaonaviService.getSheetLayouts();
      
      if (!layouts.sheet_layouts || layouts.sheet_layouts.length === 0) {
        SpreadsheetApp.getUi().alert('エラー', 'アクセス可能なシートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
        return null;
      }
      
      // シート選択肢を作成
      const sheetOptions = layouts.sheet_layouts.map((sheet, index) => {
        return `${index + 1}. ${sheet.name || 'シート名未設定'} (ID: ${sheet.id})`;
      }).join('\n');
      
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(
        'シート選択',
        `取得するシートを選択してください（番号を入力）:\n\n${sheetOptions}`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() !== ui.Button.OK) {
        return null;
      }
      
      const selectedIndex = parseInt(response.getResponseText().trim()) - 1;
      
      if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= layouts.sheet_layouts.length) {
        ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
        return null;
      }
      
      return layouts.sheet_layouts[selectedIndex];
      
    } catch (error) {
      console.error('シート選択エラー:', error);
      SpreadsheetApp.getUi().alert('エラー', 'シート一覧の取得に失敗しました: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return null;
    }
  },
  
  /**
   * 確認ダイアログを表示する
   * @param {string} title - ダイアログタイトル
   * @param {string} message - メッセージ
   * @returns {boolean} 確認結果
   */
  showConfirmDialog(title, message) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    return response === ui.Button.YES;
  },
  
  /**
   * 情報ダイアログを表示する
   * @param {string} title - ダイアログタイトル
   * @param {string} message - メッセージ
   */
  showInfoDialog(title, message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  },
  
  /**
   * エラーダイアログを表示する
   * @param {string} message - エラーメッセージ
   */
  showErrorDialog(message) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', message, ui.ButtonSet.OK);
  }
};