/**
 * 認証情報管理サービス
 */
const AuthService = {
  
  /**
   * 認証情報をチェックする
   * @returns {boolean} 認証情報が存在するかどうか
   */
  checkAuth() {
    const properties = PropertiesService.getUserProperties();
    const consumerKey = properties.getProperty('KAONAVI_CONSUMER_KEY');
    const consumerSecret = properties.getProperty('KAONAVI_CONSUMER_SECRET');
    
    return !!(consumerKey && consumerSecret);
  },

  /**
   * 認証情報を取得する
   * @returns {Object} 認証情報オブジェクト
   */
  getAuthInfo() {
    const properties = PropertiesService.getUserProperties();
    return {
      consumerKey: properties.getProperty('KAONAVI_CONSUMER_KEY'),
      consumerSecret: properties.getProperty('KAONAVI_CONSUMER_SECRET')
    };
  },

  /**
   * 認証情報を保存する
   * @param {string} consumerKey - Consumer Key
   * @param {string} consumerSecret - Consumer Secret
   */
  saveAuthInfo(consumerKey, consumerSecret) {
    const properties = PropertiesService.getUserProperties();
    properties.setProperties({
      'KAONAVI_CONSUMER_KEY': consumerKey,
      'KAONAVI_CONSUMER_SECRET': consumerSecret
    });
  },

  /**
   * 認証情報を削除する
   */
  clearAuth() {
    const properties = PropertiesService.getUserProperties();
    properties.deleteProperty('KAONAVI_CONSUMER_KEY');
    properties.deleteProperty('KAONAVI_CONSUMER_SECRET');
  },

  /**
   * 認証情報入力ダイアログを表示する
   */
  showAuthDialog() {
    const ui = SpreadsheetApp.getUi();
    
    // Consumer Key入力
    const keyResponse = ui.prompt(
      '認証情報の設定',
      'カオナビのConsumer Keyを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (keyResponse.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    const consumerKey = keyResponse.getResponseText().trim();
    if (!consumerKey) {
      ui.alert('エラー', 'Consumer Keyを入力してください。', ui.ButtonSet.OK);
      return false;
    }
    
    // Consumer Secret入力
    const secretResponse = ui.prompt(
      '認証情報の設定',
      'カオナビのConsumer Secretを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (secretResponse.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    const consumerSecret = secretResponse.getResponseText().trim();
    if (!consumerSecret) {
      ui.alert('エラー', 'Consumer Secretを入力してください。', ui.ButtonSet.OK);
      return false;
    }
    
    // 認証情報を保存
    try {
      this.saveAuthInfo(consumerKey, consumerSecret);
      ui.alert('成功', '認証情報が正常に保存されました。', ui.ButtonSet.OK);
      return true;
    } catch (error) {
      console.error('認証情報保存エラー:', error);
      ui.alert('エラー', '認証情報の保存に失敗しました: ' + error.message, ui.ButtonSet.OK);
      return false;
    }
  }
};