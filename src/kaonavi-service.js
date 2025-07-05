/**
 * カオナビAPIとの通信を担当するサービス
 */
const KaonaviService = {
  
  BASE_URL: 'https://api.kaonavi.jp/api/v2.0',
  
  /**
   * アクセストークンを取得する
   * @returns {string} アクセストークン
   */
  getAccessToken() {
    const authInfo = AuthService.getAuthInfo();
    if (!authInfo.consumerKey || !authInfo.consumerSecret) {
      throw new Error('認証情報が設定されていません');
    }
    
    // キャッシュからアクセストークンを確認
    const cache = CacheService.getUserCache();
    const cachedToken = cache.get('kaonavi_access_token');
    if (cachedToken) {
      return cachedToken;
    }
    
    // Basic認証用のcredentialsを生成
    const credentials = Utilities.base64Encode(`${authInfo.consumerKey}:${authInfo.consumerSecret}`);
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Basic ${credentials}`,
        'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
      },
      payload: 'grant_type=client_credentials'
    };
    
    try {
      const response = UrlFetchApp.fetch(`${this.BASE_URL}/token`, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode !== 200) {
        throw new Error(`トークン取得エラー: ${responseCode} - ${response.getContentText()}`);
      }
      
      const tokenData = JSON.parse(response.getContentText());
      const accessToken = tokenData.access_token;
      
      // アクセストークンをキャッシュに保存（有効期限の少し前まで）
      const expireIn = tokenData.expire_in || 3600; // デフォルト1時間
      const cacheTime = Math.max(expireIn - 300, 300); // 5分前またはデフォルト5分
      cache.put('kaonavi_access_token', accessToken, cacheTime);
      
      return accessToken;
      
    } catch (error) {
      console.error('アクセストークン取得エラー:', error);
      throw new Error(`アクセストークンの取得に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * APIリクエストを送信する
   * @param {string} endpoint - APIエンドポイント
   * @param {string} method - HTTPメソッド
   * @param {Object} payload - リクエストボディ
   * @returns {Object} APIレスポンス
   */
  makeRequest(endpoint, method = 'GET', payload = null) {
    const accessToken = this.getAccessToken();
    
    const url = `${this.BASE_URL}${endpoint}`;
    const options = {
      method: method,
      headers: {
        'Content-Type': 'application/json',
        'Kaonavi-Token': accessToken
      }
    };
    
    if (payload && method !== 'GET') {
      options.payload = JSON.stringify(payload);
    }
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode < 200 || responseCode >= 300) {
        // トークンエラーの場合、キャッシュを削除して再試行
        if (responseCode === 401) {
          const cache = CacheService.getUserCache();
          cache.remove('kaonavi_access_token');
          
          // 1回だけ再試行
          const retryToken = this.getAccessToken();
          options.headers['Kaonavi-Token'] = retryToken;
          
          const retryResponse = UrlFetchApp.fetch(url, options);
          const retryResponseCode = retryResponse.getResponseCode();
          
          if (retryResponseCode < 200 || retryResponseCode >= 300) {
            throw new Error(`APIエラー: ${retryResponseCode} - ${retryResponse.getContentText()}`);
          }
          
          return JSON.parse(retryResponse.getContentText());
        }
        
        throw new Error(`APIエラー: ${responseCode} - ${response.getContentText()}`);
      }
      
      return JSON.parse(response.getContentText());
      
    } catch (error) {
      console.error('API通信エラー:', error);
      throw new Error(`APIとの通信に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * メンバー情報レイアウトを取得する
   * @returns {Object} メンバー情報レイアウト
   */
  getMemberLayouts() {
    const cacheKey = 'member_layouts';
    const cache = CacheService.getUserCache();
    
    // キャッシュから取得を試行
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // APIから取得
    const response = this.makeRequest('/member_layouts');
    
    // キャッシュに保存（1時間）
    cache.put(cacheKey, JSON.stringify(response), 3600);
    
    return response;
  },
  
  /**
   * メンバー情報を取得する
   * @returns {Object} メンバー情報
   */
  getMemberInfo() {
    const layouts = this.getMemberLayouts();
    const members = this.makeRequest('/members');
    
    return {
      layouts: layouts,
      members: members
    };
  },
  
  /**
   * シートレイアウト一覧を取得する
   * @returns {Object} シートレイアウト一覧
   */
  getSheetLayouts() {
    const cacheKey = 'sheet_layouts';
    const cache = CacheService.getUserCache();
    
    // キャッシュから取得を試行
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
    
    // APIから取得
    const response = this.makeRequest('/sheet_layouts');
    
    // キャッシュに保存（1時間）
    cache.put(cacheKey, JSON.stringify(response), 3600);
    
    return response;
  },
  
  /**
   * 特定のシート情報を取得する
   * @param {string} sheetId - シートID
   * @returns {Object} シート情報
   */
  getSheetInfo(sheetId) {
    const layouts = this.getSheetLayouts();
    const sheetData = this.makeRequest(`/sheets/${sheetId}`);
    
    // 対応するレイアウト情報を取得
    const sheetLayout = layouts.sheet_layouts ? 
      layouts.sheet_layouts.find(layout => layout.id === sheetId) : null;
    
    return {
      layout: sheetLayout,
      data: sheetData
    };
  },
  
  /**
   * 複数のシート情報を一括取得する
   * @param {Array} sheetIds - シートIDの配列
   * @returns {Object} シート情報の配列
   */
  getMultipleSheetInfo(sheetIds) {
    const layouts = this.getSheetLayouts();
    const results = {};
    
    sheetIds.forEach(sheetId => {
      try {
        const sheetData = this.makeRequest(`/sheets/${sheetId}`);
        const sheetLayout = layouts.sheet_layouts ? 
          layouts.sheet_layouts.find(layout => layout.id === sheetId) : null;
        
        results[sheetId] = {
          layout: sheetLayout,
          data: sheetData
        };
      } catch (error) {
        console.error(`シート ${sheetId} の取得エラー:`, error);
        results[sheetId] = { error: error.message };
      }
    });
    
    return results;
  }
};