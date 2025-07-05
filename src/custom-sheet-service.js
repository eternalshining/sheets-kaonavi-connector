/**
 * カスタムシート作成機能を担当するサービス
 */
const CustomSheetService = {
  
  /**
   * 定義シートから定義情報を解析する
   * @param {Object} definitionSheet - 定義シート
   * @returns {Array} 定義情報配列
   */
  parseDefinitions(definitionSheet) {
    const definitions = [];
    
    try {
      // データ範囲を取得（A列とB列、最大1000行まで）
      const dataRange = definitionSheet.getRange('A1:B1000');
      const values = dataRange.getValues();
      
      for (let i = 0; i < values.length; i++) {
        const dataSource = values[i][0];
        const fieldName = values[i][1];
        
        // 両方の値が存在する場合のみ追加
        if (dataSource && fieldName) {
          definitions.push({
            dataSource: dataSource.toString().trim(),
            fieldName: fieldName.toString().trim()
          });
        }
      }
      
      if (definitions.length === 0) {
        throw new Error('定義シートに有効な定義が見つかりません。A列にデータソース、B列に項目名を入力してください。');
      }
      
      return definitions;
      
    } catch (error) {
      console.error('定義解析エラー:', error);
      throw new Error(`定義シートの解析に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * 定義に基づいてカスタムデータを作成する
   * @param {Array} definitions - 定義情報
   * @returns {Object} カスタムデータ
   */
  createCustomData(definitions) {
    try {
      // 必要なデータソースを特定
      const dataSources = [...new Set(definitions.map(def => def.dataSource))];
      
      // データを収集
      const collectedData = {};
      
      // メンバー情報を取得
      if (dataSources.includes('基本情報')) {
        try {
          const memberData = KaonaviService.getMemberInfo();
          collectedData['基本情報'] = this.normalizeMemberData(memberData);
        } catch (error) {
          console.error('メンバー情報取得エラー:', error);
          throw new Error(`メンバー情報の取得に失敗しました: ${error.message}`);
        }
      }
      
      // シート情報を取得
      const sheetNames = dataSources.filter(source => source !== '基本情報');
      if (sheetNames.length > 0) {
        try {
          const sheetLayouts = KaonaviService.getSheetLayouts();
          const sheetData = this.getSheetDataByNames(sheetNames, sheetLayouts);
          
          Object.keys(sheetData).forEach(sheetName => {
            collectedData[sheetName] = sheetData[sheetName];
          });
        } catch (error) {
          console.error('シート情報取得エラー:', error);
          throw new Error(`シート情報の取得に失敗しました: ${error.message}`);
        }
      }
      
      // データを結合
      const combinedData = this.combineData(collectedData, definitions);
      
      return {
        collectedData: collectedData,
        combinedData: combinedData
      };
      
    } catch (error) {
      console.error('カスタムデータ作成エラー:', error);
      throw error;
    }
  },
  
  /**
   * メンバー情報を正規化する
   * @param {Object} memberData - メンバー情報
   * @returns {Object} 正規化されたメンバー情報
   */
  normalizeMemberData(memberData) {
    const normalized = {};
    
    if (memberData.members && memberData.members.member_data) {
      memberData.members.member_data.forEach(member => {
        const memberKey = member.code || member.id || member.email;
        if (memberKey) {
          normalized[memberKey] = member;
        }
      });
    }
    
    return normalized;
  },
  
  /**
   * シート名に基づいてシートデータを取得する
   * @param {Array} sheetNames - シート名配列
   * @param {Object} sheetLayouts - シートレイアウト情報
   * @returns {Object} シートデータ
   */
  getSheetDataByNames(sheetNames, sheetLayouts) {
    const sheetData = {};
    
    if (!sheetLayouts.sheet_layouts) {
      throw new Error('シートレイアウト情報が取得できませんでした');
    }
    
    sheetNames.forEach(sheetName => {
      // シート名からシートIDを取得
      const sheetLayout = sheetLayouts.sheet_layouts.find(layout => layout.name === sheetName);
      
      if (!sheetLayout) {
        throw new Error(`シート "${sheetName}" が見つかりません`);
      }
      
      try {
        const sheetInfo = KaonaviService.getSheetInfo(sheetLayout.id);
        sheetData[sheetName] = this.normalizeSheetData(sheetInfo);
      } catch (error) {
        console.error(`シート ${sheetName} の取得エラー:`, error);
        throw new Error(`シート "${sheetName}" の取得に失敗しました: ${error.message}`);
      }
    });
    
    return sheetData;
  },
  
  /**
   * シートデータを正規化する
   * @param {Object} sheetInfo - シート情報
   * @returns {Object} 正規化されたシートデータ
   */
  normalizeSheetData(sheetInfo) {
    const normalized = {};
    
    if (sheetInfo.data && sheetInfo.data.sheet_data) {
      sheetInfo.data.sheet_data.forEach(record => {
        const memberKey = record.member_code || record.member_id || record.code;
        if (memberKey) {
          normalized[memberKey] = record;
        }
      });
    }
    
    return normalized;
  },
  
  /**
   * 収集したデータを結合する
   * @param {Object} collectedData - 収集済みデータ
   * @param {Array} definitions - 定義情報
   * @returns {Object} 結合されたデータ
   */
  combineData(collectedData, definitions) {
    const combined = {};
    
    // 全メンバーのキーを取得
    const allMemberKeys = new Set();
    Object.keys(collectedData).forEach(dataSource => {
      Object.keys(collectedData[dataSource]).forEach(memberKey => {
        allMemberKeys.add(memberKey);
      });
    });
    
    // 各メンバーについてデータを結合
    allMemberKeys.forEach(memberKey => {
      combined[memberKey] = {};
      
      Object.keys(collectedData).forEach(dataSource => {
        if (collectedData[dataSource][memberKey]) {
          combined[memberKey][dataSource] = collectedData[dataSource][memberKey];
        }
      });
    });
    
    return combined;
  },
  
  /**
   * 定義の妥当性を検証する
   * @param {Array} definitions - 定義情報
   * @param {Object} collectedData - 収集済みデータ
   * @returns {Object} 検証結果
   */
  validateDefinitions(definitions, collectedData) {
    const validation = {
      valid: true,
      errors: []
    };
    
    definitions.forEach((def, index) => {
      // データソースの存在確認
      if (!collectedData[def.dataSource]) {
        validation.valid = false;
        validation.errors.push(`行 ${index + 1}: データソース "${def.dataSource}" が見つかりません`);
      }
      
      // 項目名の存在確認
      if (collectedData[def.dataSource]) {
        const sampleData = Object.values(collectedData[def.dataSource])[0];
        if (sampleData && !sampleData.hasOwnProperty(def.fieldName)) {
          validation.valid = false;
          validation.errors.push(`行 ${index + 1}: 項目名 "${def.fieldName}" がデータソース "${def.dataSource}" に存在しません`);
        }
      }
    });
    
    return validation;
  }
};