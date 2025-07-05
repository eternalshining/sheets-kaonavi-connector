/**
 * シート操作を担当するサービス
 */
const SheetService = {
  
  /**
   * メンバー情報をシートに出力する
   * @param {Object} memberData - メンバー情報
   * @param {string} outputType - 出力方法（'new' または 'overwrite'）
   */
  outputMemberData(memberData, outputType) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet;
      
      if (outputType === 'new') {
        // 新規シートを作成
        const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
        const sheetName = `メンバー情報_${timestamp}`;
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        // 現在のシートを使用
        sheet = spreadsheet.getActiveSheet();
        sheet.clear();
      }
      
      // ヘッダーを準備
      const headers = this.extractMemberHeaders(memberData.layouts);
      
      // データを準備
      const rows = this.extractMemberRows(memberData.members, memberData.layouts);
      
      // データを書き込み
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        }
        
        // ヘッダー行をフォーマット
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#4285f4');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        
        // 列幅を自動調整
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
      }
      
    } catch (error) {
      console.error('メンバーデータ出力エラー:', error);
      throw new Error(`メンバーデータの出力に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * シート情報をシートに出力する
   * @param {Object} sheetData - シート情報
   * @param {Object} selectedSheet - 選択されたシート情報
   * @param {string} outputType - 出力方法（'new' または 'overwrite'）
   */
  outputSheetData(sheetData, selectedSheet, outputType) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet;
      
      if (outputType === 'new') {
        // 新規シートを作成
        const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
        const sheetName = `シート情報_${selectedSheet.name || 'カスタムシート'}_${timestamp}`;
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        // 現在のシートを使用
        sheet = spreadsheet.getActiveSheet();
        sheet.clear();
      }
      
      // ヘッダーを準備
      const headers = this.extractSheetHeaders(sheetData.layout);
      
      // データを準備
      const rows = this.extractSheetRows(sheetData.data, sheetData.layout);
      
      // データを書き込み
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        }
        
        // ヘッダー行をフォーマット
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#34a853');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        
        // 列幅を自動調整
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
      }
      
    } catch (error) {
      console.error('シートデータ出力エラー:', error);
      throw new Error(`シートデータの出力に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * カスタムデータをシートに出力する
   * @param {Object} customData - カスタムデータ
   * @param {Array} definitions - 定義情報
   */
  outputCustomData(customData, definitions) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
      const sheetName = `カスタムシート_${timestamp}`;
      const sheet = spreadsheet.insertSheet(sheetName);
      
      // ヘッダーを準備（定義のB列の項目名）
      const headers = definitions.map(def => def.fieldName);
      
      // データを準備
      const rows = this.extractCustomRows(customData, definitions);
      
      // データを書き込み
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
        }
        
        // ヘッダー行をフォーマット
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#ff9800');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        
        // 列幅を自動調整
        for (let i = 1; i <= headers.length; i++) {
          sheet.autoResizeColumn(i);
        }
      }
      
    } catch (error) {
      console.error('カスタムデータ出力エラー:', error);
      throw new Error(`カスタムデータの出力に失敗しました: ${error.message}`);
    }
  },
  
  /**
   * メンバー情報からヘッダーを抽出する
   * @param {Object} layouts - レイアウト情報
   * @returns {Array} ヘッダー配列
   */
  extractMemberHeaders(layouts) {
    const headers = [];
    
    if (layouts && layouts.member_layout) {
      // 基本情報
      if (layouts.member_layout.basic_fields) {
        layouts.member_layout.basic_fields.forEach(field => {
          headers.push(field.name || field.id);
        });
      }
      
      // カスタムフィールド
      if (layouts.member_layout.custom_fields) {
        layouts.member_layout.custom_fields.forEach(field => {
          headers.push(field.name || field.id);
        });
      }
    }
    
    return headers.length > 0 ? headers : ['ID', '名前', 'メールアドレス']; // デフォルトヘッダー
  },
  
  /**
   * メンバー情報からデータ行を抽出する
   * @param {Object} members - メンバー情報
   * @param {Object} layouts - レイアウト情報
   * @returns {Array} データ行配列
   */
  extractMemberRows(members, layouts) {
    const rows = [];
    
    if (members && members.member_data) {
      members.member_data.forEach(member => {
        const row = [];
        
        // 基本情報
        if (layouts && layouts.member_layout && layouts.member_layout.basic_fields) {
          layouts.member_layout.basic_fields.forEach(field => {
            row.push(member[field.id] || '');
          });
        }
        
        // カスタムフィールド
        if (layouts && layouts.member_layout && layouts.member_layout.custom_fields) {
          layouts.member_layout.custom_fields.forEach(field => {
            row.push(member[field.id] || '');
          });
        }
        
        // デフォルト値
        if (row.length === 0) {
          row.push(member.id || '', member.name || '', member.email || '');
        }
        
        rows.push(row);
      });
    }
    
    return rows;
  },
  
  /**
   * シート情報からヘッダーを抽出する
   * @param {Object} layout - レイアウト情報
   * @returns {Array} ヘッダー配列
   */
  extractSheetHeaders(layout) {
    const headers = [];
    
    if (layout && layout.fields) {
      layout.fields.forEach(field => {
        headers.push(field.name || field.id);
      });
    }
    
    return headers.length > 0 ? headers : ['データ']; // デフォルトヘッダー
  },
  
  /**
   * シート情報からデータ行を抽出する
   * @param {Object} data - シートデータ
   * @param {Object} layout - レイアウト情報
   * @returns {Array} データ行配列
   */
  extractSheetRows(data, layout) {
    const rows = [];
    
    if (data && data.sheet_data) {
      data.sheet_data.forEach(record => {
        const row = [];
        
        if (layout && layout.fields) {
          layout.fields.forEach(field => {
            row.push(record[field.id] || '');
          });
        } else {
          // レイアウト情報がない場合は、recordの値をそのまま使用
          Object.values(record).forEach(value => {
            row.push(value || '');
          });
        }
        
        rows.push(row);
      });
    }
    
    return rows;
  },
  
  /**
   * カスタムデータからデータ行を抽出する
   * @param {Object} customData - カスタムデータ
   * @param {Array} definitions - 定義情報
   * @returns {Array} データ行配列
   */
  extractCustomRows(customData, definitions) {
    const rows = [];
    
    if (customData && customData.combinedData) {
      Object.keys(customData.combinedData).forEach(memberKey => {
        const memberData = customData.combinedData[memberKey];
        const row = [];
        
        definitions.forEach(def => {
          const value = memberData[def.dataSource] && memberData[def.dataSource][def.fieldName];
          row.push(value || '');
        });
        
        rows.push(row);
      });
    }
    
    return rows;
  }
};