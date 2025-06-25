/**
 * メンバー管理システム共通ユーティリティ
 */

// 共通定数
const COL_HEADER_NAMES = {
  DEPT: "dept",
  EMAIL: "email",
  MEMBER_ID: "memberId",
  MEMBER_DATE_ID: "memberDateId",
  DATE: "date"
};
const REQUIRED_MEMBER_DATA_HEADERS = {
  DATA_SHEET: {
    INITIALIZE: [COL_HEADER_NAMES.DEPT, COL_HEADER_NAMES.EMAIL],
    UPDATE: [COL_HEADER_NAMES.DEPT, COL_HEADER_NAMES.EMAIL, COL_HEADER_NAMES.MEMBER_ID]
  },
  GANTT_SHEETS: {
    INITIALIZE: [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.DATE, COL_HEADER_NAMES.MEMBER_ID],
    UPDATE: [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.MEMBER_ID]
  }
};

/**
 * GANTT_TEMPLATEシートから値を保持する列の見出し
 * 
 * ここに指定された列は、メンバーデータの書き込み時に上書きされず、
 * テンプレートシートの元の値が保持されます。
 * 
 * 将来的に他の列も追加可能:
 * 例: "schedule", "deadline", "priority", "status" など
 */
const PRESERVE_TEMPLATE_COLUMNS = [
  COL_HEADER_NAMES.DATE
];

/**
 * エラーの詳細情報をログに出力するヘルパー関数（共通ライブラリ用）
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラーが発生したコンテキスト
 * @param {Object} additionalInfo - 追加のデバッグ情報（オプション）
 */
function logLibraryError(error, context, additionalInfo = {}) {
  console.error(`=== ライブラリエラー詳細 ===`);
  console.error(`コンテキスト: ${context}`);
  console.error(`エラーメッセージ: ${error.message}`);
  console.error(`エラータイプ: ${error.name}`);
  
  if (error.stack) {
    console.error(`スタックトレース:`);
    console.error(error.stack);
  }
  
  if (Object.keys(additionalInfo).length > 0) {
    console.error(`追加情報:`);
    console.error(JSON.stringify(additionalInfo, null, 2));
  }
  
  console.error(`=== ライブラリエラー詳細終了 ===`);
}

/**
 * メンバーリストシートからデータとヘッダーを取得する
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - メンバーリストが含まれるスプレッドシート
 * @param {Array} requiredHeaders - 必須ヘッダーのリスト
 * @returns {Object} メンバーデータとヘッダーのオブジェクト
 * @property {Array} data - メンバーデータの2次元配列
 * @property {Array} headers - ヘッダー行の配列
 */
function getMemberDataAndHeaders(spreadsheet, requiredHeaders) {
  try {
    console.log('メンバーデータとヘッダーの取得を開始します', {
      spreadsheetName: spreadsheet ? spreadsheet.getName() : 'unknown',
      requiredHeaders: requiredHeaders
    });
    
    const memberSheet = spreadsheet.getSheetByName(SHEET_NAMES.MEMBER_DATA);
    if (!memberSheet) {
      const errorMessage = `シート「${SHEET_NAMES.MEMBER_DATA}」が見つかりません。`;
      console.error(errorMessage);
      throw new Error(errorMessage);
    }
    
    console.log(`シート「${SHEET_NAMES.MEMBER_DATA}」を取得しました`);
    
    const memberDataRange = memberSheet.getDataRange();
    const memberData = memberDataRange.getValues();
    const memberHeaders = memberData[0];

    console.log('メンバーデータを取得しました', {
      dataRows: memberData.length,
      headers: memberHeaders
    });

    // ヘッダー検証を内部で実行
    validateHeaders(memberHeaders, requiredHeaders);
    
    console.log('メンバーデータとヘッダーの取得が完了しました');
    
    return {
      data: memberData,
      headers: memberHeaders
    };
  } catch (error) {
    logLibraryError(error, 'getMemberDataAndHeaders', {
      spreadsheetName: spreadsheet ? spreadsheet.getName() : 'unknown',
      requiredHeaders: requiredHeaders,
      memberDataSheetName: SHEET_NAMES.MEMBER_DATA
    });
    throw error;
  }
}

/**
 * ガントチャートのヘッダーを取得する
 * @param {SpreadsheetApp.Sheet} sheet - ガントチャートシート
 * @param {string} headerRangeA1 - ヘッダー範囲のA1記法
 * @param {Array} requiredHeaders - 必須ヘッダーのリスト
 * @returns {Object} ヘッダー情報のオブジェクト
 * @property {Array} headers - ヘッダー行の配列
 * @property {number} headerRow - ヘッダー行の行番号
 * @property {number} startCol - 開始列の列番号
 * @property {number} endCol - 終了列の列番号
 */
function getGanttHeaders(sheet, headerRangeA1, requiredHeaders) {
  const sheetName = sheet ? sheet.getName() : 'unknown';
  
  try {
    console.log(`ガントチャートヘッダーの取得を開始します`, {
      sheetName: sheetName,
      headerRangeA1: headerRangeA1,
      requiredHeaders: requiredHeaders
    });
    
    const headerRangeParts = headerRangeA1.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
    if (!headerRangeParts) {
      const errorMessage = `ヘッダー範囲の形式が不正です: ${headerRangeA1}`;
      console.error(errorMessage);
      throw new Error(errorMessage);
    }

    const headerRow = parseInt(headerRangeParts[2]);
    const startCol = sheet.getRange(headerRangeParts[1] + "1").getColumn();
    const endCol = sheet.getRange(headerRangeParts[3] + "1").getColumn();
    
    console.log('ヘッダー範囲を解析しました', {
      headerRow: headerRow,
      startCol: startCol,
      endCol: endCol
    });
    
    const headerRange = sheet.getRange(headerRow, startCol, 1, endCol - startCol + 1);
    const headers = headerRange.getValues()[0];

    console.log('ヘッダーデータを取得しました', {
      headers: headers
    });

    // ヘッダー検証を内部で実行
    validateHeaders(headers, requiredHeaders);

    console.log('ガントチャートヘッダーの取得が完了しました');

    return {
      headers,
      headerRow,
      startCol,
      endCol
    };
  } catch (error) {
    logLibraryError(error, 'getGanttHeaders', {
      sheetName: sheetName,
      headerRangeA1: headerRangeA1,
      requiredHeaders: requiredHeaders
    });
    throw error;
  }
}

/**
 * 必須ヘッダーの存在をチェックする
 * @param {Array} headers - チェックするヘッダー配列
 * @param {Array} requiredHeaders - 必須ヘッダーのリスト
 * @returns {boolean} すべての必須ヘッダーが存在する場合はtrue
 */
function validateHeaders(headers, requiredHeaders) {
  try {
    console.log('ヘッダー検証を開始します', {
      headers: headers,
      requiredHeaders: requiredHeaders
    });
    
    const missingHeaders = requiredHeaders.filter(required => !headers.includes(required));
    
    if (missingHeaders.length > 0) {
      const errorMessage = `必須ヘッダーが見つかりません: ${missingHeaders.join(', ')}
    ガントチャートのヘッダー行に${missingHeaders.join(', ')}を追加してください。`;
      console.error(errorMessage, {
        missingHeaders: missingHeaders,
        availableHeaders: headers
      });
      throw new Error(errorMessage);
    }
    
    console.log('ヘッダー検証が完了しました - すべての必須ヘッダーが存在します');
    return true;
  } catch (error) {
    logLibraryError(error, 'validateHeaders', {
      headers: headers,
      requiredHeaders: requiredHeaders
    });
    throw error;
  }
}

/**
 * メンバー情報とガントチャートの共通ヘッダーを見つける
 * @param {Array} memberHeaders - メンバー情報のヘッダー
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @returns {Array} 共通するヘッダーの配列
 */
function findCommonHeaders(memberHeaders, ganttHeaders) {
  return memberHeaders.filter(header => ganttHeaders.includes(header));
}

/**
 * メンバーヘッダーとガントヘッダーのインデックスを事前計算
 * @param {Array} memberHeaders - メンバーヘッダーの配列
 * @param {Array} ganttHeaders - ガントヘッダーの配列
 * @returns {Object} ヘッダーのインデックス
 */
function prepareHeaderIndices(memberHeaders, ganttHeaders) {
  const headerIndices = {
    member: {}, // メンバーヘッダーのインデックス
    gantt: {}, // ガントヘッダーのインデックス
  };

  memberHeaders.forEach((header, index) => {
    headerIndices.member[header] = index;
  });

  ganttHeaders.forEach((header, index) => {
    headerIndices.gantt[header] = index;
  });

  return headerIndices;
}

/**
 * memberDateIdを生成する
 * @param {string} memberId - メンバーID
 * @param {string} date - 日付
 * @returns {string} 生成されたmemberDateId
 */
function generateMemberDateId(memberId, date) {
  return `${memberId}_${date}`;
}

/**
 * メンバーIDをmemberDateIdから抽出
 * @param {string} memberDateId - memberDateId
 * @returns {string} メンバーID
 */
function extractMemberId(memberDateId) {
  try {
    if (!memberDateId || typeof memberDateId !== 'string') {
      const errorMessage = 'memberDateIdが無効です';
      console.error(errorMessage, { memberDateId: memberDateId });
      throw new Error(errorMessage);
    }
    
    const parts = memberDateId.split('_');
    if (parts.length < 2) {
      console.warn('memberDateIdの形式が期待と異なります', {
        memberDateId: memberDateId,
        expectedFormat: 'memberId_date'
      });
    }
    
    const memberId = parts[0];
    
    if (!memberId) {
      const errorMessage = 'memberDateIdからメンバーIDを抽出できませんでした';
      console.error(errorMessage, { memberDateId: memberDateId });
      throw new Error(errorMessage);
    }
    
    return memberId;
  } catch (error) {
    logLibraryError(error, 'extractMemberId', {
      memberDateId: memberDateId,
      memberDateIdType: typeof memberDateId
    });
    throw error;
  }
}

/**
 * スプレッドシート内のすべてのシートを取得
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - スプレッドシート
 * @returns {Array} シートの配列
 */
function getAllSheets(spreadsheet) {
  return spreadsheet.getSheets();
}

/**
 * 特定のシートを除外する
 * @param {Array} sheets - シートの配列
 * @param {Array} excludeNames - 除外するシート名の配列
 * @returns {Array} 除外されたシートの配列
 */
function filterSheets(sheets, excludeNames) {
  return sheets.filter(sheet => !excludeNames.includes(sheet.getName()));
}

/**
 * メンバーデータの必須項目を検証し、エラーメッセージを設定する
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Array} エラーメッセージが追加された2次元配列
 */
function validateMemberData(memberData) {
  const headers = memberData[0];
  const deptIndex = headers.indexOf(COL_HEADER_NAMES.DEPT);
  const memberIdIndex = headers.indexOf(COL_HEADER_NAMES.MEMBER_ID);
  
  if (deptIndex === -1 || memberIdIndex === -1) {
    console.log(`警告: 必要な列が見つかりません - dept: ${deptIndex}, memberId: ${memberIdIndex}`);
    return memberData;
  }
  
  return memberData.map((row, index) => {
    // ヘッダー行はそのまま返す
    if (index === 0) return row;
    
    // 行の長さを調整
    while (row.length <= Math.max(deptIndex, memberIdIndex)) {
      row.push("");
    }
    
    const dept = row[deptIndex];
    const currentMemberId = row[memberIdIndex];
    
    // 既にエラーメッセージが設定されている場合はスキップ
    if (currentMemberId && currentMemberId.includes("エラー：")) {
      return row;
    }
    
    // dept列が空または無効な場合、エラーメッセージを設定
    if (!dept || (typeof dept === 'string' && dept.trim() === '')) {
      // 有効なmemberIdが既に存在する場合は上書きしない（emailが有効だった場合）
      if (!currentMemberId || currentMemberId.trim() === '') {
        const errorMessage = "エラー：deptが記入されていないため、ガントチャートに追加できませんでした";
        row[memberIdIndex] = errorMessage;
        
        console.log(`行${index + 1}: deptが空のためエラーメッセージを設定 - 「${errorMessage}」`);
      } else {
        // 既にmemberIdがある場合は、deptエラーの情報を追加
        const errorMessage = "エラー：deptが記入されていないため、ガントチャートに追加できませんでした";
        row[memberIdIndex] = errorMessage;
        
        console.log(`行${index + 1}: deptが空のため既存のmemberIdを上書き - 「${errorMessage}」`);
      }
    }
    
    return row;
  });
}

/**
 * メンバーIDを生成してデータに追加する
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Array} メンバーIDが追加された2次元配列
 */
function generateMemberIds(memberData) {
  const headers = memberData[0];
  const emailIndex = headers.indexOf(COL_HEADER_NAMES.EMAIL);
  let memberIdIndex = headers.indexOf(COL_HEADER_NAMES.MEMBER_ID);
  
  // memberIdカラムがなければ追加
  if (memberIdIndex === -1) {
    headers.push(COL_HEADER_NAMES.MEMBER_ID);
    memberIdIndex = headers.length - 1;
  }
  
  // メールアドレスからメンバーID生成
  return memberData.map((row, index) => {
    // ヘッダー行はそのまま返す
    if (index === 0) return row;
    
    // 行の長さを調整（新しい列が追加された場合）
    while (row.length <= memberIdIndex) {
      row.push("");
    }
    
    const email = row[emailIndex];
    if (email && typeof email === 'string' && email.trim() !== '') {
      // 有効なemailがある場合、memberIdを生成
      const memberId = email.split('@')[0];
      row[memberIdIndex] = memberId;
      
      console.log(`行${index + 1}: email「${email}」からmemberId「${memberId}」を生成`);
    } else {
      // emailが空または無効な場合、エラーメッセージを設定
      const errorMessage = "エラー：emailが記入されていないため、memberIdを生成できませんでした";
      row[memberIdIndex] = errorMessage;
      
      console.log(`行${index + 1}: emailが空のためエラーメッセージを設定 - 「${errorMessage}」`);
    }
    
    return row;
  });
}

/**
 * メンバーIDをキーにしてデータをマップ化
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Object} メンバーIDをキーとしたデータのマップ
 */
function createMemberDataMap(memberData) {
  try {
    console.log('メンバーデータマップの作成を開始します', {
      memberDataLength: memberData ? memberData.length : 'undefined'
    });
    
    if (!memberData || memberData.length === 0) {
      const errorMessage = 'メンバーデータが空です';
      console.error(errorMessage);
      throw new Error(errorMessage);
    }
    
    const headers = memberData[0];
    const memberIdIndex = headers.indexOf(COL_HEADER_NAMES.MEMBER_ID);
    
    if (memberIdIndex === -1) {
      const errorMessage = `必須ヘッダー「${COL_HEADER_NAMES.MEMBER_ID}」が見つかりません`;
      console.error(errorMessage, {
        availableHeaders: headers
      });
      throw new Error(errorMessage);
    }
    
    console.log('メンバーIDインデックスを取得しました', {
      memberIdIndex: memberIdIndex,
      headers: headers
    });
    
    // メンバーIDをキーにしたオブジェクトを作成
    // ここではforループの代わりにreduceを使用して、よりシンプルに実装
    const dataMap = memberData
      .slice(1) // ヘッダー行を除外
      .reduce((dataMap, row, index) => {
        try {
          const memberId = row[memberIdIndex];
          
          // memberIdが存在し、エラーメッセージでない場合のみマップに追加
          if (memberId && !String(memberId).includes("エラー：")) {
            dataMap[memberId] = {};
            
            // 各ヘッダーに対応する値をマップに設定
            headers.forEach((header, j) => {
              dataMap[memberId][header] = row[j];
            });
            
            console.log(`メンバーデータマップに追加: ${memberId}`);
          } else if (memberId && String(memberId).includes("エラー：")) {
            console.log(`エラーメッセージを含むmemberIdをスキップ: ${memberId}`);
          }
          
          return dataMap;
        } catch (rowError) {
          logLibraryError(rowError, `メンバーデータマップ作成 - 行${index + 2}の処理`, {
            rowIndex: index + 2,
            rowData: row,
            memberId: row[memberIdIndex]
          });
          // エラーが発生した行はスキップして続行
          return dataMap;
        }
      }, {});
    
    console.log('メンバーデータマップの作成が完了しました', {
      memberCount: Object.keys(dataMap).length,
      memberIds: Object.keys(dataMap)
    });
    
    return dataMap;
  } catch (error) {
    logLibraryError(error, 'createMemberDataMap', {
      memberDataLength: memberData ? memberData.length : 'undefined',
      memberDataHeaders: memberData && memberData.length > 0 ? memberData[0] : 'undefined'
    });
    throw error;
  }
}

/**
 * 共通ヘッダーに基づいてメンバーデータをガントチャートの行にコピー
 * @param {Array} commonHeaders - 共通するヘッダーの配列
 * @param {Object} headerIndices - ヘッダーのインデックス情報
 * @param {Array} ganttRow - コピー先のガントチャート行データ
 * @param {Object} memberData - コピー元のメンバーデータ (オブジェクト形式)
 * @param {Array} excludeHeaders - コピーから除外するヘッダー（オプション）
 * @returns {Array} 更新されたガントチャート行データ
 */
function copyMemberDataToGanttRow(commonHeaders, headerIndices, ganttRow, memberData, excludeHeaders = []) {
  try {
    console.log('メンバーデータのコピーを開始します', {
      commonHeaders: commonHeaders,
      excludeHeaders: excludeHeaders,
      memberDataKeys: Object.keys(memberData || {})
    });
    
    if (!commonHeaders || !Array.isArray(commonHeaders)) {
      const errorMessage = 'commonHeadersが無効です';
      console.error(errorMessage, { commonHeaders: commonHeaders });
      throw new Error(errorMessage);
    }
    
    if (!headerIndices || !headerIndices.gantt) {
      const errorMessage = 'headerIndicesが無効です';
      console.error(errorMessage, { headerIndices: headerIndices });
      throw new Error(errorMessage);
    }
    
    if (!ganttRow || !Array.isArray(ganttRow)) {
      const errorMessage = 'ganttRowが無効です';
      console.error(errorMessage, { ganttRow: ganttRow });
      throw new Error(errorMessage);
    }
    
    if (!memberData || typeof memberData !== 'object') {
      const errorMessage = 'memberDataが無効です';
      console.error(errorMessage, { memberData: memberData });
      throw new Error(errorMessage);
    }
    
    let copiedCount = 0;
    
    commonHeaders.forEach((header, index) => {
      try {
        // 除外ヘッダーリストをチェック
        if (!excludeHeaders.includes(header)) {
          const ganttIndex = headerIndices.gantt[header];
          
          if (ganttIndex !== undefined && memberData[header] !== undefined) {
            ganttRow[ganttIndex] = memberData[header];
            copiedCount++;
          }
        }
      } catch (headerError) {
        logLibraryError(headerError, `メンバーデータコピー - ヘッダー「${header}」の処理`, {
          header: header,
          headerIndex: index,
          ganttIndex: headerIndices.gantt[header],
          memberValue: memberData[header]
        });
        // エラーが発生したヘッダーはスキップして続行
      }
    });
    
    console.log('メンバーデータのコピーが完了しました', {
      copiedCount: copiedCount,
      totalHeaders: commonHeaders.length
    });
    
    return ganttRow;
  } catch (error) {
    logLibraryError(error, 'copyMemberDataToGanttRow', {
      commonHeaders: commonHeaders,
      excludeHeaders: excludeHeaders,
      ganttRowLength: ganttRow ? ganttRow.length : 'undefined',
      memberDataKeys: memberData ? Object.keys(memberData) : 'undefined',
      headerIndicesKeys: headerIndices ? Object.keys(headerIndices) : 'undefined'
    });
    throw error;
  }
} 