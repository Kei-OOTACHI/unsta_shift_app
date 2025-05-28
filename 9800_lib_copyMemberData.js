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
    INITIALIZE: [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.DATE],
    UPDATE: [COL_HEADER_NAMES.MEMBER_DATE_ID]
  }
};
const MEMBER_DATA_SHEET_NAME = "メンバー情報";
const GANTT_TEMPLATE_SHEET_NAME = "GCテンプレ";

/**
 * メンバーリストシートからデータとヘッダーを取得する
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - メンバーリストが含まれるスプレッドシート
 * @param {Array} requiredHeaders - 必須ヘッダーのリスト
 * @returns {Object} メンバーデータとヘッダーのオブジェクト
 * @property {Array} data - メンバーデータの2次元配列
 * @property {Array} headers - ヘッダー行の配列
 */
function getMemberDataAndHeaders(spreadsheet, requiredHeaders) {
  const memberSheet = spreadsheet.getSheetByName(MEMBER_DATA_SHEET_NAME);
  if (!memberSheet) {
    throw new Error(`シート「${MEMBER_DATA_SHEET_NAME}」が見つかりません。`);
  }
  
  const memberDataRange = memberSheet.getDataRange();
  const memberData = memberDataRange.getValues();
  const memberHeaders = memberData[0];

  // ヘッダー検証を内部で実行
  validateHeaders(memberHeaders, requiredHeaders);
  
  return {
    data: memberData,
    headers: memberHeaders
  };
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
  const headerRangeParts = headerRangeA1.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
  if (!headerRangeParts) {
    throw new Error(`ヘッダー範囲の形式が不正です: ${headerRangeA1}`);
  }

  const headerRow = parseInt(headerRangeParts[2]);
  const startCol = sheet.getRange(headerRangeParts[1] + "1").getColumn();
  const endCol = sheet.getRange(headerRangeParts[3] + "1").getColumn();
  const headerRange = sheet.getRange(headerRow, startCol, 1, endCol - startCol + 1);
  const headers = headerRange.getValues()[0];

  // ヘッダー検証を内部で実行
  validateHeaders(headers, requiredHeaders);

  return {
    headers,
    headerRow,
    startCol,
    endCol
  };
}

/**
 * 必須ヘッダーの存在をチェックする
 * @param {Array} headers - チェックするヘッダー配列
 * @param {Array} requiredHeaders - 必須ヘッダーのリスト
 * @returns {boolean} すべての必須ヘッダーが存在する場合はtrue
 */
function validateHeaders(headers, requiredHeaders) {
  const missingHeaders = requiredHeaders.filter(required => !headers.includes(required));
  
  if (missingHeaders.length > 0) {
    throw new Error(`必須ヘッダーが見つかりません: ${missingHeaders.join(', ')}
    ${GANTT_TEMPLATE_SHEET_NAME}のヘッダー行に${requiredHeaders.join(', ')}を追加してください。`);
  }
  
  return true;
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
  return memberDateId.split('_')[0];
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
    
    const email = row[emailIndex];
    if (email && typeof email === 'string') {
      const memberId = email.split('@')[0];
      row[memberIdIndex] = memberId;
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
  const headers = memberData[0];
  const memberIdIndex = headers.indexOf(COL_HEADER_NAMES.MEMBER_ID);
  
  // メンバーIDをキーにしたオブジェクトを作成
  // ここではforループの代わりにreduceを使用して、よりシンプルに実装
  return memberData
    .slice(1) // ヘッダー行を除外
    .reduce((dataMap, row) => {
      const memberId = row[memberIdIndex];
      if (memberId) {
        dataMap[memberId] = {};
        
        // 各ヘッダーに対応する値をマップに設定
        headers.forEach((header, j) => {
          dataMap[memberId][header] = row[j];
        });
      }
      return dataMap;
    }, {});
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
  commonHeaders.forEach((header) => {
    // 除外ヘッダーリストをチェック
    if (!excludeHeaders.includes(header)) {
      const ganttIndex = headerIndices.gantt[header];
      
      if (ganttIndex !== undefined && memberData[header] !== undefined) {
        ganttRow[ganttIndex] = memberData[header];
      }
    }
  });
  
  return ganttRow;
} 