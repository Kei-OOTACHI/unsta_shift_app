/**
 * メンバー管理システム共通ユーティリティ
 */

// 共通定数
const MEMBER_DATA_SHEET_NAME = "メンバーリスト";
const REQUIRED_MEMBER_HEADERS = ["dept", "email"];
const REQUIRED_GANTT_HEADERS = ["memberDateId", "date"];

/**
 * メンバー情報マスタシートのデータを取得する
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - メンバー情報が含まれるスプレッドシート
 * @returns {Array} メンバー情報の2次元配列
 */
function getMemberData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(MEMBER_DATA_SHEET_NAME);
  if (!sheet) {
    throw new Error(`シート「${MEMBER_DATA_SHEET_NAME}」が見つかりません。`);
  }
  
  const dataRange = sheet.getDataRange();
  return dataRange.getValues();
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
    throw new Error(`必須ヘッダーが見つかりません: ${missingHeaders.join(', ')}`);
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
  const emailIndex = headers.indexOf('email');
  let memberIdIndex = headers.indexOf('memberId');
  
  // memberIdカラムがなければ追加
  if (memberIdIndex === -1) {
    headers.push('memberId');
    memberIdIndex = headers.length - 1;
  }
  
  // メールアドレスからメンバーID生成
  for (let i = 1; i < memberData.length; i++) {
    const email = memberData[i][emailIndex];
    if (email && typeof email === 'string') {
      const memberId = email.split('@')[0];
      memberData[i][memberIdIndex] = memberId;
    }
  }
  
  return memberData;
}

/**
 * メンバーIDをキーにしてデータをマップ化
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Object} メンバーIDをキーとしたデータのマップ
 */
function createMemberDataMap(memberData) {
  const headers = memberData[0];
  const memberIdIndex = headers.indexOf('memberId');
  const dataMap = {};
  
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const memberId = row[memberIdIndex];
    if (memberId) {
      dataMap[memberId] = {};
      
      for (let j = 0; j < headers.length; j++) {
        dataMap[memberId][headers[j]] = row[j];
      }
    }
  }
  
  return dataMap;
} 