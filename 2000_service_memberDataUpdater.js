/**
 * メンバー情報の更新を行うサービス
 *
 * このモジュールは0000_common_utils.jsに依存しているため、
 * 以下の関数を利用します:
 * - getMemberDataAndHeaders
 * - getGanttHeaders
 * - findCommonHeaders
 * - extractMemberId
 * - createMemberDataMap
 * - getAllSheets
 * - filterSheets
 */

/**
 * ガントチャートデータを更新
 * @param {Array} ganttData - ガントチャートの2次元配列
 * @param {Object} memberDataMap - メンバーデータのマップ
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @param {Array} commonHeaders - 共通するヘッダー
 * @returns {Array} 更新されたガントチャートデータ
 */
function updateGanttData(ganttData, memberDataMap, ganttHeaders, commonHeaders) {
  // ヘッダーのインデックスを事前計算
  const headerIndices = prepareHeaderIndices([], ganttHeaders);
  const memberDateIdIndex = headerIndices.gantt[COL_HEADER_NAMES.MEMBER_DATE_ID];
  
  // ヘッダー行を除いたデータ行を処理
  return ganttData.map((row, index) => {
    // ヘッダー行はそのまま返す
    if (index === 0) return row;
    
    const memberDateId = row[memberDateIdIndex];
    // 空白行はそのまま返す
    if (!memberDateId) return row;
    
    const memberId = extractMemberId(memberDateId);
    const memberData = memberDataMap[memberId];
    
    // メンバーデータが存在する場合のみ更新
    if (memberData) {
      // 共通関数を使用してデータをコピー
      copyMemberDataToGanttRow(
        commonHeaders,
        headerIndices,
        row,
        memberData,
        [COL_HEADER_NAMES.MEMBER_DATE_ID]
      );
    }
    
    return row;
  });
}

/**
 * ガントチャートシートを更新する
 * @param {SpreadsheetApp.Sheet} sheet - 更新対象のシート
 * @param {string} headerRangeA1 - ヘッダー範囲のA1記法
 * @param {Object} memberDataMap - メンバーデータのマップ
 * @param {Array} memberHeaders - メンバーデータのヘッダー
 * @returns {boolean} 更新が成功した場合はtrue
 */
function updateGanttSheet(sheet, headerRangeA1, memberDataMap, memberHeaders) {
  try {
    // ガントチャートのヘッダー行を取得
    const {
      headers: ganttHeaders,
      headerRow,
      startCol,
      endCol,
    } = getGanttHeaders(sheet, headerRangeA1, REQUIRED_MEMBER_DATA_HEADERS.GANTT_SHEETS.UPDATE);

    // 共通ヘッダーを見つける
    const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);

    if (commonHeaders.length > 0) {
      // データ範囲を取得（ヘッダー行の下から）
      const lastRow = sheet.getLastRow();
      const dataRows = lastRow - headerRow;
      if (dataRows <= 0) return false;

      const dataRange = sheet.getRange(headerRow + 1, startCol, dataRows, endCol - startCol + 1);
      const ganttData = dataRange.getValues();

      // データを更新
      const updatedData = updateGanttData(ganttData, memberDataMap, ganttHeaders, commonHeaders);

      // 更新されたデータをシートに書き込み
      dataRange.setValues(updatedData);
      return true;
    }
    return false;
  } catch (error) {
    console.error(`シート「${sheet.getName()}」の処理中にエラーが発生しました: ${error.message}`);
    return false;
  }
}

/**
 * メンバー情報の更新を実行
 */
function updateMemberDataInGanttCharts() {
  const ui = SpreadsheetApp.getUi();

  try {
    // スクリプトプロパティから情報を取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const targetUrl = scriptProperties.getProperty("GANTT_SS");
    const headerRangeA1 = scriptProperties.getProperty("HEADER_RANGE_A1");

    if (!targetUrl || !headerRangeA1) {
      throw new Error(
        "スプレッドシートURLまたはヘッダー範囲が設定されていません。先にガントチャートテンプレート複製を実行してください。"
      );
    }

    // 対象のスプレッドシートを開く
    const ganttSs = SpreadsheetApp.openByUrl(targetUrl);

    // メンバー情報の取得（コンテナバインドされているスプレッドシートから）
    const containerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const { data: memberData, headers: memberHeaders } = getMemberDataAndHeaders(
      containerSpreadsheet,
      REQUIRED_MEMBER_DATA_HEADERS.DATA_SHEET.UPDATE
    );

    // メンバーデータのマップを作成
    const memberDataMap = createMemberDataMap(memberData);

    // ガントチャートシートを取得（メンバーリストシートは除外）
    const ganttSheets = getAllSheets(ganttSs);

    // 更新
    const updatedSheets = [];
    ganttSheets.forEach((sheet) => {
      const isUpdated = updateGanttSheet(sheet, headerRangeA1, memberDataMap, memberHeaders);
      if (isUpdated) {
        updatedSheets.push(sheet.getName());
      }
    });

    if (updatedSheets.length > 0) {
      ui.alert(`以下のシートが更新されました: ${updatedSheets.join(", ")}`);
    } else {
      ui.alert("更新されたシートはありません。");
    }
  } catch (error) {
    ui.alert(`エラー: ${error.message}`);
  }
}

/**
 * メンバー管理メニューを作成
 * @param {SpreadsheetApp.Ui} ui - SpreadsheetAppのUIオブジェクト
 */
function buildMemberMenu(ui) {
  ui.createMenu("メンバー管理")
    .addItem("🟩メンバー情報を全シートに更新", "updateMemberDataInGanttCharts")
    .addToUi();
}
