/**
 * メンバー情報の更新を行うサービス
 *
 * このモジュールは0000_common_utils.jsに依存しているため、
 * 以下の関数を利用します:
 * - getMemberData
 * - validateHeaders
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
  const memberDateIdIndex = ganttHeaders.indexOf("memberDateId");

  for (let i = 1; i < ganttData.length; i++) {
    const row = ganttData[i];
    const memberDateId = row[memberDateIdIndex];

    if (!memberDateId) continue; // 空白行はスキップ

    const memberId = extractMemberId(memberDateId);
    const memberData = memberDataMap[memberId];

    if (memberData) {
      // 共通ヘッダーのデータを更新
      commonHeaders.forEach((header) => {
        if (header !== "memberDateId") {
          const ganttIndex = ganttHeaders.indexOf(header);
          if (ganttIndex !== -1) {
            row[ganttIndex] = memberData[header];
          }
        }
      });
    }
  }

  return ganttData;
}

/**
 * メンバー情報の更新を実行
 */
function updateMemberDataInGanttCharts() {
  const ui = SpreadsheetApp.getUi();

  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // メンバー情報の取得
    const memberData = getMemberData(activeSpreadsheet);
    const memberHeaders = memberData[0];

    // ヘッダー検証（memberId必須）
    const requiredHeaders = [...REQUIRED_MEMBER_HEADERS, "memberId"];
    validateHeaders(memberHeaders, requiredHeaders);

    // メンバーデータのマップを作成
    const memberDataMap = createMemberDataMap(memberData);

    // ガントチャートシートを取得（メンバーリストシートは除外）
    const allSheets = getAllSheets(activeSpreadsheet);
    const ganttSheets = filterSheets(allSheets, [MEMBER_DATA_SHEET_NAME]);

    // 更新対象のヘッダーを特定
    const updatedSheets = [];

    ganttSheets.forEach((sheet) => {
      // ヘッダー行を取得
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      const ganttHeaders = headerRange.getValues()[0];

      try {
        // ヘッダー検証（memberDateId必須）
        validateHeaders(ganttHeaders, ["memberDateId"]);

        // 共通ヘッダーを見つける
        const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);

        if (commonHeaders.length > 0) {
          // データ範囲を取得
          const dataRange = sheet.getDataRange();
          const ganttData = dataRange.getValues();

          // データを更新
          const updatedData = updateGanttData(ganttData, memberDataMap, ganttHeaders, commonHeaders);

          // 更新されたデータをシートに書き込み
          dataRange.setValues(updatedData);

          updatedSheets.push(sheet.getName());
        }
      } catch (error) {
        // このシートをスキップして次に進む
        console.error(`シート「${sheet.getName()}」の処理中にエラーが発生しました: ${error.message}`);
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
 * メニューに項目を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("メンバー管理").addItem("メンバー情報を全シートに更新", "updateMemberDataInGanttCharts").addToUi();
}
