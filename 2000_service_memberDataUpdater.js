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
 * エラーの詳細情報をログに出力するヘルパー関数
 * @param {Error} error - エラーオブジェクト
 * @param {string} context - エラーが発生したコンテキスト
 * @param {Object} additionalInfo - 追加のデバッグ情報（オプション）
 */
function logDetailedError(error, context, additionalInfo = {}) {
  console.error(`=== エラー詳細 ===`);
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
  
  console.error(`=== エラー詳細終了 ===`);
}

/**
 * ガントチャートデータを更新
 * @param {Array} ganttData - ガントチャートの2次元配列
 * @param {Object} memberDataMap - メンバーデータのマップ
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @param {Array} commonHeaders - 共通するヘッダー
 * @returns {Array} 更新されたガントチャートデータ
 */
function updateGanttData(ganttData, memberDataMap, ganttHeaders, commonHeaders) {
  try {
    // メンバーデータのヘッダーを取得（マップからサンプルを取得）
    const memberHeaders = Object.keys(memberDataMap).length > 0 
      ? Object.keys(Object.values(memberDataMap)[0]) 
      : [];
    
    // ヘッダーのインデックスを事前計算
    const headerIndices = prepareHeaderIndices(memberHeaders, ganttHeaders);
    const memberDateIdIndex = headerIndices.gantt[COL_HEADER_NAMES.MEMBER_DATE_ID];
    
    // データ行を処理（ganttDataにはヘッダー行は含まれていない）
    return ganttData.map((row, index) => {
      const memberDateId = row[memberDateIdIndex];
      if (!memberDateId) return row;
      
      const memberId = extractMemberId(memberDateId);
      const memberData = memberDataMap[memberId];
      
      if (memberData) {
        const originalRow = [...row]; // 元の行を保存
        const newRow = [...row];
        
        console.log(`行${index}: 処理前`, { originalRow, newRow, memberData, commonHeaders });
        
        copyMemberDataToGanttRow(
          commonHeaders,
          headerIndices,
          newRow,
          memberData,
          [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.DATE, COL_HEADER_NAMES.MEMBER_ID]
        );
        
        console.log(`行${index}: 処理後`, { originalRow, newRow, changed: JSON.stringify(originalRow) !== JSON.stringify(newRow) });
        
        return newRow;
      }
      
      return row;
    });
  } catch (error) {
    logDetailedError(error, 'ガントチャートデータ更新', {
      ganttDataLength: ganttData ? ganttData.length : 'undefined',
      ganttHeaders: ganttHeaders,
      commonHeaders: commonHeaders,
      memberDataMapKeys: Object.keys(memberDataMap || {})
    });
    throw error;
  }
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
  const sheetName = sheet ? sheet.getName() : 'unknown';
  
  try {
    console.log(`シート「${sheetName}」の更新を開始します`);
    
    // ガントチャートのヘッダー行を取得
    const {
      headers: ganttHeaders,
      headerRow,
      startCol,
      endCol,
    } = getGanttHeaders(sheet, headerRangeA1, REQUIRED_MEMBER_DATA_HEADERS.GANTT_SHEETS.UPDATE);

    console.log(`シート「${sheetName}」のヘッダー情報を取得しました`, {
      ganttHeaders: ganttHeaders,
      headerRow: headerRow,
      startCol: startCol,
      endCol: endCol
    });

    // 共通ヘッダーを見つける
    const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);
    console.log(`シート「${sheetName}」の共通ヘッダー:`, commonHeaders);

    if (commonHeaders.length > 0) {
      // データ範囲を取得（ヘッダー行の下から）
      const lastRow = sheet.getLastRow();
      const dataRows = lastRow - headerRow;
      
      console.log(`シート「${sheetName}」のデータ範囲情報`, {
        lastRow: lastRow,
        headerRow: headerRow,
        dataRows: dataRows
      });
      
      if (dataRows <= 0) {
        console.log(`シート「${sheetName}」にはデータ行がありません`);
        return false;
      }

      const dataRange = sheet.getRange(headerRow + 1, startCol, dataRows, endCol - startCol + 1);
      const ganttData = dataRange.getValues();

      console.log(`シート「${sheetName}」のデータを取得しました（${ganttData.length}行）`);

      // データを更新
      const updatedData = updateGanttData(ganttData, memberDataMap, ganttHeaders, commonHeaders);

      // 更新されたデータをシートに書き込み
      dataRange.setValues(updatedData);
      
      console.log(`シート「${sheetName}」の更新が完了しました`);
      return true;
    } else {
      console.log(`シート「${sheetName}」には共通ヘッダーがありません`);
      return false;
    }
  } catch (error) {
    logDetailedError(error, `シート「${sheetName}」の処理`, {
      sheetName: sheetName,
      headerRangeA1: headerRangeA1,
      memberHeaders: memberHeaders,
      memberDataMapKeys: Object.keys(memberDataMap || {})
    });
    return false;
  }
}

/**
 * メンバー情報の更新を実行
 */
function updateMemberDataInGanttCharts() {
  const ui = SpreadsheetApp.getUi();

  try {
    console.log('メンバー情報の更新処理を開始します');
    
    // スクリプトプロパティから情報を取得
    const scriptProperties = PropertiesService.getScriptProperties();
    const targetUrl = scriptProperties.getProperty("GANTT_SS");
    const headerRangeA1 = scriptProperties.getProperty("HEADER_RANGE_A1");

    console.log('スクリプトプロパティを取得しました', {
      targetUrl: targetUrl ? '設定済み' : '未設定',
      headerRangeA1: headerRangeA1 ? headerRangeA1 : '未設定'
    });

    if (!targetUrl || !headerRangeA1) {
      const errorMessage = "スプレッドシートURLまたはヘッダー範囲が設定されていません。先にガントチャートテンプレート複製を実行してください。";
      console.error(errorMessage);
      throw new Error(errorMessage);
    }

    // 対象のスプレッドシートを開く
    console.log('対象スプレッドシートを開いています...');
    const ganttSs = SpreadsheetApp.openByUrl(targetUrl);
    console.log(`対象スプレッドシート「${ganttSs.getName()}」を開きました`);

    // メンバー情報の取得（コンテナバインドされているスプレッドシートから）
    console.log('メンバー情報を取得しています...');
    const containerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const { data: memberData, headers: memberHeaders } = getMemberDataAndHeaders(
      containerSpreadsheet,
      REQUIRED_MEMBER_DATA_HEADERS.DATA_SHEET.UPDATE
    );

    console.log('メンバー情報を取得しました', {
      memberDataRows: memberData.length,
      memberHeaders: memberHeaders
    });

    // メンバーデータのマップを作成
    console.log('メンバーデータマップを作成しています...');
    const memberDataMap = createMemberDataMap(memberData);
    console.log(`メンバーデータマップを作成しました（${Object.keys(memberDataMap).length}件）`);

    // ガントチャートシートを取得（メンバーリストシートは除外）
    console.log('ガントチャートシートを取得しています...');
    const ganttSheets = getAllSheets(ganttSs);
    console.log(`ガントチャートシートを取得しました（${ganttSheets.length}シート）`);

    // 更新
    const updatedSheets = [];
    const failedSheets = [];
    
    ganttSheets.forEach((sheet) => {
      console.log(`シート「${sheet.getName()}」の処理を開始します`);
      const isUpdated = updateGanttSheet(sheet, headerRangeA1, memberDataMap, memberHeaders);
      if (isUpdated) {
        updatedSheets.push(sheet.getName());
        console.log(`シート「${sheet.getName()}」の更新に成功しました`);
      } else {
        failedSheets.push(sheet.getName());
        console.log(`シート「${sheet.getName()}」の更新をスキップしました`);
      }
    });

    console.log('全シートの処理が完了しました', {
      updatedSheets: updatedSheets,
      failedSheets: failedSheets
    });

    if (updatedSheets.length > 0) {
      const successMessage = `以下のシートが更新されました: ${updatedSheets.join(", ")}`;
      console.log(successMessage);
      ui.alert(successMessage);
    } else {
      const noUpdateMessage = "更新されたシートはありません。";
      console.log(noUpdateMessage);
      ui.alert(noUpdateMessage);
    }
  } catch (error) {
    logDetailedError(error, 'メンバー情報更新のメイン処理', {
      timestamp: new Date().toISOString()
    });
    ui.alert(`エラー: ${error.message}\n\n詳細はログを確認してください。`);
  }
}

/**
 * メンバー管理メニューを作成
 * @param {SpreadsheetApp.Ui} ui - SpreadsheetAppのUIオブジェクト
 */
function buildMemberMenu(ui) {
  return ui.createMenu("3.メンバー情報更新")
    .addItem("「2~3.メンバー情報」のデータをシフト表SSの全シートに反映", "updateMemberDataInGanttCharts");
}
