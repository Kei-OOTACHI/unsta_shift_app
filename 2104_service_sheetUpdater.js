/**
 * シート更新サービス（2101_service_sheetUpdater.js）
 *
 * このファイルは、2100_service_shiftDataMerger.jsのsetDataToSheets関数を
 * 機能ごとに分割した関数群を含みます。
 *
 * 主な機能：
 * 1. データベース系シート（RDB、CONFLICT、ERROR）の更新
 * 2. ガントチャートシートの更新
 * 3. 個別のガントシート処理（作成、データ設定、セル結合）
 *
 * パフォーマンス最適化：
 * - データ準備の事前実行
 * - 空データのバリデーション
 * - シート操作の効率化
 */

/**
 * メインのsetDataToSheets関数（リファクタリング版）
 * @param {Object} OutGanttSs - 出力ガントチャートスプレッドシート
 * @param {Object} OutMergedRdbSheet - 出力マージRDBシート
 * @param {Object} OutConflictRdbSheet - 出力コンフリクトRDBシート
 * @param {Object} OutErrorRdbSheet - 出力エラーRDBシート
 * @param {Object} ganttData - ガントデータ
 * @param {Array} rdbData - RDBデータ
 * @param {Array} conflictData - コンフリクトデータ
 * @param {Array} errorData - エラーデータ
 */
function setDataToSheets(
  OutGanttSs,
  OutMergedRdbSheet,
  OutConflictRdbSheet,
  OutErrorRdbSheet,
  ganttData,
  rdbData,
  conflictData,
  errorData
) {
  const startTime = new Date();
  const failedSheets = [];

  try {
    // データベース系シートの更新
    const sheets = { OutMergedRdbSheet, OutConflictRdbSheet, OutErrorRdbSheet };
    updateDatabaseSheets(sheets, rdbData, conflictData, errorData, failedSheets);
  } catch (error) {
    showRestorePrompt(failedSheets, "現在のスプレッドシート", startTime, error);
    throw new Error(`データ更新処理を停止しました: ${error.message}`);
  }

  // ガントチャートシートの更新
  updateGanttChartSheets(OutGanttSs, ganttData, startTime);
}

/**
 * データベース系シートを更新する
 * @param {Object} sheets - 更新対象のシートオブジェクト
 * @param {Array} rdbData - RDBデータ
 * @param {Array} conflictData - コンフリクトデータ
 * @param {Array} errorData - エラーデータ
 * @param {Array} failedSheets - 失敗したシートのリスト（参照渡し）
 */
function updateDatabaseSheets(sheets, rdbData, conflictData, errorData, failedSheets) {
  const { OutMergedRdbSheet, OutConflictRdbSheet, OutErrorRdbSheet } = sheets;

  // 処理効率化のため、先に全データの準備を行う
  const preparedRdbData = [...rdbData]; // コピーしてunshiftでの元データ変更を防ぐ
  const preparedConflictData = [...conflictData];
  const preparedErrorData = [...errorData];

  preparedRdbData.unshift(getColumnOrder(RDB_COL_INDEXES));
  preparedConflictData.unshift(getColumnOrder(CONFLICT_COL_INDEXES));
  if (preparedErrorData.length > 0) {
    preparedErrorData.unshift(getColumnOrder(ERROR_COL_INDEXES));
  }

  // RDBシートの更新
  try {
    SpreadsheetApp.getActive().toast(`${SHEET_NAMES.OUT_RDB}シートの更新を開始します...`, "処理状況");
    OutMergedRdbSheet.clear();
    if (preparedRdbData.length > 0 && preparedRdbData[0].length > 0) {
      OutMergedRdbSheet.getRange(1, 1, preparedRdbData.length, preparedRdbData[0].length).setValues(preparedRdbData);
    }
    console.log(`${SHEET_NAMES.OUT_RDB}シートの更新が完了しました`);
  } catch (error) {
    failedSheets.push(`${SHEET_NAMES.OUT_RDB}シート`);
    throw error;
  }

  // コンフリクトシートの更新
  try {
    SpreadsheetApp.getActive().toast(`${SHEET_NAMES.CONFLICT_RDB}シートの更新を開始します...`, "処理状況");
    OutConflictRdbSheet.clear();
    if (preparedConflictData.length > 0 && preparedConflictData[0].length > 0) {
      OutConflictRdbSheet.getRange(1, 1, preparedConflictData.length, preparedConflictData[0].length).setValues(
        preparedConflictData
      );
    }
    console.log(`${SHEET_NAMES.CONFLICT_RDB}シートの更新が完了しました`);
  } catch (error) {
    failedSheets.push(`${SHEET_NAMES.CONFLICT_RDB}シート`);
    throw error;
  }

  // エラーシートの更新
  try {
    SpreadsheetApp.getActive().toast(`${SHEET_NAMES.ERROR_RDB}シートの更新を開始します...`, "処理状況");
    OutErrorRdbSheet.clear();
    // エラーデータの書き込み（Ganttに存在しない部署のRDBデータ）
    if (preparedErrorData.length > 0 && preparedErrorData[0].length > 0) {
      OutErrorRdbSheet.getRange(1, 1, preparedErrorData.length, preparedErrorData[0].length).setValues(
        preparedErrorData
      );
    }
    console.log(`${SHEET_NAMES.ERROR_RDB}シートの更新が完了しました`);
  } catch (error) {
    failedSheets.push(`${SHEET_NAMES.ERROR_RDB}シート`);
    throw error;
  }
}

/**
 * ガントチャートシートを更新する
 * @param {Object} OutGanttSs - 出力ガントチャートスプレッドシート
 * @param {Object} ganttData - ガントチャートデータ
 * @param {Date} startTime - 処理開始時刻
 */
function updateGanttChartSheets(OutGanttSs, ganttData, startTime) {
  const ganttSsName = OutGanttSs.getName();
  const ganttSheets = Object.entries(ganttData);
  const totalSheets = ganttSheets.length;

  for (let i = 0; i < ganttSheets.length; i++) {
    const [sheetName, sheetData] = ganttSheets[i];

    try {
      SpreadsheetApp.getActive().toast(
        `ガントチャート「${sheetName}」の更新を開始します... (${i + 1}/${totalSheets})`,
        "処理状況"
      );

      // 空のガントデータの場合はスキップ
      if (isEmptyGanttData(sheetData.ganttShiftValues)) {
        continue;
      }

      processGanttSheet(OutGanttSs, sheetName, sheetData, ganttSsName);

      console.log(`ガントチャート「${ganttSsName}」のシート「${sheetName}」の作成が完了しました`);
      SpreadsheetApp.getActive().toast(
        `ガントチャート「${sheetName}」の更新が完了しました (${i + 1}/${totalSheets})`,
        "処理状況"
      );
    } catch (error) {
      showRestorePrompt(
        [`シート「${sheetName}」`],
        `ガントチャートスプレッドシート「${ganttSsName}」`,
        startTime,
        error
      );
      throw new Error(`ガントチャート更新処理を停止しました: ${error.message}`);
    }
  }
}

/**
 * 個別のガントシートを処理する
 * @param {Object} OutGanttSs - 出力ガントチャートスプレッドシート
 * @param {string} sheetName - シート名
 * @param {Object} sheetData - シートデータ
 * @param {string} ganttSsName - ガントスプレッドシート名
 */
function processGanttSheet(OutGanttSs, sheetName, sheetData, ganttSsName) {
  const { ganttHeaderValues, ganttShiftValues, ganttHeaderBgs, ganttShiftBgs, firstDataRowOffset, firstDataColOffset } =
    sheetData;

  // 既存のシートを削除して新しいシートを作成
  const newSheet = recreateGanttSheet(OutGanttSs, sheetName);

  // ガントデータを統合してシートに設定
  setGanttDataToSheet(
    newSheet,
    ganttHeaderValues,
    ganttShiftValues,
    ganttHeaderBgs,
    ganttShiftBgs,
    firstDataColOffset,
    firstDataRowOffset
  );

  // セル結合処理
  applyGanttCellMerging(newSheet, ganttShiftValues, firstDataRowOffset, firstDataColOffset);
}

/**
 * 空のガントデータかどうかを判定する
 * @param {Array} ganttShiftValues - ガントシフト値
 * @returns {boolean} 空のデータの場合true
 */
function isEmptyGanttData(ganttShiftValues) {
  return (
    !ganttShiftValues ||
    ganttShiftValues.length === 0 ||
    (ganttShiftValues.length === 1 && ganttShiftValues[0].length === 0)
  );
}

/**
 * ガントシートを削除して新しいシートを作成する
 * @param {Object} OutGanttSs - 出力ガントチャートスプレッドシート
 * @param {string} sheetName - シート名
 * @returns {Object} 新しく作成されたシート
 */
function recreateGanttSheet(OutGanttSs, sheetName) {
  // 既存のシートを削除
  const existingSheet = OutGanttSs.getSheetByName(sheetName);
  if (existingSheet) {
    // シートが1枚しかない場合、削除前に一時的なシートを追加
    if (OutGanttSs.getSheets().length === 1) {
      OutGanttSs.insertSheet("TemporarySheet");
    }
    OutGanttSs.deleteSheet(existingSheet);
  }

  // テンプレートシートをコピーして新しいシートを作成
  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.GANTT_TEMPLATE);
  if (!templateSheet) {
    throw new Error(`テンプレートシート「${SHEET_NAMES.GANTT_TEMPLATE}」が見つかりません。`);
  }

  // テンプレートをコピーして新しいシートを作成
  const newSheet = templateSheet.copyTo(OutGanttSs);
  newSheet.setName(sheetName);

  // 一時的なシートが存在する場合は削除
  const tempSheet = OutGanttSs.getSheetByName("TemporarySheet");
  if (tempSheet) {
    OutGanttSs.deleteSheet(tempSheet);
  }

  return newSheet;
}

/**
 * ガントデータをシートに設定する
 * @param {Object} sheet - 対象シート
 * @param {Array} ganttHeaderValues - ガントヘッダー値
 * @param {Array} ganttShiftValues - ガントシフト値
 * @param {Array} ganttHeaderBgs - ガントヘッダー背景色
 * @param {Array} ganttShiftBgs - ガントシフト背景色
 * @param {number} firstDataColOffset - 最初のデータ列オフセット
 * @param {number} firstDataRowOffset - 最初のデータ行オフセット
 */
function setGanttDataToSheet(
  sheet,
  ganttHeaderValues,
  ganttShiftValues,
  ganttHeaderBgs,
  ganttShiftBgs,
  firstDataColOffset,
  firstDataRowOffset
) {
  // mergeGanttData関数を使用してヘッダーとシフトデータを組み合わせ
  const { values: fullValues, backgrounds: fullBgs } = mergeGanttData(
    ganttHeaderValues,
    ganttShiftValues,
    ganttHeaderBgs,
    ganttShiftBgs,
    firstDataColOffset,
    firstDataRowOffset
  );

  // 完全なデータをシートに設定
  try {
    if (fullValues && fullValues.length > 0 && fullValues[0].length > 0) {
      const fullRange = sheet.getRange(1, 1, fullValues.length, fullValues[0].length);

      // 行の高さが変わらないように、現在の行の高さを保存
      const currentRowHeights = [];
      for (let i = 1; i <= fullValues.length; i++) {
        currentRowHeights.push(sheet.getRowHeight(i));
      }
      fullRange.breakApart();
      // データと背景色を設定
      fullRange.setValues(fullValues);
      fullRange.setBackgrounds(fullBgs);

      // 行の高さを元に戻す
      for (let i = 0; i < currentRowHeights.length; i++) {
        sheet.setRowHeight(i + 1, currentRowHeights[i]);
      }
    }
  } catch (error) {
    throw new Error(`完全データ設定中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * ガントシートにセル結合を適用する
 * @param {Object} sheet - 対象シート
 * @param {Array} ganttShiftValues - ガントシフト値
 * @param {number} firstDataRowOffset - 最初のデータ行オフセット
 * @param {number} firstDataColOffset - 最初のデータ列オフセット
 */
function applyGanttCellMerging(sheet, ganttShiftValues, firstDataRowOffset, firstDataColOffset) {
  try {
    const startRow = firstDataRowOffset + 1; // 1-indexedに変換
    const startCol = firstDataColOffset + 1; // 1-indexedに変換
    const headerRange = sheet.getRange(startRow, 1, ganttShiftValues.length, startCol - 1);
    mergeSameValuesVertically(sheet, headerRange);
    const shiftRange = sheet.getRange(startRow, startCol, ganttShiftValues.length, ganttShiftValues[0].length);
    mergeSameValuesHorizontally(sheet, shiftRange);
  } catch (e) {
    throw new Error(`セル結合処理でエラーが発生しました: ${e.message}`);
  }
}

// エラー発生時の復元案内を表示する関数
function showRestorePrompt(failedSheets, targetDescription, startTime, error) {
  const formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  let message =
    "■ データ更新処理でエラーが発生しました\n" +
    "【失敗箇所】\n" +
    targetDescription +
    "の以下のシート:\n" +
    failedSheets.map((sheet) => "・" + sheet).join("\n") +
    "\n\n" +
    "【エラー詳細】\n" +
    error.message +
    "\n\n" +
    "【復元方法】\n" +
    formattedStartTime +
    "\n" +
    "以下の手順で履歴から復元してください:\n" +
    "1. 対象のスプレッドシートを開く\n" +
    "2. ファイルメニュー → 「バージョン履歴」 → 「バージョン履歴を表示」を選択\n" +
    "3. 処理開始時刻(" +
    formattedStartTime +
    ")より前の最新バージョンを選択\n" +
    "4. 「このバージョンを復元」をクリック\n" +
    "復元完了後、問題を修正してから再度処理を実行してください。";

  Browser.msgBox("データ更新エラー - 履歴からの復元が必要", message, Browser.Buttons.OK);

  // ログにも出力
  console.error("=== データ更新処理エラー ===");
  console.error(`失敗箇所: ${targetDescription}`);
  console.error(`失敗シート: ${failedSheets.join(", ")}`);
  console.error(`処理開始時刻: ${formattedStartTime}`);
  console.error(`エラー: ${error.message}`);
  console.error("Stack trace:", error.stack);
}
