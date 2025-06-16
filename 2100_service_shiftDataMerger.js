// 列・行インデックスの直接参照が可能です:
// RDB_COL_INDEXES.dept, GANTT_COL_INDEXES.firstData, GANTT_ROW_INDEXES.timeScale等

const SHEET_NAMES = {
  IN_RDB: "4.登録予定_入力データ",
  OUT_RDB: "4.登録済み_出力データ",
  CONFLICT_RDB: "4.登録失敗_重複データ",
  ERROR_RDB: "4.登録失敗_エラーデータ",
  GANTT_TEMPLATE: "1~2.GCテンプレ",
};

function buildShiftDataMergerMenu(ui) {
  return ui.createMenu("シフトデータ統合").addItem("シフトデータを統合", "main");
}

function main() {
  // 名前付き範囲の設定確認
  validateAllNamedRanges();

  // 名前付き範囲からインデックスを初期化
  initializeColumnIndexes();

  const ganttSsUrl = PropertiesService.getScriptProperties().getProperty("GANTT_SS");
  const InGanttSs = SpreadsheetApp.openByUrl(ganttSsUrl);
  const InRdbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.IN_RDB);
  const OutGanttSs = SpreadsheetApp.openByUrl(ganttSsUrl);
  const OutMergedRdbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.OUT_RDB);
  const OutConflictRdbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFLICT_RDB);
  const OutErrorRdbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ERROR_RDB);

  integrateShiftData(InRdbSheet, InGanttSs, OutMergedRdbSheet, OutGanttSs, OutConflictRdbSheet, OutErrorRdbSheet);
}

function integrateShiftData(
  InRdbSheet,
  InGanttSs,
  OutMergedRdbSheet,
  OutGanttSs,
  OutConflictRdbSheet,
  OutErrorRdbSheet
) {
  SpreadsheetApp.getActive().toast("Ganttデータの取得とグループ化を開始します...", "処理状況");
  const ganttDataGrpedByDept = getAllGanttSeetDataAndGrpBySheetName(InGanttSs);
  SpreadsheetApp.getActive().toast("RDBデータの検証と分離を開始します...", "処理状況");
  const { validRdbData, invalidRdbData } = validateAndSeparateRdbData(getRdbData(InRdbSheet));
  SpreadsheetApp.getActive().toast("RDBデータの部署ごとのグループ化を開始します...", "処理状況");
  const rdbDataGrpedByDept = groupByDept(validRdbData, RDB_COL_INDEXES.dept);

  // 処理対象の部署（Ganttに存在する部署のみ）
  const validDepartments = new Set([...Object.keys(ganttDataGrpedByDept)]);

  let newGanttValues = {};
  let newGanttBgs = {};
  let newRdbData = [];
  let conflictData = [];
  let errorData = [...invalidRdbData]; // 無効なRDBデータを初期エラーデータとして追加
  let failedDepartments = [];

  // 処理開始の通知
  SpreadsheetApp.getActive().toast("シフトデータ統合処理を開始します...", "処理開始");

  // RDBのみに存在する部署をエラーデータとして収集
  Object.entries(rdbDataGrpedByDept).forEach(([deptKey, rdbData]) => {
    if (!validDepartments.has(deptKey)) {
      // Ganttに存在しない部署のRDBデータはエラーとして扱う
      errorData = errorData.concat(
        rdbData.map((row) =>
          // sourceにSHEET_NAMES.IN_RDBを設定し、errorMessageにエラーメッセージを設定
          row.concat(SHEET_NAMES.IN_RDB, `部署名${deptKey}のシートがガントチャートSSに見つかりまりませんでした。`)
        )
      );
    }
  });

  // 有効な部署（Ganttに存在する部署）のみを処理
  const validDeptsArray = Array.from(validDepartments);
  const results = validDeptsArray.map((deptKey, index) => {
    // 処理中の部署を通知
    SpreadsheetApp.getActive().toast(`処理中: ${deptKey} (${index + 1}/${validDeptsArray.length})`, "進捗状況");

    const hasRdbData = rdbDataGrpedByDept.hasOwnProperty(deptKey);

    let rdbDataForProcessing = [];
    let ganttDataForProcessing = ganttDataGrpedByDept[deptKey];

    if (hasRdbData) {
      // 両方にデータがある場合
      rdbDataForProcessing = rdbDataGrpedByDept[deptKey];
    } else {
      // Ganttのみにデータがある場合
      rdbDataForProcessing = []; // 空のRDBデータ
    }

    // processDepartment関数を呼び出し
    return processDepartment(deptKey, rdbDataForProcessing, ganttDataForProcessing);
  });

  results.forEach((result) => {
    if (result && result.success) {
      newGanttValues[result.dept] = {
        ganttHeaderValues: result.ganttHeaderValues,
        ganttShiftValues: result.ganttShiftValues,
        ganttHeaderBgs: result.ganttHeaderBgs,
        ganttShiftBgs: result.ganttShiftBgs,
        firstDataColOffset: result.firstDataColOffset,
        firstDataRowOffset: result.firstDataRowOffset,
      };
      newRdbData = newRdbData.concat(result.rdbData);
      conflictData = conflictData.concat(result.conflictData);
      errorData = errorData.concat(result.errorData || []); // エラーデータを追加
    } else if (result && !result.success) {
      failedDepartments.push(result.dept);
    }
  });

  if (failedDepartments.length > 0) {
    console.warn("以下の局の処理に失敗しました:", failedDepartments.join(", "));
    SpreadsheetApp.getActive().toast(`処理失敗: ${failedDepartments.join(", ")}`, "エラー");
  }

  // データ書き込み開始の通知
  SpreadsheetApp.getActive().toast("データの書き込みを開始します...", "処理状況");

  setDataToSheets(
    OutGanttSs,
    OutMergedRdbSheet,
    OutConflictRdbSheet,
    OutErrorRdbSheet,
    newGanttValues,
    newRdbData,
    conflictData,
    errorData
  );
  // rebuildSheets(
  //   OutGanttSs,
  //   OutMergedRdbSheet,
  //   OutConflictRdbSheet,
  //   OutErrorRdbSheet,
  //   newGanttValues,
  //   newRdbData,
  //   conflictData,
  //   errorData
  // );

  // 処理完了の通知
  SpreadsheetApp.getActive().toast("シフトデータ統合処理が完了しました！", "完了");
}

function validateAndSeparateRdbData(rdbData) {
  const validRdbData = [];
  const invalidRdbData = [];

  if (rdbData.length === 0) {
    return { validRdbData, invalidRdbData };
  }

  // ヘッダー行は常に有効として追加
  validRdbData.push(rdbData[0]);

  // データ行のバリデーション（1行目以降）
  for (let i = 1; i < rdbData.length; i++) {
    const row = rdbData[i];
    const errorMessages = [];

    // 必須フィールドのバリデーション（インデックスは0ベース）
    const memberDateId = row[RDB_COL_INDEXES.memberDateId];
    const startTimeValue = row[RDB_COL_INDEXES.startTime];
    const endTimeValue = row[RDB_COL_INDEXES.endTime];
    const dept = row[RDB_COL_INDEXES.dept];

    // memberDateIdのバリデーション
    if (!memberDateId || memberDateId.toString().trim() === "") {
      errorMessages.push("memberDateIdが空です");
    }

    // startTimeのバリデーション
    let startTime = null;
    try {
      startTime = parseTimeToDate(startTimeValue);
    } catch (error) {
      errorMessages.push("startTimeが無効または空です");
    }

    // endTimeのバリデーション
    let endTime = null;
    try {
      endTime = parseTimeToDate(endTimeValue);
    } catch (error) {
      errorMessages.push("endTimeが無効または空です");
    }

    // startTimeとendTimeの順序チェック
    if (
      startTime instanceof Date &&
      endTime instanceof Date &&
      !isNaN(startTime.getTime()) &&
      !isNaN(endTime.getTime()) &&
      startTime >= endTime
    ) {
      errorMessages.push("startTimeがendTime以降の時刻です");
    }

    // deptのバリデーション
    if (!dept || dept.toString().trim() === "") {
      errorMessages.push("deptが空です");
    }

    // エラーがある場合は無効データとして分類
    if (errorMessages.length > 0) {
      const errorRow = row.concat(SHEET_NAMES.IN_RDB, errorMessages.join("、"));
      invalidRdbData.push(errorRow);
    } else {
      validRdbData.push(row);
    }
  }

  return { validRdbData, invalidRdbData };
}

function processDepartment(deptKey, rdbData, ganttData) {
  try {
    const dept = deptKey; // deptKeyは既に部署名になっている
    const { values: ganttValues, backgrounds: ganttBgs } = ganttData;

    // ガントチャートのヘッダーとシフトデータを分割
    const {
      ganttHeaderValues,
      ganttShiftValues,
      ganttHeaderBgs,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      firstDataColOffset,
      firstDataRowOffset,
    } = splitGanttData(ganttValues, ganttBgs);

    const { validShiftsMap, conflictShiftObjs, errorShifts } = convert2dAryToObjsAndJoin(
      ganttShiftValues,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      rdbData,
      dept
    );

    const {
      ganttValues: deptGanttValues,
      ganttBgs: deptGanttBgs,
      rdbData: deptRdbData,
      conflictData: deptConflictData,
      errorData: transformerErrorData,
    } = convertObjsTo2dAry(validShiftsMap, conflictShiftObjs, timeHeaders, memberDateIdHeaders);

    // エラーデータを統合（元のerrorShiftsとtransformerからのエラーデータ）
    const deptErrorData = [
      // 元のエラーシフトデータ（ガントチャートからのエラー）
      ...errorShifts.map((shiftObj) => 
        getColumnOrder(ERROR_COL_INDEXES).map((key) => {
          // startTimeとendTimeはh:mm形式の文字列に変換
          if (key === 'startTime' || key === 'endTime') {
            return formatTimeToHHMM(shiftObj[key]);
          }
          return shiftObj[key];
        })
      ),
      // convertObjsTo2dAryで検出されたエラーデータ（memberDateIdが見つからない）
      ...transformerErrorData
    ];

    // ヘッダー情報も含めて返す（新しいシート作成用）
    return {
      success: true,
      dept,
      ganttHeaderValues, // ヘッダー情報を追加
      ganttShiftValues: deptGanttValues, // シフトデータ
      ganttHeaderBgs, // ヘッダー背景色を追加
      ganttShiftBgs: deptGanttBgs, // シフトデータ背景色
      rdbData: deptRdbData,
      conflictData: deptConflictData,
      errorData: deptErrorData,
      firstDataColOffset, // firstDataの列オフセット
      firstDataRowOffset, // firstDataの行オフセット
    };
  } catch (error) {
    console.error(`Error processing department ${deptKey}:`, error);
    console.error("Stack trace:", error.stack);
    return {
      success: false,
      dept: deptKey,
      error: error.toString(),
    };
  }
}

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
    // データベース、コンフリクト、エラーシートのクリアと更新
    try {
      OutMergedRdbSheet.getDataRange().clearContent();
      rdbData.unshift(getColumnOrder(RDB_COL_INDEXES));
      OutMergedRdbSheet.getRange(1, 1, rdbData.length, rdbData[0].length).setValues(rdbData);
      console.log(`${SHEET_NAMES.OUT_RDB}シートの更新が完了しました`);
      SpreadsheetApp.getActive().toast(`${SHEET_NAMES.OUT_RDB}シートの更新が完了しました`, "更新完了");
    } catch (error) {
      failedSheets.push(`${SHEET_NAMES.OUT_RDB}シート`);
      throw error;
    }

    try {
      OutConflictRdbSheet.getDataRange().clearContent();
      conflictData.unshift(getColumnOrder(CONFLICT_COL_INDEXES));
      OutConflictRdbSheet.getRange(1, 1, conflictData.length, conflictData[0].length).setValues(conflictData);
      console.log(`${SHEET_NAMES.CONFLICT_RDB}シートの更新が完了しました`);
      SpreadsheetApp.getActive().toast(`${SHEET_NAMES.CONFLICT_RDB}シートの更新が完了しました`, "更新完了");
    } catch (error) {
      failedSheets.push(`${SHEET_NAMES.CONFLICT_RDB}シート`);
      throw error;
    }

    try {
      OutErrorRdbSheet.getDataRange().clearContent();
      // エラーデータの書き込み（Ganttに存在しない部署のRDBデータ）
      if (errorData.length > 0) {
        errorData.unshift(getColumnOrder(ERROR_COL_INDEXES));
        OutErrorRdbSheet.getRange(1, 1, errorData.length, errorData[0].length).setValues(errorData);
      }
      console.log(`${SHEET_NAMES.ERROR_RDB}シートの更新が完了しました`);
      SpreadsheetApp.getActive().toast(`${SHEET_NAMES.ERROR_RDB}シートの更新が完了しました`, "更新完了");
    } catch (error) {
      failedSheets.push(`${SHEET_NAMES.ERROR_RDB}シート`);
      throw error;
    }
  } catch (error) {
    showRestorePrompt(failedSheets, "現在のスプレッドシート", startTime, error);
    throw new Error(`データ更新処理を停止しました: ${error.message}`);
  }

  // ガントチャートの各シートを処理
  const ganttSsName = OutGanttSs.getName();

  for (const [sheetName, sheetData] of Object.entries(ganttData)) {
    try {
      const { ganttShiftValues: shiftValues, ganttShiftBgs: shiftBgs, firstDataRowOffset, firstDataColOffset } = sheetData;

      // 空のガントデータの場合はスキップ
      if (!shiftValues || shiftValues.length === 0 || (shiftValues.length === 1 && shiftValues[0].length === 0)) {
        continue;
      }

      // 既存のシートを取得
      const targetSheet = OutGanttSs.getSheetByName(sheetName);
      if (!targetSheet) {
        console.warn(`シート「${sheetName}」が見つかりません。スキップします。`);
        continue;
      }

      const startRow = firstDataRowOffset + 1; // 1-indexedに変換
      const startCol = firstDataColOffset + 1; // 1-indexedに変換

      // firstData列から右側かつfirstData行から下側の範囲を取得
      const shiftRange = targetSheet.getRange(startRow, startCol, shiftValues.length, shiftValues[0].length);

      Logger.log("シフトデータ範囲: " + shiftRange.getA1Notation());

      // firstData以降の全範囲の既存結合を解除とクリア（ヘッダー部分は保持）
      try {
        // firstData以降の全範囲の結合を解除
        const mergedRanges = shiftRange.getMergedRanges();
        mergedRanges.forEach((range) => range.breakApart());
        shiftRange.clear();
        // firstData以降の全範囲をクリア（ヘッダー部分は保持）
      } catch (e) {
        throw new Error(`対象範囲の結合解除・クリア処理でエラーが発生しました: ${e.message}`);
      }
      try {
        // firstDataの位置からシフトデータを設定
        shiftRange.setValues(shiftValues);
        shiftRange.setBackgrounds(shiftBgs);
      } catch (error) {
        throw new Error(`シフトデータ設定中にエラーが発生しました: ${error.message}`);
      }

      // 結合処理を安全に実行（シフトデータ範囲のみ）
      try {
        mergeSameValuesHorizontally(targetSheet, shiftRange);
        // mergeSameValuesVertically(targetSheet, shiftRange);
      } catch (e) {
        throw new Error(`セル結合処理でエラーが発生しました: ${e.message}`);
      }

      console.log(`ガントチャート「${ganttSsName}」のシート「${sheetName}」の更新が完了しました`);
      SpreadsheetApp.getActive().toast(`ガントチャート「${ganttSsName}」のシート「${sheetName}」の更新が完了しました`, "更新完了");
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

// Dateオブジェクトをh:mm形式の文字列に変換する
function formatTimeToHHMM(dateValue) {
  if (!dateValue || !(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
    return "";
  }
  
  const hours = dateValue.getHours().toString().padStart(2, '0');
  const minutes = dateValue.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

// h:mm形式の時刻文字列またはDateオブジェクトを正しいDateオブジェクトに変換する
function parseTimeToDate(timeValue) {
  if (timeValue instanceof Date) {
    // 既にDateオブジェクトの場合はそのまま返す
    return timeValue;
  }
  
  if (typeof timeValue === 'string') {
    // h:mmまたはhh:mm形式の文字列の場合
    const timeMatch = timeValue.match(/^(\d{1,2}):(\d{2})$/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1], 10);
      const minutes = parseInt(timeMatch[2], 10);
      
      // 1970年1月1日00:00分のDateオブジェクトを生成
      const date = new Date(1970, 0, 1, hours, minutes, 0, 0);
      return date;
    }
  }
  
  // その他の場合は通常のDateコンストラクタを試す
  const date = new Date(timeValue);
  if (isNaN(date.getTime())) {
    throw new Error(`無効な時刻形式です: ${timeValue}`);
  }
  return date;
}
