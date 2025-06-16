// 列・行インデックスの直接参照が可能です:
// RDB_COL_INDEXES.dept, GANTT_COL_INDEXES.firstData, GANTT_ROW_INDEXES.timeScale等

const SHEET_NAMES = {
  IN_RDB: "Input",
  OUT_RDB: "シフトDB",
  CONFLICT_RDB: "重複データ",
  ERROR_RDB: "エラーデータ",
  GANTT_TEMPLATE: "GCテンプレ",
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
    const startTime = row[RDB_COL_INDEXES.startTime];
    const endTime = row[RDB_COL_INDEXES.endTime];
    const dept = row[RDB_COL_INDEXES.dept];

    // memberDateIdのバリデーション
    if (!memberDateId || memberDateId.toString().trim() === "") {
      errorMessages.push("memberDateIdが空です");
    }

    // startTimeのバリデーション
    if (!startTime || !(startTime instanceof Date) || isNaN(startTime.getTime())) {
      errorMessages.push("startTimeが無効または空です");
    }

    // endTimeのバリデーション
    if (!endTime || !(endTime instanceof Date) || isNaN(endTime.getTime())) {
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
    } = convertObjsTo2dAry(validShiftsMap, conflictShiftObjs, timeHeaders, memberDateIdHeaders);

    // エラーデータを直接変換
    const deptErrorData = errorShifts.map((shiftObj) => getColumnOrder(ERROR_COL_INDEXES).map((key) => shiftObj[key]));

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
