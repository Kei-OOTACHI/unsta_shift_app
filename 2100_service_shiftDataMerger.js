// 列・行インデックスの直接参照が可能です:
// RDB_COL_INDEXES.dept, GANTT_COL_INDEXES.firstData, GANTT_ROW_INDEXES.timeScale等

const SHEET_NAMES = { IN_RDB: "Input", OUT_RDB: "シフトDB", CONFLICT_RDB: "重複データ",ERROR_RDB: "エラーデータ" };

function buildShiftDataMergerMenu(ui) {
  return ui.createMenu("シフトデータ統合")
    .addItem("シフトデータを統合", "main");
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

  integrateShiftData(
    InRdbSheet,
    InGanttSs,
    OutMergedRdbSheet,
    OutGanttSs,
    OutConflictRdbSheet,
    OutErrorRdbSheet
  );
}

function integrateShiftData(
  InRdbSheet,
  InGanttSs,
  OutMergedRdbSheet,
  OutGanttSs,
  OutConflictRdbSheet,
  OutErrorRdbSheet
) {
  const ganttDataGrpedByDept = getAllGanttSeetDataAndGrpBySheetName(InGanttSs);
  const rdbDataGrpedByDept = groupByDept(getRdbData(InRdbSheet), RDB_COL_INDEXES.dept);

  // 処理対象の部署（Ganttに存在する部署のみ）
  const validDepartments = new Set([
    ...Object.keys(ganttDataGrpedByDept)
  ]);

  let newGanttValues = {};
  let newGanttBgs = {};
  let newRdbData = [];
  let conflictData = [];
  let errorData = [];
  let failedDepartments = [];

  // RDBのみに存在する部署をエラーデータとして収集
  Object.entries(rdbDataGrpedByDept).forEach(([deptKey, rdbData]) => {
    if (!validDepartments.has(deptKey)) {
      // Ganttに存在しない部署のRDBデータはエラーとして扱う
      errorData = errorData.concat(rdbData);
    }
  });

  // 有効な部署（Ganttに存在する部署）のみを処理
  const results = Array.from(validDepartments).map((deptKey) => {
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
    return processDepartment(
      deptKey,
      rdbDataForProcessing,
      ganttDataForProcessing
    );
  });

  results.forEach((result) => {
    if (result && result.success) {
      newGanttValues[result.dept] = result.ganttValues;
      newGanttBgs[result.dept] = result.ganttBgs;
      newRdbData = newRdbData.concat(result.rdbData);
      conflictData = conflictData.concat(result.conflictData);
    } else if (result && !result.success) {
      failedDepartments.push(result.dept);
    }
  });

  if (failedDepartments.length > 0) {
    console.warn("以下の局の処理に失敗しました:", failedDepartments.join(", "));
  }

  setDataToSheets(
    OutGanttSs,
    OutMergedRdbSheet,
    OutConflictRdbSheet,
    OutErrorRdbSheet,
    newGanttValues,
    newGanttBgs,
    newRdbData,
    conflictData,
    errorData
  );
}

function processDepartment(
  deptKey,
  rdbData,
  ganttData
) {
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

    const { validShiftsMap, conflictShiftObjs } = convert2dAryToObjsAndJoin(
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
    } = convertObjsTo2dAry(
      validShiftsMap,
      conflictShiftObjs,
      timeHeaders
    );

    // ガントチャートのヘッダーとシフトデータを結合
    const { values: mergedValues, backgrounds: mergedBgs } = mergeGanttData(
      ganttHeaderValues,
      deptGanttValues,
      ganttHeaderBgs,
      deptGanttBgs,
      firstDataColOffset,
      firstDataRowOffset
    );

    return {
      success: true,
      dept,
      ganttValues: mergedValues,
      ganttBgs: mergedBgs,
      rdbData: deptRdbData,
      conflictData: deptConflictData,
    };
  } catch (error) {
    console.error(`Error processing department ${deptKey}:`, error);
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
  ganttValues,
  ganttBgs,
  rdbData,
  conflictData,
  errorData
) {
  try {
    // 現在のデータをバックアップ
    const backup = {
      rdbData: OutMergedRdbSheet.getDataRange().getValues(),
      conflictData: OutConflictRdbSheet.getDataRange().getValues(),
      errorData: OutErrorRdbSheet.getDataRange().getValues(),
      ganttSheets: {}
    };

    // ガントチャートの各シートのバックアップ
    Object.keys(ganttValues).forEach(sheetName => {
      const sheet = OutGanttSs.getSheetByName(sheetName);
      if (sheet) {
        backup.ganttSheets[sheetName] = {
          values: sheet.getDataRange().getValues(),
          backgrounds: sheet.getDataRange().getBackgrounds()
        };
      }
    });

    // データベース、コンフリクト、エラーシートのクリアと更新
    OutMergedRdbSheet.clear();
    OutConflictRdbSheet.clear();
    OutErrorRdbSheet.clear();
    
    rdbData.unshift(getColumnOrder(RDB_COL_INDEXES));
    OutMergedRdbSheet.getRange(1, 1, rdbData.length, rdbData[0].length).setValues(rdbData);
    
    conflictData.unshift(getColumnOrder(CONFLICT_COL_INDEXES));
    OutConflictRdbSheet.getRange(1, 1, conflictData.length, conflictData[0].length).setValues(conflictData);
    
    // エラーデータの書き込み（Ganttに存在しない部署のRDBデータ）
    if (errorData.length > 0) {
      errorData.unshift(getColumnOrder(RDB_COL_INDEXES));
      OutErrorRdbSheet.getRange(1, 1, errorData.length, errorData[0].length).setValues(errorData);
    }

    // ガントチャートの各シートを処理
    Object.entries(ganttValues).forEach(([sheetName, sheetValues]) => {
      // 空のガントデータの場合はスキップ
      if (!sheetValues || sheetValues.length === 0 || 
          (sheetValues.length === 1 && sheetValues[0].length === 0)) {
        return;
      }
      
      const targetSheet = OutGanttSs.getSheetByName(sheetName)?.clear() || OutGanttSs.insertSheet(sheetName);
      const sheetBgs = ganttBgs[sheetName];

      // データを設定
      const range = targetSheet.getRange(1, 1, sheetValues.length, sheetValues[0].length);
      range.setValues(sheetValues);
      range.setBackgrounds(sheetBgs);
    });

  } catch (error) {
    console.error('データの更新中にエラーが発生しました:', error);
    
    // エラーが発生した場合、バックアップから復元
    try {
      // データベース、コンフリクト、エラーシートの復元
      OutMergedRdbSheet.clear();
      OutConflictRdbSheet.clear();
      OutErrorRdbSheet.clear();
      OutMergedRdbSheet.getRange(1, 1, backup.rdbData.length, backup.rdbData[0].length).setValues(backup.rdbData);
      OutConflictRdbSheet.getRange(1, 1, backup.conflictData.length, backup.conflictData[0].length).setValues(backup.conflictData);
      OutErrorRdbSheet.getRange(1, 1, backup.errorData.length, backup.errorData[0].length).setValues(backup.errorData);

      // ガントチャートの各シートの復元
      Object.entries(backup.ganttSheets).forEach(([sheetName, sheetData]) => {
        const targetSheet = OutGanttSs.getSheetByName(sheetName)?.clear() || OutGanttSs.insertSheet(sheetName);
        const range = targetSheet.getRange(1, 1, sheetData.values.length, sheetData.values[0].length);
        range.setValues(sheetData.values);
        range.setBackgrounds(sheetData.backgrounds);
      });

      throw new Error('データの更新に失敗しました。元の状態に戻しました。詳細: ' + error.message);
    } catch (restoreError) {
      throw new Error('データの更新に失敗し、元の状態への復元にも失敗しました。詳細: ' + restoreError.message);
    }
  }
}
