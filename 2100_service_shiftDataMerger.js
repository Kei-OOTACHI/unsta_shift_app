const SS_URLS = { GANTT_CHART: "", DATABASE: "", LOG: "" };
const SHEET_NAMES = { RDB: "データベース", CONFLICT_RDB: "コンフリクトデータ" };
const DEPARTMENTS = {
  A: "会場整備局",
  B: "参加対応局",
  C: "開発局",
  D: "企画局",
  E: "広報制作局",
  F: "渉外局",
  G: "総務局",
  H: "財務局",
};

function main() {
  const { rdbColManager, ganttColManager, ganttRowManager, conflictColManager } = convertToNewFormat();

  const InRdbSheet = SpreadsheetApp.openByUrl(SS_URLS.DATABASE).getSheetByName(SHEET_NAMES.RDB);
  const InGanttSs = SpreadsheetApp.openByUrl(SS_URLS.GANTT_CHART);
  const OutMergedRdbSheet = SpreadsheetApp.openByUrl(SS_URLS.DATABASE).getSheetByName(SHEET_NAMES.RDB);
  const OutGanttSs = SpreadsheetApp.openByUrl(SS_URLS.GANTT_CHART);
  const OutConflictRdbSheet = SpreadsheetApp.openByUrl(SS_URLS.DATABASE).getSheetByName(SHEET_NAMES.CONFLICT_RDB);

  integrateShiftData(
    InRdbSheet,
    InGanttSs,
    OutMergedRdbSheet,
    OutGanttSs,
    OutConflictRdbSheet,
    rdbColManager,
    ganttColManager,
    ganttRowManager,
    conflictColManager
  );
}

function integrateShiftData(
  InRdbSheet,
  InGanttSs,
  OutMergedRdbSheet,
  OutGanttSs,
  OutConflictRdbSheet,
  rdbColManager,
  ganttColManager,
  ganttRowManager,
  conflictColManager
) {
  const ganttDataGrpedByDept = getAllGanttSeetDataAndGrpBySheetName(InGanttSs);
  const rdbDataGrpedByDept = groupeByMemIdInitial(getRdbData(InRdbSheet), rdbColManager.getColumnIndex("memberDateId"));

  let newGanttValues = {};
  let newGanttBgs = {};
  let newRdbData = [];
  let conflictData = [];
  let failedDepartments = [];

  const results = Object.entries(rdbDataGrpedByDept).map(([deptKey, rdbData]) => {
    return processDepartment(
      deptKey,
      rdbData,
      ganttDataGrpedByDept,
      rdbColManager,
      ganttColManager,
      ganttRowManager,
      conflictColManager
    );
  });

  results.forEach((result) => {
    if (result.success) {
      newGanttValues[result.dept] = result.ganttValues;
      newGanttBgs[result.dept] = result.ganttBgs;
      newRdbData = newRdbData.concat(result.rdbData);
      conflictData = conflictData.concat(result.conflictData);
    } else {
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
    newGanttValues,
    newGanttBgs,
    newRdbData,
    conflictData,
    rdbColManager,
    ganttColManager,
    ganttRowManager,
    conflictColManager
  );
}

function processDepartment(
  deptKey,
  rdbData,
  ganttDataGrpedByDept,
  rdbColManager,
  ganttColManager,
  ganttRowManager,
  conflictColManager
) {
  try {
    const dept = DEPARTMENTS[deptKey];
    const { values: ganttValues, backgrounds: ganttBgs } = ganttDataGrpedByDept[dept];

    // ガントチャートのヘッダーとシフトデータを分割
    const {
      ganttHeaderValues,
      ganttShiftValues,
      ganttHeaderBgs,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      originalColOrder,
      originalRowOrder,
    } = splitGanttData(ganttValues, ganttBgs, ganttColManager, ganttRowManager);

    const { validShiftsMap, conflictShiftObjs } = convert2dAryToObjsAndJoin(
      ganttShiftValues,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      rdbData,
      rdbColManager
    );
 
    const {
      ganttValues: deptGanttValues,
      ganttBgs: deptGanttBgs,
      rdbData: deptRdbData,
      conflictData: deptConflictData,
    } = convertObjsTo2dAry(
      validShiftsMap,
      conflictShiftObjs,
      timeHeaders,
      rdbColManager,
      ganttColManager,
      conflictColManager
    );

    // ガントチャートのヘッダーとシフトデータを結合
    const { values: mergedValues, backgrounds: mergedBgs } = mergeGanttData(
      ganttHeaderValues,
      deptGanttValues,
      ganttHeaderBgs,
      deptGanttBgs,
      ganttColManager,
      ganttRowManager,
      originalColOrder,
      originalRowOrder
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
      dept: DEPARTMENTS[deptKey],
      error: error.toString(),
    };
  }
}

function setDataToSheets(
  OutGanttSs,
  OutMergedRdbSheet,
  OutConflictRdbSheet,
  ganttValues,
  ganttBgs,
  rdbData,
  conflictData,
  rdbColManager,
  ganttColManager,
  ganttRowManager,
  conflictColManager
) {
  try {
    // 現在のデータをバックアップ
    const backup = {
      rdbData: OutMergedRdbSheet.getDataRange().getValues(),
      conflictData: OutConflictRdbSheet.getDataRange().getValues(),
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

    // データベースとコンフリクトシートのクリアと更新
    OutMergedRdbSheet.clear();
    OutConflictRdbSheet.clear();
    rdbData.unshift(rdbColManager.config.order);
    OutMergedRdbSheet.getRange(1, 1, rdbData.length, rdbData[0].length).setValues(rdbData);
    conflictData.unshift(conflictColManager.config.order);
    OutConflictRdbSheet.getRange(1, 1, conflictData.length, conflictData[0].length).setValues(conflictData);

    // ガントチャートの各シートを処理
    Object.entries(ganttValues).forEach(([sheetName, sheetValues]) => {
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
      // データベースとコンフリクトシートの復元
      OutMergedRdbSheet.clear();
      OutConflictRdbSheet.clear();
      OutMergedRdbSheet.getRange(1, 1, backup.rdbData.length, backup.rdbData[0].length).setValues(backup.rdbData);
      OutConflictRdbSheet.getRange(1, 1, backup.conflictData.length, backup.conflictData[0].length).setValues(backup.conflictData);

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
