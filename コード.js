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

/** 指定したスプレッドシート内のすべてのシートから、ガントデータと背景データを取得、シート名ごとに分類
 *
 * @param {Spreadsheet} ganttSs - ガントチャートのデータが含まれる Google スプレッドシートオブジェクト
 * @return {Object} シート名をキーとし、各シートのガントデータと背景データを含むオブジェクト
 *                  {
 *                      "シート名1": { data: [[...]], bg: [[...]] },
 *                      "シート名2": { data: [[...]], bg: [[...]] },
 *                      ...
 *                  }
 */
function getAllGanttSeetDataAndGrpBySheetName(ganttSs) {
  const sheets = ganttSs.getSheets();

  return sheets.reduce((acc, sheet) => {
    const sheetName = sheet.getName();
    const { values, backgrounds } = getGanttSeetData(sheet);

    // シート名をキーにしてデータを格納
    acc[sheetName] = { values, backgrounds };
    return acc;
  }, {});
}

// ① 結合セルを分割し、各セルに元の値・背景色を反映
function getGanttSeetData(sourceSheet) {
  const sourceRange = sourceSheet.getDataRange();
  const values = sourceRange.getValues();
  const backgrounds = sourceRange.getBackgrounds();
  const mergedRanges = sourceRange.getMergedRanges();

  // 結合範囲ごとに、最初のセルの値と背景色を全セルに反映
  mergedRanges.forEach((mergedRange) => {
    const startRow = mergedRange.getRow() - 1; // 0-indexed
    const startColumn = mergedRange.getColumn() - 1; // 0-indexed
    const rowCount = mergedRange.getNumRows();
    const columnCount = mergedRange.getNumColumns();

    const mergedValue = values[startRow][startColumn];
    const mergedBackground = backgrounds[startRow][startColumn];

    // 指定範囲の全セルに値と背景色を適用（縦横両方向に対応）
    for (let i = startRow; i < startRow + rowCount; i++) {
      for (let j = startColumn; j < startColumn + columnCount; j++) {
        values[i][j] = mergedValue;
        backgrounds[i][j] = mergedBackground;
      }
    }
  });

  return { values, backgrounds };
}

/** 2次元配列を指定した列の先頭一文字で分類
 *
 * @param {Array} data - 2次元配列
 * @param {number} colIndex - 分類に使用する列のインデックス（0始まり）各セルの値は先頭一文字に局固有のアルファベット、その後ろに数字が続く
 * @return {Object} 先頭一文字をキーとしたオブジェクト。各キーに該当する行の配列が格納される。
 */
function groupeByMemIdInitial(data, colIndex) {
  return data.reduce((acc, row) => {
    if (row.length > colIndex) {
      const key = String(row[colIndex]).charAt(0);
      // 既にキーが存在していればその配列に追加、存在しなければ新たな配列を作成
      acc[key] = acc[key] ? [...acc[key], row] : [row];
    }
    return acc;
  }, {});
}

// SSの指定したシートのデータを２次元配列として抽出
function getRdbData(rdbSheet) {
  const rdbData = rdbSheet.getRange(2, 1, rdbSheet.getLastRow() - 1, rdbSheet.getLastColumn());

  // 空の行を除外
  const filterEmptyRows = (data) => data.filter((row) => row.some((cell) => cell !== ""));

  return filterEmptyRows(rdbData);
}

function splitGanttData(ganttValues, ganttBgs, ganttColManager, ganttRowManager) {
  const firstDataCol = ganttColManager.getColumnIndex("firstData");
  const firstDataRow = ganttRowManager.getColumnIndex("firstData");

  // シフトデータ部分
  const ganttShiftValues = ganttValues.slice(firstDataRow).map((row) => row.slice(firstDataCol));
  const ganttShiftBgs = ganttBgs.slice(firstDataRow).map((row) => row.slice(firstDataCol));

  // ヘッダー部分（「の形）
  const ganttHeaderValues = [];
  const ganttHeaderBgs = [];

  // 上部ヘッダー行（全列を含む）
  for (let i = 0; i < firstDataRow; i++) {
    ganttHeaderValues.push([...ganttValues[i]]);
    ganttHeaderBgs.push([...ganttBgs[i]]);
  }

  // 左側ヘッダー列（firstDataRow行目以降、firstDataCol列までのデータ）
  for (let i = firstDataRow; i < ganttValues.length; i++) {
    ganttHeaderValues.push(ganttValues[i].slice(0, firstDataCol));
    ganttHeaderBgs.push(ganttBgs[i].slice(0, firstDataCol));
  }

  // ganttColManagerのorderを修正（firstDataより前の要素を削除）
  const originalColOrder = [...ganttColManager.config.order];
  const firstDataColIndex = ganttColManager.config.order.indexOf("firstData");
  const adjustedColOrder = ganttColManager.config.order.slice(firstDataColIndex);

  // ganttRowManagerのorderを修正（firstDataより前の要素を削除）
  const originalRowOrder = [...ganttRowManager.config.order];
  const firstDataRowIndex = ganttRowManager.config.order.indexOf("firstData");
  const adjustedRowOrder = ganttRowManager.config.order.slice(firstDataRowIndex);

  // 修正したorderで設定を更新
  ganttColManager.config.order = adjustedColOrder;
  ganttRowManager.config.order = adjustedRowOrder;

  // インデックスを再初期化
  ganttColManager.initializeIndexes();
  ganttRowManager.initializeIndexes();

  // timescale,memberDateIdのリストを作成
  const timeRow = ganttRowManager.getColumnIndex("timeScale");
  const memberDateIdCol = ganttColManager.getColumnIndex("memberDateId");

  const timeHeaders = ganttValues[timeRow].slice(firstDataCol);
  const memberDateIdHeaders = ganttValues.slice(firstDataRow).map((row) => row[memberDateIdCol]);

  return {
    ganttHeaderValues,
    ganttShiftValues,
    ganttHeaderBgs,
    ganttShiftBgs,
    timeHeaders,
    memberDateIdHeaders,
    originalColOrder,
    originalRowOrder,
  };
}

function convertObjsTo2dAry(
  validShiftsMap,
  conflictShiftObjs,
  timeHeaders,
  rdbColManager,
  ganttColManager,
  conflictColManager
) {
  // rdbDataとconflictDataのヘッダー行を追加
  const rdbData = [rdbColManager.config.order.slice()];
  const conflictData = [conflictColManager.config.order.slice()];
  
  // Mapからrdbデータを直接生成（中間変換なし）
  const processedShiftIds = new Set();
  
  // ganttData用のmemberMap（既に作成済み）
  const ganttValueMap = new Map();
  // 背景色用のmemberBgMap（新規追加）
  const ganttBgMap = new Map();
  
  // 各メンバーのシフト情報を処理
  for (const [memberId, timeMap] of validShiftsMap.entries()) {
    // 各時間スロットごとに処理
    for (const [timeKey, shiftInfo] of timeMap.entries()) {
      // まだ処理していないシフトIDの場合のみrdbDataに追加
      if (!processedShiftIds.has(shiftInfo.shiftId)) {
        const rdbRow = rdbColManager.config.order.map(key => shiftInfo[key]);
        rdbData.push(rdbRow);
        processedShiftIds.add(shiftInfo.shiftId);
        
        // ganttData用のデータも準備
        if (!ganttValueMap.has(shiftInfo.memberDateId)) {
          ganttValueMap.set(shiftInfo.memberDateId, Array(timeHeaders.length).fill(""));
          ganttBgMap.set(shiftInfo.memberDateId, Array(timeHeaders.length).fill("#FFFFFF")); // 背景色の初期値は白
        }
        
        const timeRow = ganttValueMap.get(shiftInfo.memberDateId);
        const bgRow = ganttBgMap.get(shiftInfo.memberDateId);
        const startIndex = findTimeIndex(timeHeaders, shiftInfo.startTime);
        const endIndex = findTimeIndex(timeHeaders, shiftInfo.endTime);
        
        if (startIndex !== -1 && endIndex !== -1) {
          for (let i = startIndex; i < endIndex; i++) {
            timeRow[i] = shiftInfo.job;
            // 背景色も設定
            bgRow[i] = shiftInfo.background || "#FFFFFF";
          }
        }
      }
    }
  }

  // マップから直接ganttDataに変換
  const ganttValues = Array.from(ganttValueMap.values());
  // 背景色の2次元配列も生成
  const ganttBgs = Array.from(ganttBgMap.values());

  // コンフリクトデータを処理
  conflictShiftObjs.forEach((shiftObj) => {
    const conflictRow = conflictColManager.config.order.map((key) => shiftObj[key]);
    conflictData.push(conflictRow);
  });

  return {
    ganttValues,
    ganttBgs,
    rdbData,
    conflictData,
  };
}

// 時間ヘッダー配列から指定時間に最も近いインデックスを見つける
function findTimeIndex(timeHeaders, time) {
  const timeStr = time.toISOString().slice(11, 16);
  for (let i = 0; i < timeHeaders.length; i++) {
    const headerTime = new Date(timeHeaders[i]).toISOString().slice(11, 16);
    if (headerTime === timeStr) {
      return i;
    }
  }
  return -1;
}

function mergeGanttData(
  ganttHeaderValues,
  ganttShiftValues,
  ganttHeaderBgs,
  ganttShiftBgs,
  ganttColManager,
  ganttRowManager,
  originalColOrder,
  originalRowOrder
) {
  const firstDataCol = ganttColManager.getColumnIndex("firstData");
  const firstDataRow = ganttRowManager.getColumnIndex("firstData");

  // 結合後のデータを格納する配列
  const mergedValues = [];
  const mergedBgs = [];

  // 上部ヘッダー行を追加（そのまま）
  for (let i = 0; i < firstDataRow; i++) {
    mergedValues.push(ganttHeaderValues[i]);
    mergedBgs.push(ganttHeaderBgs[i]);
  }

  // 左側ヘッダー列とシフトデータを結合して追加
  for (let i = 0; i < ganttShiftValues.length; i++) {
    const headerRow = ganttHeaderValues[i + firstDataRow];
    const bgHeaderRow = ganttHeaderBgs[i + firstDataRow];

    mergedValues.push([...headerRow, ...ganttShiftValues[i]]);
    mergedBgs.push([...bgHeaderRow, ...ganttShiftBgs[i]]);
  }

  // managerのorderを元に戻す
  ganttColManager.config.order = originalColOrder;
  ganttRowManager.config.order = originalRowOrder;

  // インデックスを再初期化
  ganttColManager.initializeIndexes();
  ganttRowManager.initializeIndexes();

  return {
    values: mergedValues,
    backgrounds: mergedBgs,
  };
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
