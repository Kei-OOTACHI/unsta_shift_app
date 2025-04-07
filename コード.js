const GC_COL_ORDER = ["a", "memberDateId", "firstData"]; // ガントチャートの見出し行数
const GC_ROW_ORDER = ["hour", "minute", "timeHeaders", "firstData"]; // ガントチャートの見出し行数
const DB_COL_ROLE = { job: 1, memberDateId: 2, startTime: 3, endTime: 4, cellColor: 5, sorce: 6 }
const DB_COL_ORDER = ["job", "memberDateId", "startTime", "endTime", "background", "sorce"];
const SS_URLS = { GANTT_CHART: "", DATABASE: "", LOG: "" }
const SHEET_NAMES = { DB: "データベース" }
// データベースの列順

/** シフトデータを統合 */
function integrateShiftData() {
    const ganttSs = SpreadsheetApp.openByUrl(SS_URLS.GANTT_CHART);
    const dbSs = SpreadsheetApp.openByUrl(SS_URLS.DATABASE);
    const loggingSs = SpreadsheetApp.openByUrl(SS_URLS.LOG);
    const ganttDataAndBgGrpedByDept = getAllGanttDataAndGrpBySheetName(ganttSs);
    const dbDataGrpedByDept = groupeByMemIdInitial(getDbData(dbSs,SHEET_NAMES.DB), DB_COL_ROLE.memberDateId - 1);

    //以下を局ごとにループで処理
    const mergetShiftsGrpedByDept=processObjectEntries(ganttDataAndBgGrpedByDept,dbDataGrpedByDept);
    const timeHeaders = ganttDataGrpedByDept[GC_ROW_ORDER.indexOf("timeHeaders")].slice(GC_COL_ORDER.indexOf("firstData"));
    const timeIndexMap = new Map(timeHeaders.map((time, index) => [new Date(time).toISOString().slice(11, 16), index]));
    const shiftsFromGantt = extractGanttShifts(ganttDataGrpedByDept, ganttBgGrpedByDept, timeHeaders);
    const shiftsFromDb = extractDbShifts(dbDataGrpedByDept);

    let allShifts = [...shiftsFromDb, ...shiftsFromGantt];
    let { validShifts, doubleBookings } = detectDoubleBookings(allShifts);

    updateSheets(ganttSheet, dbSheet, conflictSheet, validShifts, doubleBookings, timeHeaders, timeIndexMap);
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
function getAllGanttDataAndGrpBySheetName(ganttSs) {
    const sheets = ganttSs.getSheets();

    return sheets.reduce((acc, sheet) => {
        const sheetName = sheet.getName();
        const [ganttData, ganttBg] = getGanttDataAndBg(sheet);

        // シート名をキーにしてデータを格納
        acc[sheetName] = { data: ganttData, bg: ganttBg };
        return acc;
    }, {});
}

// ① 結合セルを分割し、各セルに元の値・背景色を反映
function getGanttDataAndBg(sourceSheet) {
    const sourceRange = sourceSheet.getDataRange();
    const values = sourceRange.getValues();
    const backgrounds = sourceRange.getBackgrounds();
    const mergedRanges = sourceRange.getMergedRanges();

    // 結合範囲ごとに、最初のセルの値と背景色を全セルに反映
    mergedRanges.forEach(mergedRange => {
        const startRow = mergedRange.getRow() - 1;       // 0-indexed
        const startColumn = mergedRange.getColumn() - 1;   // 0-indexed
        const rowCount = mergedRange.getNumRows();
        const columnCount = mergedRange.getNumColumns();

        const mergedValue = values[startRow][startColumn];
        const mergedBackground = backgrounds[startRow][startColumn];

        // 指定範囲の全セルに値と背景色を適用
        for (let i = startRow; i < startRow + rowCount; i++) {
            values[i].fill(mergedValue, startColumn, startColumn + columnCount);
            backgrounds[i].fill(mergedBackground, startColumn, startColumn + columnCount);
        }
    });

    return [values, backgrounds];
}

/** 2次元配列を指定した列の先頭一文字で分類
 *
 * @param {Array} data - 2次元配列
 * @param {number} colIndex - 分類に使用する列のインデックス（0始まり）
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
function getDbData(dbSs,sheetName) {
    const dbSheet = dbSs.getSheetByName(sheetName);
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn());

    // 空の行を除外
    const filterEmptyRows = (data) => data.filter(row => row.some(cell => cell !== ""));

    return filterEmptyRows(dbData);
}

/** シートを更新 */
function updateSheets(ganttSheet, dbSheet, conflictSheet, validShifts, doubleBookings, timeHeaders, timeIndexMap) {
    let ganttData = [[...Array(GC_COL_ORDER.indexOf("firstData")).fill(""), ...timeHeaders]];
    let dbData = [Object.keys(DB_COL_ROLE).sort((a, b) => { DB_COL_ROLE[a] - DB_COL_ROLE[b] })];
    let conflictData = [["社員番号", "業務", "開始時間", "終了時間", "出典"]];

    let groupedShifts = validShifts.reduce((acc, shift) => {
        acc[shift.memberDateId] = acc[shift.memberDateId] || Array(timeHeaders.length).fill("");
        let startIndex = timeIndexMap.get(shift.startTime.toISOString().slice(11, 16));
        let endIndex = timeIndexMap.get(shift.endTime.toISOString().slice(11, 16));
        for (let i = startIndex; i < endIndex; i++) {
            acc[shift.memberDateId][i] = shift.job;
        }
        dbData.push(dbData[0].map(key => shift[key]));
        return acc;
    }, {});

    Object.entries(groupedShifts).forEach(([memId, rowData]) => {
        ganttData.push([...Array(GC_COL_ORDER.indexOf("firstData") - 1).fill(""), memId, ...rowData]);
    });

    doubleBookings.forEach(shift => {
        conflictData.push([shift.memberDateId, shift.job, shift.startTime, shift.endTime, shift.source]);
    });

    ganttSheet.clear();
    dbSheet.clear();
    conflictSheet.clear();

    ganttSheet.getRange(1, 1, ganttData.length, ganttData[0].length).setValues(ganttData);
    dbSheet.getRange(1, 1, dbData.length, dbData[0].length).setValues(dbData);
    conflictSheet.getRange(1, 1, conflictData.length, conflictData[0].length).setValues(conflictData);
}
