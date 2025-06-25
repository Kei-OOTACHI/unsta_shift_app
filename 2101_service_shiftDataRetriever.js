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
    const ganttSsName = ganttSs.getName();
  
    return sheets.reduce((acc, sheet) => {
      const sheetName = sheet.getName();
      
      // 処理中のシート名をトーストメニューで通知
      SpreadsheetApp.getActive().toast(`ガントチャート「${ganttSsName}」のシート「${sheetName}」のデータを取得中...`, "進捗状況");

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
  
  /** 2次元配列を指定した列の値で分類
   *
   * @param {Array} data - 2次元配列（1行目はヘッダー）
   * @param {number} colIndex - 分類に使用する列のインデックス（0始まり）各セルの値は部署名
   * @return {Object} 部署名をキーとしたオブジェクト。各キーに該当する行の配列が格納される。
   */
  function groupByDept(data, colIndex) {
    // ヘッダー行をスキップして処理
    return data.slice(1).reduce((acc, row) => {
      if (row.length > colIndex) {
        const key = String(row[colIndex]);
        // 既にキーが存在していればその配列に追加、存在しなければ新たな配列を作成
        acc[key] = acc[key] ? [...acc[key], row] : [row];
      }
      return acc;
    }, {});
  }

// SSの指定したシートのデータを２次元配列として抽出
function getRdbData(rdbSheet) {
    const rdbData = rdbSheet.getRange(1, 1, rdbSheet.getLastRow(), rdbSheet.getLastColumn()).getValues();
  
    // 空の行を除外
    const filterEmptyRows = (data) => data.filter((row) => row.some((cell) => cell !== ""));
  
    return filterEmptyRows(rdbData);
  }
    