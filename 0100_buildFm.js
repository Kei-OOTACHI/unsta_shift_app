function buildFmMenu(ui) {
  ui.createMenu("FMシート作成ツール")
    .addItem("時間軸を設定", "setTimescale")
    .addItem("行セットを複製", "duplicateRows")
    .addToUi();
}

async function setTimescale() {
  const startCell = promptRangeSelection("タイムスケールを挿入開始するセルを選択。選択したらOKボタンを押下。選びなおす場合はキャンセルボタンを押下し、再実行。");
  if (!startCell) return; // キャンセル時の処理
  
  // フィールド設定
  const fieldConfigs = [
    {
      id: "startTime",
      label: "開始時刻",
      type: "time",
      required: true,
    },
    {
      id: "endTime",
      label: "終了時刻",
      type: "time",
      required: true,
    },
    {
      id: "interval",
      label: "時間間隔",
      type: "number",
      required: true,
    },
  ];
  
  try {
    const result = await showCustomInputDialog(fieldConfigs, "シフトの開始時刻、終了時刻、時間間隔を入力");
    const startTime = result.startTime;
    const endTime = result.endTime;
    const interval = result.interval;
    const timescale = buildTimescaleArray(startTime, endTime, interval);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(startCell.getRow(), startCell.getColumn(), 1, timescale.length);
    range.setValues([timescale]);
  } catch (error) {
    console.error(error);
  }
}

async function duplicateRows() {
  const orgRange = promptRangeSelection("複製する行セットを選択...");
  if (!orgRange) return;

  const fieldConfigs = [{
    id: "times",
    label: "複製する行数",
    type: "number",
    required: true,
  }];

  try {
    const result = await showCustomInputDialog(fieldConfigs, "行セットを何回複製するか入力");
    const times = result.times;
    duplicateSelectedRowsWithFormatting(times, orgRange);
  } catch (error) {
    console.error(error);
  }
}

function promptRangeSelection(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const promptResponse = ui.prompt(message, ui.ButtonSet.OK_CANCEL);
  if (promptResponse.getSelectedButton() == ui.Button.OK) {
    return sheet.getActiveRange();
  }
  return null; // キャンセル時はnullを返す
}

function buildTimescaleArray(startTime, endTime, interval) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 開始時刻と終了時刻をDateオブジェクトに変換
  const start = new Date();
  const end = new Date();

  // 開始時刻を設定
  const startParts = startTime.split(":");
  start.setHours(parseInt(startParts[0]));
  start.setMinutes(parseInt(startParts[1]));

  // 終了時刻を設定
  const endParts = endTime.split(":");
  end.setHours(parseInt(endParts[0]));
  end.setMinutes(parseInt(endParts[1]));

  // 時間間隔を分に変換
  const intervalMinutes = parseInt(interval);

  // 時刻を格納する配列を作成
  const timeValues = [];
  let currentTime = new Date(start);

  // 時刻を配列に追加
  while (currentTime <= end) {
    const timeString = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), "HH:mm");
    timeValues.push([timeString]); // 2次元配列にするために配列で囲む

    // 時間をインクリメント
    currentTime.setMinutes(currentTime.getMinutes() + intervalMinutes);
  }

  return timeValues;
}

function duplicateSelectedRowsWithFormatting(times, selectedRange) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = sheet.getLastColumn();
  const orgFirstRow = selectedRange.getFirstRow();
  const orgLastRow = selectedRange.getLastRow();
  const orgRows = sheet.getRange(orgFirstRow, 1, orgLastRow - orgFirstRow + 1, lastCol);

  const numRows = orgRows.getNumRows();
  const numColumns = orgRows.getNumColumns();
  const values = orgRows.getValues();
  const backgroundColors = orgRows.getBackgrounds();
  const fontColors = orgRows.getFontColors();
  const fontWeights = orgRows.getFontWeights();
  const fontStyles = orgRows.getFontStyles();
  const borders = orgRows.getBorder();

  // 行の高さを取得
  const rowHeights = [];
  for (let i = 0; i < numRows; i++) {
    rowHeights.push(sheet.getRowHeight(orgRows.getRow() + i));
  }

  // 結合されたセルの情報を取得
  const mergedRanges = [];
  const mergedRangesInSelection = orgRows.getMergedRanges();
  mergedRangesInSelection.forEach((mergedRange) => {
    const relativeRow = mergedRange.getRow() - orgRows.getRow();
    const relativeColumn = mergedRange.getColumn() - orgRows.getColumn();
    mergedRanges.push({
      row: relativeRow,
      column: relativeColumn,
      numRows: mergedRange.getNumRows(),
      numColumns: mergedRange.getNumColumns(),
    });
  });

  for (let i = 0; i < times; i++) {
    const startRow = orgRows.getLastRow() + 1 + i * numRows;
    sheet.insertRowsAfter(orgRows.getLastRow() + i * numRows, numRows);
    const targetRange = sheet.getRange(startRow, orgRows.getColumn(), numRows, numColumns);
    targetRange.setValues(values);
    targetRange.setBackgrounds(backgroundColors);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
    targetRange.setFontStyles(fontStyles);
    targetRange.setBorder(
      borders.top,
      borders.left,
      borders.bottom,
      borders.right,
      borders.vertical,
      borders.horizontal
    );

    // 行の高さを設定
    for (let j = 0; j < numRows; j++) {
      sheet.setRowHeight(startRow + j, rowHeights[j]);
    }

    // 結合されたセルを適用
    mergedRanges.forEach((mergedRange) => {
      const mergeStartRow = startRow + mergedRange.row;
      const mergeStartColumn = orgRows.getColumn() + mergedRange.column;
      sheet.getRange(mergeStartRow, mergeStartColumn, mergedRange.numRows, mergedRange.numColumns).merge();
    });
  }
}
