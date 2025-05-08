function buildFmMenu(ui) {
  ui.createMenu("FMシート作成ツール")
    .addItem("時間軸を設定", "setTimescale")
    .addItem("行セットを複製", "duplicateRows")
    .addToUi();
}

// --- 時間軸設定機能 ---
function setTimescale() {
  const startCell = promptRangeSelection(
    "タイムスケールを挿入開始するセルは、現在選択されているセルで問題ないですか。\n  問題なければ「OK」を押下。\n  選びなおす場合は「キャンセル」を押下し、再実行。"
  );
  if (!startCell) return;

  const fieldConfigs = [
    { id: "startTime", label: "開始時刻", type: "time", required: true },
    { id: "endTime", label: "終了時刻", type: "time", required: true },
    { id: "interval", label: "時間間隔", type: "number", required: true },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: 'シフト時間設定',
    message: 'シフトの開始時刻、終了時刻、時間間隔を入力',
    onSubmitFuncName: 'processTimescaleInput',
    onCancelFuncName: 'handleDialogCancel',
    context: { startCellA1: startCell.getA1Notation() }
  });
}

// 時間軸設定ダイアログの送信処理 (グローバル関数)
function processTimescaleInput(formData, context) {
  try {
    const startCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(context.startCellA1);
    const startTime = formData.startTime;
    const endTime = formData.endTime;
    const interval = formData.interval;
    const timescale = buildTimescaleArray(startTime, endTime, interval);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(startCell.getRow(), startCell.getColumn(), 1, timescale.length);
    range.setValues([timescale]);
  } catch (error) {
    console.error('時間軸設定エラー:', error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

// --- 行複製機能 ---
function duplicateRows() {
  const orgRange = promptRangeSelection("複製する行セットは、現在選択されている行で問題ないですか。\n  問題なければ「OK」を押下。\n  選びなおす場合は「キャンセル」を押下し、再実行。");
  if (!orgRange) return;

  const fieldConfigs = [
    { id: "times", label: "複製する行数", type: "number", required: true },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: '行複製設定',
    message: '行セットを何回複製するか入力',
    onSubmitFuncName: 'processDuplicateRowsInput',
    context: { orgRangeA1: orgRange.getA1Notation() }
  });
}

// 行複製ダイアログの送信処理 (グローバル関数)
function processDuplicateRowsInput(formData, context) {
  try {
    // getRange(A1Notation)で取得したオブジェクトではなく、正しくシート上のRangeオブジェクトを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const orgRangeA1 = context.orgRangeA1;
    Logger.log(orgRangeA1);
    const orgRange = sheet.getRange(orgRangeA1);
    
    const times = formData.times;
    duplicateSelectedRowsWithFormatting(times, orgRange);
  } catch (error) {
    console.error('行複製エラー:', error);
    console.error('エラー詳細:', error.stack);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}

// --- 共通処理 ---

// ダイアログキャンセル時の共通処理 (グローバル関数)
function handleDialogCancel(context) {
  console.log('ダイアログがキャンセルされました。 Context:', context);
  // 必要に応じて追加の処理を記述
}

// 範囲選択プロンプト
function promptRangeSelection(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const selectedRange = sheet.getActiveRange();
  
  // 列全体が選択されているかチェック
  if (!selectedRange.isStartRowBounded()) {
    ui.alert(
      "エラー",
      "列全体（A列やAA:AZ列など）が選択されています。\n行を選択してから再度実行してください。",
      ui.ButtonSet.OK
    );
    return null;
  }
  
  const res = ui.alert("範囲の選択", message, ui.ButtonSet.OK_CANCEL);
  if (res == ui.Button.OK) {
    Logger.log(selectedRange.getA1Notation() + " was selected");
    return selectedRange;
  } else {
    Logger.log("canceled");
    return null;
  }
}

// 時間軸配列生成
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

// 書式付き行複製
function duplicateSelectedRowsWithFormatting(times, selectedRange) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = sheet.getLastColumn();
  Logger.log(selectedRange.getA1Notation()+"の範囲を複製します。");
  const orgFirstRow = selectedRange.getRow();
  const orgLastRow = selectedRange.getLastRow();
  const orgRows = sheet.getRange(orgFirstRow, 1, orgLastRow - orgFirstRow + 1, lastCol);

  const numRows = orgRows.getNumRows();
  const numColumns = orgRows.getNumColumns();

  // 行の高さを取得
  const rowHeights = [];
  for (let i = 0; i < numRows; i++) {
    rowHeights.push(sheet.getRowHeight(orgRows.getRow() + i));
  }

  for (let i = 0; i < times; i++) {
    const startRow = orgRows.getLastRow() + 1 + i * numRows;
    sheet.insertRowsAfter(orgRows.getLastRow() + i * numRows, numRows);
    const targetRange = sheet.getRange(startRow, orgRows.getColumn(), numRows, numColumns);
    
    // 値、書式、結合セルを一度にコピー
    orgRows.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL);

    // 行の高さを設定
    for (let j = 0; j < numRows; j++) {
      sheet.setRowHeight(startRow + j, rowHeights[j]);
    }
  }
}

