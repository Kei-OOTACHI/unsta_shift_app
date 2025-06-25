function buildFmMenu(ui) {
  return ui.createMenu("1.シフト表テンプレ作成支援")
    .addItem("時間軸を設定", "setTimescale")
    .addItem("行セットを複製", "duplicateRows");
}

// --- 時間軸設定機能 ---
function setTimescale() {
  validateNamedRange(RANGE_NAMES.TIME_SCALE);
  const timescaleRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(RANGE_NAMES.TIME_SCALE);

  const fieldConfigs = [
    { id: "startTime", label: "開始時刻", type: "time", required: true },
    { id: "endTime", label: "終了時刻", type: "time", required: true },
    { id: "interval", label: "時間間隔", type: "number", required: true },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: "シフト時間設定",
    message: "シフトの開始時刻、終了時刻、時間間隔を入力",
    onSubmitFuncName: "processTimescaleInput",
    onCancelFuncName: "handleDialogCancel",
    context: { timescaleA1: timescaleRange.getA1Notation() },
  });
}

// 時間軸設定ダイアログの送信処理 (グローバル関数)
function processTimescaleInput(formData, context) {
  try {
    const timescaleRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(context.timescaleA1);
    const startTime = formData.startTime;
    const endTime = formData.endTime;
    const interval = formData.interval;
    const timescale = buildTimescaleArray(startTime, endTime, interval);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(timescaleRange.getRow(), timescaleRange.getColumn(), 1, timescale.length);
    range.setValues([timescale]);
  } catch (error) {
    console.error("時間軸設定エラー:", error);
    console.error("エラー詳細:", error.stack);
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + error.message);
  }
}

// --- 行複製機能 ---
function duplicateRows() {
  const orgRange = promptRangeSelection(
    "複製する行セットは、現在選択されている行で問題ないですか。\n  問題なければ「OK」を押下。\n  選びなおす場合は「キャンセル」を押下し、再実行。"
  );
  if (!orgRange) return;

  const fieldConfigs = [{ id: "times", label: "複製する行数", type: "number", required: true }];

  showCustomDialog({
    fields: fieldConfigs,
    title: "行複製設定",
    message: "行セットを何回複製するか入力",
    onSubmitFuncName: "processDuplicateRowsInput",
    context: { orgRangeA1: orgRange.getA1Notation() },
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
    console.error("行複製エラー:", error);
    console.error("エラー詳細:", error.stack);
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + error.message);
  }
}

// --- 共通処理 ---

// ダイアログキャンセル時の共通処理 (グローバル関数)
function handleDialogCancel(context) {
  console.log("ダイアログがキャンセルされました。 Context:", context);
  // 必要に応じて追加の処理を記述
}

// 時間軸配列生成
function buildTimescaleArray(startTime, endTime, interval) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 開始時刻と終了時刻をDateオブジェクトに変換
  const createTimeObject = (timeStr) => {
    const [hours, minutes] = timeStr.split(":").map(Number);
    const date = new Date(Date.UTC(1970, 0, 1, hours, minutes, 0, 0)); // UTCでDateオブジェクトを作成
    return date;
  };

  const start = createTimeObject(startTime);
  const end = createTimeObject(endTime);

  // 時間間隔を分に変換
  const intervalMinutes = parseInt(interval);

  // 時刻を格納する配列を作成
  const timeValues = [];
  let currentTime = start;

  // 時刻を配列に追加
  while (currentTime < end) {
    const timeString = Utilities.formatDate(currentTime, "UTC", "HH:mm"); // UTCを指定してフォーマット
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
  Logger.log(selectedRange.getA1Notation() + "の範囲を複製します。");
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
