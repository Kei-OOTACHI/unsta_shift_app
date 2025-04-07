/**
 * オブジェクトの各エントリを順々に関数に引き渡して加工する関数
 *
 * @param {Object} dataObj - 加工対象のオブジェクト（キー: 一意の識別子, 値: 加工対象のデータ）
 * @param {Function} processFunction - 各データに適用する処理関数 (引数: key, value)
 * @return {Object} 加工後のデータを格納したオブジェクト
 */
function processObjectEntries(dataObj, processFunction) {
    return Object.entries(dataObj).reduce((acc, [key, value]) => {
        acc[key] = processFunction(key, value);
        return acc;
    }, {});
}


// ② unmergeGantで分割後、各行ごとに横方向で同じ値・背景色が続くセルをグループとしてまとめる関数
function extractGanttShifts(ganttData, ganttBg,timeHeaders) {
  let shifts = [];
  // 各行を走査
  for (let i = 0; i < ganttData.length; i++) {
    let row = ganttData[i];
    let bgRow = ganttBg[i];
    let j = 1;//memberDateIdの見出し列をスキップするので1スタート
    while (j < row.length) {
      // ガントチャートの棒セルの判定:
      // 「値が入力されている」または「背景色が白(#ffffff)以外」であれば対象とする
      if (row[j] !== "" || (bgRow[j] && bgRow[j] !== "#FFFFFF")) {
        let startCol = j;
        let cellValue = row[j];
        let cellBg = bgRow[j];
        // 横方向に同じ値＆背景色が続く間、グループに追加
        while (j < row.length && row[j] === cellValue && (bgRow[j] && bgRow[j] === cellBg)) {
          j++;
        }
        let endCol = j - 1;
        shifts.push({
          memberDateId: row[GC_COL_ORDER.indexOf("memberDateId")],//一番左にメンバー番号が記載されている
          startTime: new Date(timeHeaders[startCol - 1]),  //memberDateIdの1列分、timeheadersがずれている
          startTime: new Date(timeHeaders[endCol]),  //一つ後の時間が終了時刻なので±0
          job: cellValue,
          background: cellBg,
          sorce: "Gantt"
        });
      } else {
        j++;
      }
    }
  }

  return shifts;
}

/** シフトデータを抽出 */
function extractDbShifts(dbData) {
  let shifts = [];

  dbData.slice(1).forEach(row => {
    shifts.push({ source: "DB", job: row[0], memberDateId: row[1], startTime: new Date(row[2]), endTime: new Date(row[3]), background: row[4] });
  });

  return shifts;
}

/** ダブルブッキング検出 */
function detectDoubleBookings(shifts) {
  let validShifts = [], doubleBookings = [];
  let groupedShifts = shifts.reduce((acc, shift) => {
    acc[shift.memberDateId] = acc[shift.memberDateId] || [];
    acc[shift.memberDateId].push(shift);
    return acc;
  }, {});

  for (let memberDateId in groupedShifts) {
    let sortedShifts = groupedShifts[memberDateId].sort((a, b) => a.startTime - b.startTime);
    let overlaps = [];
    sortedShifts.forEach((shift, i, arr) => {
      if (i > 0 && arr[i - 1].endTime > shift.startTime) {
        if (!overlaps.includes(arr[i - 1])) overlaps.push(arr[i - 1]);
        if (!overlaps.includes(shift)) overlaps.push(shift);
      }
    });
    overlaps = [...new Set(overlaps)];//setオブジェクトにすることで重複削除。バカ便利。
    validShifts.push(...sortedShifts.filter(shift => !overlaps.includes(shift)));
    doubleBookings.push(...overlaps);
  }
  return { validShifts, doubleBookings };
}