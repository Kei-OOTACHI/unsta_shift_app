/**
 * メンバー情報を使用してガントチャートテンプレートを複製します
 *
 * このモジュールは0000_common_utils.jsに依存しているため、
 * 以下の関数を利用します:
 * - getMemberData
 * - validateHeaders
 * - findCommonHeaders
 * - generateMemberIds
 * - generateMemberDateId
 */

/**
 * メンバーデータを部署ごとにグループ化
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Object} 部署をキーとしたメンバーデータのマップ
 */
function groupMemberDataByDept(memberData) {
  const headers = memberData[0];
  const deptIndex = headers.indexOf("dept");
  const groupedData = {};

  for (let i = 1; i < memberData.length; i++) {
    const dept = memberData[i][deptIndex];
    if (!dept) continue;

    if (!groupedData[dept]) {
      groupedData[dept] = [headers.slice()]; // ヘッダー行を含める
    }

    groupedData[dept].push(memberData[i].slice());
  }

  return groupedData;
}

/**
 * ガントチャートのヘッダー行を取得
 * @param {SpreadsheetApp.Range} ganttRange - ガントチャートの範囲
 * @returns {Array} ヘッダー行の配列
 */
function getGanttHeaders(ganttRange) {
  return ganttRange.getValues()[0];
}

/**
 * ガントチャートテンプレートを複製して部署ごとのシートを作成
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - スプレッドシート
 * @param {SpreadsheetApp.Sheet} templateSheet - テンプレートシート
 * @param {string} dept - 部署名
 * @returns {SpreadsheetApp.Sheet} 作成されたシート
 */
function createDeptSheet(spreadsheet, templateSheet, dept) {
  const newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(dept);
  return newSheet;
}

/**
 * ガントチャート用のデータを準備する
 * @param {Array} memberData - メンバーデータの2次元配列
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @param {Array} commonHeaders - 共通ヘッダー
 * @param {number} daysPerMember - 一人あたりの日数
 * @param {boolean} insertBlankLine - メンバー間に空白行を挿入するか
 * @returns {Array} ガントチャート用の2次元配列
 */
function prepareGanttData(memberData, ganttHeaders, commonHeaders, daysPerMember, insertBlankLine) {
  // インデックスを事前計算
  const headerIndices = {
    member: {}, // メンバーヘッダーのインデックス
    gantt: {}, // ガントヘッダーのインデックス
  };

  const memberHeaders = memberData[0];

  // メンバーヘッダーのインデックスをキャッシュ
  memberHeaders.forEach((header, index) => {
    headerIndices.member[header] = index;
  });

  // ガントヘッダーのインデックスをキャッシュ
  ganttHeaders.forEach((header, index) => {
    headerIndices.gantt[header] = index;
  });

  const memberIdIndex = headerIndices.member["memberId"];
  const memberDateIdIndex = headerIndices.gantt["memberDateId"];
  const dateIndex = headerIndices.gantt["date"];

  const resultData = [];

  // メンバーごとに処理（ループ内での検索を減らす）
  for (let i = 1; i < memberData.length; i++) {
    const memberRow = memberData[i];
    const memberId = memberRow[memberIdIndex];

    // 日付ごとに行を生成
    const memberRows = [];
    for (let day = 1; day <= daysPerMember; day++) {
      const newRow = new Array(ganttHeaders.length).fill("");
      const date = `day${day}`;

      // memberDateIdの設定
      newRow[memberDateIdIndex] = generateMemberDateId(memberId, date);

      // 日付の設定
      newRow[dateIndex] = date;

      // 共通ヘッダーの値をコピー（キャッシュされたインデックスを使用）
      commonHeaders.forEach((header) => {
        if (header !== "memberDateId" && header !== "date") {
          newRow[headerIndices.gantt[header]] = memberRow[headerIndices.member[header]];
        }
      });

      memberRows.push(newRow);
    }

    // 結果に追加
    resultData.push(...memberRows);

    // メンバー間に空白行を挿入（オプション）
    if (insertBlankLine && i < memberData.length - 1) {
      resultData.push(new Array(ganttHeaders.length).fill(""));
    }
  }

  return resultData;
}

/**
 * メインの処理を実行する関数
 */
function createGanttChartsWithMemberData() {
  // ガントチャートテンプレートの範囲選択
  const ganttRange = promptRangeSelection("ガントチャートテンプレートの範囲を選択してください");
  if (!ganttRange) return; // キャンセルされた場合

  // ダイアログでパラメータを取得
  const fieldConfigs = [
    {
      id: "targetUrl",
      label: "対象スプレッドシートのURL",
      type: "text",
      required: true,
    },
    {
      id: "daysPerMember",
      label: "一人あたりの日数",
      type: "number",
      required: true,
    },
    {
      id: "insertBlankLine",
      label: "メンバー間に空白行を挿入する",
      type: "checkbox",
      required: false,
    },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: "ガントチャート作成パラメータ",
    message: "ガントチャートテンプレート複製のパラメータを入力してください",
    onSubmitFuncName: "handleGanttDialogSubmit",
    context: { ganttRange: ganttRange.getA1Notation() },
  });
}

/**
 * ダイアログのフォーム送信時に呼び出されるコールバック関数
 * @param {Object} formData - フォームから送信されたデータ
 * @param {Object} context - コンテキスト情報
 */
function handleGanttDialogSubmit(formData, context) {
  const ui = SpreadsheetApp.getUi();
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    const targetUrl = formData.targetUrl;
    const daysPerMember = parseInt(formData.daysPerMember);
    const insertBlankLine = !!formData.insertBlankLine; // チェックボックスの値をブール値に変換

    if (isNaN(daysPerMember) || daysPerMember <= 0) {
      ui.alert("有効な日数を入力してください");
      return;
    }

    // ガントチャート範囲を取得
    const ganttRange = activeSheet.getRange(context.ganttRange);

    // 2. メンバー情報の取得
    const spreadsheet = SpreadsheetApp.openByUrl(targetUrl);
    const memberData = getMemberData(spreadsheet);
    const memberHeaders = memberData[0];

    // 3. ヘッダー検証
    validateHeaders(memberHeaders, REQUIRED_MEMBER_HEADERS);

    // 4. メンバーID生成
    const memberDataWithIds = generateMemberIds(memberData);

    // 5. 部署ごとにグループ化
    const groupedMemberData = groupMemberDataByDept(memberDataWithIds);

    // 6. ガントチャートヘッダー取得
    const ganttHeaders = getGanttHeaders(ganttRange);

    // 7. ガントチャートヘッダー検証
    validateHeaders(ganttHeaders, REQUIRED_GANTT_HEADERS);

    // 8. 共通ヘッダーの特定と通知
    const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);
    ui.alert(`以下のフィールドが転記されます: ${commonHeaders.join(", ")}`);

    // 9. 部署ごとにシートを作成して処理
    Object.keys(groupedMemberData).forEach((dept) => {
      const deptMemberData = groupedMemberData[dept];
      const newSheet = createDeptSheet(spreadsheet, activeSheet, dept);

      // 10 & 11. memberDateId生成とデータ準備
      const preparedData = prepareGanttData(
        deptMemberData,
        ganttHeaders,
        commonHeaders,
        daysPerMember,
        insertBlankLine
      );

      // 12. データのセット
      const ganttRange = newSheet.getRange(context.ganttRange);
      const targetRange = newSheet.getRange(
        ganttRange.getRow() + 1, // ヘッダー行の次から
        ganttRange.getColumn(),
        preparedData.length,
        ganttHeaders.length
      );
      targetRange.setValues(preparedData);
    });

    ui.alert("処理が完了しました");
  } catch (error) {
    ui.alert(`エラー: ${error.message}`);
  }
}

/**
 * ヘッダーの順序に基づいてメンバーデータを並べ替える
 * @param {Array} dataArray - 元のデータ配列
 * @param {Array} headerOrder - ヘッダーの順序
 * @returns {Array} 並べ替えられたデータ配列
 */
function sortMemberDataByHeaders(dataArray, headerOrder) {
  // 1行目を見出しとして取得
  const headers = dataArray[0];

  // 新しい配列の初期化
  const sortedArray = [];

  // 新しい見出し行を作成
  const newHeaders = headerOrder.map((header) => {
    if (headers.includes(header)) {
      return header;
    } else {
      return ""; // 空白列を挿入
    }
  });

  // 新しい見出し行を追加
  sortedArray.push(newHeaders);

  // データ行を並べ替え
  for (let i = 1; i < dataArray.length; i++) {
    const row = dataArray[i];
    const newRow = headerOrder.map((header) => {
      const index = headers.indexOf(header);
      if (index !== -1) {
        return row[index];
      } else {
        return ""; // 空白列を挿入
      }
    });
    sortedArray.push(newRow);
  }

  return sortedArray;
}

/**
 * メンバーデータの行を複製し、空白行を挿入
 * @param {Array} dataArray - 元のデータ配列
 * @param {number} duplicateCount - 複製回数
 * @param {boolean} insertBlankLine - 空白行を挿入するかどうか
 * @returns {Array} 複製されたデータ配列
 */
function duplicateMemberDataRows(dataArray, duplicateCount, insertBlankLine) {
  const resultArray = [];

  dataArray.forEach((row) => {
    // 各行を指定された回数だけ複製
    for (let i = 0; i < duplicateCount; i++) {
      resultArray.push([...row]); // スプレッド演算子で配列をコピー
    }

    // 複製された行のまとまりの間に空白行を挿入
    if (insertBlankLine) {
      resultArray.push(new Array(row.length).fill(""));
    }
  });

  // 最後に追加された空白行を削除（必要であれば）
  if (insertBlankLine && resultArray.length > 0) {
    resultArray.pop();
  }

  return resultArray;
}
