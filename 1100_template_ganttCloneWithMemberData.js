/**
 * メンバー情報を使用してガントチャートテンプレートを複製します
 *
 * このモジュールは0000_common_utils.jsに依存しているため、
 * 以下の関数を利用します:
 * - getMemberDataAndHeaders
 * - getGanttHeaders
 * - findCommonHeaders
 * - generateMemberIds
 * - generateMemberDateId
 */

function buildGanttMenu(ui) {
  return ui.createMenu("ガントチャート作成")
    .addItem("局ごとのガントチャートシートを作成", "promptUserForGanttChartInfo");
}
/**
 * メンバーデータを部署ごとにグループ化し、オブジェクト形式に変換
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Object} 部署をキーとしたメンバーデータのマップ
 */
function groupMemberDataByDept(memberData) {
  const headers = memberData[0];
  const deptIndex = headers.indexOf(COL_HEADER_NAMES.DEPT);
  const groupedData = {};

  // ヘッダーからオブジェクトのプロパティ名を設定
  for (let i = 1; i < memberData.length; i++) {
    const dept = memberData[i][deptIndex];
    if (!dept) continue;

    if (!groupedData[dept]) {
      groupedData[dept] = {
        headers: headers.slice(),  // ヘッダー行を保存
        members: new Map(),        // メンバーをMapで保持して順序を維持
      };
    }

    // メンバー行をオブジェクトに変換
    const memberObj = {};
    headers.forEach((header, j) => {
      memberObj[header] = memberData[i][j];
    });

    // 部署グループのMapにメンバーオブジェクトを追加（キーにインデックスを使用して順序保持）
    groupedData[dept].members.set(i, memberObj);
  }

  return groupedData;
}

/**
 * ガントチャートテンプレートを複製して部署ごとのシートを作成
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - スプレッドシート
 * @param {SpreadsheetApp.Sheet} templateSheet - テンプレートシート
 * @param {string} dept - 部署名
 * @returns {SpreadsheetApp.Sheet} 作成されたシート
 */
function createDeptSheet(spreadsheet, templateSheet, dept) {
  const existingsheet = spreadsheet.getSheetByName(dept);
  if (existingsheet) SpreadsheetApp.deleteSheet(existingsheet);

  const newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(dept);
  return newSheet;
}

/**
 * ガントチャート用のデータを準備する
 * @param {Object} deptData - 部署のメンバーデータ（オブジェクト形式）
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @param {Array} commonHeaders - 共通ヘッダー
 * @param {number} daysPerMember - 一人あたりの日数
 * @param {boolean} insertBlankLine - メンバー間に空白行を挿入するか
 * @returns {Array} ガントチャート用の2次元配列
 */
function prepareGanttData(deptData, ganttHeaders, commonHeaders, daysPerMember, insertBlankLine) {
  const headerIndices = prepareHeaderIndices(deptData.headers, ganttHeaders);
  const memberDateIdIndex = headerIndices.gantt[COL_HEADER_NAMES.MEMBER_DATE_ID];

  // Mapの値（メンバーオブジェクト）は挿入順で取得されるため、ソート不要
  const members = Array.from(deptData.members.values());
  
  // 各メンバーごとのベース行を準備
  const memberBaseRows = members.map(memberObj => {
    const memberId = memberObj[COL_HEADER_NAMES.MEMBER_ID];
    
    // 基本行を作成
    const baseRow = new Array(ganttHeaders.length).fill("");
    
    // day1のmemberDateIdを設定（複製後に修正する）
    baseRow[memberDateIdIndex] = generateMemberDateId(memberId, "day1");
    
    // 共通関数を使用してメンバー情報をコピー
    copyMemberDataToGanttRow(
      commonHeaders,
      headerIndices,
      baseRow,
      memberObj,
      [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.DATE]
    );
    
    return baseRow;
  });
  
  // 全メンバー分のベース行を結合し、一度にduplicateMemberDataRowsで複製する
  // これにより、メンバー間の空白行も適切に挿入される
  const allMemberRows = duplicateMemberDataRows(memberBaseRows, daysPerMember, insertBlankLine);
  
  // 複製された各行のmemberDateIdを修正
  let currentMemberIndex = 0;
  let dayCounter = 1;
  
  for (let i = 0; i < allMemberRows.length; i++) {
    const row = allMemberRows[i];
    
    // 空白行はスキップ
    if (row.every(cell => cell === "")) {
      currentMemberIndex++;
      dayCounter = 1;
      continue;
    }
    
    // 現在のメンバーIDを取得
    const memberId = members[currentMemberIndex][COL_HEADER_NAMES.MEMBER_ID];
    
    // memberDateIdを更新
    row[memberDateIdIndex] = generateMemberDateId(memberId, `day${dayCounter}`);
    
    // 日数カウンターを更新
    dayCounter++;
    if (dayCounter > daysPerMember) {
      dayCounter = 1;
      currentMemberIndex++;
    }
  }
  
  return allMemberRows;
}

/**
 * メインの処理を実行する関数
 */
function promptUserForGanttChartInfo() {
  // ガントチャートテンプレートの範囲選択
  const ganttHeaderRange = promptRangeSelection(
    "ガントチャートの見出しは、現在選択されている範囲で問題ないですか。\n  問題なければ「OK」を押下。\n  選びなおす場合は「キャンセル」を押下し、再実行。"
  );
  if (!ganttHeaderRange) return; // キャンセルされた場合

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
    context: { ganttHeaderRange: ganttHeaderRange.getA1Notation() },
  });
}

/**
 * ダイアログのフォーム送信時に呼び出されるコールバック関数
 * @param {Object} formData - フォームから送信されたデータ
 * @param {Object} context - コンテキスト情報
 */
function handleGanttDialogSubmit(formData, context) {
  const targetUrl = formData.targetUrl;
  const daysPerMember = parseInt(formData.daysPerMember);
  if (isNaN(daysPerMember) || daysPerMember <= 0) {
    SpreadsheetApp.getUi().alert("有効な日数を入力してください");
    return;
  }
  const insertBlankLine = !!formData.insertBlankLine; // チェックボックスの値をブール値に変換
  createGanttChartsWithMemberData(targetUrl, daysPerMember, insertBlankLine, context.ganttHeaderRange);
  // スクリプトプロパティにURLとヘッダー範囲を保存
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("GANTT_SS", targetUrl);
  scriptProperties.setProperty("HEADER_RANGE_A1", context.ganttHeaderRange);
}

/**
 * 部署ごとのガントチャートシートを作成
 * @param {string} targetUrl - 対象スプレッドシートのURL
 * @param {SpreadsheetApp.Sheet} ganttTemplateSheet - ガントチャートテンプレートシート
 * @param {string} dept - 部署名
 * @param {Object} deptData - 部署のメンバーデータ（オブジェクト形式）
 * @param {Array} ganttHeaders - ガントチャートのヘッダー
 * @param {Array} commonHeaders - 共通ヘッダー
 * @param {number} daysPerMember - 一人あたりの日数
 * @param {boolean} insertBlankLine - メンバー間に空白行を挿入するか
 * @param {string} ganttHeaderRangeA1 - ガントチャートのヘッダー範囲
 */
function createDeptGanttSheet(
  targetUrl,
  ganttTemplateSheet,
  dept,
  deptData,
  ganttHeaders,
  commonHeaders,
  daysPerMember,
  insertBlankLine,
  ganttHeaderRangeA1
) {
  const targetSs = SpreadsheetApp.openByUrl(targetUrl);
  const newSheet = createDeptSheet(targetSs, ganttTemplateSheet, dept);
  
  // memberDateId生成とデータ準備
  const preparedData = prepareGanttData(deptData, ganttHeaders, commonHeaders, daysPerMember, insertBlankLine);

  // データのセット
  const ganttRange = newSheet.getRange(ganttHeaderRangeA1);
  const targetRange = newSheet.getRange(
    ganttRange.getRow() + 1, // ヘッダー行の次から
    ganttRange.getColumn(),
    preparedData.length,
    ganttHeaders.length
  );
  targetRange.setValues(preparedData);
}

/**
 * メンバーデータを使用してガントチャートを作成
 * @param {string} targetUrl - 対象スプレッドシートのURL
 * @param {number} daysPerMember - 一人あたりの日数
 * @param {boolean} insertBlankLine - メンバー間に空白行を挿入するか
 * @param {string} ganttHeaderRange - ガントチャートのヘッダー範囲
 */
function createGanttChartsWithMemberData(targetUrl, daysPerMember, insertBlankLine, ganttHeaderRangeA1) {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ganttTemplateSheet = activeSpreadsheet.getSheetByName(GANTT_TEMPLATE_SHEET_NAME);

    // 2. メンバー情報の取得
    const { data: memberData, headers: memberHeaders } = getMemberDataAndHeaders(
      activeSpreadsheet,
      REQUIRED_MEMBER_DATA_HEADERS.DATA_SHEET.INITIALIZE
    );

    // 4. メンバーID生成
    const memberDataWithIds = generateMemberIds(memberData);

    // 5. 部署ごとにグループ化してオブジェクト形式に変換
    const groupedMemberData = groupMemberDataByDept(memberDataWithIds);

    // 6. ガントチャートヘッダー取得
    const { headers: ganttHeaders } = getGanttHeaders(
      ganttTemplateSheet,
      ganttHeaderRangeA1,
      REQUIRED_MEMBER_DATA_HEADERS.GANTT_SHEETS.INITIALIZE
    );

    // 8. 共通ヘッダーの特定と通知
    const ui = SpreadsheetApp.getUi();
    const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);
    ui.alert(`以下のフィールドが転記されます: ${commonHeaders.join(", ")}`);

    // 9. 部署ごとにシートを作成して処理
    Object.keys(groupedMemberData).forEach((dept) => {
      createDeptGanttSheet(
        targetUrl,
        ganttTemplateSheet,
        dept,
        groupedMemberData[dept],
        ganttHeaders,
        commonHeaders,
        daysPerMember,
        insertBlankLine,
        ganttHeaderRangeA1
      );
    });

    ui.alert("処理が完了しました");
  } catch (error) {
    ui.alert(`エラー: ${error.message}`);
  }
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

  // 最後に追加された空白行を削除（今後ガントチャートの下にも何か記入するようであればコメントアウトを解除）
  // if (insertBlankLine && resultArray.length > 0) {
  //   resultArray.pop();
  // }

  return resultArray;
}
