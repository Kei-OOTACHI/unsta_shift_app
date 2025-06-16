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

/**
 * ガントチャート作成メニューを構築
 * @param {SpreadsheetApp.Ui} ui - SpreadsheetAppのUIオブジェクト
 * @returns {SpreadsheetApp.Menu} 構築されたメニューオブジェクト
 */
function buildGanttMenu(ui) {
  return ui.createMenu("2.シフト表SS作成")
    .addItem("シフト表SSに「1~2.シフト表テンプレ」をコピーして、「2~3.メンバー情報」のデータを局ごとに展開", "promptUserForGanttChartInfo");
}

/**
 * ユーザーにガントチャート作成のパラメータを入力させるダイアログを表示する関数
 */
function promptUserForGanttChartInfo() {
  // ガントチャートテンプレートの範囲選択
  validateNamedRange(RANGE_NAMES.GANTT_HEADER_ROW);
  const ganttHeaderRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(RANGE_NAMES.GANTT_HEADER_ROW);

  // ダイアログでパラメータを取得
  const fieldConfigs = [
    {
      id: "targetUrl",
      label: "空のSSのURLを入力（ここに入力したURLのSSがシフト表になります）",
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
      label: "メンバー間に空白行を挿入するなら☑",
      type: "boolean",
      required: false,
    },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: "シフト表SS作成パラメータ",
    message: "シフト表SS作成のパラメータを入力してください",
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
  try {
    Logger.log("ダイアログフォーム送信処理を開始");
    Logger.log("フォームデータ:", JSON.stringify(formData));
    Logger.log("コンテキスト:", JSON.stringify(context));
    
    const targetUrl = formData.targetUrl;
    const daysPerMember = parseInt(formData.daysPerMember);
    if (isNaN(daysPerMember) || daysPerMember <= 0) {
      const errorMessage = "有効な日数を入力してください";
      Logger.log(errorMessage);
      SpreadsheetApp.getUi().alert(errorMessage);
      return;
    }
    const insertBlankLine = !!formData.insertBlankLine; // チェックボックスの値をブール値に変換
    
    Logger.log(`処理パラメータ: targetUrl=${targetUrl}, daysPerMember=${daysPerMember}, insertBlankLine=${insertBlankLine}`);
    
    createGanttChartsWithMemberData(targetUrl, daysPerMember, insertBlankLine, context.ganttHeaderRange);
    
    // スクリプトプロパティにURLとヘッダー範囲を保存
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("GANTT_SS", targetUrl);
    scriptProperties.setProperty("HEADER_RANGE_A1", context.ganttHeaderRange);
    
    Logger.log("スクリプトプロパティに設定を保存しました");
  } catch (error) {
    // 詳細なエラー情報をログに出力
    Logger.log("=== handleGanttDialogSubmit エラー詳細情報 ===");
    Logger.log("エラーメッセージ:", error.message);
    Logger.log("エラー名:", error.name);
    Logger.log("スタックトレース:", error.stack);
    
    console.error("=== handleGanttDialogSubmit エラー詳細情報 ===");
    console.error("エラーメッセージ:", error.message);
    console.error("エラー名:", error.name);
    console.error("スタックトレース:", error.stack);
    console.error("エラーオブジェクト全体:", error);
    
    // エラーを再スローして、元のスタックトレースを保持
    throw error;
  }
}

/**
 * メンバーデータを使用してガントチャートを作成
 * @param {string} targetUrl - 対象スプレッドシートのURL
 * @param {number} daysPerMember - 一人あたりの日数
 * @param {boolean} insertBlankLine - メンバー間に空白行を挿入するか
 * @param {string} ganttHeaderRange - ガントチャートのヘッダー範囲
 */
function createGanttChartsWithMemberData(targetUrl, daysPerMember, insertBlankLine, ganttHeaderRangeA1) {
  const ui = SpreadsheetApp.getUi();  // ui変数を定義
  try {
    Logger.log("ガントチャート作成処理を開始します");
    Logger.log(`パラメータ: targetUrl=${targetUrl}, daysPerMember=${daysPerMember}, insertBlankLine=${insertBlankLine}, ganttHeaderRangeA1=${ganttHeaderRangeA1}`);
    
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ganttTemplateSheet = activeSpreadsheet.getSheetByName(SHEET_NAMES.GANTT_TEMPLATE);

    // 2. メンバー情報の取得
    Logger.log("メンバー情報を取得中...");
    const { data: memberData, headers: memberHeaders } = getMemberDataAndHeaders(
      activeSpreadsheet,
      REQUIRED_MEMBER_DATA_HEADERS.DATA_SHEET.INITIALIZE
    );

    // 4. メンバーID生成
    Logger.log("メンバーIDを生成中...");
    const memberDataWithIds = generateMemberIds(memberData);

    // 4.1. メンバー情報シートにmemberIdを書き戻し
    Logger.log("メンバー情報シートにmemberIdを書き戻し中...");
    updateMemberDataSheetWithIds(activeSpreadsheet, memberDataWithIds);

    // 5. 部署ごとにグループ化してオブジェクト形式に変換
    Logger.log("部署ごとにメンバーデータをグループ化中...");
    const groupedMemberData = groupMemberDataByDept(memberDataWithIds);

    // 6. ガントチャートヘッダー取得
    Logger.log("ガントチャートヘッダーを取得中...");
    const { headers: ganttHeaders } = getGanttHeaders(
      ganttTemplateSheet,
      ganttHeaderRangeA1,
      REQUIRED_MEMBER_DATA_HEADERS.GANTT_SHEETS.INITIALIZE
    );

    // 8. 共通ヘッダーの特定と通知
    const commonHeaders = findCommonHeaders(memberHeaders, ganttHeaders);
    const message = `以下のフィールドが転記されます: ${commonHeaders.join(", ")}`;
    Logger.log(message);
    ui.alert(message);

    // 9. 部署ごとにシートを作成して処理
    Logger.log("部署ごとのシート作成を開始...");
    Object.keys(groupedMemberData).forEach((dept) => {
      Logger.log(`部署「${dept}」のシートを作成中...`);
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

    const successMessage = "処理が完了しました";
    Logger.log(successMessage);
    ui.alert(successMessage);
  } catch (error) {
    // 詳細なエラー情報をログに出力
    Logger.log("=== エラー詳細情報 ===");
    Logger.log("エラーメッセージ:", error.message);
    Logger.log("エラー名:", error.name);
    Logger.log("スタックトレース:", error.stack);
    Logger.log("エラーオブジェクト全体:", JSON.stringify(error, Object.getOwnPropertyNames(error)));
    
    console.error("=== エラー詳細情報 ===");
    console.error("エラーメッセージ:", error.message);
    console.error("エラー名:", error.name);
    console.error("スタックトレース:", error.stack);
    console.error("エラーオブジェクト全体:", error);
    
    const errorMessage = `エラー: ${error.message}`;
    ui.alert(errorMessage + "\n\n詳細は実行ログを確認してください。");
    
    // エラーを再スローして、元のスタックトレースを保持
    throw error;
  }
}

/**
 * メンバーデータを部署ごとにグループ化し、オブジェクト形式に変換
 * @param {Array} memberData - メンバー情報の2次元配列
 * @returns {Object} 部署をキーとしたメンバーデータのマップ
 */
function groupMemberDataByDept(memberData) {
  const headers = memberData[0];
  const deptIndex = headers.indexOf(COL_HEADER_NAMES.DEPT);
  const memberIdIndex = headers.indexOf(COL_HEADER_NAMES.MEMBER_ID);
  const groupedData = {};

  Logger.log("メンバーデータのヘッダー:", JSON.stringify(headers));
  Logger.log(`部署インデックス: ${deptIndex}, メンバーIDインデックス: ${memberIdIndex}`);

  // ヘッダーからオブジェクトのプロパティ名を設定
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const dept = row[deptIndex];
    if (!dept) continue;

    if (!groupedData[dept]) {
      groupedData[dept] = {
        headers: headers.slice(),
        members: new Map()
      };
      Logger.log(`部署「${dept}」のグループを作成しました`);
    }

    // メンバー行をオブジェクトに変換（一度のループで処理）
    const memberObj = headers.reduce((obj, header, j) => {
      obj[header] = row[j];
      return obj;
    }, {});

    // メンバーIDが存在するかチェック
    if (!memberObj[COL_HEADER_NAMES.MEMBER_ID]) {
      Logger.log(`警告: 行${i}のメンバーIDが空です:`, JSON.stringify(memberObj));
    } else {
      Logger.log(`部署「${dept}」にメンバー「${memberObj[COL_HEADER_NAMES.MEMBER_ID]}」を追加`);
    }

    groupedData[dept].members.set(i, memberObj);
  }

  // 各部署のメンバー数をログ出力
  Object.keys(groupedData).forEach(dept => {
    Logger.log(`部署「${dept}」のメンバー数: ${groupedData[dept].members.size}`);
  });

  // メンバーデータの内容を確認
  Logger.log("=== メンバーデータの内容確認 ===");
  Object.keys(groupedData).forEach(dept => {
    const members = Array.from(groupedData[dept].members.values());
    Logger.log(`部署「${dept}」のメンバー数:`, members.length);
    members.forEach((member, index) => {
      Logger.log(`部署「${dept}」のメンバー${index}:`, JSON.stringify(member));
      if (!member[COL_HEADER_NAMES.MEMBER_ID]) {
        Logger.log(`警告: 部署「${dept}」のメンバー${index}にmemberIdが存在しません`);
      }
    });
  });

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
  if (existingsheet) {
    // シートが1枚しかない場合は削除しない
    if (spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(existingsheet);
    } else {
      // シートが1枚しかない場合は、既存のシートをクリアして再利用
      existingsheet.clear();
      existingsheet.setName("sheet1");
    }
  }

  const newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(dept);
  return newSheet;
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
  Logger.log(`部署「${dept}」のシート作成を開始`);
  Logger.log(`部署データの構造:`, JSON.stringify({
    hasHeaders: !!deptData.headers,
    headersLength: deptData.headers ? deptData.headers.length : 0,
    hasMembers: !!deptData.members,
    membersSize: deptData.members ? deptData.members.size : 0
  }));
  
  // メンバーデータの詳細をログ出力
  if (deptData.members) {
    const membersList = Array.from(deptData.members.values());
    Logger.log(`部署「${dept}」のメンバー一覧:`);
    membersList.forEach((member, index) => {
      Logger.log(`  メンバー${index}: ${JSON.stringify(member)}`);
    });
  }
  
  const targetSs = SpreadsheetApp.openByUrl(targetUrl);
  const newSheet = createDeptSheet(targetSs, ganttTemplateSheet, dept);
  
  // memberDateId生成とデータ準備
  Logger.log(`部署「${dept}」のガントデータを準備中...`);
  const preparedData = prepareGanttData(deptData, ganttHeaders, commonHeaders, daysPerMember, insertBlankLine);
  Logger.log(`部署「${dept}」の準備されたデータ行数: ${preparedData.length}`);

  // データのセット
  const ganttRange = newSheet.getRange(ganttHeaderRangeA1);
  const targetRange = newSheet.getRange(
    ganttRange.getRow() + 1, // ヘッダー行の次から
    ganttRange.getColumn(),
    preparedData.length,
    ganttHeaders.length
  );
  targetRange.setValues(preparedData);
  
  Logger.log(`部署「${dept}」のシート作成が完了しました`);
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
  const memberIdIndex = headerIndices.gantt[COL_HEADER_NAMES.MEMBER_ID];
  const members = Array.from(deptData.members.values());
  
  Logger.log(`ガントヘッダーインデックス: memberDateId=${memberDateIdIndex}, memberId=${memberIdIndex}`);
  
  // メンバーが存在しない場合は空の配列を返す
  if (!members || members.length === 0) {
    return [];
  }
  
  // 各メンバーごとのベース行を準備（Mapを使用して効率的に処理）
  const memberBaseRows = members.map(memberObj => {
    if (!memberObj || !memberObj[COL_HEADER_NAMES.MEMBER_ID]) {
      Logger.log("無効なメンバーデータをスキップ: " + JSON.stringify(memberObj));
      return null;
    }

    const memberId = memberObj[COL_HEADER_NAMES.MEMBER_ID];
    const baseRow = new Array(ganttHeaders.length).fill("");
    baseRow[memberDateIdIndex] = generateMemberDateId(memberId, "day1");
    
    // memberIdを展開
    if (memberIdIndex !== undefined) {
      baseRow[memberIdIndex] = memberId;
      Logger.log(`メンバーID「${memberId}」をガントチャート行に設定`);
    }
    
    copyMemberDataToGanttRow(
      commonHeaders,
      headerIndices,
      baseRow,
      memberObj,
      [COL_HEADER_NAMES.MEMBER_DATE_ID, COL_HEADER_NAMES.DATE]
    );
    
    return baseRow;
  }).filter(row => row !== null); // nullの行を除外
  
  // 有効な行が存在しない場合は空の配列を返す
  if (memberBaseRows.length === 0) {
    return [];
  }
  
  // 全メンバー分のベース行を結合
  const allMemberRows = duplicateMemberDataRows(memberBaseRows, daysPerMember, insertBlankLine);
  
  // 複製された各行のmemberDateIdとmemberIdを修正（効率的な処理）
  let currentMemberIndex = 0;
  let dayCounter = 1;
  
  return allMemberRows.map(row => {
    if (row.every(cell => cell === "")) {
      currentMemberIndex++;
      dayCounter = 1;
      return row;
    }
    
    const currentMember = members[currentMemberIndex];
    if (!currentMember || !currentMember[COL_HEADER_NAMES.MEMBER_ID]) {
      Logger.log(`警告: 無効なメンバーデータ (インデックス: ${currentMemberIndex}):`, JSON.stringify(currentMember));
      return row;
    }
    
    const memberId = currentMember[COL_HEADER_NAMES.MEMBER_ID];
    row[memberDateIdIndex] = generateMemberDateId(memberId, `day${dayCounter}`);
    
    // memberIdを各行に展開
    if (memberIdIndex !== undefined) {
      row[memberIdIndex] = memberId;
    }
    
    dayCounter++;
    if (dayCounter > daysPerMember) {
      dayCounter = 1;
      currentMemberIndex++;
    }
    
    // 現在のメンバーインデックスが有効かチェック
    if (currentMemberIndex >= members.length) {
      Logger.log(`警告: 無効なメンバーインデックス ${currentMemberIndex} (メンバー数: ${members.length})`);
      return row;
    }
    
    return row;
  });
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
  const emptyRow = new Array(dataArray[0].length).fill("");
  
  dataArray.forEach(row => {
    // 各行を指定された回数だけ複製（Array.fromを使用して効率的に処理）
    resultArray.push(...Array.from({ length: duplicateCount }, () => [...row]));
    
    if (insertBlankLine) {
      resultArray.push([...emptyRow]);
    }
  });
  
  return resultArray;
}

/**
 * メンバー情報シートにmemberIdを書き戻す
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - スプレッドシート
 * @param {Array} memberDataWithIds - memberIdが追加されたメンバーデータ
 */
function updateMemberDataSheetWithIds(spreadsheet, memberDataWithIds) {
  try {
    const memberSheet = spreadsheet.getSheetByName(SHEET_NAMES.MEMBER_DATA);
    if (!memberSheet) {
      Logger.log(`警告: シート「${SHEET_NAMES.MEMBER_DATA}」が見つかりません`);
      return;
    }

    // 現在のデータ範囲を取得
    const currentRange = memberSheet.getDataRange();
    const currentData = currentRange.getValues();
    
    // 新しいデータの行数と列数を確認
    const newRowCount = memberDataWithIds.length;
    const newColCount = memberDataWithIds[0].length;
    const currentRowCount = currentData.length;
    const currentColCount = currentData[0].length;
    
    Logger.log(`メンバー情報シート更新: 現在の行数=${currentRowCount}, 列数=${currentColCount}, 新しい行数=${newRowCount}, 列数=${newColCount}`);
    
    // 必要に応じてシートのサイズを調整
    if (newRowCount > currentRowCount) {
      memberSheet.insertRows(currentRowCount + 1, newRowCount - currentRowCount);
    }
    if (newColCount > currentColCount) {
      memberSheet.insertColumns(currentColCount + 1, newColCount - currentColCount);
    }
    
    // 新しいデータを設定
    const targetRange = memberSheet.getRange(1, 1, newRowCount, newColCount);
    targetRange.setValues(memberDataWithIds);
    
    Logger.log("メンバー情報シートにmemberIdを書き戻しました");
  } catch (error) {
    Logger.log("メンバー情報シート更新エラー:", error.message);
    console.error("メンバー情報シート更新エラー:", error);
  }
}
