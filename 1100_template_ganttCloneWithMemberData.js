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
    // Logger.log("=== handleGanttDialogSubmit エラー詳細情報 ===");
    // Logger.log("エラーメッセージ:", error.message);
    // Logger.log("エラー名:", error.name);
    // Logger.log("スタックトレース:", error.stack);
    
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

    // 4.1. メンバーデータの必須項目検証
    Logger.log("メンバーデータの必須項目を検証中...");
    const validatedMemberData = validateMemberData(memberDataWithIds);

    // 4.2. メンバー情報シートにmemberIdを書き戻し
    Logger.log("メンバー情報シートにmemberIdを書き戻し中...");
    updateMemberDataSheetWithIds(activeSpreadsheet, validatedMemberData);

    // 4.3. エラー行の確認と通知
    const errorRows = validatedMemberData.slice(1).filter((row, index) => {
      const memberIdIndex = validatedMemberData[0].indexOf(COL_HEADER_NAMES.MEMBER_ID);
      const memberId = row[memberIdIndex];
      return memberId && memberId.includes("エラー：");
    });

    if (errorRows.length > 0) {
      // エラーの種類別に集計
      const emailErrors = errorRows.filter(row => {
        const memberIdIndex = validatedMemberData[0].indexOf(COL_HEADER_NAMES.MEMBER_ID);
        const memberId = row[memberIdIndex];
        return memberId && memberId.includes("emailが記入されていない");
      });
      
      const deptErrors = errorRows.filter(row => {
        const memberIdIndex = validatedMemberData[0].indexOf(COL_HEADER_NAMES.MEMBER_ID);
        const memberId = row[memberIdIndex];
        return memberId && memberId.includes("deptが記入されていない");
      });

      let errorMessage = `${errorRows.length}行で必須項目が不足しているため、ガントチャート作成から除外されます。\n`;
      if (emailErrors.length > 0) {
        errorMessage += `- email未記入: ${emailErrors.length}行\n`;
      }
      if (deptErrors.length > 0) {
        errorMessage += `- dept未記入: ${deptErrors.length}行\n`;
      }
      errorMessage += `該当行にエラーメッセージを記入しました。`;
      
      Logger.log(errorMessage);
      // ui.alert("注意", errorMessage, ui.ButtonSet.OK);
    }

    // 5. 部署ごとにグループ化してオブジェクト形式に変換
    Logger.log("部署ごとにメンバーデータをグループ化中...");
    const groupedMemberData = groupMemberDataByDept(validatedMemberData);

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
    
    if (!dept || (typeof dept === 'string' && dept.trim() === '')) {
      continue;
    }

    // メンバー行をオブジェクトに変換（一度のループで処理）
    const memberObj = headers.reduce((obj, header, j) => {
      obj[header] = row[j];
      return obj;
    }, {});

    // 元の行番号を保存（デバッグ用）
    memberObj._originalRowIndex = i;

    // memberIdがエラーメッセージかチェック
    const memberId = memberObj[COL_HEADER_NAMES.MEMBER_ID];
    
    if (memberId && memberId.includes("エラー：")) {
      Logger.log(`行${i + 1}をスキップ: ${memberId}`);
      continue; // エラーメッセージが含まれる行はスキップ
    }

    // メンバーIDが存在するかチェック
    if (!memberId || (typeof memberId === 'string' && memberId.trim() === '')) {
      Logger.log(`警告: 行${i + 1}のメンバーIDが空です:`, JSON.stringify(memberObj));
      continue;
    }

    if (!groupedData[dept]) {
      groupedData[dept] = {
        headers: headers.slice(),
        members: [] // 順序を保持するために配列を使用
      };
      Logger.log(`部署「${dept}」のグループを作成しました`);
    }

    Logger.log(`部署「${dept}」にメンバー「${memberId}」を追加`);

    // 配列にプッシュして順序を保持
    groupedData[dept].members.push(memberObj);
  }

  // 各部署のメンバー数をログ出力
  Object.keys(groupedData).forEach(dept => {
    Logger.log(`部署「${dept}」のメンバー数: ${groupedData[dept].members.length}`);
  });

  // メンバーデータの内容を確認
  Logger.log("=== メンバーデータの内容確認 ===");
  Object.keys(groupedData).forEach(dept => {
    const members = groupedData[dept].members;
    Logger.log(`部署「${dept}」のメンバー数:`, members.length);
    members.forEach((member, index) => {
      Logger.log(`部署「${dept}」のメンバー${index} (元の行${member._originalRowIndex}):`, JSON.stringify(member));
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
    membersLength: deptData.members ? deptData.members.length : 0
  }));
  
  // メンバーデータの詳細をログ出力
  if (deptData.members) {
    const membersList = deptData.members;
    Logger.log(`部署「${dept}」のメンバー一覧:`);
    membersList.forEach((member, index) => {
      Logger.log(`  メンバー${index} (元の行${member._originalRowIndex}): ${JSON.stringify(member)}`);
    });
  }
  
  const targetSs = SpreadsheetApp.openByUrl(targetUrl);
  const newSheet = createDeptSheet(targetSs, ganttTemplateSheet, dept);
  
  // 保護する列のインデックスを取得
  const preserveColumnIndices = PRESERVE_TEMPLATE_COLUMNS.map(columnName => ({
    name: columnName,
    index: ganttHeaders.indexOf(columnName)
  })).filter(col => col.index !== -1);
  
  Logger.log(`保護する列: ${preserveColumnIndices.map(col => `${col.name}(${col.index})`).join(', ')}`);
  
  // memberDateId生成とデータ準備
  Logger.log(`部署「${dept}」のガントデータを準備中...`);
  const preparedData = prepareGanttData(deptData, ganttHeaders, commonHeaders, daysPerMember, insertBlankLine);
  Logger.log(`部署「${dept}」の準備されたデータ行数: ${preparedData.length}`);

  // ガントチャート範囲の取得
  const ganttRange = newSheet.getRange(ganttHeaderRangeA1);
  const dataStartRow = ganttRange.getRow() + 1; // ヘッダー行の次から
  const dataStartCol = ganttRange.getColumn();
  
  // 保護する列の値を事前に保存
  const originalColumnValues = {};
  if (preserveColumnIndices.length > 0 && preparedData.length > 0) {
    Logger.log("GANTT_TEMPLATEから保護対象列の値を保存中...");
    
    preserveColumnIndices.forEach(col => {
      // 新しいシートでの対象列の範囲を取得（テンプレートから複製された値）
      const columnRange = newSheet.getRange(
        dataStartRow,
        dataStartCol + col.index,
        preparedData.length,
        1
      );
      const columnValues = columnRange.getValues();
      originalColumnValues[col.name] = columnValues;
      
      Logger.log(`${col.name}列の値を保存しました: ${columnValues.length}行`);
      Logger.log(`保存された${col.name}列の最初の5行:`, columnValues.slice(0, 5).map(row => row[0]));
      
      // 列の値が空の場合の警告
      const emptyCount = columnValues.filter(row => !row[0]).length;
      if (emptyCount > 0) {
        Logger.log(`警告: ${col.name}列の${emptyCount}行が空です`);
      }
    });
  }

  // メンバーデータをセット（保護対象列も一時的に上書き）
  const targetRange = newSheet.getRange(
    dataStartRow,
    dataStartCol,
    preparedData.length,
    ganttHeaders.length
  );
  targetRange.setValues(preparedData);
  
  // 保護対象列の値を復元
  if (preserveColumnIndices.length > 0 && Object.keys(originalColumnValues).length > 0) {
    Logger.log("保護対象列の値を復元中...");
    
    preserveColumnIndices.forEach(col => {
      if (originalColumnValues[col.name]) {
        const columnRange = newSheet.getRange(
          dataStartRow,
          dataStartCol + col.index,
          preparedData.length,
          1
        );
        columnRange.setValues(originalColumnValues[col.name]);
        Logger.log(`${col.name}列の値を復元しました`);
      }
    });
  }
  
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
  const members = deptData.members; // 既に配列なのでそのまま使用
  
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
    
    // メンバーデータをコピー（memberDateIdと保護対象列は除外）
    const excludeHeaders = [COL_HEADER_NAMES.MEMBER_DATE_ID, ...PRESERVE_TEMPLATE_COLUMNS];
    copyMemberDataToGanttRow(
      commonHeaders,
      headerIndices,
      baseRow,
      memberObj,
      excludeHeaders
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
  
  return allMemberRows.map((row, rowIndex) => {
    if (row.every(cell => cell === "")) {
      // 空白行では currentMemberIndex を増加させない（メンバー処理完了時に既に増加済み）
      dayCounter = 1;
      return row;
    }
    
    if (currentMemberIndex >= members.length) {
      Logger.log(`警告: 無効なメンバーインデックス ${currentMemberIndex} (メンバー数: ${members.length})`);
      return row;
    }
    
    const currentMember = members[currentMemberIndex];
    if (!currentMember || !currentMember[COL_HEADER_NAMES.MEMBER_ID]) {
      Logger.log(`警告: 無効なメンバーデータ (インデックス: ${currentMemberIndex}):`, JSON.stringify(currentMember));
      return row;
    }
    
    const memberId = currentMember[COL_HEADER_NAMES.MEMBER_ID];
    const memberDateId = generateMemberDateId(memberId, `day${dayCounter}`);
    
    Logger.log(`行${rowIndex}: メンバー${currentMemberIndex} (${memberId}, 元の行${currentMember._originalRowIndex}) の ${dayCounter}日目`);
    
    row[memberDateIdIndex] = memberDateId;
    
    // memberIdを各行に展開
    if (memberIdIndex !== undefined) {
      row[memberIdIndex] = memberId;
    }
    
    // メンバーデータを各行に再コピー（順序を保証するため、保護対象列は除外）
    const excludeHeaders = [COL_HEADER_NAMES.MEMBER_DATE_ID, ...PRESERVE_TEMPLATE_COLUMNS];
    copyMemberDataToGanttRow(
      commonHeaders,
      headerIndices,
      row,
      currentMember,
      excludeHeaders
    );
    
    dayCounter++;
    if (dayCounter > daysPerMember) {
      dayCounter = 1;
      currentMemberIndex++;
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

/**
 * promptUserForGanttChartInfo関数をデバッグするためのテスト関数
 * カスタムダイアログからの戻り値をシミュレートして、handleGanttDialogSubmitを直接呼び出す
 */
function testPromptUserForGanttChartInfo() {
  try {
    Logger.log("=== promptUserForGanttChartInfo テスト関数開始 ===");
    
    // ガントチャートテンプレートの範囲選択（実際の関数と同じロジック）
    validateNamedRange(RANGE_NAMES.GANTT_HEADER_ROW);
    const ganttHeaderRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(RANGE_NAMES.GANTT_HEADER_ROW);
    
    // ダミーのフォームデータ（カスタムダイアログから返されるJSONをシミュレート）
    const dummyFormData = {
      targetUrl: "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit", // 実際のURLに変更してください
      daysPerMember: 5,
      insertBlankLine: true
    };
    
    // ダミーのコンテキストデータ
    const dummyContext = {
      ganttHeaderRange: ganttHeaderRange.getA1Notation()
    };
    
    Logger.log("テスト用ダミーデータ:");
    Logger.log("フォームデータ:", JSON.stringify(dummyFormData));
    Logger.log("コンテキスト:", JSON.stringify(dummyContext));
    
    // 実際のコールバック関数を呼び出し
    handleGanttDialogSubmit(dummyFormData, dummyContext);
    
    Logger.log("=== promptUserForGanttChartInfo テスト関数完了 ===");
    
  } catch (error) {

    
    console.error("=== promptUserForGanttChartInfo テスト関数エラー ===");
    console.error("エラーメッセージ:", error.message);
    console.error("エラー名:", error.name);
    console.error("スタックトレース:", error.stack);
    console.error("エラーオブジェクト全体:", error);
    
    // UIにもエラーを表示
    SpreadsheetApp.getUi().alert(`テスト関数エラー: ${error.message}\n\n詳細は実行ログを確認してください。`);
    
    throw error;
  }
}

/**
 * テスト用：異なるパラメータでのバリエーションテスト
 */
function testPromptUserForGanttChartInfoVariations() {
  const testCases = [
    {
      name: "基本ケース",
      formData: {
        targetUrl: "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_1/edit",
        daysPerMember: 3,
        insertBlankLine: false
      }
    },
    {
      name: "空白行挿入ありケース",
      formData: {
        targetUrl: "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_2/edit",
        daysPerMember: 7,
        insertBlankLine: true
      }
    },
    {
      name: "最小日数ケース",
      formData: {
        targetUrl: "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID_3/edit",
        daysPerMember: 1,
        insertBlankLine: false
      }
    }
  ];
  
  Logger.log("=== promptUserForGanttChartInfo バリエーションテスト開始 ===");
  
  try {
    // ガントチャートテンプレートの範囲選択
    validateNamedRange(RANGE_NAMES.GANTT_HEADER_ROW);
    const ganttHeaderRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(RANGE_NAMES.GANTT_HEADER_ROW);
    
    const dummyContext = {
      ganttHeaderRange: ganttHeaderRange.getA1Notation()
    };
    
    testCases.forEach((testCase, index) => {
      Logger.log(`--- テストケース ${index + 1}: ${testCase.name} ---`);
      Logger.log("フォームデータ:", JSON.stringify(testCase.formData));
      
      try {
        handleGanttDialogSubmit(testCase.formData, dummyContext);
        Logger.log(`テストケース ${index + 1} 完了`);
      } catch (error) {
        Logger.log(`テストケース ${index + 1} エラー:`, error.message);
        console.error(`テストケース ${index + 1} エラー:`, error);
      }
    });
    
    Logger.log("=== promptUserForGanttChartInfo バリエーションテスト完了 ===");
    
  } catch (error) {
    Logger.log("=== バリエーションテスト全体エラー ===");
    Logger.log("エラーメッセージ:", error.message);
    console.error("バリエーションテスト全体エラー:", error);
    
    SpreadsheetApp.getUi().alert(`バリエーションテストエラー: ${error.message}\n\n詳細は実行ログを確認してください。`);
  }
}
