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

/**
 * カスタムダイアログシステムで使用するスクリプトプロパティのプレフィックス
 */
const DIALOG_PROP_PREFIX = "DIALOG_CALLBACK_";

/**
 * カスタムダイアログを表示する関数
 *
 * @param {Object} options ダイアログのオプション
 * @param {Array} options.fields フィールド定義の配列
 * @param {String} options.title ダイアログのタイトル
 * @param {String} options.message ダイアログのメッセージ
 * @param {Number} options.width ダイアログの幅（省略可）
 * @param {Number} options.height ダイアログの高さ（省略可）
 * @param {String} options.onSubmitFuncName 送信時に呼び出す**グローバル関数名**（文字列）
 * @param {String} options.onCancelFuncName キャンセル時に呼び出す**グローバル関数名**（文字列、省略可）
 * @param {Object} options.context 送信/キャンセル関数に渡す追加情報（省略可、シリアライズ可能である必要あり）
 */
function showCustomDialog(options) {
  // 必須オプションのチェック
  if (!options.fields || !options.onSubmitFuncName) {
    throw new Error("フィールド定義(fields)と送信時コールバック関数名(onSubmitFuncName)は必須です");
  }

  // デフォルト値の設定
  const dialogOptions = {
    title: options.title || "データ入力",
    message: options.message || "",
    width: options.width || 400,
    height: options.height || 300,
    fields: options.fields,
    onSubmitFuncName: options.onSubmitFuncName,
    onCancelFuncName: options.onCancelFuncName || null, // 省略時はnull
    context: options.context || {},
  };

  // 一意のキーを生成
  const dialogKey = DIALOG_PROP_PREFIX + new Date().getTime();

  // スクリプトプロパティに情報を保存
  const props = PropertiesService.getScriptProperties();
  const dataToStore = JSON.stringify({
    submitFunc: dialogOptions.onSubmitFuncName,
    cancelFunc: dialogOptions.onCancelFuncName,
    context: dialogOptions.context,
    expiresAt: new Date().getTime() + 10 * 60 * 1000, // 10分で有効期限切れ
  });
  props.setProperty(dialogKey, dataToStore);

  // テンプレートを作成
  const htmlTemplate = HtmlService.createTemplateFromFile("3101_view_CustomPopupTemplate");

  // テンプレートに変数を渡す
  htmlTemplate.fields = dialogOptions.fields;
  htmlTemplate.message = dialogOptions.message;
  htmlTemplate.dialogKey = dialogKey; // キーをHTMLに渡す

  // HTMLを評価して取得
  const htmlOutput = htmlTemplate
    .evaluate()
    .setWidth(dialogOptions.width)
    .setHeight(dialogOptions.height)
    .setTitle(dialogOptions.title);

  // モーダルダイアログとして表示
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogOptions.title);
}

/**
 * フォームから送信されたデータを処理する中央ハンドラ関数
 *
 * @param {Object} formData フォームデータ（キャンセルの場合はnull）
 * @param {String} dialogKey ダイアログキー
 * @param {String} action アクション（"submit" または "cancel"）
 * @return {Object} 処理結果
 */
function handleDialogResponse(formData, dialogKey, action) {
  const props = PropertiesService.getScriptProperties();
  const storedData = props.getProperty(dialogKey);

  // 不要になったプロパティは削除
  props.deleteProperty(dialogKey);

  if (!storedData) {
    console.error("ダイアログ情報が見つかりません:", dialogKey);
    return { success: false, error: "ダイアログ情報が見つかりません。タイムアウトした可能性があります。" };
  }

  try {
    const parsedData = JSON.parse(storedData);

    // 有効期限チェック
    if (parsedData.expiresAt < new Date().getTime()) {
      console.error("ダイアログ情報が期限切れです:", dialogKey);
      return { success: false, error: "ダイアログセッションがタイムアウトしました。" };
    }

    // 実行する関数名を決定
    const funcName = action === "submit" ? parsedData.submitFunc : parsedData.cancelFunc;

    if (funcName && typeof this[funcName] === "function") {
      // グローバルスコープから関数を見つけて実行
      if (action === "submit") {
        this[funcName](formData, parsedData.context); // contextも渡す
      } else {
        this[funcName](parsedData.context); // contextも渡す
      }
      return { success: true };
    } else if (funcName) {
      console.error("指定されたコールバック関数が見つかりません:", funcName);
      return { success: false, error: "コールバック関数が見つかりません: " + funcName };
    } else if (action === "cancel") {
      // キャンセル関数が指定されていない場合は正常終了
      return { success: true };
    } else {
      return { success: false, error: "実行すべきコールバック関数がありません" };
    }
  } catch (error) {
    console.error("ダイアログ応答処理エラー:", dialogKey, error);
    return { success: false, error: "エラーが発生しました: " + String(error) };
  }
}

/**
 * 古いダイアログプロパティをクリーンアップする関数（トリガーなどで定期実行推奨）
 */
function cleanupDialogProperties() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const now = new Date().getTime();

  for (const key in allProps) {
    if (key.startsWith(DIALOG_PROP_PREFIX)) {
      try {
        const parsedData = JSON.parse(allProps[key]);
        if (parsedData.expiresAt < now) {
          props.deleteProperty(key);
          console.log("期限切れのダイアログプロパティを削除:", key);
        }
      } catch (e) {
        // 不正なJSONの場合は削除
        props.deleteProperty(key);
        console.log("不正なダイアログプロパティを削除:", key);
      }
    }
  }
}

/**
 * 使用例：シフト時間入力ダイアログ
 */
function showShiftTimeInputDialog(startCell) {
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

  showCustomDialog({
    fields: fieldConfigs,
    title: "シフト時間設定",
    message: "シフトの開始時刻、終了時刻、時間間隔を入力",
    onSubmitFuncName: "handleFormSubmit",
    onCancelFuncName: "handleFormCancel",
    context: { startCell: startCell },
  });
}

/**
 * 使用例：行複製ダイアログ
 */
function showRowDuplicationDialog(orgRange) {
  const fieldConfigs = [
    {
      id: "times",
      label: "複製する行数",
      type: "number",
      required: true,
    },
  ];

  showCustomDialog({
    fields: fieldConfigs,
    title: "行複製設定",
    message: "行セットを何回複製するか入力",
    onSubmitFuncName: "handleFormSubmit",
    onCancelFuncName: "handleFormCancel",
    context: { orgRange: orgRange },
  });
}
