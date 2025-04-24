/**
 * 動的なポップアップウィンドウを表示して入力値を取得する関数
 * @param {Array} fieldConfigs - 入力フィールド設定の配列
 * @param {String} message - ポップアップウィンドウの上部に表示するメッセージ
 * @return {Object} - ユーザー入力値のオブジェクト（Promiseとして使用可）
 */
function showCustomInputDialog(fieldConfigs, message = '') {
  return new Promise((resolve, reject) => {
    try {
      // グローバル変数に保存するためのキーを生成
      const dialogKey = 'dialog_' + new Date().getTime();
      
      // コールバック関数を登録
      CacheService.getScriptCache().put(dialogKey, JSON.stringify({
        status: 'pending',
        data: null
      }), 600); // 10分間有効
      
      // HTMLテンプレートを作成
      const htmlTemplate = HtmlService.createTemplateFromFile('CustomPopupTemplate');
      
      // テンプレートに変数を渡す
      htmlTemplate.fieldConfigs = fieldConfigs;
      htmlTemplate.dialogKey = dialogKey;
      htmlTemplate.message = message;
      
      // HTMLを評価して取得
      const htmlOutput = htmlTemplate.evaluate()
        .setWidth(400)
        .setHeight(300)
        .setTitle('データ入力');
      
      // モーダルダイアログとして表示
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'データ入力');
      
      // ダイアログの結果を定期的にチェック
      const checkInterval = 1000; // 1秒ごとにチェック
      const maxChecks = 600; // 最大チェック回数（10分）
      let checkCount = 0;
      
      const intervalId = setInterval(() => {
        checkCount++;
        const cacheData = CacheService.getScriptCache().get(dialogKey);
        
        if (cacheData) {
          const dialogData = JSON.parse(cacheData);
          
          if (dialogData.status === 'completed') {
            clearInterval(intervalId);
            resolve(dialogData.data);
          } else if (dialogData.status === 'failed') {
            clearInterval(intervalId);
            reject(new Error(dialogData.error || '不明なエラーが発生しました'));
          }
        }
        
        // タイムアウト
        if (checkCount >= maxChecks) {
          clearInterval(intervalId);
          reject(new Error('ダイアログ応答のタイムアウト'));
        }
      }, checkInterval);
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * フォームから送信されたデータを処理する関数
 * @param {Object} formData - フォームから送信されたデータ
 * @param {String} dialogKey - ダイアログを識別するキー
 * @return {Object} - 処理結果
 */
function processFormData(formData, dialogKey) {
  try {
    // ここで必要な処理を行う（例：スプレッドシートに保存など）
    
    // 結果をキャッシュに保存
    CacheService.getScriptCache().put(dialogKey, JSON.stringify({
      status: 'completed',
      data: formData
    }), 600);
    
    return { success: true, data: formData };
  } catch (error) {
    // エラー情報をキャッシュに保存
    CacheService.getScriptCache().put(dialogKey, JSON.stringify({
      status: 'failed',
      error: error.message
    }), 600);
    
    return { success: false, error: error.message };
  }
}

/**
 * 使用例
 */
function testCustomDialog() {
  const fieldConfigs = [
    {
      id: "job",
      label: "業務内容",
      type: "string",
      required: true,
      value: ""
    },
    {
      id: "memberDateId",
      label: "メンバーID",
      type: "string",
      required: true,
      value: "9901"
    },
    {
      id: "email",
      label: "メールアドレス",
      type: "email",
      required: true,
      value: ""
    },
    {
      id: "website",
      label: "ウェブサイト",
      type: "url",
      required: false,
      value: ""
    },
    {
      id: "age",
      label: "年齢",
      type: "number",
      required: false,
      min: 18,
      max: 100,
      value: ""
    },
    {
      id: "startDate",
      label: "開始日",
      type: "date",
      required: false,
      value: ""
    },
    {
      id: "phone",
      label: "電話番号",
      type: "tel",
      required: false,
      value: ""
    }
  ];
  
  // メッセージ付きでダイアログを表示
  showCustomInputDialog(fieldConfigs, '新しいメンバー情報を入力してください')
    .then(result => {
      console.log('入力されたデータ:', result);
      // ここで結果を処理
    })
    .catch(error => {
      console.error('エラー:', error);
    });
} 