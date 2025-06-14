// 列・行インデックス定数オブジェクト
let RDB_COL_INDEXES = {
  dept: "",
  memberDateId: "",
  startTime: "",
  endTime: "",
  job: "",
  background: "",
};

let GANTT_COL_INDEXES = {
  memberDateId: "",
  firstData: "",
};

let GANTT_ROW_INDEXES = {
  timeScale: "",
  firstData: "",
};

let CONFLICT_COL_INDEXES = {
  dept: "",
  memberDateId: "",
  startTime: "",
  endTime: "",
  job: "",
  background: "",
  source: "",
};

// 全ての名前付き範囲を確認する関数
function validateAllNamedRanges() {
  const requiredRanges = ["rdbHeaderRow", "ganttHeaderRow", "timeScale", "firstData"];
  
  try {
    for (const rangeName of requiredRanges) {
      validateNamedRange(rangeName);
    }
    
    Browser.msgBox("確認完了", "全ての名前付き範囲の確認が完了しました。処理を続行します。", Browser.Buttons.OK);
    
  } catch (error) {
    Browser.msgBox("エラー", `名前付き範囲の確認でエラーが発生しました:\n${error.message}`, Browser.Buttons.OK);
    throw error;
  }
}

// 名前付き範囲の設定確認関数
function validateNamedRange(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 名前付き範囲の存在確認
    const namedRange = ss.getRangeByName(rangeName);
    
    if (!namedRange) {
      // 名前付き範囲が存在しない場合
      const message = `名前付き範囲「${rangeName}」が定義されていません。\n\n` +
                     `メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」を設定してください。\n\n` +
                     `設定後、スクリプトを再実行してください。`;
      
      Browser.msgBox("名前付き範囲が未定義", message, Browser.Buttons.OK);
      throw new Error(`処理を中止します: 名前付き範囲「${rangeName}」が設定されていません`);
    }
    
    // 名前付き範囲をアクティブ化してユーザーに確認
    namedRange.activate();
    
    const message = `現在選択されている範囲が「${rangeName}」の名前付き範囲として実行して問題ないですか？\n\n` +
                   `修正したい場合は「いいえ」を選択し、メニューバーの「データ」>「名前付き範囲」から設定を修正してください。`;
    
    const response = Browser.msgBox("名前付き範囲の確認", message, Browser.Buttons.YES_NO);
    
    if (response === "no") {
      const retryMessage = `「いいえ」が選択されました。\n\n` +
                          `メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」の設定を修正後、スクリプトを再実行してください。`;
      
      Browser.msgBox("名前付き範囲の修正", retryMessage, Browser.Buttons.OK);
      throw new Error(`処理を中止します: 名前付き範囲「${rangeName}」の修正が必要です`);
    }
    
    return true;
    
  } catch (error) {
    console.error(`名前付き範囲「${rangeName}」の確認中にエラーが発生しました:`, error);
    throw error;
  }
}

// 名前付き範囲からインデックスを設定する関数
function initializeColumnIndexes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // RDB_COL_INDEXESの設定
    const rdbHeaderRowRange = ss.getRangeByName("rdbHeaderRow");
    if (rdbHeaderRowRange) {
      const headerValues = rdbHeaderRowRange.getValues()[0]; // 1行目のみ取得
      const startCol = rdbHeaderRowRange.getColumn();
      
      Object.keys(RDB_COL_INDEXES).forEach(key => {
        const colIndex = headerValues.indexOf(key);
        if (colIndex !== -1) {
          RDB_COL_INDEXES[key] = startCol + colIndex - 1; // スプレッドシートは1ベース、配列は0ベース
        }
      });
    }
    
    // GANTT_COL_INDEXESのmemberDateId設定
    const ganttHeaderRowRange = ss.getRangeByName("ganttHeaderRow");
    if (ganttHeaderRowRange) {
      const headerValues = ganttHeaderRowRange.getValues()[0];
      const startCol = ganttHeaderRowRange.getColumn();
      const memberDateIdIndex = headerValues.indexOf("memberDateId");
      if (memberDateIdIndex !== -1) {
        GANTT_COL_INDEXES.memberDateId = startCol + memberDateIdIndex - 1;
      }
    }
    
    // GANTT_ROW_INDEXESのtimeScale設定
    const timeScaleRange = ss.getRangeByName("timeScale");
    if (timeScaleRange) {
      GANTT_ROW_INDEXES.timeScale = timeScaleRange.getRow();
    }
    
    // firstData範囲の設定
    const firstDataRange = ss.getRangeByName("firstData");
    if (firstDataRange) {
      GANTT_COL_INDEXES.firstData = firstDataRange.getColumn();
      GANTT_ROW_INDEXES.firstData = firstDataRange.getRow();
    }
    
    // CONFLICT_COL_INDEXESの設定（RDB_COL_INDEXESと同じ値を使用）
    Object.keys(CONFLICT_COL_INDEXES).forEach(key => {
      // sourceプロパティの場合は、未設定なら最後の列として追加
      const isSourceKey = key === 'source';
      const isKeyMissing = !RDB_COL_INDEXES.hasOwnProperty(key) || RDB_COL_INDEXES[key] === '';
      
      if (isSourceKey && isKeyMissing) {
        const maxIndex = Math.max(...Object.values(RDB_COL_INDEXES));
        CONFLICT_COL_INDEXES[key] = maxIndex + 1;
      } else if (RDB_COL_INDEXES.hasOwnProperty(key)) {
        CONFLICT_COL_INDEXES[key] = RDB_COL_INDEXES[key];
      }
    });
    
  } catch (error) {
    console.error("インデックスの初期化中にエラーが発生しました:", error);
    throw new Error("名前付き範囲からのインデックス取得に失敗しました: " + error.message);
  }
}

// インデックスオブジェクトから列順序配列を生成する関数
function getColumnOrder(indexes) {
  // 最大のインデックス値を取得
  const maxIndex = Math.max(...Object.values(indexes));

  // インデックス値をキーとしてプロパティ名を格納するマップを作成
  const indexToKey = new Map();
  Object.entries(indexes).forEach(([key, index]) => {
    indexToKey.set(index, key);
  });

  // 0からmaxIndexまでの配列を生成し、欠番には""をセット
  return Array.from({ length: maxIndex + 1 }, (_, i) => {
    return indexToKey.has(i) ? indexToKey.get(i) : "";
  });
}

// 従来のスプレッドシート依存関数は削除済み
// 全てのインデックスは定数オブジェクトで管理
