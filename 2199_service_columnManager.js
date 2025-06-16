const RANGE_NAMES = {
  RDB_HEADER_ROW: "登録予定_入力データシート_ヘッダー部分",
  GANTT_HEADER_ROW: "GCテンプレシート_ヘッダー部分",
  TIME_SCALE: "GCテンプレシート_時間軸部分",
  FIRST_DATA: "GCテンプレシート_シフトデータ部分の一番左上のセル",
};

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

let ERROR_COL_INDEXES = {
  dept: "",
  memberDateId: "",
  startTime: "",
  endTime: "",
  job: "",
  background: "",
  source: "",
  errorMessage: "",
};

// 全ての名前付き範囲を確認する関数
function validateAllNamedRanges() {
  const requiredRanges = [RANGE_NAMES.RDB_HEADER_ROW, RANGE_NAMES.GANTT_HEADER_ROW, RANGE_NAMES.TIME_SCALE, RANGE_NAMES.FIRST_DATA];

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
      const message =
        `名前付き範囲「${rangeName}」が定義されていません。\n\n` +
        `メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」を設定してください。\n\n` +
        `設定後、スクリプトを再実行してください。`;

      Browser.msgBox("名前付き範囲が未定義", message, Browser.Buttons.OK);
      throw new Error(`処理を中止します: 名前付き範囲「${rangeName}」が設定されていません`);
    }

    // 名前付き範囲をアクティブ化してユーザーに確認
    namedRange.activate();
    SpreadsheetApp.flush(); // 変更を即座に反映

    const message =
      `範囲「${rangeName}」は現在選択されている範囲で問題ないですか？

      修正したい場合は「いいえ」を選択し、メニューバーの「データ」>「名前付き範囲」から設定を修正してください。`;

    const response = Browser.msgBox("名前付き範囲の確認", message, Browser.Buttons.YES_NO);

    if (response === "no") {
      const retryMessage =
        `「いいえ」が選択されました。
        
        メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」の設定を修正後、スクリプトを再実行してください。`;

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
    const rdbHeaderRowRange = ss.getRangeByName(RANGE_NAMES.RDB_HEADER_ROW);
    if (rdbHeaderRowRange) {
      const headerValues = rdbHeaderRowRange.getValues()[0]; // 1行目のみ取得
      const startCol = rdbHeaderRowRange.getColumn() - 1; // 0-indexed配列用に-1

      if (headerValues.some((value) => !RDB_COL_INDEXES.hasOwnProperty(value))) {
        const unknownValues = headerValues.filter((value) => !RDB_COL_INDEXES.hasOwnProperty(value));
        const message =
          `範囲「${RANGE_NAMES.RDB_HEADER_ROW}」に不要な見出し「${unknownValues.join("、")}」が含まれています。
          不要な見出しを削除してメニューバーの「データ」>「名前付き範囲」から範囲「${RANGE_NAMES.RDB_HEADER_ROW}」を選びなおしてください。
          不要な見出しを削除したらスクリプトを再実行してください。`;

        Browser.msgBox(`範囲「${RANGE_NAMES.RDB_HEADER_ROW}」の確認`, message, Browser.Buttons.OK);
        throw new Error(`処理を中止します: 範囲「${RANGE_NAMES.RDB_HEADER_ROW}」の設定に問題があります。`);
      }

      Object.keys(RDB_COL_INDEXES).forEach((key) => {
        const colIndex = headerValues.indexOf(key);

        if (colIndex === -1) {
          const message =
            `「${key}」が範囲「${RANGE_NAMES.RDB_HEADER_ROW}」に含まれていません。
            メニューバーの「データ」>「名前付き範囲」から範囲「${RANGE_NAMES.RDB_HEADER_ROW}」を選びなおしてください。`;

          Browser.msgBox(`範囲「${RANGE_NAMES.RDB_HEADER_ROW}」の確認`, message, Browser.Buttons.OK);
          throw new Error(`処理を中止します: 範囲「${RANGE_NAMES.RDB_HEADER_ROW}」の設定に問題があります`);
        } else {
          RDB_COL_INDEXES[key] = startCol + colIndex; // startColが既に0ベースなので、そのまま加算
        }
      });
    }

    // GANTT_COL_INDEXESのmemberDateId設定
    const ganttHeaderRowRange = ss.getRangeByName(RANGE_NAMES.GANTT_HEADER_ROW);
    if (ganttHeaderRowRange) {
      const headerValues = ganttHeaderRowRange.getValues()[0];
      const startCol = ganttHeaderRowRange.getColumn() - 1; // 0-indexed配列用に-1
      const memberDateIdIndex = headerValues.indexOf("memberDateId");
      
      if (memberDateIdIndex !== -1) {
        GANTT_COL_INDEXES.memberDateId = startCol + memberDateIdIndex; // startColが既に0ベースなので、そのまま加算
      }else{
        const message =
          `「memberDateId」が範囲「${RANGE_NAMES.GANTT_HEADER_ROW}」に含まれていません。
          メニューバーの「データ」>「名前付き範囲」から範囲「${RANGE_NAMES.GANTT_HEADER_ROW}」を選びなおしてください。`;
        Browser.msgBox(`範囲「${RANGE_NAMES.GANTT_HEADER_ROW}」の確認`, message, Browser.Buttons.OK);
        throw new Error(`処理を中止します: 範囲「${RANGE_NAMES.GANTT_HEADER_ROW}」の設定に問題があります`);
      }
    }

    // GANTT_ROW_INDEXESのtimeScale設定
    const timeScaleRange = ss.getRangeByName(RANGE_NAMES.TIME_SCALE);
    if (timeScaleRange) {
      GANTT_ROW_INDEXES.timeScale = timeScaleRange.getRow() - 1; // 0-indexed配列用に-1
    }

    // firstData範囲の設定
    const firstDataRange = ss.getRangeByName(RANGE_NAMES.FIRST_DATA);
    if (firstDataRange) {
      GANTT_COL_INDEXES.firstData = firstDataRange.getColumn() - 1; // 0-indexed配列用に-1
      GANTT_ROW_INDEXES.firstData = firstDataRange.getRow() - 1; // 0-indexed配列用に-1
    }

    // CONFLICT_COL_INDEXESの設定（RDB_COL_INDEXESと同じ値を使用）
    Object.keys(CONFLICT_COL_INDEXES).forEach((key) => {
      // sourceプロパティの場合は、未設定なら最後の列として追加
      if (key === "source") {
        const maxIndex = Math.max(...Object.values(RDB_COL_INDEXES));
        CONFLICT_COL_INDEXES[key] = maxIndex + 1;
      } else {
        CONFLICT_COL_INDEXES[key] = RDB_COL_INDEXES[key];
      }
    });

    // ERROR_COL_INDEXESの設定（CONFLICT_COL_INDEXESと同じ値を使用）
    Object.keys(ERROR_COL_INDEXES).forEach((key) => {
      if (key === "errorMessage") {
        const maxIndex = Math.max(...Object.values(CONFLICT_COL_INDEXES));
        ERROR_COL_INDEXES[key] = maxIndex + 1;
      } else {
        ERROR_COL_INDEXES[key] = CONFLICT_COL_INDEXES[key];
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
