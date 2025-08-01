/**
 * ConfigManager - 設定とカラム管理を担当するクラス
 * 
 * 責任:
 * - 名前付き範囲の管理
 * - カラムインデックスの管理
 * - 設定の検証
 * - 初期化処理
 */
class ConfigManager {
  constructor() {
    this.rangeNames = {
      RDB_HEADER_ROW: "登録予定_入力データシート_ヘッダー部分",
      GANTT_HEADER_ROW: "シフト表テンプレシート_ヘッダー部分",
      TIME_SCALE: "シフト表テンプレシート_時間軸部分",
      FIRST_DATA: "シフト表テンプレシート_シフトデータ部分の一番左上のセル",
    };

    this.sheetNames = {
      IN_RDB: "4.登録予定_入力データ",
      OUT_RDB: "4.登録済み_出力データ",
      CONFLICT_RDB: "4.登録失敗_重複データ",
      ERROR_RDB: "4.登録失敗_エラーデータ",
      GANTT_TEMPLATE: "1~2.シフト表テンプレ",
      MEMBER_DATA: "2~3.メンバー情報",
    };

    this.columnIndexes = {
      RDB: {
        dept: null,
        memberDateId: null,
        startTime: null,
        endTime: null,
        job: null,
        background: null,
      },
      GANTT: {
        memberDateId: null,
        firstData: null,
      },
      ROW: {
        timeScale: null,
        firstData: null,
      },
      CONFLICT: {
        dept: null,
        memberDateId: null,
        startTime: null,
        endTime: null,
        job: null,
        background: null,
        source: null,
      },
      ERROR: {
        dept: null,
        memberDateId: null,
        startTime: null,
        endTime: null,
        job: null,
        background: null,
        source: null,
        errorMessage: null,
      },
    };

    this.isInitialized = false;
  }

  /**
   * 設定の初期化
   * @returns {ConfigManager} 初期化済みのConfigManagerインスタンス
   */
  static initialize() {
    const config = new ConfigManager();
    config.validateAllNamedRanges();
    config.initializeColumnIndexes();
    config.isInitialized = true;
    return config;
  }

  /**
   * 全ての名前付き範囲を検証
   * @throws {ValidationError} 検証に失敗した場合
   */
  validateAllNamedRanges() {
    const requiredRanges = Object.values(this.rangeNames);

    try {
      for (const rangeName of requiredRanges) {
        this.validateNamedRange(rangeName);
      }
      
      NotificationService.showSuccess("全ての名前付き範囲の確認が完了しました。処理を続行します。");
    } catch (error) {
      throw new ValidationError(`名前付き範囲の確認でエラーが発生しました: ${error.message}`);
    }
  }

  /**
   * 個別の名前付き範囲を検証
   * @param {string} rangeName - 検証する範囲名
   * @throws {ValidationError} 検証に失敗した場合
   */
  validateNamedRange(rangeName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      const namedRange = ss.getRangeByName(rangeName);

      if (!namedRange) {
        const message = 
          `名前付き範囲「${rangeName}」が定義されていません。\n\n` +
          `メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」を設定してください。\n\n` +
          `設定後、スクリプトを再実行してください。`;

        Browser.msgBox("名前付き範囲が未定義", message, Browser.Buttons.OK);
        throw new ValidationError(`名前付き範囲「${rangeName}」が設定されていません`);
      }

      namedRange.activate();
      SpreadsheetApp.flush();

      const message = 
        `範囲「${rangeName}」は現在選択されている範囲で問題ないですか？\n\n` +
        `修正したい場合は「いいえ」を選択し、メニューバーの「データ」>「名前付き範囲」から設定を修正してください。`;

      const response = Browser.msgBox("名前付き範囲の確認", message, Browser.Buttons.YES_NO);

      if (response === "no") {
        const retryMessage = 
          `「いいえ」が選択されました。\n\n` +
          `メニューバーの「データ」>「名前付き範囲」から範囲「${rangeName}」の設定を修正後、スクリプトを再実行してください。`;

        Browser.msgBox("名前付き範囲の修正", retryMessage, Browser.Buttons.OK);
        throw new ValidationError(`名前付き範囲「${rangeName}」の修正が必要です`);
      }

      return true;
    } catch (error) {
      console.error(`名前付き範囲「${rangeName}」の確認中にエラーが発生しました:`, error);
      throw error;
    }
  }

  /**
   * 名前付き範囲からカラムインデックスを初期化
   * @throws {ValidationError} 初期化に失敗した場合
   */
  initializeColumnIndexes() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      this.initializeRdbColumnIndexes(ss);
      this.initializeGanttColumnIndexes(ss);
      this.initializeRowIndexes(ss);
      this.initializeConflictAndErrorIndexes();
    } catch (error) {
      console.error("インデックスの初期化中にエラーが発生しました:", error);
      throw new ValidationError("名前付き範囲からのインデックス取得に失敗しました: " + error.message);
    }
  }

  /**
   * RDBカラムインデックスの初期化
   * @param {Spreadsheet} ss - スプレッドシートオブジェクト
   */
  initializeRdbColumnIndexes(ss) {
    const rdbHeaderRowRange = ss.getRangeByName(this.rangeNames.RDB_HEADER_ROW);
    if (!rdbHeaderRowRange) return;

    const headerValues = rdbHeaderRowRange.getValues()[0];
    const startCol = rdbHeaderRowRange.getColumn() - 1;

    this.validateRdbHeaders(headerValues);

    Object.keys(this.columnIndexes.RDB).forEach((key) => {
      const colIndex = headerValues.indexOf(key);
      if (colIndex === -1) {
        throw new ValidationError(`「${key}」が範囲「${this.rangeNames.RDB_HEADER_ROW}」に含まれていません`);
      }
      this.columnIndexes.RDB[key] = startCol + colIndex;
    });
  }

  /**
   * RDBヘッダーの検証
   * @param {Array} headerValues - ヘッダー値の配列
   */
  validateRdbHeaders(headerValues) {
    const unknownValues = headerValues.filter(value => 
      !this.columnIndexes.RDB.hasOwnProperty(value)
    );

    if (unknownValues.length > 0) {
      const message = 
        `範囲「${this.rangeNames.RDB_HEADER_ROW}」に不要な見出し「${unknownValues.join("、")}」が含まれています。\n` +
        `不要な見出しを削除してメニューバーの「データ」>「名前付き範囲」から範囲「${this.rangeNames.RDB_HEADER_ROW}」を選びなおしてください。`;

      Browser.msgBox(`範囲「${this.rangeNames.RDB_HEADER_ROW}」の確認`, message, Browser.Buttons.OK);
      throw new ValidationError(`範囲「${this.rangeNames.RDB_HEADER_ROW}」の設定に問題があります`);
    }
  }

  /**
   * Ganttカラムインデックスの初期化
   * @param {Spreadsheet} ss - スプレッドシートオブジェクト
   */
  initializeGanttColumnIndexes(ss) {
    const ganttHeaderRowRange = ss.getRangeByName(this.rangeNames.GANTT_HEADER_ROW);
    if (!ganttHeaderRowRange) return;

    const headerValues = ganttHeaderRowRange.getValues()[0];
    const startCol = ganttHeaderRowRange.getColumn() - 1;
    const memberDateIdIndex = headerValues.indexOf("memberDateId");

    if (memberDateIdIndex === -1) {
      throw new ValidationError(`「memberDateId」が範囲「${this.rangeNames.GANTT_HEADER_ROW}」に含まれていません`);
    }

    this.columnIndexes.GANTT.memberDateId = startCol + memberDateIdIndex;

    const firstDataRange = ss.getRangeByName(this.rangeNames.FIRST_DATA);
    if (firstDataRange) {
      this.columnIndexes.GANTT.firstData = firstDataRange.getColumn() - 1;
    }
  }

  /**
   * 行インデックスの初期化
   * @param {Spreadsheet} ss - スプレッドシートオブジェクト
   */
  initializeRowIndexes(ss) {
    const timeScaleRange = ss.getRangeByName(this.rangeNames.TIME_SCALE);
    if (timeScaleRange) {
      this.columnIndexes.ROW.timeScale = timeScaleRange.getRow() - 1;
    }

    const firstDataRange = ss.getRangeByName(this.rangeNames.FIRST_DATA);
    if (firstDataRange) {
      this.columnIndexes.ROW.firstData = firstDataRange.getRow() - 1;
    }
  }

  /**
   * ConflictとErrorカラムインデックスの初期化
   */
  initializeConflictAndErrorIndexes() {
    Object.keys(this.columnIndexes.CONFLICT).forEach((key) => {
      if (key === "source") {
        const maxIndex = Math.max(...Object.values(this.columnIndexes.RDB));
        this.columnIndexes.CONFLICT[key] = maxIndex + 1;
      } else {
        this.columnIndexes.CONFLICT[key] = this.columnIndexes.RDB[key];
      }
    });

    Object.keys(this.columnIndexes.ERROR).forEach((key) => {
      if (key === "errorMessage") {
        const maxIndex = Math.max(...Object.values(this.columnIndexes.CONFLICT));
        this.columnIndexes.ERROR[key] = maxIndex + 1;
      } else {
        this.columnIndexes.ERROR[key] = this.columnIndexes.CONFLICT[key];
      }
    });
  }

  /**
   * カラムインデックスから列順序配列を生成
   * @param {Object} indexes - インデックスオブジェクト
   * @returns {Array} 列順序配列
   */
  getColumnOrder(indexes) {
    const maxIndex = Math.max(...Object.values(indexes));
    const indexToKey = new Map();
    
    Object.entries(indexes).forEach(([key, index]) => {
      indexToKey.set(index, key);
    });

    return Array.from({ length: maxIndex + 1 }, (_, i) => {
      return indexToKey.has(i) ? indexToKey.get(i) : "";
    });
  }

  /**
   * 初期化状態の確認
   * @throws {Error} 初期化されていない場合
   */
  ensureInitialized() {
    if (!this.isInitialized) {
      throw new Error("ConfigManagerが初期化されていません。ConfigManager.initialize()を呼び出してください。");
    }
  }

  /**
   * 設定情報の取得
   * @returns {Object} 設定情報オブジェクト
   */
  getConfig() {
    this.ensureInitialized();
    return {
      rangeNames: { ...this.rangeNames },
      sheetNames: { ...this.sheetNames },
      columnIndexes: JSON.parse(JSON.stringify(this.columnIndexes)),
    };
  }

  /**
   * 特定のカラムインデックスの取得
   * @param {string} type - インデックスタイプ（RDB, GANTT, ROW, CONFLICT, ERROR）
   * @returns {Object} カラムインデックスオブジェクト
   */
  getColumnIndexes(type) {
    this.ensureInitialized();
    if (!this.columnIndexes[type]) {
      throw new Error(`無効なインデックスタイプ: ${type}`);
    }
    return { ...this.columnIndexes[type] };
  }

  /**
   * シート名の取得
   * @returns {Object} シート名オブジェクト
   */
  getSheetNames() {
    return { ...this.sheetNames };
  }
}

// 従来のグローバル変数との互換性を保つための関数
function validateAllNamedRanges() {
  const config = ConfigManager.initialize();
  config.validateAllNamedRanges();
}

function initializeColumnIndexes() {
  const config = ConfigManager.initialize();
  config.initializeColumnIndexes();
}

function getColumnOrder(indexes) {
  const config = new ConfigManager();
  return config.getColumnOrder(indexes);
} 