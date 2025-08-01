/**
 * データ取得サービス
 * 
 * 責任:
 * - ガントチャートデータの取得
 * - RDBデータの取得
 * - 結合セルの処理
 * - 部署ごとのデータ分類
 * - データ構造の正規化
 */
class DataRetriever {
  constructor(configManager) {
    this.config = configManager;
    this.rowIndexes = configManager.getColumnIndexes('ROW');
    this.ganttIndexes = configManager.getColumnIndexes('GANTT');
    this.rdbIndexes = configManager.getColumnIndexes('RDB');
  }

  /**
   * すべてのガントチャートデータを取得し、シート名でグループ化
   * @param {Spreadsheet} ganttSs - ガントチャートスプレッドシート
   * @returns {Object} シート名をキーとしたガントデータのマップ
   */
  getAllGanttSheetDataAndGroupByName(ganttSs) {
    const sheets = ganttSs.getSheets();
    const ganttSsName = ganttSs.getName();
    const result = {};

    sheets.forEach((sheet, index) => {
      const sheetName = sheet.getName();
      
      try {
        // 処理中のシート名を通知
        NotificationService.showDetailedProgress(
          `ガントチャート「${ganttSsName}」のシート「${sheetName}」のデータを取得中...`,
          index + 1,
          sheets.length
        );

        const ganttData = this.getGanttSheetData(sheet);
        result[sheetName] = ganttData;
        
      } catch (error) {
        console.error(`シート「${sheetName}」のデータ取得中にエラーが発生しました:`, error);
        throw new DataProcessingError(`シート「${sheetName}」のデータ取得に失敗しました: ${error.message}`, { sheetName });
      }
    });

    return result;
  }

  /**
   * 単一のガントチャートシートからデータを取得
   * @param {Sheet} sourceSheet - ソースシート
   * @returns {GanttChartData} ガントチャートデータ
   */
  getGanttSheetData(sourceSheet) {
    const sourceRange = sourceSheet.getDataRange();
    const values = sourceRange.getValues();
    const backgrounds = sourceRange.getBackgrounds();
    const mergedRanges = sourceRange.getMergedRanges();

    // 結合範囲ごとに、最初のセルの値と背景色を全セルに反映
    this.expandMergedCells(values, backgrounds, mergedRanges);

    return new GanttChartData(values, backgrounds);
  }

  /**
   * 結合セルを展開して、各セルに元の値と背景色を反映
   * @param {Array} values - 値の2次元配列
   * @param {Array} backgrounds - 背景色の2次元配列
   * @param {Array} mergedRanges - 結合範囲の配列
   */
  expandMergedCells(values, backgrounds, mergedRanges) {
    mergedRanges.forEach((mergedRange) => {
      const startRow = mergedRange.getRow() - 1; // 0-indexed
      const startColumn = mergedRange.getColumn() - 1; // 0-indexed
      const rowCount = mergedRange.getNumRows();
      const columnCount = mergedRange.getNumColumns();

      const mergedValue = values[startRow][startColumn];
      const mergedBackground = backgrounds[startRow][startColumn];

      // 指定範囲の全セルに値と背景色を適用（縦横両方向に対応）
      for (let i = startRow; i < startRow + rowCount; i++) {
        for (let j = startColumn; j < startColumn + columnCount; j++) {
          if (i < values.length && j < values[i].length) {
            values[i][j] = mergedValue;
            backgrounds[i][j] = mergedBackground;
          }
        }
      }
    });
  }

  /**
   * RDBシートからデータを取得
   * @param {Sheet} rdbSheet - RDBシート
   * @returns {Array} RDBデータの2次元配列
   */
  getRdbData(rdbSheet) {
    try {
      const lastRow = rdbSheet.getLastRow();
      const lastColumn = rdbSheet.getLastColumn();
      
      if (lastRow === 0 || lastColumn === 0) {
        console.log("RDBシートにデータがありません");
        return [];
      }

      const rdbData = rdbSheet.getRange(1, 1, lastRow, lastColumn).getValues();
      
      // 空の行を除外
      return this.filterEmptyRows(rdbData);
      
    } catch (error) {
      console.error("RDBデータの取得中にエラーが発生しました:", error);
      throw new DataProcessingError(`RDBデータの取得に失敗しました: ${error.message}`, { sheetName: rdbSheet.getName() });
    }
  }

  /**
   * 空の行をフィルタリング
   * @param {Array} data - 2次元配列データ
   * @returns {Array} フィルタリング済みデータ
   */
  filterEmptyRows(data) {
    return data.filter((row) => row.some((cell) => cell !== ""));
  }

  /**
   * 指定した列の値でデータを部署別にグループ化
   * @param {Array} data - 2次元配列（1行目はヘッダー）
   * @param {number} colIndex - 分類に使用する列のインデックス（0始まり）
   * @returns {Object} 部署名をキーとしたオブジェクト
   */
  groupByDepartment(data, colIndex) {
    if (!data || data.length <= 1) {
      return {};
    }

    // ヘッダー行をスキップして処理
    return data.slice(1).reduce((acc, row) => {
      if (row.length > colIndex && row[colIndex]) {
        const key = String(row[colIndex]).trim();
        
        if (key) {
          // 既にキーが存在していればその配列に追加、存在しなければ新たな配列を作成
          if (!acc[key]) {
            acc[key] = [];
          }
          acc[key].push(row);
        }
      }
      return acc;
    }, {});
  }

  /**
   * データ取得の統計情報を作成
   * @param {Object} ganttDataByDept - 部署別ガントデータ
   * @param {Object} rdbDataByDept - 部署別RDBデータ
   * @returns {Object} 統計情報
   */
  createRetrievalStatistics(ganttDataByDept, rdbDataByDept) {
    const ganttDepartments = Object.keys(ganttDataByDept);
    const rdbDepartments = Object.keys(rdbDataByDept);
    const commonDepartments = ganttDepartments.filter(dept => rdbDepartments.includes(dept));
    const ganttOnlyDepartments = ganttDepartments.filter(dept => !rdbDepartments.includes(dept));
    const rdbOnlyDepartments = rdbDepartments.filter(dept => !ganttDepartments.includes(dept));

    // ガントシフト数の計算
    const ganttShiftCounts = Object.entries(ganttDataByDept).reduce((acc, [dept, data]) => {
      acc[dept] = this.countGanttShifts(data);
      return acc;
    }, {});

    // RDBシフト数の計算
    const rdbShiftCounts = Object.entries(rdbDataByDept).reduce((acc, [dept, data]) => {
      acc[dept] = data.length;
      return acc;
    }, {});

    return {
      ganttDepartments,
      rdbDepartments,
      commonDepartments,
      ganttOnlyDepartments,
      rdbOnlyDepartments,
      ganttShiftCounts,
      rdbShiftCounts,
      totalGanttShifts: Object.values(ganttShiftCounts).reduce((sum, count) => sum + count, 0),
      totalRdbShifts: Object.values(rdbShiftCounts).reduce((sum, count) => sum + count, 0)
    };
  }

  /**
   * ガントデータからシフト個数をカウント
   * @param {GanttChartData} ganttData - ガントチャートデータ
   * @returns {number} シフト個数
   */
  countGanttShifts(ganttData) {
    if (!ganttData || ganttData.isEmpty()) {
      return 0;
    }

    const { values, backgrounds } = ganttData;
    let shiftCount = 0;

    // firstDataの位置を取得（0ベースのインデックス）
    const firstDataRow = this.rowIndexes.firstData;
    const firstDataCol = this.ganttIndexes.firstData;

    // シフトデータ部分を取得
    const shiftValues = values.slice(firstDataRow).map((row) => row.slice(firstDataCol));
    const shiftBgs = backgrounds.slice(firstDataRow).map((row) => row.slice(firstDataCol));

    // 値または背景色でシフトデータを判定してカウント
    shiftValues.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const bgColor = shiftBgs[rowIndex] ? shiftBgs[rowIndex][colIndex] : "#ffffff";
        const hasValue = cell !== "" && cell !== null && cell !== undefined;
        const hasNonWhiteBg = this.isNonWhiteBackground(bgColor);

        // 値が入っているか、背景色が白以外の場合はシフトデータとしてカウント
        if (hasValue || hasNonWhiteBg) {
          shiftCount++;
        }
      });
    });

    return shiftCount;
  }

  /**
   * 背景色が白以外かどうかを判定
   * @param {string} bgColor - 背景色
   * @returns {boolean} 白以外の場合true
   */
  isNonWhiteBackground(bgColor) {
    if (!bgColor) return false;
    
    const normalizedColor = bgColor.toLowerCase();
    return normalizedColor !== "#ffffff" && 
           normalizedColor !== "#fff" && 
           normalizedColor !== "white" &&
           normalizedColor !== "";
  }

  /**
   * データ取得状況のサマリーを出力
   * @param {Object} statistics - 統計情報
   */
  logRetrievalSummary(statistics) {
    console.log("=== データ取得サマリー ===");
    console.log(`ガントチャート部署数: ${statistics.ganttDepartments.length}`);
    console.log(`RDB部署数: ${statistics.rdbDepartments.length}`);
    console.log(`共通部署数: ${statistics.commonDepartments.length}`);
    console.log(`ガントのみ部署数: ${statistics.ganttOnlyDepartments.length}`);
    console.log(`RDBのみ部署数: ${statistics.rdbOnlyDepartments.length}`);
    console.log(`総ガントシフト数: ${statistics.totalGanttShifts}`);
    console.log(`総RDBシフト数: ${statistics.totalRdbShifts}`);

    if (statistics.ganttOnlyDepartments.length > 0) {
      console.log(`ガントのみ部署: ${statistics.ganttOnlyDepartments.join(", ")}`);
    }

    if (statistics.rdbOnlyDepartments.length > 0) {
      console.log(`RDBのみ部署: ${statistics.rdbOnlyDepartments.join(", ")}`);
    }
  }

  /**
   * プロパティサービスからガントチャートスプレッドシートのURLを取得
   * @returns {string} スプレッドシートURL
   * @throws {ConfigurationError} URLが設定されていない場合
   */
  getGanttSpreadsheetUrl() {
    const url = PropertiesService.getScriptProperties().getProperty("GANTT_SS");
    
    if (!url) {
      throw new ConfigurationError(
        "ガントチャートスプレッドシートのURLが設定されていません。スクリプトプロパティで「GANTT_SS」を設定してください。",
        "GANTT_SS"
      );
    }
    
    return url;
  }

  /**
   * URLからスプレッドシートを開く
   * @param {string} url - スプレッドシートURL
   * @returns {Spreadsheet} スプレッドシートオブジェクト
   * @throws {DataProcessingError} スプレッドシートを開けない場合
   */
  openSpreadsheetByUrl(url) {
    try {
      return SpreadsheetApp.openByUrl(url);
    } catch (error) {
      throw new DataProcessingError(
        `スプレッドシートを開けませんでした: ${error.message}`,
        { url }
      );
    }
  }

  /**
   * シート名からシートを取得
   * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
   * @param {string} sheetName - シート名
   * @returns {Sheet} シートオブジェクト
   * @throws {DataProcessingError} シートが見つからない場合
   */
  getSheetByName(spreadsheet, sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new DataProcessingError(
        `シート「${sheetName}」が見つかりません`,
        { sheetName, spreadsheetName: spreadsheet.getName() }
      );
    }
    
    return sheet;
  }

  /**
   * データ取得処理の実行
   * @param {ShiftDataRequest} request - データ取得リクエスト
   * @returns {Object} 取得結果
   */
  retrieveAllData(request) {
    try {
      // ガントチャートデータの取得
      NotificationService.showProgress("Ganttデータの取得とグループ化を開始します...");
      const ganttDataGroupedByDept = this.getAllGanttSheetDataAndGroupByName(request.ganttSs);

      // RDBデータの取得
      NotificationService.showProgress("RDBデータの取得を開始します...");
      const rdbData = this.getRdbData(request.rdbSheet);

      // RDBデータの部署別グループ化
      NotificationService.showProgress("RDBデータの部署ごとのグループ化を開始します...");
      const rdbDataGroupedByDept = this.groupByDepartment(rdbData, this.rdbIndexes.dept);

      // 統計情報の作成
      const statistics = this.createRetrievalStatistics(ganttDataGroupedByDept, rdbDataGroupedByDept);
      this.logRetrievalSummary(statistics);

      return {
        ganttDataGroupedByDept,
        rdbData,
        rdbDataGroupedByDept,
        statistics
      };

    } catch (error) {
      console.error("データ取得処理中にエラーが発生しました:", error);
      throw error;
    }
  }

  /**
   * データの整合性チェック
   * @param {Object} retrievedData - 取得データ
   * @returns {ValidationResult} チェック結果
   */
  validateRetrievedData(retrievedData) {
    const result = new ValidationResult();
    const { ganttDataGroupedByDept, rdbData, statistics } = retrievedData;

    // ガントデータのチェック
    if (!ganttDataGroupedByDept || Object.keys(ganttDataGroupedByDept).length === 0) {
      result.addError("ガントチャートデータが取得できませんでした");
    }

    // RDBデータのチェック
    if (!rdbData || rdbData.length === 0) {
      result.addWarning("RDBデータが取得できませんでした");
    }

    // 処理可能な部署のチェック
    if (statistics.commonDepartments.length === 0) {
      result.addError("処理可能な部署が見つかりません");
    }

    // RDBのみの部署がある場合の警告
    if (statistics.rdbOnlyDepartments.length > 0) {
      result.addWarning(`RDBのみに存在する部署があります: ${statistics.rdbOnlyDepartments.join(", ")}`);
    }

    if (result.errors.length === 0) {
      result.markSuccess();
    }

    return result;
  }
}

/**
 * データキャッシュマネージャー
 * 
 * 大きなデータセットのキャッシュを管理し、
 * 同じデータへの重複アクセスを避ける
 */
class DataCacheManager {
  constructor() {
    this.cache = new Map();
    this.maxCacheSize = 50; // 最大キャッシュ数
    this.cacheStats = {
      hits: 0,
      misses: 0,
      evictions: 0
    };
  }

  /**
   * キャッシュからデータを取得
   * @param {string} key - キャッシュキー
   * @returns {*} キャッシュされたデータ、または null
   */
  get(key) {
    if (this.cache.has(key)) {
      this.cacheStats.hits++;
      const cached = this.cache.get(key);
      // アクセス時刻を更新（LRU用）
      cached.lastAccessed = new Date();
      return cached.data;
    }
    
    this.cacheStats.misses++;
    return null;
  }

  /**
   * データをキャッシュに保存
   * @param {string} key - キャッシュキー
   * @param {*} data - 保存するデータ
   */
  set(key, data) {
    // キャッシュサイズ制限チェック
    if (this.cache.size >= this.maxCacheSize) {
      this.evictLeastRecentlyUsed();
    }

    this.cache.set(key, {
      data: data,
      createdAt: new Date(),
      lastAccessed: new Date(),
      size: this.estimateDataSize(data)
    });
  }

  /**
   * 最近使用されていないデータを削除
   */
  evictLeastRecentlyUsed() {
    let oldestKey = null;
    let oldestTime = new Date();

    for (const [key, cached] of this.cache.entries()) {
      if (cached.lastAccessed < oldestTime) {
        oldestTime = cached.lastAccessed;
        oldestKey = key;
      }
    }

    if (oldestKey) {
      this.cache.delete(oldestKey);
      this.cacheStats.evictions++;
    }
  }

  /**
   * データサイズを推定
   * @param {*} data - データ
   * @returns {number} 推定サイズ（バイト）
   */
  estimateDataSize(data) {
    try {
      return JSON.stringify(data).length * 2; // 大まかな推定
    } catch (error) {
      return 1000; // デフォルト値
    }
  }

  /**
   * キャッシュ統計の取得
   * @returns {Object} キャッシュ統計
   */
  getStats() {
    const totalRequests = this.cacheStats.hits + this.cacheStats.misses;
    const hitRate = totalRequests > 0 ? (this.cacheStats.hits / totalRequests) * 100 : 0;

    return {
      ...this.cacheStats,
      totalRequests,
      hitRate: Math.round(hitRate * 100) / 100,
      cacheSize: this.cache.size,
      maxCacheSize: this.maxCacheSize
    };
  }

  /**
   * キャッシュをクリア
   */
  clear() {
    this.cache.clear();
    this.cacheStats = {
      hits: 0,
      misses: 0,
      evictions: 0
    };
  }
}

/**
 * データサイズ計算ユーティリティ
 */
class DataSizeUtils {
  /**
   * 2次元配列のサイズを計算
   * @param {Array} data - 2次元配列
   * @returns {Object} サイズ情報
   */
  static calculate2DArraySize(data) {
    if (!Array.isArray(data)) {
      return { rows: 0, columns: 0, cells: 0 };
    }

    const rows = data.length;
    const columns = rows > 0 ? Math.max(...data.map(row => Array.isArray(row) ? row.length : 0)) : 0;
    const cells = data.reduce((total, row) => total + (Array.isArray(row) ? row.length : 0), 0);

    return { rows, columns, cells };
  }

  /**
   * データセットのメモリ使用量を推定
   * @param {Object} dataset - データセット
   * @returns {Object} メモリ使用量情報
   */
  static estimateMemoryUsage(dataset) {
    const sizes = {};
    let totalSize = 0;

    for (const [key, data] of Object.entries(dataset)) {
      let size = 0;
      
      if (Array.isArray(data)) {
        const arrayInfo = this.calculate2DArraySize(data);
        size = arrayInfo.cells * 50; // セルあたり50バイトと推定
      } else if (typeof data === 'object') {
        try {
          size = JSON.stringify(data).length * 2;
        } catch (error) {
          size = 1000; // エラーの場合はデフォルト値
        }
      } else {
        size = String(data).length * 2;
      }

      sizes[key] = size;
      totalSize += size;
    }

    return {
      individual: sizes,
      total: totalSize,
      totalMB: Math.round((totalSize / (1024 * 1024)) * 100) / 100
    };
  }
} 