/**
 * シフトデータプロセッサー - メインオーケストレーター
 * 
 * 責任:
 * - 全体的なフローの制御
 * - サービス間の連携
 * - エラーハンドリング
 * - 進捗管理
 */
class ShiftDataProcessor {
  constructor() {
    this.configManager = new ConfigManager();
    this.validationService = new ValidationService(this.configManager);
    this.dataRetriever = new DataRetriever(this.configManager);
    this.dataTransformer = new DataTransformer(this.configManager);
    this.dataMapper = new DataMapper(this.configManager);
    this.sheetUpdater = new SheetUpdater(this.configManager);
    this.errorHandler = ErrorHandler;
    this.notificationService = NotificationService;
  }

  /**
   * シフトデータ統合処理の実行
   * @param {ShiftDataRequest} request - 処理リクエスト
   * @returns {ShiftDataResult} 処理結果
   */
  processShiftDataIntegration(request) {
    const result = new ShiftDataResult();
    
    try {
      // 設定の初期化と検証
      this.initializeAndValidateConfiguration();
      
      // データ取得
      const retrievedData = this.retrieveData(request);
      
      // データ検証
      this.validateData(retrievedData);
      
      // データマッピングと競合解決
      const mappedData = this.mapAndResolveConflicts(retrievedData);
      
      // データ変換
      const transformedData = this.transformData(mappedData);
      
      // シート更新
      const updateResults = this.updateSheets(transformedData, request);
      
      // 結果の検証と報告
      this.validateAndReportResults(updateResults);
      
      result.markSuccess();
      result.data = updateResults;
      
      this.notificationService.showSuccess("シフトデータの統合処理が完了しました");
      
    } catch (error) {
      this.errorHandler.handleError(error, result);
      this.notificationService.showError("シフトデータの統合処理中にエラーが発生しました", error.message);
    }
    
    return result;
  }

  /**
   * 設定の初期化と検証
   */
  initializeAndValidateConfiguration() {
    this.notificationService.showProgress("設定の初期化中...");
    
    this.configManager.ensureInitialized();
    
    const validationResult = this.validationService.validateConfiguration();
    if (!validationResult.success) {
      throw new ConfigurationError("設定の検証に失敗しました", validationResult.errors);
    }
  }

  /**
   * データ取得フェーズ
   * @param {ShiftDataRequest} request - 処理リクエスト
   * @returns {Object} 取得されたデータ
   */
  retrieveData(request) {
    this.notificationService.showProgress("データ取得を開始します...");
    
    const retrievedData = this.dataRetriever.retrieveAllData(request);
    
    const validationResult = this.dataRetriever.validateRetrievedData(retrievedData);
    if (!validationResult.success) {
      throw new DataProcessingError("データ取得の検証に失敗しました", validationResult.errors);
    }
    
    return retrievedData;
  }

  /**
   * データ検証フェーズ
   * @param {Object} retrievedData - 取得されたデータ
   */
  validateData(retrievedData) {
    this.notificationService.showProgress("データ検証を開始します...");
    
    // RDBデータの検証
    const rdbValidation = this.validationService.validateAndSeparateRdbData(retrievedData.rdbData);
    if (!rdbValidation.success) {
      throw new DataProcessingError("RDBデータの検証に失敗しました", rdbValidation.errors);
    }
    
    // 部署データの検証
    const ganttDepts = Object.keys(retrievedData.ganttDataGroupedByDept);
    const rdbDepts = Object.keys(retrievedData.rdbDataGroupedByDept);
    const deptValidation = this.validationService.validateDepartments(ganttDepts, rdbDepts);
    
    if (!deptValidation.success && deptValidation.validData.length === 0) {
      throw new DataProcessingError("処理可能な部署が見つかりません", deptValidation.errors);
    }
  }

  /**
   * データマッピングと競合解決フェーズ
   * @param {Object} retrievedData - 取得されたデータ
   * @returns {Object} マッピングされたデータ
   */
  mapAndResolveConflicts(retrievedData) {
    this.notificationService.showProgress("データマッピングと競合解決を開始します...");
    
    const ganttDepts = Object.keys(retrievedData.ganttDataGroupedByDept);
    const rdbDepts = Object.keys(retrievedData.rdbDataGroupedByDept);
    const deptList = ganttDepts.filter(dept => rdbDepts.includes(dept));
    
    const mappedData = this.dataMapper.executeMapping(
      retrievedData.ganttDataGroupedByDept,
      retrievedData.rdbDataGroupedByDept,
      deptList
    );
    
    const validationResult = this.dataMapper.validateMappingResult(mappedData);
    if (!validationResult.success) {
      throw new DataProcessingError("データマッピングの検証に失敗しました", validationResult.errors);
    }
    
    return mappedData;
  }

  /**
   * データ変換フェーズ
   * @param {Object} mappedData - マッピングされたデータ
   * @returns {Object} 変換されたデータ
   */
  transformData(mappedData) {
    this.notificationService.showProgress("データ変換を開始します...");
    
    // ガントチャートデータの分割（最初のガントデータを使用）
    const firstGanttSheet = Object.values(mappedData.ganttDataByDept || {})[0];
    const splitGanttData = this.dataTransformer.splitGanttData(
      firstGanttSheet.values,
      firstGanttSheet.backgrounds
    );
    
    // 変換の実行
    const transformedData = this.dataTransformer.transformData(
      splitGanttData,
      mappedData.validShiftsMap,
      mappedData.conflictShifts
    );
    
    // ガントデータの統合
    const mergedGanttData = this.dataTransformer.mergeGanttData(
      splitGanttData.ganttHeaderValues,
      transformedData.ganttValues,
      splitGanttData.ganttHeaderBgs,
      transformedData.ganttBgs,
      splitGanttData.firstDataColOffset,
      splitGanttData.firstDataRowOffset
    );
    
    return {
      ...transformedData,
      mergedGanttData
    };
  }

  /**
   * シート更新フェーズ
   * @param {Object} transformedData - 変換されたデータ
   * @param {ShiftDataRequest} request - 処理リクエスト
   * @returns {Object} 更新結果
   */
  updateSheets(transformedData, request) {
    this.notificationService.showProgress("シート更新を開始します...");
    
    const updateResults = [];
    
    try {
      // ガントチャートシートの更新
      if (transformedData.mergedGanttData) {
        const ganttSheet = this.dataRetriever.getSheetByName(
          request.ganttSs,
          this.configManager.getSheetNames().GANTT_TEMPLATE
        );
        
        const ganttResult = this.sheetUpdater.updateGanttSheet(
          ganttSheet,
          transformedData.mergedGanttData.values,
          transformedData.mergedGanttData.backgrounds
        );
        updateResults.push(ganttResult);
      }
      
      // RDBシートの更新
      const rdbSheetUpdates = this.prepareRdbSheetUpdates(transformedData, request);
      const rdbResults = this.sheetUpdater.updateMultipleRdbSheets(rdbSheetUpdates);
      updateResults.push(...rdbResults.updates);
      
      // 更新統計の作成
      const statistics = this.sheetUpdater.createUpdateStatistics(updateResults);
      this.sheetUpdater.logUpdateSummary(statistics);
      
      return {
        updateResults,
        statistics,
        transformedData
      };
      
    } catch (error) {
      throw new DataProcessingError(`シート更新中にエラーが発生しました: ${error.message}`, { updateResults });
    }
  }

  /**
   * RDBシート更新の準備
   * @param {Object} transformedData - 変換されたデータ
   * @param {ShiftDataRequest} request - 処理リクエスト
   * @returns {Object} シート更新情報
   */
  prepareRdbSheetUpdates(transformedData, request) {
    const sheetNames = this.configManager.getSheetNames();
    const sheetUpdates = {};
    
    // ヘッダー行の準備
    const headerRow = this.configManager.getRdbHeaderRow();
    
    // 出力シートの準備
    if (transformedData.rdbData && transformedData.rdbData.length > 0) {
      sheetUpdates[sheetNames.OUT_RDB] = {
        sheet: this.dataRetriever.getSheetByName(request.spreadsheet, sheetNames.OUT_RDB),
        data: [headerRow, ...transformedData.rdbData]
      };
    }
    
    // 競合シートの準備
    if (transformedData.conflictData && transformedData.conflictData.length > 0) {
      const conflictHeaderRow = this.configManager.getConflictHeaderRow();
      sheetUpdates[sheetNames.CONFLICT_RDB] = {
        sheet: this.dataRetriever.getSheetByName(request.spreadsheet, sheetNames.CONFLICT_RDB),
        data: [conflictHeaderRow, ...transformedData.conflictData]
      };
    }
    
    // エラーシートの準備
    if (transformedData.errorData && transformedData.errorData.length > 0) {
      const errorHeaderRow = this.configManager.getErrorHeaderRow();
      sheetUpdates[sheetNames.ERROR_RDB] = {
        sheet: this.dataRetriever.getSheetByName(request.spreadsheet, sheetNames.ERROR_RDB),
        data: [errorHeaderRow, ...transformedData.errorData]
      };
    }
    
    return sheetUpdates;
  }

  /**
   * 結果の検証と報告フェーズ
   * @param {Object} updateResults - 更新結果
   */
  validateAndReportResults(updateResults) {
    this.notificationService.showProgress("結果の検証と報告を開始します...");
    
    // データ整合性チェック
    const integrityData = this.extractIntegrityData(updateResults);
    const integrityResult = this.validationService.validateDataIntegrity(integrityData);
    
    // 最終結果の報告
    this.notificationService.showSuccess(integrityResult.message);
    
    if (integrityResult.hasErrors) {
      this.notificationService.showWarning("データ整合性の問題が検出されました", integrityResult.details.join("\n"));
    }
  }

  /**
   * 整合性チェック用データの抽出
   * @param {Object} updateResults - 更新結果
   * @returns {Object} 整合性チェック用データ
   */
  extractIntegrityData(updateResults) {
    // 実装は具体的な更新結果の構造に依存
    return {
      inputGanttShiftCount: 0,
      inputRdbShiftCount: 0,
      outputGanttShiftCount: 0,
      outputMergedRdbShiftCount: 0,
      outputConflictShiftCount: 0,
      outputErrorShiftCount: 0
    };
  }
}

/**
 * 時間ユーティリティクラス
 */
class TimeUtils {
  /**
   * 時間文字列をDateオブジェクトに変換
   * @param {*} timeValue - 時間値
   * @returns {Date} Dateオブジェクト
   */
  static parseTimeToDate(timeValue) {
    if (!timeValue) {
      throw new Error("時間値が空です");
    }

    // 既にDateオブジェクトの場合
    if (timeValue instanceof Date) {
      return timeValue;
    }

    // 文字列の場合
    if (typeof timeValue === 'string') {
      const timeMatch = timeValue.match(/^(\d{1,2}):(\d{2})$/);
      if (timeMatch) {
        const hours = parseInt(timeMatch[1], 10);
        const minutes = parseInt(timeMatch[2], 10);
        const date = new Date();
        date.setHours(hours, minutes, 0, 0);
        return date;
      }
    }

    // その他の形式を試行
    const date = new Date(timeValue);
    if (isNaN(date.getTime())) {
      throw new Error(`無効な時間形式です: ${timeValue}`);
    }

    return date;
  }

  /**
   * DateオブジェクトをHH:MM形式に変換
   * @param {Date} date - Dateオブジェクト
   * @returns {string} HH:MM形式の文字列
   */
  static formatTimeToHHMM(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      throw new Error("無効なDateオブジェクトです");
    }

    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    return `${hours}:${minutes}`;
  }

  /**
   * 時間範囲の重複を判定
   * @param {Date} start1 - 範囲1の開始時間
   * @param {Date} end1 - 範囲1の終了時間
   * @param {Date} start2 - 範囲2の開始時間
   * @param {Date} end2 - 範囲2の終了時間
   * @returns {boolean} 重複する場合true
   */
  static isTimeRangeOverlapping(start1, end1, start2, end2) {
    return start1 < end2 && start2 < end1;
  }

  /**
   * 時間差を分単位で計算
   * @param {Date} startTime - 開始時間
   * @param {Date} endTime - 終了時間
   * @returns {number} 分数
   */
  static getTimeDifferenceInMinutes(startTime, endTime) {
    return Math.round((endTime.getTime() - startTime.getTime()) / (1000 * 60));
  }
}

/**
 * カスタムエラークラス群
 */
class ValidationError extends Error {
  constructor(message, field = null) {
    super(message);
    this.name = 'ValidationError';
    this.field = field;
  }
}

class ConfigurationError extends Error {
  constructor(message, configKey = null) {
    super(message);
    this.name = 'ConfigurationError';
    this.configKey = configKey;
  }
}

class DataProcessingError extends Error {
  constructor(message, context = {}) {
    super(message);
    this.name = 'DataProcessingError';
    this.context = context;
  }
}

/**
 * ファクトリークラス - オブジェクトの生成を管理
 */
class ShiftDataProcessorFactory {
  /**
   * 設定に基づいてプロセッサーを作成
   * @param {Object} config - 設定オブジェクト
   * @returns {ShiftDataProcessor} プロセッサーインスタンス
   */
  static createProcessor(config = {}) {
    const processor = new ShiftDataProcessor();
    
    // カスタム設定の適用
    if (config.customRangeNames) {
      processor.configManager.updateRangeNames(config.customRangeNames);
    }
    
    if (config.customSheetNames) {
      processor.configManager.updateSheetNames(config.customSheetNames);
    }
    
    return processor;
  }

  /**
   * 開発/テスト用プロセッサーを作成
   * @returns {ShiftDataProcessor} テスト用プロセッサー
   */
  static createTestProcessor() {
    return this.createProcessor({
      customSheetNames: {
        IN_RDB: "Test_InputData",
        OUT_RDB: "Test_OutputData",
        CONFLICT_RDB: "Test_ConflictData",
        ERROR_RDB: "Test_ErrorData"
      }
    });
  }
} 