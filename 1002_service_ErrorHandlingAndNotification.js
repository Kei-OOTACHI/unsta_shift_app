/**
 * エラー処理と通知サービス
 * 
 * 責任:
 * - エラーの分類と処理
 * - ユーザー通知
 * - ログ出力
 * - 復旧支援
 */

/**
 * 通知サービス
 */
class NotificationService {
  /**
   * 成功メッセージを表示
   * @param {string} message - 表示するメッセージ
   * @param {string} title - タイトル（オプション）
   */
  static showSuccess(message, title = "成功") {
    Browser.msgBox(title, message, Browser.Buttons.OK);
    SpreadsheetApp.getActive().toast(message, "完了");
    console.log(`[SUCCESS] ${message}`);
  }

  /**
   * 警告メッセージを表示
   * @param {string} message - 表示するメッセージ
   * @param {string} title - タイトル（オプション）
   */
  static showWarning(message, title = "警告") {
    Browser.msgBox(title, message, Browser.Buttons.OK);
    SpreadsheetApp.getActive().toast(message, "警告");
    console.warn(`[WARNING] ${message}`);
  }

  /**
   * エラーメッセージを表示
   * @param {string} message - 表示するメッセージ
   * @param {string} title - タイトル（オプション）
   */
  static showError(message, title = "エラー") {
    Browser.msgBox(title, message, Browser.Buttons.OK);
    SpreadsheetApp.getActive().toast(message, "エラー");
    console.error(`[ERROR] ${message}`);
  }

  /**
   * バリデーションエラーメッセージを表示
   * @param {string} message - 表示するメッセージ
   * @param {string} field - エラーが発生したフィールド
   */
  static showValidationError(message, field = null) {
    const fullMessage = field ? `${field}: ${message}` : message;
    this.showError(fullMessage, "バリデーションエラー");
  }

  /**
   * 処理エラーメッセージを表示
   * @param {string} message - 表示するメッセージ
   */
  static showProcessingError(message) {
    this.showError(message, "処理エラー");
  }

  /**
   * 設定エラーメッセージを表示
   * @param {string} message - 表示するメッセージ
   * @param {string} configKey - エラーが発生した設定キー
   */
  static showConfigurationError(message, configKey = null) {
    const fullMessage = configKey ? `設定「${configKey}」: ${message}` : message;
    this.showError(fullMessage, "設定エラー");
  }

  /**
   * 一般的なエラーメッセージを表示
   * @param {string} message - 表示するメッセージ
   */
  static showGenericError(message) {
    this.showError(message, "予期しないエラー");
  }

  /**
   * 進捗状況を表示
   * @param {string} message - 表示するメッセージ
   * @param {string} title - タイトル（オプション）
   */
  static showProgress(message, title = "処理状況") {
    SpreadsheetApp.getActive().toast(message, title);
    console.log(`[PROGRESS] ${message}`);
  }

  /**
   * 詳細な進捗状況を表示
   * @param {string} message - 表示するメッセージ
   * @param {number} current - 現在の進捗
   * @param {number} total - 総数
   */
  static showDetailedProgress(message, current, total) {
    const progressMessage = `${message} (${current}/${total})`;
    this.showProgress(progressMessage);
  }

  /**
   * 整合性チェック結果を表示
   * @param {Object} result - チェック結果
   */
  static showIntegrityCheckResult(result) {
    const { hasErrors, message, details } = result;
    
    if (hasErrors) {
      const alertMessage = `${message}\n\n【詳細】\n${details.join('\n')}\n\nデータの整合性に問題があります。\nエンジニアに連絡して確認を依頼してください。`;
      Browser.msgBox("データ整合性エラー", alertMessage, Browser.Buttons.OK);
    } else {
      const successMessage = `${message}\n\n【整合性チェック結果】\n✅ すべての整合性チェックが正常に完了しました`;
      Browser.msgBox("データ整合性チェック完了", successMessage, Browser.Buttons.OK);
    }
  }
}

/**
 * エラーハンドラー
 */
class ErrorHandler {
  /**
   * エラーを処理する
   * @param {Error} error - 処理するエラー
   * @param {Object} context - エラーのコンテキスト情報
   */
  static handle(error, context = {}) {
    console.error("Error occurred:", error);
    console.error("Context:", context);
    console.error("Stack trace:", error.stack);

    // エラータイプによる分岐処理
    if (error instanceof ValidationError) {
      this.handleValidationError(error, context);
    } else if (error instanceof DataProcessingError) {
      this.handleDataProcessingError(error, context);
    } else if (error instanceof ConfigurationError) {
      this.handleConfigurationError(error, context);
    } else if (error instanceof SheetUpdateError) {
      this.handleSheetUpdateError(error, context);
    } else {
      this.handleGenericError(error, context);
    }
  }

  /**
   * バリデーションエラーを処理する
   * @param {ValidationError} error - バリデーションエラー
   * @param {Object} context - コンテキスト情報
   */
  static handleValidationError(error, context) {
    NotificationService.showValidationError(error.message, error.field);
    
    // 必要に応じて追加の処理を実装
    if (context.showInstructions) {
      this.showValidationInstructions(error.field);
    }
  }

  /**
   * データ処理エラーを処理する
   * @param {DataProcessingError} error - データ処理エラー
   * @param {Object} context - コンテキスト情報
   */
  static handleDataProcessingError(error, context) {
    NotificationService.showProcessingError(error.message);
    
    // データ損失の可能性がある場合は復旧支援を提供
    if (context.requiresRecovery) {
      this.showRecoveryOptions(context);
    }
  }

  /**
   * 設定エラーを処理する
   * @param {ConfigurationError} error - 設定エラー
   * @param {Object} context - コンテキスト情報
   */
  static handleConfigurationError(error, context) {
    NotificationService.showConfigurationError(error.message, error.configKey);
    
    // 設定修正のガイダンスを提供
    if (error.configKey) {
      this.showConfigurationGuidance(error.configKey);
    }
  }

  /**
   * シート更新エラーを処理する
   * @param {SheetUpdateError} error - シート更新エラー
   * @param {Object} context - コンテキスト情報
   */
  static handleSheetUpdateError(error, context) {
    NotificationService.showError(error.message, "シート更新エラー");
    
    // バックアップからの復旧支援を提供
    if (context.startTime) {
      this.showRestorePrompt(
        context.failedSheets || [error.sheetName],
        context.targetDescription || "スプレッドシート",
        context.startTime,
        error
      );
    }
  }

  /**
   * 一般的なエラーを処理する
   * @param {Error} error - 一般的なエラー
   * @param {Object} context - コンテキスト情報
   */
  static handleGenericError(error, context) {
    NotificationService.showGenericError(error.message);
    
    // 予期しないエラーの場合は詳細情報を提供
    console.error("予期しないエラーが発生しました:", {
      error: error,
      context: context,
      timestamp: new Date().toISOString()
    });
  }

  /**
   * バリデーション修正の指示を表示
   * @param {string} field - エラーが発生したフィールド
   */
  static showValidationInstructions(field) {
    const instructions = this.getValidationInstructions(field);
    if (instructions) {
      NotificationService.showWarning(instructions, "修正方法");
    }
  }

  /**
   * バリデーション修正の指示を取得
   * @param {string} field - フィールド名
   * @returns {string|null} 修正指示
   */
  static getValidationInstructions(field) {
    const instructionMap = {
      'memberDateId': 'memberDateIdが空の場合は、該当する行のデータを確認してください。',
      'startTime': 'startTimeは h:mm 形式で入力してください（例: 9:00）。',
      'endTime': 'endTimeは h:mm 形式で入力してください（例: 17:00）。',
      'dept': '部署名が空の場合は、正しい部署名を入力してください。',
      'timeRange': 'startTimeがendTime以降になっています。正しい時間範囲を入力してください。'
    };
    
    return instructionMap[field] || null;
  }

  /**
   * 設定修正のガイダンスを表示
   * @param {string} configKey - 設定キー
   */
  static showConfigurationGuidance(configKey) {
    const guidance = this.getConfigurationGuidance(configKey);
    if (guidance) {
      NotificationService.showWarning(guidance, "設定修正ガイダンス");
    }
  }

  /**
   * 設定修正のガイダンスを取得
   * @param {string} configKey - 設定キー
   * @returns {string|null} ガイダンス
   */
  static getConfigurationGuidance(configKey) {
    const guidanceMap = {
      'namedRange': 'メニューバーの「データ」>「名前付き範囲」から設定を確認してください。',
      'spreadsheetUrl': 'スクリプトプロパティでスプレッドシートのURLを確認してください。',
      'columnMapping': 'ヘッダー行の列名が正しく設定されているか確認してください。'
    };
    
    return guidanceMap[configKey] || null;
  }

  /**
   * 復旧支援のオプションを表示
   * @param {Object} context - コンテキスト情報
   */
  static showRecoveryOptions(context) {
    const message = 
      "データ処理中にエラーが発生しました。\n\n" +
      "【復旧オプション】\n" +
      "1. 処理を再試行する\n" +
      "2. 部分的なデータを確認する\n" +
      "3. バックアップから復元する\n\n" +
      "どのオプションを選択しますか？";
    
    // 実際の復旧処理はここで実装
    NotificationService.showWarning(message, "復旧支援");
  }

  /**
   * 復元案内を表示
   * @param {Array} failedSheets - 失敗したシートの一覧
   * @param {string} targetDescription - 対象の説明
   * @param {Date} startTime - 処理開始時刻
   * @param {Error} error - 発生したエラー
   */
  static showRestorePrompt(failedSheets, targetDescription, startTime, error) {
    const formattedStartTime = Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    const message = 
      "■ データ更新処理でエラーが発生しました\n" +
      "【失敗箇所】\n" +
      targetDescription + "の以下のシート:\n" +
      failedSheets.map(sheet => "・" + sheet).join("\n") + "\n\n" +
      "【エラー詳細】\n" +
      error.message + "\n\n" +
      "【復元方法】\n" +
      "処理開始時刻: " + formattedStartTime + "\n\n" +
      "以下の手順で履歴から復元してください:\n" +
      "1. 対象のスプレッドシートを開く\n" +
      "2. ファイルメニュー → 「バージョン履歴」 → 「バージョン履歴を表示」を選択\n" +
      "3. 処理開始時刻(" + formattedStartTime + ")より前の最新バージョンを選択\n" +
      "4. 「このバージョンを復元」をクリック\n" +
      "復元完了後、問題を修正してから再度処理を実行してください。";

    Browser.msgBox("データ更新エラー - 履歴からの復元が必要", message, Browser.Buttons.OK);

    // ログにも出力
    console.error("=== データ更新処理エラー ===");
    console.error(`失敗箇所: ${targetDescription}`);
    console.error(`失敗シート: ${failedSheets.join(", ")}`);
    console.error(`処理開始時刻: ${formattedStartTime}`);
    console.error(`エラー: ${error.message}`);
    console.error("Stack trace:", error.stack);
  }

  /**
   * エラー統計情報を取得
   * @param {Array} errors - エラーの配列
   * @returns {Object} エラー統計
   */
  static getErrorStatistics(errors) {
    const statistics = {
      total: errors.length,
      byType: {},
      byField: {},
      timeline: []
    };

    errors.forEach(error => {
      // エラータイプ別の集計
      const errorType = error.constructor.name;
      statistics.byType[errorType] = (statistics.byType[errorType] || 0) + 1;

      // フィールド別の集計（ValidationErrorの場合）
      if (error instanceof ValidationError && error.field) {
        statistics.byField[error.field] = (statistics.byField[error.field] || 0) + 1;
      }

      // タイムライン
      statistics.timeline.push({
        timestamp: new Date(),
        type: errorType,
        message: error.message
      });
    });

    return statistics;
  }

  /**
   * エラーレポートを生成
   * @param {Array} errors - エラーの配列
   * @returns {string} エラーレポート
   */
  static generateErrorReport(errors) {
    const statistics = this.getErrorStatistics(errors);
    
    let report = "=== エラーレポート ===\n\n";
    report += `総エラー数: ${statistics.total}\n\n`;
    
    if (Object.keys(statistics.byType).length > 0) {
      report += "【エラータイプ別】\n";
      Object.entries(statistics.byType).forEach(([type, count]) => {
        report += `- ${type}: ${count}件\n`;
      });
      report += "\n";
    }
    
    if (Object.keys(statistics.byField).length > 0) {
      report += "【フィールド別】\n";
      Object.entries(statistics.byField).forEach(([field, count]) => {
        report += `- ${field}: ${count}件\n`;
      });
      report += "\n";
    }
    
    report += "【詳細】\n";
    errors.forEach((error, index) => {
      report += `${index + 1}. ${error.constructor.name}: ${error.message}\n`;
    });
    
    return report;
  }
}

/**
 * パフォーマンス監視サービス
 */
class PerformanceMonitor {
  constructor() {
    this.startTime = null;
    this.checkpoints = [];
  }

  /**
   * 監視開始
   * @param {string} operationName - 操作名
   */
  start(operationName) {
    this.startTime = new Date();
    this.operationName = operationName;
    this.checkpoints = [];
    console.log(`[PERFORMANCE] Started: ${operationName}`);
  }

  /**
   * チェックポイントを追加
   * @param {string} checkpointName - チェックポイント名
   */
  checkpoint(checkpointName) {
    if (!this.startTime) {
      console.warn("Performance monitor not started");
      return;
    }

    const now = new Date();
    const elapsed = now - this.startTime;
    const checkpoint = {
      name: checkpointName,
      timestamp: now,
      elapsed: elapsed
    };

    this.checkpoints.push(checkpoint);
    console.log(`[PERFORMANCE] Checkpoint: ${checkpointName} (${elapsed}ms)`);
  }

  /**
   * 監視終了
   * @returns {Object} パフォーマンス結果
   */
  end() {
    if (!this.startTime) {
      console.warn("Performance monitor not started");
      return null;
    }

    const endTime = new Date();
    const totalElapsed = endTime - this.startTime;
    
    const result = {
      operationName: this.operationName,
      totalElapsed: totalElapsed,
      checkpoints: this.checkpoints,
      averageCheckpointTime: this.checkpoints.length > 0 ? 
        this.checkpoints.reduce((sum, cp) => sum + cp.elapsed, 0) / this.checkpoints.length : 0
    };

    console.log(`[PERFORMANCE] Completed: ${this.operationName} (${totalElapsed}ms)`);
    return result;
  }
}

/**
 * デバッグユーティリティ
 */
class DebugUtils {
  /**
   * オブジェクトを詳細にログ出力
   * @param {*} obj - 出力するオブジェクト
   * @param {string} label - ラベル
   */
  static logObject(obj, label = "Object") {
    console.log(`[DEBUG] ${label}:`, JSON.stringify(obj, null, 2));
  }

  /**
   * 配列の統計情報を出力
   * @param {Array} array - 配列
   * @param {string} label - ラベル
   */
  static logArrayStats(array, label = "Array") {
    console.log(`[DEBUG] ${label} Stats:`, {
      length: array.length,
      firstItem: array[0],
      lastItem: array[array.length - 1],
      sample: array.slice(0, 3)
    });
  }

  /**
   * 関数の実行時間を測定
   * @param {Function} fn - 実行する関数
   * @param {string} label - ラベル
   * @returns {*} 関数の戻り値
   */
  static measureTime(fn, label = "Function") {
    const start = new Date();
    const result = fn();
    const elapsed = new Date() - start;
    console.log(`[DEBUG] ${label} execution time: ${elapsed}ms`);
    return result;
  }
} 