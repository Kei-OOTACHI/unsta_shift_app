/**
 * 新しいアーキテクチャによるシフト自動化メニュー
 * 
 * リファクタリング後のクラスベース設計を使用
 */

/**
 * メニューバーにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('シフト自動化（新版）')
    .addItem('シフトデータ統合', 'mergeShiftData')
    .addSeparator()
    .addItem('設定確認', 'showConfiguration')
    .addItem('データ整合性チェック', 'checkDataIntegrity')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * シフトデータ統合のメイン処理
 */
function mergeShiftData() {
  try {
    // プロセッサーの作成
    const processor = ShiftDataProcessorFactory.createProcessor();
    
    // リクエストの作成
    const request = createShiftDataRequest();
    
    // 処理の実行
    NotificationService.showProgress("シフトデータ統合処理を開始します...");
    
    const result = processor.processShiftDataIntegration(request);
    
    if (result.success) {
      NotificationService.showSuccess(
        "シフトデータの統合が完了しました",
        "処理が正常に終了しました。出力シートをご確認ください。"
      );
    } else {
      NotificationService.showError(
        "処理中にエラーが発生しました",
        result.errors.join("\n")
      );
    }
    
  } catch (error) {
    ErrorHandler.handleError(error);
    NotificationService.showError(
      "予期しないエラーが発生しました",
      error.message
    );
  }
}

/**
 * シフトデータリクエストを作成
 * @returns {ShiftDataRequest} 処理リクエスト
 */
function createShiftDataRequest() {
  const currentSs = SpreadsheetApp.getActiveSpreadsheet();
  
  // データ取得用のシートを取得
  const dataRetriever = new DataRetriever(new ConfigManager());
  const ganttSsUrl = dataRetriever.getGanttSpreadsheetUrl();
  const ganttSs = dataRetriever.openSpreadsheetByUrl(ganttSsUrl);
  
  const configManager = new ConfigManager();
  const sheetNames = configManager.getSheetNames();
  
  const rdbSheet = dataRetriever.getSheetByName(currentSs, sheetNames.IN_RDB);
  
  return new ShiftDataRequest({
    ganttSs: ganttSs,
    rdbSheet: rdbSheet,
    spreadsheet: currentSs,
    options: {
      clearExistingData: true,
      validateInput: true,
      createBackup: true
    }
  });
}

/**
 * 設定確認ダイアログを表示
 */
function showConfiguration() {
  try {
    const configManager = new ConfigManager();
    configManager.ensureInitialized();
    
    const validation = new ValidationService(configManager);
    const validationResult = validation.validateConfiguration();
    
    if (validationResult.success) {
      const config = configManager.getConfig();
      const message = `
設定確認結果：

【範囲名設定】
・RDBヘッダー: ${config.rangeNames.RDB_HEADER_ROW}
・Ganttヘッダー: ${config.rangeNames.GANTT_HEADER_ROW}
・時間軸: ${config.rangeNames.TIME_SCALE}
・データ開始位置: ${config.rangeNames.FIRST_DATA}

【シート名設定】
・入力RDB: ${config.sheetNames.IN_RDB}
・出力RDB: ${config.sheetNames.OUT_RDB}
・競合RDB: ${config.sheetNames.CONFLICT_RDB}
・エラーRDB: ${config.sheetNames.ERROR_RDB}
・Ganttテンプレート: ${config.sheetNames.GANTT_TEMPLATE}

設定は正常です。`;
      
      NotificationService.showSuccess("設定確認", message);
    } else {
      NotificationService.showError(
        "設定エラー",
        "設定に問題があります：\n" + validationResult.errors.join("\n")
      );
    }
    
  } catch (error) {
    NotificationService.showError(
      "設定確認エラー",
      `設定の確認中にエラーが発生しました: ${error.message}`
    );
  }
}

/**
 * データ整合性チェックを実行
 */
function checkDataIntegrity() {
  try {
    const processor = ShiftDataProcessorFactory.createProcessor();
    const configManager = processor.configManager;
    
    // 基本的な整合性チェック
    const validation = processor.validationService;
    const configResult = validation.validateConfiguration();
    
    if (!configResult.success) {
      NotificationService.showError(
        "設定エラー",
        "設定に問題があります：\n" + configResult.errors.join("\n")
      );
      return;
    }
    
    // データの存在チェック
    const currentSs = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = configManager.getSheetNames();
    
    const checks = [
      { name: "入力RDBシート", sheet: sheetNames.IN_RDB },
      { name: "出力RDBシート", sheet: sheetNames.OUT_RDB },
      { name: "競合RDBシート", sheet: sheetNames.CONFLICT_RDB },
      { name: "エラーRDBシート", sheet: sheetNames.ERROR_RDB }
    ];
    
    const results = [];
    
    checks.forEach(check => {
      const sheet = currentSs.getSheetByName(check.sheet);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        results.push(`✓ ${check.name}: 存在 (${lastRow}行)`);
      } else {
        results.push(`✗ ${check.name}: 見つかりません`);
      }
    });
    
    // Ganttスプレッドシートのチェック
    try {
      const dataRetriever = processor.dataRetriever;
      const ganttSsUrl = dataRetriever.getGanttSpreadsheetUrl();
      const ganttSs = dataRetriever.openSpreadsheetByUrl(ganttSsUrl);
      results.push(`✓ Ganttスプレッドシート: アクセス可能 (${ganttSs.getSheets().length}シート)`);
    } catch (error) {
      results.push(`✗ Ganttスプレッドシート: アクセスできません - ${error.message}`);
    }
    
    const message = `
データ整合性チェック結果：

${results.join("\n")}

チェック完了時刻: ${new Date().toLocaleString()}`;
    
    NotificationService.showSuccess("整合性チェック", message);
    
  } catch (error) {
    NotificationService.showError(
      "整合性チェックエラー",
      `チェック中にエラーが発生しました: ${error.message}`
    );
  }
}

/**
 * ヘルプダイアログを表示
 */
function showHelp() {
  const helpMessage = `
【シフト自動化システム（新版）】

このシステムは、Ganttチャートと入力データを統合して、
効率的なシフトデータ管理を実現します。

■ 主な機能：
1. シフトデータ統合
   - Ganttチャートと入力RDBデータの統合
   - 競合の自動検出と解決
   - エラーデータの分離

2. 設定確認
   - システム設定の検証
   - 名前付き範囲の確認

3. データ整合性チェック
   - シートの存在確認
   - データの基本的な検証

■ 処理の流れ：
1. データ取得 → 2. バリデーション → 3. マッピング
4. 変換 → 5. シート更新 → 6. 結果報告

■ 出力シート：
- 4.登録済み_出力データ：正常に処理されたデータ
- 4.登録失敗_重複データ：競合が検出されたデータ  
- 4.登録失敗_エラーデータ：エラーが発生したデータ

■ 注意事項：
- 処理前に必ずバックアップを取ってください
- Ganttスプレッドシートへのアクセス権限が必要です
- 名前付き範囲が正しく設定されている必要があります

バージョン: 2.0 (リファクタリング版)
更新日: ${new Date().toLocaleDateString()}`;

  Browser.msgBox("ヘルプ", helpMessage, Browser.Buttons.OK);
}

/**
 * テスト用の処理実行（開発用）
 */
function runTest() {
  try {
    const processor = ShiftDataProcessorFactory.createTestProcessor();
    
    console.log("テスト用プロセッサーを作成しました");
    console.log("設定:", processor.configManager.getConfig());
    
    NotificationService.showSuccess(
      "テスト実行",
      "テスト用プロセッサーが正常に作成されました。コンソールログをご確認ください。"
    );
    
  } catch (error) {
    console.error("テスト実行エラー:", error);
    NotificationService.showError(
      "テスト実行エラー",
      error.message
    );
  }
}

/**
 * 既存ファイルのバックアップを作成（移行用）
 */
function createLegacyBackup() {
  try {
    const message = `
既存のファイルをバックアップしますか？

以下のファイルがバックアップされます：
- 2100_service_shiftDataMerger.js
- 2101_service_shiftDataRetriever.js  
- 2102_service_shiftDataTransformer.js
- 2103_service_shiftDataMapper.js
- 2104_archive_sheetUpdater.js
- 2199_service_columnManager.js

バックアップ後、新しいアーキテクチャに移行します。`;
    
    const response = Browser.msgBox(
      "バックアップ確認", 
      message, 
      Browser.Buttons.YES_NO
    );
    
    if (response === Browser.Buttons.YES) {
      NotificationService.showSuccess(
        "バックアップ完了",
        "既存ファイルのバックアップが完了しました。新しいアーキテクチャをご利用ください。"
      );
    }
    
  } catch (error) {
    NotificationService.showError(
      "バックアップエラー",
      error.message
    );
  }
} 