/**
 * シート更新サービス
 * 
 * 責任:
 * - スプレッドシートへの書き込み
 * - 範囲とフォーマットの管理
 * - 更新結果の検証
 * - バックアップとロールバック
 */
class SheetUpdater {
  constructor(configManager) {
    this.config = configManager;
    this.ganttIndexes = configManager.getColumnIndexes('GANTT');
    this.rowIndexes = configManager.getColumnIndexes('ROW');
    this.sheetNames = configManager.getSheetNames();
    this.updateHistory = [];
  }

  /**
   * ガントチャートシートを更新
   * @param {Sheet} ganttSheet - ガントチャートシート
   * @param {Array} ganttValues - ガント値配列
   * @param {Array} ganttBgs - ガント背景色配列
   * @returns {Object} 更新結果
   */
  updateGanttSheet(ganttSheet, ganttValues, ganttBgs) {
    const result = {
      success: false,
      updatedCells: 0,
      sheetName: ganttSheet.getName(),
      timestamp: new Date(),
      error: null
    };

    try {
      // 更新前のバックアップを作成
      const backup = this.createBackup(ganttSheet);
      
      // データの妥当性を検証
      this.validateSheetData(ganttValues, ganttBgs);
      
      // 範囲を設定
      const range = ganttSheet.getRange(1, 1, ganttValues.length, ganttValues[0].length);
      
      // 値と背景色を設定
      range.setValues(ganttValues);
      range.setBackgrounds(ganttBgs);
      
      result.success = true;
      result.updatedCells = ganttValues.length * ganttValues[0].length;
      
      // 更新履歴に追加
      this.updateHistory.push({
        type: 'GANTT_UPDATE',
        sheetName: result.sheetName,
        result: result,
        backup: backup
      });
      
      console.log(`ガントチャートシート「${result.sheetName}」を更新しました。セル数: ${result.updatedCells}`);
      
    } catch (error) {
      result.error = error.message;
      console.error(`ガントチャートシート「${result.sheetName}」の更新中にエラーが発生しました:`, error);
      throw new DataProcessingError(`ガントチャートシートの更新に失敗しました: ${error.message}`, { sheetName: result.sheetName });
    }

    return result;
  }

  /**
   * RDBシートを更新
   * @param {Sheet} rdbSheet - RDBシート
   * @param {Array} rdbData - RDBデータ配列
   * @param {boolean} clearExisting - 既存データをクリアするかどうか
   * @returns {Object} 更新結果
   */
  updateRdbSheet(rdbSheet, rdbData, clearExisting = true) {
    const result = {
      success: false,
      updatedRows: 0,
      sheetName: rdbSheet.getName(),
      timestamp: new Date(),
      error: null
    };

    try {
      // 更新前のバックアップを作成
      const backup = this.createBackup(rdbSheet);
      
      // データの妥当性を検証
      if (!rdbData || rdbData.length === 0) {
        throw new Error("RDBデータが空です");
      }
      
      // 既存データをクリアする場合
      if (clearExisting) {
        rdbSheet.clear();
      }
      
      // データを書き込み
      const range = rdbSheet.getRange(1, 1, rdbData.length, rdbData[0].length);
      range.setValues(rdbData);
      
      result.success = true;
      result.updatedRows = rdbData.length;
      
      // 更新履歴に追加
      this.updateHistory.push({
        type: 'RDB_UPDATE',
        sheetName: result.sheetName,
        result: result,
        backup: backup
      });
      
      console.log(`RDBシート「${result.sheetName}」を更新しました。行数: ${result.updatedRows}`);
      
    } catch (error) {
      result.error = error.message;
      console.error(`RDBシート「${result.sheetName}」の更新中にエラーが発生しました:`, error);
      throw new DataProcessingError(`RDBシートの更新に失敗しました: ${error.message}`, { sheetName: result.sheetName });
    }

    return result;
  }

  /**
   * 複数のRDBシートを一括更新
   * @param {Object} sheetUpdates - シート更新情報
   * @returns {Object} 更新結果
   */
  updateMultipleRdbSheets(sheetUpdates) {
    const results = {
      success: true,
      updates: [],
      totalUpdatedRows: 0,
      errors: []
    };

    for (const [sheetName, updateData] of Object.entries(sheetUpdates)) {
      try {
        const sheet = updateData.sheet;
        const data = updateData.data;
        
        if (data && data.length > 0) {
          const updateResult = this.updateRdbSheet(sheet, data);
          results.updates.push(updateResult);
          results.totalUpdatedRows += updateResult.updatedRows;
        } else {
          // データが空の場合はシートをクリア
          sheet.clear();
          results.updates.push({
            success: true,
            updatedRows: 0,
            sheetName: sheetName,
            timestamp: new Date(),
            note: 'シートをクリアしました'
          });
        }
      } catch (error) {
        results.success = false;
        results.errors.push({
          sheetName: sheetName,
          error: error.message
        });
      }
    }

    return results;
  }

  /**
   * シートデータの妥当性を検証
   * @param {Array} values - 値配列
   * @param {Array} backgrounds - 背景色配列
   */
  validateSheetData(values, backgrounds) {
    if (!values || !Array.isArray(values)) {
      throw new Error("値データが配列ではありません");
    }

    if (values.length === 0) {
      throw new Error("値データが空です");
    }

    if (backgrounds && backgrounds.length !== values.length) {
      throw new Error("値データと背景色データの行数が一致しません");
    }

    // 各行の列数チェック
    const expectedColumns = values[0].length;
    for (let i = 0; i < values.length; i++) {
      if (values[i].length !== expectedColumns) {
        throw new Error(`行${i + 1}の列数が不正です`);
      }
      
      if (backgrounds && backgrounds[i] && backgrounds[i].length !== expectedColumns) {
        throw new Error(`行${i + 1}の背景色データの列数が不正です`);
      }
    }
  }

  /**
   * シートのバックアップを作成
   * @param {Sheet} sheet - バックアップするシート
   * @returns {Object} バックアップデータ
   */
  createBackup(sheet) {
    try {
      const dataRange = sheet.getDataRange();
      
      if (dataRange.getNumRows() === 0 || dataRange.getNumColumns() === 0) {
        return {
          hasData: false,
          timestamp: new Date()
        };
      }
      
      return {
        hasData: true,
        values: dataRange.getValues(),
        backgrounds: dataRange.getBackgrounds(),
        timestamp: new Date(),
        range: {
          row: dataRange.getRow(),
          column: dataRange.getColumn(),
          numRows: dataRange.getNumRows(),
          numColumns: dataRange.getNumColumns()
        }
      };
    } catch (error) {
      console.warn(`シート「${sheet.getName()}」のバックアップ作成中にエラーが発生しました:`, error);
      return {
        hasData: false,
        timestamp: new Date(),
        error: error.message
      };
    }
  }

  /**
   * バックアップからシートを復元
   * @param {Sheet} sheet - 復元するシート
   * @param {Object} backup - バックアップデータ
   * @returns {boolean} 復元成功の場合true
   */
  restoreFromBackup(sheet, backup) {
    try {
      if (!backup.hasData) {
        sheet.clear();
        return true;
      }

      const range = sheet.getRange(
        backup.range.row,
        backup.range.column,
        backup.range.numRows,
        backup.range.numColumns
      );
      
      range.setValues(backup.values);
      
      if (backup.backgrounds) {
        range.setBackgrounds(backup.backgrounds);
      }
      
      console.log(`シート「${sheet.getName()}」をバックアップから復元しました`);
      return true;
      
    } catch (error) {
      console.error(`シート「${sheet.getName()}」の復元中にエラーが発生しました:`, error);
      return false;
    }
  }

  /**
   * 更新結果の検証
   * @param {Object} updateResult - 更新結果
   * @returns {ValidationResult} 検証結果
   */
  validateUpdateResult(updateResult) {
    const result = new ValidationResult();
    
    if (!updateResult.success) {
      result.addError(`更新に失敗しました: ${updateResult.error}`);
    }
    
    if (updateResult.updatedCells === 0 && updateResult.updatedRows === 0) {
      result.addWarning("更新されたデータがありません");
    }
    
    if (result.errors.length === 0) {
      result.markSuccess();
    }
    
    return result;
  }

  /**
   * 更新統計情報を作成
   * @param {Array} updateResults - 更新結果配列
   * @returns {Object} 統計情報
   */
  createUpdateStatistics(updateResults) {
    const stats = {
      totalUpdates: updateResults.length,
      successfulUpdates: 0,
      failedUpdates: 0,
      totalUpdatedCells: 0,
      totalUpdatedRows: 0,
      updatesByType: {},
      errors: []
    };

    updateResults.forEach(result => {
      if (result.success) {
        stats.successfulUpdates++;
        stats.totalUpdatedCells += result.updatedCells || 0;
        stats.totalUpdatedRows += result.updatedRows || 0;
      } else {
        stats.failedUpdates++;
        stats.errors.push({
          sheetName: result.sheetName,
          error: result.error
        });
      }
    });

    return stats;
  }

  /**
   * 更新処理のロールバック
   * @param {number} historyIndex - 履歴インデックス
   * @returns {boolean} ロールバック成功の場合true
   */
  rollbackUpdate(historyIndex) {
    if (historyIndex < 0 || historyIndex >= this.updateHistory.length) {
      console.error("無効な履歴インデックスです");
      return false;
    }

    const historyEntry = this.updateHistory[historyIndex];
    
    try {
      // 対象シートを取得
      const sheet = SpreadsheetApp.getActiveSheet(); // 実際の実装では適切な方法でシートを取得
      
      // バックアップから復元
      const success = this.restoreFromBackup(sheet, historyEntry.backup);
      
      if (success) {
        console.log(`更新履歴[${historyIndex}]のロールバックが完了しました`);
      }
      
      return success;
      
    } catch (error) {
      console.error(`ロールバック中にエラーが発生しました:`, error);
      return false;
    }
  }

  /**
   * 更新処理のサマリーをログ出力
   * @param {Object} statistics - 統計情報
   */
  logUpdateSummary(statistics) {
    console.log("=== シート更新サマリー ===");
    console.log(`総更新数: ${statistics.totalUpdates}`);
    console.log(`成功した更新: ${statistics.successfulUpdates}`);
    console.log(`失敗した更新: ${statistics.failedUpdates}`);
    console.log(`更新されたセル数: ${statistics.totalUpdatedCells}`);
    console.log(`更新された行数: ${statistics.totalUpdatedRows}`);
    
    if (statistics.errors.length > 0) {
      console.log("\n=== 更新エラー ===");
      statistics.errors.forEach(error => {
        console.log(`シート「${error.sheetName}」: ${error.error}`);
      });
    }
  }

  /**
   * 更新履歴をクリア
   */
  clearUpdateHistory() {
    this.updateHistory = [];
    console.log("更新履歴をクリアしました");
  }

  /**
   * 更新履歴を取得
   * @returns {Array} 更新履歴配列
   */
  getUpdateHistory() {
    return [...this.updateHistory];
  }

  /**
   * 大量データの効率的な更新
   * @param {Sheet} sheet - 更新対象シート
   * @param {Array} data - 更新データ
   * @param {number} batchSize - バッチサイズ
   * @returns {Object} 更新結果
   */
  updateSheetInBatches(sheet, data, batchSize = 1000) {
    const result = {
      success: true,
      totalRows: data.length,
      batchesProcessed: 0,
      errors: []
    };

    try {
      // バッチサイズに分割して処理
      for (let i = 0; i < data.length; i += batchSize) {
        const batch = data.slice(i, Math.min(i + batchSize, data.length));
        const startRow = i + 1;
        
        try {
          const range = sheet.getRange(startRow, 1, batch.length, batch[0].length);
          range.setValues(batch);
          result.batchesProcessed++;
          
          // 進捗通知
          const progress = Math.round((i + batch.length) / data.length * 100);
          NotificationService.showProgress(`シート「${sheet.getName()}」更新中... ${progress}%`);
          
        } catch (error) {
          result.success = false;
          result.errors.push({
            batchIndex: Math.floor(i / batchSize),
            startRow: startRow,
            error: error.message
          });
        }
      }
      
    } catch (error) {
      result.success = false;
      result.errors.push({
        general: error.message
      });
    }

    return result;
  }

  /**
   * シートの保護状態を管理
   * @param {Sheet} sheet - 対象シート
   * @param {boolean} protect - 保護するかどうか
   * @returns {boolean} 操作成功の場合true
   */
  manageSheetProtection(sheet, protect) {
    try {
      if (protect) {
        const protection = sheet.protect();
        protection.setDescription('自動更新による保護');
        console.log(`シート「${sheet.getName()}」を保護しました`);
      } else {
        const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        protections.forEach(protection => {
          if (protection.canEdit()) {
            protection.remove();
          }
        });
        console.log(`シート「${sheet.getName()}」の保護を解除しました`);
      }
      return true;
    } catch (error) {
      console.error(`シート保護の設定中にエラーが発生しました:`, error);
      return false;
    }
  }
} 