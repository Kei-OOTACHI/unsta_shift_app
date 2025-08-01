/**
 * データ変換サービス
 * 
 * 責任:
 * - ガントチャートデータの分割
 * - シフトオブジェクトから2次元配列への変換
 * - 時間ヘッダーの処理
 * - データ構造の正規化
 */
class DataTransformer {
  constructor(configManager) {
    this.config = configManager;
    this.ganttIndexes = configManager.getColumnIndexes('GANTT');
    this.rowIndexes = configManager.getColumnIndexes('ROW');
    this.rdbIndexes = configManager.getColumnIndexes('RDB');
    this.conflictIndexes = configManager.getColumnIndexes('CONFLICT');
    this.errorIndexes = configManager.getColumnIndexes('ERROR');
  }

  /**
   * ガントチャートデータをヘッダーとシフトデータに分割
   * @param {Array} ganttValues - ガントチャートの値
   * @param {Array} ganttBgs - ガントチャートの背景色
   * @returns {Object} 分割されたガントデータ
   */
  splitGanttData(ganttValues, ganttBgs) {
    const firstDataCol = this.ganttIndexes.firstData;
    const firstDataRow = this.rowIndexes.firstData;

    // シフトデータ部分
    const ganttShiftValues = ganttValues.slice(firstDataRow).map((row) => row.slice(firstDataCol));
    const ganttShiftBgs = ganttBgs.slice(firstDataRow).map((row) => row.slice(firstDataCol));

    // ヘッダー部分（L字形）
    const ganttHeaderValues = [];
    const ganttHeaderBgs = [];

    // 上部ヘッダー行（全列を含む）
    for (let i = 0; i < firstDataRow; i++) {
      ganttHeaderValues.push([...ganttValues[i]]);
      ganttHeaderBgs.push([...ganttBgs[i]]);
    }

    // 左側ヘッダー列（firstDataRow行目以降、firstDataCol列までのデータ）
    for (let i = firstDataRow; i < ganttValues.length; i++) {
      ganttHeaderValues.push(ganttValues[i].slice(0, firstDataCol));
      ganttHeaderBgs.push(ganttBgs[i].slice(0, firstDataCol));
    }

    // 時間軸とmemberDateIdの取得
    const timeHeaders = this.extractTimeHeaders(ganttValues, firstDataCol);
    const memberDateIdHeaders = this.extractMemberDateIdHeaders(ganttValues, firstDataRow);

    return {
      ganttHeaderValues,
      ganttShiftValues,
      ganttHeaderBgs,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      firstDataColOffset: firstDataCol,
      firstDataRowOffset: firstDataRow,
    };
  }

  /**
   * 時間ヘッダーを抽出し、終了時間計算用の要素を追加
   * @param {Array} ganttValues - ガントチャートの値
   * @param {number} firstDataCol - 最初のデータ列
   * @returns {Array} 時間ヘッダー配列
   */
  extractTimeHeaders(ganttValues, firstDataCol) {
    const originalTimeRow = this.rowIndexes.timeScale;
    const originalTimeHeaders = ganttValues[originalTimeRow].slice(firstDataCol);
    
    // 時間間隔を計算して最後に1つ追加の時間要素を作成
    const timeHeaders = [...originalTimeHeaders];
    
    if (originalTimeHeaders.length >= 2) {
      // 最後の2つの時間から間隔を計算
      const lastTime = new Date(originalTimeHeaders[originalTimeHeaders.length - 1]);
      const prevTime = new Date(originalTimeHeaders[originalTimeHeaders.length - 2]);
      const timeDiff = lastTime.getTime() - prevTime.getTime();
      const nextTime = new Date(lastTime.getTime() + timeDiff);
      timeHeaders.push(nextTime);
    } else if (originalTimeHeaders.length === 1) {
      // 時間要素が1つしかない場合は1時間後を追加
      const lastTime = new Date(originalTimeHeaders[0]);
      const nextTime = new Date(lastTime.getTime() + 60 * 60 * 1000);
      timeHeaders.push(nextTime);
    }
    
    return timeHeaders;
  }

  /**
   * memberDateIdヘッダーを抽出
   * @param {Array} ganttValues - ガントチャートの値
   * @param {number} firstDataRow - 最初のデータ行
   * @returns {Array} memberDateIdヘッダー配列
   */
  extractMemberDateIdHeaders(ganttValues, firstDataRow) {
    const originalMemberDateIdCol = this.ganttIndexes.memberDateId;
    return ganttValues.slice(firstDataRow).map((row) => row[originalMemberDateIdCol]);
  }

  /**
   * シフトオブジェクトのマップを2次元配列に変換
   * @param {Map} validShiftsMap - 有効なシフトデータのマップ
   * @param {Array} conflictShiftObjs - 競合シフトオブジェクトの配列
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @param {Array} memberDateIdHeaders - memberDateIdヘッダー配列
   * @returns {Object} 変換結果
   */
  convertObjectsTo2DArray(validShiftsMap, conflictShiftObjs, timeHeaders, memberDateIdHeaders) {
    const result = {
      ganttValues: [],
      ganttBgs: [],
      rdbData: [],
      conflictData: [],
      errorData: []
    };

    // 元の時間ヘッダー長（endTime計算用に追加されたものを除外）
    const originalTimeHeadersLength = timeHeaders.length - 1;
    
    // memberDateIdHeadersのSetを作成（高速検索用）
    const memberDateIdHeadersSet = new Set(memberDateIdHeaders);
    
    // 各メンバーのシフト情報を処理
    const ganttValueMap = new Map();
    const ganttBgMap = new Map();
    const processedShiftIds = new Set();
    
    for (const [memberDateId, timeMap] of validShiftsMap.entries()) {
      // memberDateIdがmemberDateIdHeadersに存在しない場合のチェック
      if (!memberDateIdHeadersSet.has(memberDateId)) {
        // このmemberDateIdのシフトデータをエラーデータとして出力
        for (const [timeKey, shiftInfo] of timeMap.entries()) {
          const errorRow = this.createErrorRow(shiftInfo, `memberDateId「${memberDateId}」がガントチャートのヘッダーに見つかりません`);
          result.errorData.push(errorRow);
        }
        continue;
      }
      
      // 各時間スロットごとに処理
      for (const [timeKey, shiftInfo] of timeMap.entries()) {
        // まだ処理していないシフトIDの場合のみrdbDataに追加
        if (!processedShiftIds.has(shiftInfo.shiftId)) {
          const rdbRow = this.createRdbRow(shiftInfo);
          result.rdbData.push(rdbRow);
          processedShiftIds.add(shiftInfo.shiftId);
          
          // ganttData用のデータも準備
          if (!ganttValueMap.has(shiftInfo.memberDateId)) {
            ganttValueMap.set(shiftInfo.memberDateId, Array(originalTimeHeadersLength).fill(""));
            ganttBgMap.set(shiftInfo.memberDateId, Array(originalTimeHeadersLength).fill("#FFFFFF"));
          }
          
          this.fillGanttTimeSlots(
            ganttValueMap.get(shiftInfo.memberDateId),
            ganttBgMap.get(shiftInfo.memberDateId),
            shiftInfo,
            timeHeaders,
            originalTimeHeadersLength
          );
        }
      }
    }

    // 元のmemberDateIdHeadersの順序を保持してganttDataを生成
    result.ganttValues = this.createOrderedGanttValues(memberDateIdHeaders, ganttValueMap, originalTimeHeadersLength);
    result.ganttBgs = this.createOrderedGanttBgs(memberDateIdHeaders, ganttBgMap, originalTimeHeadersLength);

    // コンフリクトデータを処理
    result.conflictData = this.createConflictDataArray(conflictShiftObjs);

    return result;
  }

  /**
   * ガントチャートの時間スロットにシフト情報を設定
   * @param {Array} timeRow - 時間行の配列
   * @param {Array} bgRow - 背景色行の配列
   * @param {Object} shiftInfo - シフト情報
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @param {number} originalTimeHeadersLength - 元の時間ヘッダー長
   */
  fillGanttTimeSlots(timeRow, bgRow, shiftInfo, timeHeaders, originalTimeHeadersLength) {
    const startIndex = this.findTimeIndex(timeHeaders, shiftInfo.startTime);
    const endIndex = this.findTimeIndex(timeHeaders, shiftInfo.endTime);
    
    if (startIndex !== -1 && endIndex !== -1) {
      // 元の列数の範囲内でのみシフトデータを設定
      for (let i = startIndex; i < Math.min(endIndex, originalTimeHeadersLength); i++) {
        timeRow[i] = shiftInfo.job;
        bgRow[i] = shiftInfo.background || "#FFFFFF";
      }
    }
  }

  /**
   * 時間ヘッダー配列から指定時間に最も近いインデックスを検索
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @param {Date} time - 検索する時間
   * @returns {number} インデックス（見つからない場合は-1）
   */
  findTimeIndex(timeHeaders, time) {
    const timeStr = time.toISOString().slice(11, 16);
    for (let i = 0; i < timeHeaders.length; i++) {
      const headerTime = new Date(timeHeaders[i]).toISOString().slice(11, 16);
      if (headerTime === timeStr) {
        return i;
      }
    }
    return -1;
  }

  /**
   * RDBデータ行を作成
   * @param {Object} shiftInfo - シフト情報
   * @returns {Array} RDB行データ
   */
  createRdbRow(shiftInfo) {
    const columnOrder = this.config.getColumnOrder(this.rdbIndexes);
    return columnOrder.map(key => {
      if (key === 'startTime' || key === 'endTime') {
        return TimeUtils.formatTimeToHHMM(shiftInfo[key]);
      }
      return shiftInfo[key] || "";
    });
  }

  /**
   * エラーデータ行を作成
   * @param {Object} shiftInfo - シフト情報
   * @param {string} errorMessage - エラーメッセージ
   * @returns {Array} エラー行データ
   */
  createErrorRow(shiftInfo, errorMessage) {
    const columnOrder = this.config.getColumnOrder(this.errorIndexes);
    return columnOrder.map(key => {
      if (key === 'startTime' || key === 'endTime') {
        return TimeUtils.formatTimeToHHMM(shiftInfo[key]);
      } else if (key === 'errorMessage') {
        return errorMessage || "";
      }
      return shiftInfo[key] || "";
    });
  }

  /**
   * コンフリクトデータ配列を作成
   * @param {Array} conflictShiftObjs - コンフリクトシフトオブジェクトの配列
   * @returns {Array} コンフリクトデータ配列
   */
  createConflictDataArray(conflictShiftObjs) {
    const columnOrder = this.config.getColumnOrder(this.conflictIndexes);
    return conflictShiftObjs.map((shiftObj) => {
      return columnOrder.map((key) => {
        if (key === 'startTime' || key === 'endTime') {
          return TimeUtils.formatTimeToHHMM(shiftObj[key]);
        }
        return shiftObj[key] || "";
      });
    });
  }

  /**
   * 順序を保持したガント値配列を作成
   * @param {Array} memberDateIdHeaders - memberDateIdヘッダー配列
   * @param {Map} ganttValueMap - ガント値マップ
   * @param {number} originalTimeHeadersLength - 元の時間ヘッダー長
   * @returns {Array} ガント値配列
   */
  createOrderedGanttValues(memberDateIdHeaders, ganttValueMap, originalTimeHeadersLength) {
    return memberDateIdHeaders.map(memberDateId => {
      if (ganttValueMap.has(memberDateId)) {
        return ganttValueMap.get(memberDateId);
      } else {
        return Array(originalTimeHeadersLength).fill("");
      }
    });
  }

  /**
   * 順序を保持したガント背景色配列を作成
   * @param {Array} memberDateIdHeaders - memberDateIdヘッダー配列
   * @param {Map} ganttBgMap - ガント背景色マップ
   * @param {number} originalTimeHeadersLength - 元の時間ヘッダー長
   * @returns {Array} ガント背景色配列
   */
  createOrderedGanttBgs(memberDateIdHeaders, ganttBgMap, originalTimeHeadersLength) {
    return memberDateIdHeaders.map(memberDateId => {
      if (ganttBgMap.has(memberDateId)) {
        return ganttBgMap.get(memberDateId);
      } else {
        return Array(originalTimeHeadersLength).fill("#FFFFFF");
      }
    });
  }

  /**
   * ガントヘッダーとシフトデータを統合
   * @param {Array} ganttHeaderValues - ガントヘッダー値
   * @param {Array} ganttShiftValues - ガントシフト値
   * @param {Array} ganttHeaderBgs - ガントヘッダー背景色
   * @param {Array} ganttShiftBgs - ガントシフト背景色
   * @param {number} firstDataColOffset - 最初のデータ列オフセット
   * @param {number} firstDataRowOffset - 最初のデータ行オフセット
   * @returns {Object} 統合されたガントデータ
   */
  mergeGanttData(
    ganttHeaderValues,
    ganttShiftValues,
    ganttHeaderBgs,
    ganttShiftBgs,
    firstDataColOffset,
    firstDataRowOffset
  ) {
    const mergedValues = [];
    const mergedBgs = [];

    // 上部ヘッダー行を追加
    for (let i = 0; i < firstDataRowOffset; i++) {
      mergedValues.push([...ganttHeaderValues[i]]);
      mergedBgs.push([...ganttHeaderBgs[i]]);
    }

    // 左側ヘッダー列とシフトデータを結合して追加
    for (let i = 0; i < ganttShiftValues.length; i++) {
      const headerRow = ganttHeaderValues[i + firstDataRowOffset];
      const bgHeaderRow = ganttHeaderBgs[i + firstDataRowOffset];

      mergedValues.push([...headerRow, ...ganttShiftValues[i]]);
      mergedBgs.push([...bgHeaderRow, ...ganttShiftBgs[i]]);
    }

    return {
      values: mergedValues,
      backgrounds: mergedBgs,
    };
  }

  /**
   * 2次元配列データの統計情報を作成
   * @param {Array} ganttValues - ガント値配列
   * @param {Array} rdbData - RDBデータ配列
   * @param {Array} conflictData - コンフリクトデータ配列
   * @param {Array} errorData - エラーデータ配列
   * @returns {Object} 統計情報
   */
  createTransformationStatistics(ganttValues, rdbData, conflictData, errorData) {
    return {
      ganttCells: this.count2DArrayCells(ganttValues),
      rdbRows: rdbData.length,
      conflictRows: conflictData.length,
      errorRows: errorData.length,
      totalOutputRows: rdbData.length + conflictData.length + errorData.length
    };
  }

  /**
   * 2次元配列のセル数をカウント
   * @param {Array} array2D - 2次元配列
   * @returns {number} セル数
   */
  count2DArrayCells(array2D) {
    if (!Array.isArray(array2D)) return 0;
    
    return array2D.reduce((total, row) => {
      if (Array.isArray(row)) {
        return total + row.filter(cell => cell !== "" && cell !== null && cell !== undefined).length;
      }
      return total;
    }, 0);
  }

  /**
   * データ変換処理のサマリーをログ出力
   * @param {Object} statistics - 統計情報
   */
  logTransformationSummary(statistics) {
    console.log("=== データ変換サマリー ===");
    console.log(`ガントセル数: ${statistics.ganttCells}`);
    console.log(`RDB行数: ${statistics.rdbRows}`);
    console.log(`コンフリクト行数: ${statistics.conflictRows}`);
    console.log(`エラー行数: ${statistics.errorRows}`);
    console.log(`総出力行数: ${statistics.totalOutputRows}`);
  }

  /**
   * データ変換の実行
   * @param {Object} splitGanttData - 分割されたガントデータ
   * @param {Map} validShiftsMap - 有効なシフトマップ
   * @param {Array} conflictShiftObjs - コンフリクトシフトオブジェクト
   * @returns {Object} 変換結果
   */
  transformData(splitGanttData, validShiftsMap, conflictShiftObjs) {
    const { timeHeaders, memberDateIdHeaders } = splitGanttData;
    
    // オブジェクトから2次元配列への変換
    const transformResult = this.convertObjectsTo2DArray(
      validShiftsMap,
      conflictShiftObjs,
      timeHeaders,
      memberDateIdHeaders
    );

    // 統計情報の作成
    const statistics = this.createTransformationStatistics(
      transformResult.ganttValues,
      transformResult.rdbData,
      transformResult.conflictData,
      transformResult.errorData
    );

    this.logTransformationSummary(statistics);

    return {
      ...transformResult,
      statistics
    };
  }

  /**
   * 時間範囲の検証
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @returns {ValidationResult} 検証結果
   */
  validateTimeRange(timeHeaders) {
    const result = new ValidationResult();
    
    if (!timeHeaders || timeHeaders.length === 0) {
      result.addError("時間ヘッダーが空です");
      return result;
    }

    // 時間の順序チェック
    for (let i = 1; i < timeHeaders.length; i++) {
      const prevTime = new Date(timeHeaders[i - 1]);
      const currentTime = new Date(timeHeaders[i]);
      
      if (isNaN(prevTime.getTime()) || isNaN(currentTime.getTime())) {
        result.addError(`無効な時間データが含まれています: ${timeHeaders[i - 1]}, ${timeHeaders[i]}`);
      } else if (prevTime >= currentTime) {
        result.addError(`時間の順序が不正です: ${timeHeaders[i - 1]} >= ${timeHeaders[i]}`);
      }
    }

    if (result.errors.length === 0) {
      result.markSuccess();
    }

    return result;
  }

  /**
   * 変換品質の評価
   * @param {Object} inputData - 入力データ
   * @param {Object} outputData - 出力データ
   * @returns {Object} 品質評価結果
   */
  evaluateTransformationQuality(inputData, outputData) {
    const quality = {
      dataIntegrity: true,
      warnings: [],
      errors: [],
      metrics: {}
    };

    // データ整合性チェック
    const inputShiftCount = this.countShiftsInInput(inputData);
    const outputShiftCount = this.countShiftsInOutput(outputData);
    
    quality.metrics.inputShiftCount = inputShiftCount;
    quality.metrics.outputShiftCount = outputShiftCount;
    quality.metrics.lossRate = inputShiftCount > 0 ? ((inputShiftCount - outputShiftCount) / inputShiftCount) * 100 : 0;

    if (inputShiftCount !== outputShiftCount) {
      quality.dataIntegrity = false;
      quality.errors.push(`シフト数不一致: 入力${inputShiftCount}, 出力${outputShiftCount}`);
    }

    // 変換効率の評価
    quality.metrics.transformationEfficiency = quality.metrics.lossRate < 5 ? "高" : 
                                               quality.metrics.lossRate < 15 ? "中" : "低";

    return quality;
  }

  /**
   * 入力データのシフト数をカウント
   * @param {Object} inputData - 入力データ
   * @returns {number} シフト数
   */
  countShiftsInInput(inputData) {
    // 実装は具体的な入力データ構造に依存
    return 0;
  }

  /**
   * 出力データのシフト数をカウント
   * @param {Object} outputData - 出力データ
   * @returns {number} シフト数
   */
  countShiftsInOutput(outputData) {
    return outputData.rdbData ? outputData.rdbData.length : 0;
  }
} 