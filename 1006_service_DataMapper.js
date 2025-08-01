/**
 * データマッピングサービス
 * 
 * 責任:
 * - シフトデータの分析と抽出
 * - シフトの競合検出
 * - 部署別データの統合
 * - データの正規化
 */
class DataMapper {
  constructor(configManager) {
    this.config = configManager;
    this.ganttIndexes = configManager.getColumnIndexes('GANTT');
    this.rowIndexes = configManager.getColumnIndexes('ROW');
    this.rdbIndexes = configManager.getColumnIndexes('RDB');
  }

  /**
   * 部署別のシフト統合とマッピング
   * @param {Object} ganttDataByDept - 部署別ガントデータ
   * @param {Object} rdbDataByDept - 部署別RDBデータ
   * @param {Array} deptList - 処理対象の部署リスト
   * @returns {Object} マッピング結果
   */
  mapShiftsByDepartment(ganttDataByDept, rdbDataByDept, deptList) {
    const mappingResult = {
      allShifts: [],
      conflictShifts: [],
      validShiftsMap: new Map(),
      statistics: {
        totalShifts: 0,
        validShifts: 0,
        conflictShifts: 0,
        errorShifts: 0,
        departmentCounts: {}
      }
    };

    deptList.forEach(dept => {
      // 部署別データのマッピング
      const departmentMapping = this.mapDepartmentData(
        ganttDataByDept[dept],
        rdbDataByDept[dept],
        dept
      );

      // 結果の統合
      mappingResult.allShifts.push(...departmentMapping.allShifts);
      mappingResult.conflictShifts.push(...departmentMapping.conflictShifts);
      
      // 有効なシフトマップの統合
      for (const [memberDateId, timeMap] of departmentMapping.validShiftsMap.entries()) {
        if (!mappingResult.validShiftsMap.has(memberDateId)) {
          mappingResult.validShiftsMap.set(memberDateId, new Map());
        }
        for (const [timeKey, shiftInfo] of timeMap.entries()) {
          mappingResult.validShiftsMap.get(memberDateId).set(timeKey, shiftInfo);
        }
      }

      // 統計情報の更新
      mappingResult.statistics.departmentCounts[dept] = departmentMapping.statistics;
      mappingResult.statistics.totalShifts += departmentMapping.statistics.totalShifts;
      mappingResult.statistics.validShifts += departmentMapping.statistics.validShifts;
      mappingResult.statistics.conflictShifts += departmentMapping.statistics.conflictShifts;
      mappingResult.statistics.errorShifts += departmentMapping.statistics.errorShifts;
    });

    return mappingResult;
  }

  /**
   * 単一部署のデータマッピング
   * @param {GanttChartData} ganttData - ガントチャートデータ
   * @param {Array} rdbData - RDBデータ
   * @param {string} dept - 部署名
   * @returns {Object} 部署別マッピング結果
   */
  mapDepartmentData(ganttData, rdbData, dept) {
    const result = {
      allShifts: [],
      conflictShifts: [],
      validShiftsMap: new Map(),
      statistics: {
        totalShifts: 0,
        validShifts: 0,
        conflictShifts: 0,
        errorShifts: 0
      }
    };

    try {
      // ガントチャートからシフトデータを抽出
      const ganttShifts = this.extractShiftsFromGanttData(ganttData, dept);
      
      // RDBデータをシフトオブジェクトに変換
      const rdbShifts = this.convertRdbDataToShifts(rdbData, dept);
      
      // 全シフトを統合
      const allShifts = [...ganttShifts, ...rdbShifts];
      result.allShifts = allShifts;
      result.statistics.totalShifts = allShifts.length;

      // 競合検出とマッピング
      const conflictAnalysis = this.detectConflicts(allShifts);
      result.conflictShifts = conflictAnalysis.conflictShifts;
      result.validShiftsMap = conflictAnalysis.validShiftsMap;
      
      result.statistics.validShifts = conflictAnalysis.validShifts;
      result.statistics.conflictShifts = conflictAnalysis.conflictShifts.length;
      
      console.log(`部署「${dept}」の処理結果:`, result.statistics);

    } catch (error) {
      console.error(`部署「${dept}」のマッピング中にエラーが発生しました:`, error);
      result.statistics.errorShifts += 1;
    }

    return result;
  }

  /**
   * ガントチャートからシフトデータを抽出
   * @param {GanttChartData} ganttData - ガントチャートデータ
   * @param {string} dept - 部署名
   * @returns {Array} シフトオブジェクトの配列
   */
  extractShiftsFromGanttData(ganttData, dept) {
    if (!ganttData || ganttData.isEmpty()) {
      return [];
    }

    const shifts = [];
    const { values, backgrounds } = ganttData;
    
    // 時間軸とmemberDateIdの取得
    const timeHeaders = this.extractTimeHeaders(values);
    const memberDateIdHeaders = this.extractMemberDateIdHeaders(values);

    // シフトデータ部分の抽出
    const firstDataRow = this.rowIndexes.firstData;
    const firstDataCol = this.ganttIndexes.firstData;
    
    const shiftValues = values.slice(firstDataRow).map(row => row.slice(firstDataCol));
    const shiftBgs = backgrounds.slice(firstDataRow).map(row => row.slice(firstDataCol));

    // 各セルを検査してシフトデータを抽出
    shiftValues.forEach((row, rowIndex) => {
      const memberDateId = memberDateIdHeaders[rowIndex];
      
      if (!memberDateId) return;

      row.forEach((cell, colIndex) => {
        const bgColor = shiftBgs[rowIndex] ? shiftBgs[rowIndex][colIndex] : "#ffffff";
        
        if (this.isShiftData(cell, bgColor)) {
          const shiftObj = this.createShiftObject(
            cell,
            bgColor,
            memberDateId,
            timeHeaders,
            colIndex,
            dept,
            'GANTT'
          );
          
          if (shiftObj) {
            shifts.push(shiftObj);
          }
        }
      });
    });

    return shifts;
  }

  /**
   * セルがシフトデータかどうかを判定
   * @param {*} cell - セルの値
   * @param {string} bgColor - 背景色
   * @returns {boolean} シフトデータの場合true
   */
  isShiftData(cell, bgColor) {
    const hasValue = cell !== "" && cell !== null && cell !== undefined;
    const hasNonWhiteBg = this.isNonWhiteBackground(bgColor);
    return hasValue || hasNonWhiteBg;
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
   * シフトオブジェクトを作成
   * @param {*} cell - セルの値
   * @param {string} bgColor - 背景色
   * @param {string} memberDateId - メンバー日付ID
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @param {number} colIndex - 列インデックス
   * @param {string} dept - 部署名
   * @param {string} source - データソース
   * @returns {Object|null} シフトオブジェクト
   */
  createShiftObject(cell, bgColor, memberDateId, timeHeaders, colIndex, dept, source) {
    if (colIndex >= timeHeaders.length - 1) {
      return null; // 終了時間計算用の追加要素は除外
    }

    const startTime = new Date(timeHeaders[colIndex]);
    const endTime = new Date(timeHeaders[colIndex + 1]);
    
    if (isNaN(startTime.getTime()) || isNaN(endTime.getTime())) {
      return null;
    }

    const shiftId = this.generateShiftId(memberDateId, startTime, endTime, source);
    
    return {
      shiftId,
      memberDateId,
      startTime,
      endTime,
      job: String(cell || ""),
      background: bgColor || "#ffffff",
      dept,
      source
    };
  }

  /**
   * シフトIDを生成
   * @param {string} memberDateId - メンバー日付ID
   * @param {Date} startTime - 開始時間
   * @param {Date} endTime - 終了時間
   * @param {string} source - データソース
   * @returns {string} シフトID
   */
  generateShiftId(memberDateId, startTime, endTime, source) {
    const startStr = startTime.toISOString().slice(11, 16);
    const endStr = endTime.toISOString().slice(11, 16);
    return `${memberDateId}_${startStr}_${endStr}_${source}`;
  }

  /**
   * RDBデータをシフトオブジェクトに変換
   * @param {Array} rdbData - RDBデータ
   * @param {string} dept - 部署名
   * @returns {Array} シフトオブジェクトの配列
   */
  convertRdbDataToShifts(rdbData, dept) {
    if (!rdbData || rdbData.length === 0) {
      return [];
    }

    return rdbData.map((row, index) => {
      try {
        const memberDateId = row[this.rdbIndexes.memberDateId];
        const startTime = TimeUtils.parseTimeToDate(row[this.rdbIndexes.startTime]);
        const endTime = TimeUtils.parseTimeToDate(row[this.rdbIndexes.endTime]);
        const job = row[this.rdbIndexes.job] || "";
        const background = row[this.rdbIndexes.background] || "#ffffff";
        
        const shiftId = this.generateShiftId(memberDateId, startTime, endTime, 'RDB');
        
        return {
          shiftId,
          memberDateId,
          startTime,
          endTime,
          job,
          background,
          dept,
          source: 'RDB'
        };
      } catch (error) {
        console.error(`RDBデータ行${index}の変換中にエラーが発生しました:`, error);
        return null;
      }
    }).filter(shift => shift !== null);
  }

  /**
   * 時間ヘッダーを抽出
   * @param {Array} values - ガントチャートの値
   * @returns {Array} 時間ヘッダー配列
   */
  extractTimeHeaders(values) {
    const timeRow = this.rowIndexes.timeScale;
    const firstDataCol = this.ganttIndexes.firstData;
    
    const originalTimeHeaders = values[timeRow].slice(firstDataCol);
    const timeHeaders = [...originalTimeHeaders];
    
    // 終了時間計算用の追加要素を作成
    if (originalTimeHeaders.length >= 2) {
      const lastTime = new Date(originalTimeHeaders[originalTimeHeaders.length - 1]);
      const prevTime = new Date(originalTimeHeaders[originalTimeHeaders.length - 2]);
      const timeDiff = lastTime.getTime() - prevTime.getTime();
      const nextTime = new Date(lastTime.getTime() + timeDiff);
      timeHeaders.push(nextTime);
    }
    
    return timeHeaders;
  }

  /**
   * memberDateIdヘッダーを抽出
   * @param {Array} values - ガントチャートの値
   * @returns {Array} memberDateIdヘッダー配列
   */
  extractMemberDateIdHeaders(values) {
    const firstDataRow = this.rowIndexes.firstData;
    const memberDateIdCol = this.ganttIndexes.memberDateId;
    
    return values.slice(firstDataRow).map(row => row[memberDateIdCol]);
  }

  /**
   * 競合検出とマッピング
   * @param {Array} allShifts - 全シフトデータ
   * @returns {Object} 競合検出結果
   */
  detectConflicts(allShifts) {
    const conflictMap = new Map();
    const validShiftsMap = new Map();
    const conflictShifts = [];
    let validShifts = 0;

    // memberDateIdと時間範囲でグループ化
    const timeSlotMap = new Map();
    
    allShifts.forEach(shift => {
      const memberDateId = shift.memberDateId;
      const timeSlotKey = this.createTimeSlotKey(shift.startTime, shift.endTime);
      const conflictKey = `${memberDateId}_${timeSlotKey}`;
      
      if (!timeSlotMap.has(conflictKey)) {
        timeSlotMap.set(conflictKey, []);
      }
      timeSlotMap.get(conflictKey).push(shift);
    });

    // 競合検出
    for (const [conflictKey, shifts] of timeSlotMap.entries()) {
      if (shifts.length > 1) {
        // 複数のシフトがある場合は競合
        const selectedShift = this.selectShiftFromConflicts(shifts);
        const conflictedShifts = shifts.filter(shift => shift.shiftId !== selectedShift.shiftId);
        
        // 選択されたシフトを有効として追加
        this.addToValidShiftsMap(validShiftsMap, selectedShift);
        validShifts++;
        
        // 競合シフトを追加
        conflictShifts.push(...conflictedShifts);
        
      } else {
        // 単一のシフトは有効
        const shift = shifts[0];
        this.addToValidShiftsMap(validShiftsMap, shift);
        validShifts++;
      }
    }

    return {
      validShiftsMap,
      conflictShifts,
      validShifts
    };
  }

  /**
   * 時間スロットキーを作成
   * @param {Date} startTime - 開始時間
   * @param {Date} endTime - 終了時間
   * @returns {string} 時間スロットキー
   */
  createTimeSlotKey(startTime, endTime) {
    const startStr = startTime.toISOString().slice(11, 16);
    const endStr = endTime.toISOString().slice(11, 16);
    return `${startStr}_${endStr}`;
  }

  /**
   * 競合シフトから選択する
   * @param {Array} shifts - 競合するシフト配列
   * @returns {Object} 選択されたシフト
   */
  selectShiftFromConflicts(shifts) {
    // 優先順位: RDB > GANTT
    const rdbShifts = shifts.filter(shift => shift.source === 'RDB');
    if (rdbShifts.length > 0) {
      return rdbShifts[0]; // RDBが優先
    }
    
    return shifts[0]; // それ以外は最初のものを選択
  }

  /**
   * 有効なシフトマップに追加
   * @param {Map} validShiftsMap - 有効なシフトマップ
   * @param {Object} shift - シフトオブジェクト
   */
  addToValidShiftsMap(validShiftsMap, shift) {
    const memberDateId = shift.memberDateId;
    const timeSlotKey = this.createTimeSlotKey(shift.startTime, shift.endTime);
    
    if (!validShiftsMap.has(memberDateId)) {
      validShiftsMap.set(memberDateId, new Map());
    }
    
    validShiftsMap.get(memberDateId).set(timeSlotKey, shift);
  }

  /**
   * マッピング結果の統計情報を作成
   * @param {Object} mappingResult - マッピング結果
   * @returns {Object} 統計情報
   */
  createMappingStatistics(mappingResult) {
    const stats = {
      totalShifts: mappingResult.statistics.totalShifts,
      validShifts: mappingResult.statistics.validShifts,
      conflictShifts: mappingResult.statistics.conflictShifts,
      errorShifts: mappingResult.statistics.errorShifts,
      conflictRate: 0,
      validRate: 0,
      departmentBreakdown: mappingResult.statistics.departmentCounts
    };

    if (stats.totalShifts > 0) {
      stats.conflictRate = Math.round((stats.conflictShifts / stats.totalShifts) * 100 * 100) / 100;
      stats.validRate = Math.round((stats.validShifts / stats.totalShifts) * 100 * 100) / 100;
    }

    return stats;
  }

  /**
   * マッピング結果のサマリーをログ出力
   * @param {Object} statistics - 統計情報
   */
  logMappingSummary(statistics) {
    console.log("=== データマッピングサマリー ===");
    console.log(`総シフト数: ${statistics.totalShifts}`);
    console.log(`有効シフト数: ${statistics.validShifts}`);
    console.log(`競合シフト数: ${statistics.conflictShifts}`);
    console.log(`エラーシフト数: ${statistics.errorShifts}`);
    console.log(`有効率: ${statistics.validRate}%`);
    console.log(`競合率: ${statistics.conflictRate}%`);
    
    console.log("\n=== 部署別統計 ===");
    Object.entries(statistics.departmentBreakdown).forEach(([dept, deptStats]) => {
      console.log(`${dept}: 総${deptStats.totalShifts}, 有効${deptStats.validShifts}, 競合${deptStats.conflictShifts}`);
    });
  }

  /**
   * データマッピングの実行
   * @param {Object} ganttDataByDept - 部署別ガントデータ
   * @param {Object} rdbDataByDept - 部署別RDBデータ
   * @param {Array} deptList - 処理対象部署リスト
   * @returns {Object} マッピング結果
   */
  executeMapping(ganttDataByDept, rdbDataByDept, deptList) {
    try {
      NotificationService.showProgress("シフトデータのマッピングを開始します...");
      
      const mappingResult = this.mapShiftsByDepartment(ganttDataByDept, rdbDataByDept, deptList);
      const statistics = this.createMappingStatistics(mappingResult);
      
      this.logMappingSummary(statistics);
      
      return {
        ...mappingResult,
        statistics
      };
      
    } catch (error) {
      console.error("データマッピング中にエラーが発生しました:", error);
      throw error;
    }
  }

  /**
   * マッピング結果の検証
   * @param {Object} mappingResult - マッピング結果
   * @returns {ValidationResult} 検証結果
   */
  validateMappingResult(mappingResult) {
    const result = new ValidationResult();
    
    // 基本的な整合性チェック
    if (mappingResult.statistics.totalShifts === 0) {
      result.addWarning("処理対象のシフトデータがありません");
    }
    
    // 競合率の高さをチェック
    if (mappingResult.statistics.conflictRate > 50) {
      result.addWarning(`競合率が高すぎます: ${mappingResult.statistics.conflictRate}%`);
    }
    
    // エラー率のチェック
    if (mappingResult.statistics.errorShifts > 0) {
      result.addError(`エラーシフトが${mappingResult.statistics.errorShifts}件発生しました`);
    }
    
    // 有効なシフトが存在するかチェック
    if (mappingResult.statistics.validShifts === 0) {
      result.addError("有効なシフトデータがありません");
    }
    
    if (result.errors.length === 0) {
      result.markSuccess();
    }
    
    return result;
  }
} 