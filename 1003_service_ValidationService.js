/**
 * バリデーションサービス
 * 
 * 責任:
 * - RDBデータのバリデーション
 * - ガントチャートデータのバリデーション
 * - 時間データのバリデーション
 * - 部署データのバリデーション
 * - バリデーション結果の管理
 */
class ValidationService {
  constructor(configManager) {
    this.config = configManager;
    this.columnIndexes = configManager.getColumnIndexes('RDB');
    this.sheetNames = configManager.getSheetNames();
  }

  /**
   * RDBデータを検証して有効・無効に分離
   * @param {Array} rdbData - RDBデータの2次元配列
   * @returns {ValidationResult} 検証結果
   */
  validateAndSeparateRdbData(rdbData) {
    const result = new ValidationResult();
    
    if (!rdbData || rdbData.length === 0) {
      result.markSuccess();
      return result;
    }

    // ヘッダー行は常に有効として追加
    if (rdbData.length > 0) {
      result.addValidData(rdbData[0]);
    }

    // データ行のバリデーション（1行目以降）
    for (let i = 1; i < rdbData.length; i++) {
      const row = rdbData[i];
      const validationResult = this.validateRdbRow(row, i);
      
      if (validationResult.isValid) {
        result.addValidData(row);
      } else {
        // エラーメッセージと共に無効データとして追加
        const errorRow = row.concat(this.sheetNames.IN_RDB, validationResult.errorMessage);
        result.addInvalidData(errorRow);
      }
    }

    result.markSuccess();
    return result;
  }

  /**
   * 単一のRDB行を検証
   * @param {Array} row - 検証する行データ
   * @param {number} rowIndex - 行インデックス
   * @returns {Object} 検証結果
   */
  validateRdbRow(row, rowIndex) {
    const errors = [];
    const result = {
      isValid: true,
      errorMessage: "",
      rowIndex: rowIndex
    };

    try {
      // 必須フィールドのバリデーション
      const memberDateId = row[this.columnIndexes.memberDateId];
      const startTimeValue = row[this.columnIndexes.startTime];
      const endTimeValue = row[this.columnIndexes.endTime];
      const dept = row[this.columnIndexes.dept];

      // memberDateIdのバリデーション
      if (!this.validateMemberDateId(memberDateId)) {
        errors.push("memberDateIdが空です");
      }

      // startTimeのバリデーション
      let startTime = null;
      try {
        startTime = this.validateAndParseTime(startTimeValue, "startTime");
      } catch (error) {
        errors.push("startTimeが無効または空です");
      }

      // endTimeのバリデーション
      let endTime = null;
      try {
        endTime = this.validateAndParseTime(endTimeValue, "endTime");
      } catch (error) {
        errors.push("endTimeが無効または空です");
      }

      // startTimeとendTimeの順序チェック
      if (startTime && endTime && !this.validateTimeRange(startTime, endTime)) {
        errors.push("startTimeがendTime以降の時刻です");
      }

      // deptのバリデーション
      if (!this.validateDepartment(dept)) {
        errors.push("deptが空です");
      }

      // 結果の設定
      if (errors.length > 0) {
        result.isValid = false;
        result.errorMessage = errors.join("、");
      }

    } catch (error) {
      result.isValid = false;
      result.errorMessage = `行${rowIndex + 1}の検証中にエラーが発生しました: ${error.message}`;
    }

    return result;
  }

  /**
   * memberDateIdを検証
   * @param {*} memberDateId - 検証するmemberDateId
   * @returns {boolean} 検証結果
   */
  validateMemberDateId(memberDateId) {
    return memberDateId && memberDateId.toString().trim() !== "";
  }

  /**
   * 部署名を検証
   * @param {*} dept - 検証する部署名
   * @returns {boolean} 検証結果
   */
  validateDepartment(dept) {
    return dept && dept.toString().trim() !== "";
  }

  /**
   * 時間データを検証し、Dateオブジェクトに変換
   * @param {*} timeValue - 検証する時間値
   * @param {string} fieldName - フィールド名
   * @returns {Date} 検証済みのDateオブジェクト
   * @throws {ValidationError} 検証に失敗した場合
   */
  validateAndParseTime(timeValue, fieldName) {
    if (!timeValue) {
      throw new ValidationError(`${fieldName}が空です`, fieldName);
    }

    try {
      return TimeUtils.parseTimeToDate(timeValue);
    } catch (error) {
      throw new ValidationError(`${fieldName}が無効な時刻形式です: ${timeValue}`, fieldName);
    }
  }

  /**
   * 時間範囲を検証
   * @param {Date} startTime - 開始時間
   * @param {Date} endTime - 終了時間
   * @returns {boolean} 検証結果
   */
  validateTimeRange(startTime, endTime) {
    if (!(startTime instanceof Date) || !(endTime instanceof Date)) {
      return false;
    }

    if (isNaN(startTime.getTime()) || isNaN(endTime.getTime())) {
      return false;
    }

    return startTime < endTime;
  }

  /**
   * シフトデータオブジェクトを検証
   * @param {ShiftData} shiftData - 検証するシフトデータ
   * @returns {ValidationResult} 検証結果
   */
  validateShiftData(shiftData) {
    const result = new ValidationResult();
    
    if (!shiftData) {
      result.addError("シフトデータが空です");
      return result;
    }

    const errors = shiftData.validate();
    
    if (errors.length > 0) {
      errors.forEach(error => result.addError(error));
      result.addInvalidData(shiftData);
    } else {
      result.addValidData(shiftData);
    }

    if (errors.length === 0) {
      result.markSuccess();
    }

    return result;
  }

  /**
   * ガントチャートデータを検証
   * @param {GanttChartData} ganttData - 検証するガントチャートデータ
   * @param {string} sheetName - シート名
   * @returns {ValidationResult} 検証結果
   */
  validateGanttData(ganttData, sheetName) {
    const result = new ValidationResult();
    
    if (!ganttData) {
      result.addError(`シート「${sheetName}」のガントチャートデータが空です`);
      return result;
    }

    // 空のデータチェック
    if (ganttData.isEmpty()) {
      result.addWarning(`シート「${sheetName}」にデータがありません`);
      result.markSuccess();
      return result;
    }

    // データ構造の検証
    if (!ganttData.values || !Array.isArray(ganttData.values)) {
      result.addError(`シート「${sheetName}」のデータ構造が不正です`);
      return result;
    }

    if (!ganttData.backgrounds || !Array.isArray(ganttData.backgrounds)) {
      result.addError(`シート「${sheetName}」の背景色データが不正です`);
      return result;
    }

    // 行数と列数の整合性チェック
    if (ganttData.values.length !== ganttData.backgrounds.length) {
      result.addError(`シート「${sheetName}」のデータと背景色の行数が一致しません`);
      return result;
    }

    // 各行の列数チェック
    for (let i = 0; i < ganttData.values.length; i++) {
      if (ganttData.values[i].length !== ganttData.backgrounds[i].length) {
        result.addError(`シート「${sheetName}」の${i + 1}行目のデータと背景色の列数が一致しません`);
        return result;
      }
    }

    result.addValidData(ganttData);
    result.markSuccess();
    return result;
  }

  /**
   * 時間ヘッダーを検証
   * @param {Array} timeHeaders - 時間ヘッダー配列
   * @returns {ValidationResult} 検証結果
   */
  validateTimeHeaders(timeHeaders) {
    const result = new ValidationResult();
    
    if (!timeHeaders || !Array.isArray(timeHeaders)) {
      result.addError("時間ヘッダーが配列ではありません");
      return result;
    }

    if (timeHeaders.length === 0) {
      result.addError("時間ヘッダーが空です");
      return result;
    }

    // 各時間ヘッダーの検証
    for (let i = 0; i < timeHeaders.length; i++) {
      const timeHeader = timeHeaders[i];
      
      try {
        const date = new Date(timeHeader);
        if (isNaN(date.getTime())) {
          result.addError(`時間ヘッダー[${i}]が無効な日付です: ${timeHeader}`);
        }
      } catch (error) {
        result.addError(`時間ヘッダー[${i}]の検証中にエラーが発生しました: ${error.message}`);
      }
    }

    // 時間の順序チェック
    for (let i = 1; i < timeHeaders.length; i++) {
      const prevTime = new Date(timeHeaders[i - 1]);
      const currentTime = new Date(timeHeaders[i]);
      
      if (prevTime >= currentTime) {
        result.addError(`時間ヘッダーの順序が不正です: ${timeHeaders[i - 1]} >= ${timeHeaders[i]}`);
      }
    }

    if (result.errors.length === 0) {
      result.addValidData(timeHeaders);
      result.markSuccess();
    }

    return result;
  }

  /**
   * memberDateIdヘッダーを検証
   * @param {Array} memberDateIdHeaders - memberDateIdヘッダー配列
   * @returns {ValidationResult} 検証結果
   */
  validateMemberDateIdHeaders(memberDateIdHeaders) {
    const result = new ValidationResult();
    
    if (!memberDateIdHeaders || !Array.isArray(memberDateIdHeaders)) {
      result.addError("memberDateIdヘッダーが配列ではありません");
      return result;
    }

    if (memberDateIdHeaders.length === 0) {
      result.addError("memberDateIdヘッダーが空です");
      return result;
    }

    // 重複チェック
    const uniqueIds = new Set();
    const duplicates = [];
    
    memberDateIdHeaders.forEach((id, index) => {
      if (id && id.toString().trim() !== "") {
        if (uniqueIds.has(id)) {
          duplicates.push(`${id} (行${index + 1})`);
        } else {
          uniqueIds.add(id);
        }
      }
    });

    if (duplicates.length > 0) {
      result.addWarning(`重複するmemberDateIdがあります: ${duplicates.join(", ")}`);
    }

    result.addValidData(memberDateIdHeaders);
    result.markSuccess();
    return result;
  }

  /**
   * 部署データを検証
   * @param {Array} ganttDepartments - ガントチャートの部署一覧
   * @param {Array} rdbDepartments - RDBの部署一覧
   * @returns {ValidationResult} 検証結果
   */
  validateDepartments(ganttDepartments, rdbDepartments) {
    const result = new ValidationResult();
    
    const ganttDeptSet = new Set(ganttDepartments);
    const rdbDeptSet = new Set(rdbDepartments);
    
    // 有効な部署（両方に存在する部署）
    const validDepartments = new Set([...ganttDeptSet].filter(dept => rdbDeptSet.has(dept)));
    
    // ガントチャートのみに存在する部署
    const ganttOnlyDepartments = [...ganttDeptSet].filter(dept => !rdbDeptSet.has(dept));
    
    // RDBのみに存在する部署
    const rdbOnlyDepartments = [...rdbDeptSet].filter(dept => !ganttDeptSet.has(dept));
    
    // 結果の設定
    result.addValidData(Array.from(validDepartments));
    
    if (ganttOnlyDepartments.length > 0) {
      result.addWarning(`ガントチャートのみに存在する部署: ${ganttOnlyDepartments.join(", ")}`);
    }
    
    if (rdbOnlyDepartments.length > 0) {
      result.addError(`RDBのみに存在する部署: ${rdbOnlyDepartments.join(", ")}`);
      result.addInvalidData(rdbOnlyDepartments);
    }
    
    if (validDepartments.size === 0) {
      result.addError("処理可能な部署が見つかりません");
    } else {
      result.markSuccess();
    }
    
    return result;
  }

  /**
   * 処理リクエストを検証
   * @param {ShiftDataRequest} request - 処理リクエスト
   * @returns {ValidationResult} 検証結果
   */
  validateProcessingRequest(request) {
    const result = new ValidationResult();
    
    try {
      request.validate();
      result.addValidData(request);
      result.markSuccess();
    } catch (error) {
      result.addError(error.message);
      result.addInvalidData(request);
    }
    
    return result;
  }

  /**
   * 設定の整合性を検証
   * @returns {ValidationResult} 検証結果
   */
  validateConfiguration() {
    const result = new ValidationResult();
    
    try {
      this.config.ensureInitialized();
      
      // 各インデックスの有効性チェック
      const rdbIndexes = this.config.getColumnIndexes('RDB');
      const ganttIndexes = this.config.getColumnIndexes('GANTT');
      const rowIndexes = this.config.getColumnIndexes('ROW');
      
      // 必須インデックスの存在チェック
      const requiredRdbIndexes = ['memberDateId', 'startTime', 'endTime', 'dept'];
      const requiredGanttIndexes = ['memberDateId', 'firstData'];
      const requiredRowIndexes = ['timeScale', 'firstData'];
      
      requiredRdbIndexes.forEach(key => {
        if (rdbIndexes[key] === null || rdbIndexes[key] === undefined) {
          result.addError(`RDBインデックス「${key}」が設定されていません`);
        }
      });
      
      requiredGanttIndexes.forEach(key => {
        if (ganttIndexes[key] === null || ganttIndexes[key] === undefined) {
          result.addError(`Ganttインデックス「${key}」が設定されていません`);
        }
      });
      
      requiredRowIndexes.forEach(key => {
        if (rowIndexes[key] === null || rowIndexes[key] === undefined) {
          result.addError(`行インデックス「${key}」が設定されていません`);
        }
      });
      
      if (result.errors.length === 0) {
        result.addValidData(this.config.getConfig());
        result.markSuccess();
      }
      
    } catch (error) {
      result.addError(`設定の検証中にエラーが発生しました: ${error.message}`);
    }
    
    return result;
  }

  /**
   * データの整合性を検証
   * @param {Object} integrityData - 整合性チェック用データ
   * @returns {Object} 整合性チェック結果
   */
  validateDataIntegrity(integrityData) {
    const {
      inputGanttShiftCount,
      inputRdbShiftCount,
      outputGanttShiftCount,
      outputMergedRdbShiftCount,
      outputConflictShiftCount,
      outputErrorShiftCount
    } = integrityData;

    const result = {
      hasErrors: false,
      message: "",
      details: []
    };

    // 入力データの個数
    const inputTotal = inputGanttShiftCount + inputRdbShiftCount;
    const outputTotal = outputMergedRdbShiftCount + outputConflictShiftCount + outputErrorShiftCount;

    // 結果メッセージの構築
    let message = "■ シフトデータ個数チェック結果\n\n";
    message += "【入力データ】\n";
    message += `・InGanttSs のシフトデータ: ${inputGanttShiftCount}個\n`;
    message += `・InRdbSheet のシフトデータ: ${inputRdbShiftCount}個\n`;
    message += `・入力合計: ${inputTotal}個\n\n`;
    message += "【出力データ】\n";
    message += `・OutGanttSs に書き込み: ${outputGanttShiftCount}個\n`;
    message += `・OutMergedRdbSheet に書き込み: ${outputMergedRdbShiftCount}個\n`;
    message += `・OutConflictRdbSheet に書き込み: ${outputConflictShiftCount}個\n`;
    message += `・OutErrorRdbSheet に書き込み: ${outputErrorShiftCount}個\n`;
    message += `・出力合計: ${outputTotal}個\n`;

    result.message = message;

    // 整合性チェック
    const errors = [];

    // チェック1: OutGanttSsとOutMergedRdbSheetの個数が一致するか
    if (outputGanttShiftCount !== outputMergedRdbShiftCount) {
      const error = `OutGanttSs(${outputGanttShiftCount}個)とOutMergedRdbSheet(${outputMergedRdbShiftCount}個)の個数が一致しません`;
      errors.push(error);
    }

    // チェック2: 入力合計と出力合計が一致するか
    if (inputTotal !== outputTotal) {
      const error = `入力合計(${inputTotal}個)と出力合計(${outputTotal}個)が一致しません`;
      errors.push(error);
    }

    if (errors.length > 0) {
      result.hasErrors = true;
      result.details = errors;
    }

    return result;
  }
}

/**
 * バリデーションルールの管理
 */
class ValidationRules {
  /**
   * 時間形式のバリデーションルール
   * @param {*} timeValue - 検証する時間値
   * @returns {boolean} 検証結果
   */
  static isValidTimeFormat(timeValue) {
    if (!timeValue) return false;
    
    // 既にDateオブジェクトの場合
    if (timeValue instanceof Date) {
      return !isNaN(timeValue.getTime());
    }
    
    // 文字列の場合
    if (typeof timeValue === 'string') {
      const timeMatch = timeValue.match(/^(\d{1,2}):(\d{2})$/);
      if (timeMatch) {
        const hours = parseInt(timeMatch[1], 10);
        const minutes = parseInt(timeMatch[2], 10);
        return hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59;
      }
    }
    
    return false;
  }

  /**
   * memberDateIdのバリデーションルール
   * @param {*} memberDateId - 検証するmemberDateId
   * @returns {boolean} 検証結果
   */
  static isValidMemberDateId(memberDateId) {
    if (!memberDateId) return false;
    
    const str = memberDateId.toString().trim();
    return str.length > 0 && str.length <= 50; // 適切な長さ制限
  }

  /**
   * 部署名のバリデーションルール
   * @param {*} dept - 検証する部署名
   * @returns {boolean} 検証結果
   */
  static isValidDepartment(dept) {
    if (!dept) return false;
    
    const str = dept.toString().trim();
    return str.length > 0 && str.length <= 100; // 適切な長さ制限
  }

  /**
   * 背景色のバリデーションルール
   * @param {*} background - 検証する背景色
   * @returns {boolean} 検証結果
   */
  static isValidBackground(background) {
    if (!background) return true; // 省略可能
    
    const str = background.toString().trim();
    // 16進数カラーコードの形式チェック
    return /^#[0-9A-Fa-f]{6}$/.test(str) || str.toLowerCase() === 'white';
  }

  /**
   * 職種のバリデーションルール
   * @param {*} job - 検証する職種
   * @returns {boolean} 検証結果
   */
  static isValidJob(job) {
    if (!job) return true; // 省略可能
    
    const str = job.toString().trim();
    return str.length <= 50; // 適切な長さ制限
  }
} 