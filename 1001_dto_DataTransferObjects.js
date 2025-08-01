/**
 * データ転送オブジェクト（DTO）クラス群
 * 
 * 複雑な引数リストを構造化されたオブジェクトに置き換え、
 * データの受け渡しを明確にし、コードの可読性を向上させる
 */

/**
 * 処理リクエストの基底クラス
 */
class BaseRequest {
  constructor(data = {}) {
    this.timestamp = new Date();
    this.requestId = this.generateRequestId();
    Object.assign(this, data);
  }

  generateRequestId() {
    return `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  validate() {
    // 子クラスでオーバーライドして実装
    return true;
  }
}

/**
 * 処理結果の基底クラス
 */
class BaseResult {
  constructor() {
    this.timestamp = new Date();
    this.success = false;
    this.errors = [];
    this.warnings = [];
  }

  addError(error) {
    this.errors.push(error);
    this.success = false;
  }

  addWarning(warning) {
    this.warnings.push(warning);
  }

  hasErrors() {
    return this.errors.length > 0;
  }

  hasWarnings() {
    return this.warnings.length > 0;
  }

  markSuccess() {
    this.success = true;
  }
}

/**
 * シフトデータ処理リクエスト
 */
class ShiftDataRequest extends BaseRequest {
  constructor({
    rdbSheet,
    ganttSs,
    outputSheets,
    processingOptions = {}
  }) {
    super();
    this.rdbSheet = rdbSheet;
    this.ganttSs = ganttSs;
    this.outputSheets = outputSheets;
    this.processingOptions = processingOptions;
  }

  validate() {
    if (!this.rdbSheet) {
      throw new ValidationError("RDBシートが指定されていません");
    }
    if (!this.ganttSs) {
      throw new ValidationError("ガントチャートスプレッドシートが指定されていません");
    }
    if (!this.outputSheets) {
      throw new ValidationError("出力シートが指定されていません");
    }
    return true;
  }
}

/**
 * 出力シート情報
 */
class OutputSheets {
  constructor(outMergedRdbSheet, outConflictRdbSheet, outErrorRdbSheet) {
    this.outMergedRdbSheet = outMergedRdbSheet;
    this.outConflictRdbSheet = outConflictRdbSheet;
    this.outErrorRdbSheet = outErrorRdbSheet;
  }
}

/**
 * シフトデータ処理結果
 */
class ShiftDataResult extends BaseResult {
  constructor() {
    super();
    this.ganttData = new Map();
    this.rdbData = [];
    this.conflictData = [];
    this.errorData = [];
    this.processedDepartments = [];
    this.failedDepartments = [];
  }

  addGanttData(department, data) {
    this.ganttData.set(department, data);
  }

  addRdbData(data) {
    this.rdbData.push(...data);
  }

  addConflictData(data) {
    this.conflictData.push(...data);
  }

  addErrorData(data) {
    this.errorData.push(...data);
  }

  addProcessedDepartment(department) {
    this.processedDepartments.push(department);
  }

  addFailedDepartment(department) {
    this.failedDepartments.push(department);
  }

  hasGanttData() {
    return this.ganttData.size > 0;
  }

  getTotalDataCount() {
    return this.rdbData.length + this.conflictData.length + this.errorData.length;
  }
}

/**
 * 部署処理リクエスト
 */
class DepartmentProcessingRequest extends BaseRequest {
  constructor({
    department,
    rdbData,
    ganttData,
    timeHeaders,
    memberDateIdHeaders
  }) {
    super();
    this.department = department;
    this.rdbData = rdbData;
    this.ganttData = ganttData;
    this.timeHeaders = timeHeaders;
    this.memberDateIdHeaders = memberDateIdHeaders;
  }

  validate() {
    if (!this.department) {
      throw new ValidationError("部署名が指定されていません");
    }
    if (!this.ganttData) {
      throw new ValidationError("ガントチャートデータが指定されていません");
    }
    return true;
  }
}

/**
 * 部署処理結果
 */
class DepartmentProcessingResult extends BaseResult {
  constructor() {
    super();
    this.department = null;
    this.ganttHeaderValues = [];
    this.ganttShiftValues = [];
    this.ganttHeaderBgs = [];
    this.ganttShiftBgs = [];
    this.rdbData = [];
    this.conflictData = [];
    this.errorData = [];
    this.firstDataColOffset = 0;
    this.firstDataRowOffset = 0;
  }

  setDepartment(department) {
    this.department = department;
  }

  setGanttData(headerValues, shiftValues, headerBgs, shiftBgs) {
    this.ganttHeaderValues = headerValues;
    this.ganttShiftValues = shiftValues;
    this.ganttHeaderBgs = headerBgs;
    this.ganttShiftBgs = shiftBgs;
  }

  setOffsets(colOffset, rowOffset) {
    this.firstDataColOffset = colOffset;
    this.firstDataRowOffset = rowOffset;
  }

  setRdbData(data) {
    this.rdbData = data;
  }

  setConflictData(data) {
    this.conflictData = data;
  }

  setErrorData(data) {
    this.errorData = data;
  }
}

/**
 * ガントチャートデータ
 */
class GanttChartData {
  constructor(values, backgrounds) {
    this.values = values;
    this.backgrounds = backgrounds;
  }

  isEmpty() {
    return !this.values || this.values.length === 0 || 
           (this.values.length === 1 && this.values[0].length === 0);
  }
}

/**
 * シフトデータ
 */
class ShiftData {
  constructor({
    memberDateId,
    startTime,
    endTime,
    job,
    dept,
    background = "#FFFFFF",
    source = "Unknown"
  }) {
    this.memberDateId = memberDateId;
    this.startTime = startTime;
    this.endTime = endTime;
    this.job = job;
    this.dept = dept;
    this.background = background;
    this.source = source;
    this.shiftId = this.generateShiftId();
  }

  generateShiftId() {
    const startTime = this.startTime instanceof Date ? this.startTime.getTime() : this.startTime;
    const endTime = this.endTime instanceof Date ? this.endTime.getTime() : this.endTime;
    return `${this.source.toLowerCase()}_${this.memberDateId}_${startTime}_${endTime}_${this.job}`;
  }

  validate() {
    const errors = [];

    if (!this.memberDateId || this.memberDateId.toString().trim() === "") {
      errors.push("memberDateIdが空です");
    }

    if (!this.startTime) {
      errors.push("startTimeが空です");
    }

    if (!this.endTime) {
      errors.push("endTimeが空です");
    }

    if (this.startTime && this.endTime && this.startTime >= this.endTime) {
      errors.push("startTimeがendTime以降の時刻です");
    }

    if (!this.dept || this.dept.toString().trim() === "") {
      errors.push("deptが空です");
    }

    return errors;
  }

  isValid() {
    return this.validate().length === 0;
  }

  toRdbArray(columnOrder) {
    return columnOrder.map(key => {
      if (key === 'startTime' || key === 'endTime') {
        return this.formatTimeToHHMM(this[key]);
      }
      return this[key] || "";
    });
  }

  toConflictArray(columnOrder) {
    return columnOrder.map(key => {
      if (key === 'startTime' || key === 'endTime') {
        return this.formatTimeToHHMM(this[key]);
      }
      return this[key] || "";
    });
  }

  toErrorArray(columnOrder, errorMessage) {
    return columnOrder.map(key => {
      if (key === 'startTime' || key === 'endTime') {
        return this.formatTimeToHHMM(this[key]);
      }
      if (key === 'errorMessage') {
        return errorMessage || "";
      }
      return this[key] || "";
    });
  }

  formatTimeToHHMM(timeValue) {
    if (!timeValue || !(timeValue instanceof Date) || isNaN(timeValue.getTime())) {
      return "";
    }
    
    const hours = timeValue.getHours().toString().padStart(2, '0');
    const minutes = timeValue.getMinutes().toString().padStart(2, '0');
    return `${hours}:${minutes}`;
  }
}

/**
 * バリデーション結果
 */
class ValidationResult extends BaseResult {
  constructor() {
    super();
    this.validData = [];
    this.invalidData = [];
  }

  addValidData(data) {
    this.validData.push(data);
  }

  addInvalidData(data) {
    this.invalidData.push(data);
  }

  getValidCount() {
    return this.validData.length;
  }

  getInvalidCount() {
    return this.invalidData.length;
  }

  getTotalCount() {
    return this.validData.length + this.invalidData.length;
  }
}

/**
 * 時間処理ユーティリティ
 */
class TimeUtils {
  /**
   * 時間値をDateオブジェクトに変換
   * @param {*} timeValue - 時間値
   * @returns {Date} Dateオブジェクト
   */
  static parseTimeToDate(timeValue) {
    if (timeValue instanceof Date) {
      return timeValue;
    }
    
    if (typeof timeValue === 'string') {
      const timeMatch = timeValue.match(/^(\d{1,2}):(\d{2})$/);
      if (timeMatch) {
        const hours = parseInt(timeMatch[1], 10);
        const minutes = parseInt(timeMatch[2], 10);
        return new Date(1970, 0, 1, hours, minutes, 0, 0);
      }
    }
    
    const date = new Date(timeValue);
    if (isNaN(date.getTime())) {
      throw new ValidationError(`無効な時刻形式です: ${timeValue}`);
    }
    return date;
  }

  /**
   * Dateオブジェクトをh:mm形式の文字列に変換
   * @param {Date} dateValue - Dateオブジェクト
   * @returns {string} h:mm形式の文字列
   */
  static formatTimeToHHMM(dateValue) {
    if (!dateValue || !(dateValue instanceof Date) || isNaN(dateValue.getTime())) {
      return "";
    }
    
    const hours = dateValue.getHours().toString().padStart(2, '0');
    const minutes = dateValue.getMinutes().toString().padStart(2, '0');
    return `${hours}:${minutes}`;
  }
}

/**
 * エラークラス群
 */
class ValidationError extends Error {
  constructor(message, field = null) {
    super(message);
    this.name = "ValidationError";
    this.field = field;
  }
}

class DataProcessingError extends Error {
  constructor(message, data = null) {
    super(message);
    this.name = "DataProcessingError";
    this.data = data;
  }
}

class ConfigurationError extends Error {
  constructor(message, configKey = null) {
    super(message);
    this.name = "ConfigurationError";
    this.configKey = configKey;
  }
}

class SheetUpdateError extends Error {
  constructor(message, sheetName = null) {
    super(message);
    this.name = "SheetUpdateError";
    this.sheetName = sheetName;
  }
} 