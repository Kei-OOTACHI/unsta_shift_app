function convert2dAryToObjsAndJoin(ganttValue, ganttBg, timeHeaders, memberDateIdHeaders, rdbData, deptName) {
    // memberIdをキーとする時間枠データのマップを作成（ガントチャートのデータ）
    const { validShiftsMap: shiftsFromGanttMap, errorShifts: ganttErrorShifts } = gantt2dAryToMap(ganttValue, ganttBg, timeHeaders, memberDateIdHeaders, deptName);
  
    // RDBからのデータもmemberIdをキー、時間枠をマップとして変換
    const shiftsFromRdbMap = rdb2dAryToMap(rdbData, timeHeaders);
  
    // 重複検出と解決
    const { validShiftsMap, conflictShiftObjs } = detectConflictsWithMap(
      shiftsFromGanttMap,
      shiftsFromRdbMap,
      timeHeaders
    );
  
    // エラーデータとコンフリクトデータを分離して返す
    return { 
      validShiftsMap, 
      conflictShiftObjs, 
      errorShifts: ganttErrorShifts 
    };
  }
  
  // ガントチャートデータをmemberIdをキーとするマップに変換
  function gantt2dAryToMap(ganttValue, ganttBg, timeHeaders, memberDateIdHeaders, deptName) {
    // memberIdをキーとし、時間ごとのシフト情報を格納するマップ
    const shiftsMap = new Map();
    const errorShifts = [];
  
    // 各行を走査
    for (let i = 0; i < ganttValue.length; i++) {
      const row = ganttValue[i];
      const bgRow = ganttBg[i];
      const memberId = memberDateIdHeaders[i];
      const errorMessages = [];
  
      // memberIdのバリデーション
      if (!memberId || memberId.toString().trim() === "") {
        // memberIdが無効な場合、この行にシフトデータがあるかチェック
        const hasShiftData = row.some(cell => cell !== "" && cell !== null && cell !== undefined);
        if (hasShiftData) {
          errorMessages.push("memberDateIdが空または無効です");
        } else {
          continue; // シフトデータがない場合はスキップ
        }
      }
  
      // エラーメッセージがある場合はエラーシフトとして記録
      if (errorMessages.length > 0) {
        // 行全体の情報を使ってエラーシフトを作成
        const errorShift = {
          job: "",
          dept: deptName,
          background: "",
          source: "Gantt",
          memberDateId: memberId || "",
          startTime: "",
          endTime: "",
          errorMessage: errorMessages.join("、")
        };
        errorShifts.push(errorShift);
        continue; // エラーがある行は処理をスキップ
      }

      // このメンバーのマップがなければ初期化
      if (!shiftsMap.has(memberId)) {
        shiftsMap.set(memberId, new Map());
      }
  
      // このメンバーの時間枠ごとのマップ
      const memberTimeMap = shiftsMap.get(memberId);
  
      let j = 0;
      while (j < row.length) {
                // ガントチャートの棒セルを検出
        // 値が存在するか、背景色が白以外の場合にシフトデータとして処理
        const cellValue = row[j];
        const cellBg = bgRow[j];
        const hasValue = cellValue !== "" && cellValue !== null && cellValue !== undefined;
        const hasNonWhiteBg = cellBg && cellBg.toLowerCase() !== "#ffffff";
        
        if (hasValue || hasNonWhiteBg) {
          const startCol = j;

          // 横方向に同じ値＆背景色が続く間、同じシフトとして扱う
          while (j < row.length && row[j] === cellValue && bgRow[j] && bgRow[j] === cellBg) {
            j++;
          }
  
          const endCol = j - 1;
          
          // 空のシフトデータ（値が空で背景色が白）は除外
          const isEmpty = (!cellValue || cellValue === "") && 
                         (!cellBg || cellBg.toLowerCase() === "#ffffff");
          if (isEmpty) {
            continue;
          }
          
          // 時間のバリデーション
          if (startCol >= timeHeaders.length || endCol + 1 >= timeHeaders.length) {
            // 時間範囲が無効な場合はエラーシフトとして記録
            const errorShift = {
              job: cellValue,
              dept: deptName,
              background: cellBg,
              source: "Gantt",
              memberDateId: memberId,
              startTime: "",
              endTime: "",
              errorMessage: "時間範囲が無効です（timeHeadersの範囲外）"
            };
            errorShifts.push(errorShift);
            continue;
          }
          
          const startTime = new Date(timeHeaders[startCol]);
          const endTime = new Date(timeHeaders[endCol + 1]);
          
          // 時間の有効性チェック
          if (isNaN(startTime.getTime()) || isNaN(endTime.getTime())) {
            const errorShift = {
              job: cellValue,
              dept: deptName,
              background: cellBg,
              source: "Gantt",
              memberDateId: memberId,
              startTime: timeHeaders[startCol],
              endTime: timeHeaders[endCol + 1],
              errorMessage: "startTimeまたはendTimeが無効な日付です"
            };
            errorShifts.push(errorShift);
            continue;
          }
          
          // startTimeとendTimeの順序チェック
          if (startTime >= endTime) {
            const errorShift = {
              job: cellValue,
              dept: deptName,
              background: cellBg,
              source: "Gantt",
              memberDateId: memberId,
              startTime,
              endTime,
              errorMessage: "startTimeがendTime以降の時刻です"
            };
            errorShifts.push(errorShift);
            continue;
          }
  
          // シフト全体の情報を作成
          const shiftInfo = {
            job: cellValue,
            dept: deptName,
            background: cellBg,
            source: "Gantt",
            memberDateId: memberId,
            startTime,
            endTime,
            // 元のシフト識別用のID
            shiftId: `gantt_${memberId}_${startTime.getTime()}_${endTime.getTime()}_${cellValue}`,
          };
  
          // 時間範囲内の各時間スロットに値を設定（同じシフトIDを持たせる）
          for (let k = startCol; k <= endCol; k++) {
            const timeKey = timeHeaders[k];
            memberTimeMap.set(timeKey, shiftInfo);
          }
        } else {
          j++;
        }
      }
    }
  
    return { validShiftsMap: shiftsMap, errorShifts };
  }
  
  // RDBデータをmemberIdをキーとするマップに変換
  function rdb2dAryToMap(rdbData, timeHeaders) {
    const shiftsMap = new Map();
    
    // 列インデックスを直接参照
    const memberDateIdIndex = RDB_COL_INDEXES.memberDateId;
    const deptIndex = RDB_COL_INDEXES.dept;
    const jobIndex = RDB_COL_INDEXES.job;
    const startTimeIndex = RDB_COL_INDEXES.startTime;
    const endTimeIndex = RDB_COL_INDEXES.endTime;
    const backgroundIndex = RDB_COL_INDEXES.background;
    
    // ヘッダー行をスキップ
    for (let i = 1; i < rdbData.length; i++) {
      const row = rdbData[i];
      const memberId = row[memberDateIdIndex];
      const dept = row[deptIndex];
      const job = row[jobIndex];
      const startTime = new Date(row[startTimeIndex]);
      const endTime = new Date(row[endTimeIndex]);
      const background = row[backgroundIndex];
      
      if (!memberId) continue;
      
      // このメンバーのマップがなければ初期化
      if (!shiftsMap.has(memberId)) {
        shiftsMap.set(memberId, new Map());
      }
      
      const memberTimeMap = shiftsMap.get(memberId);
      
      // シフト全体の情報を作成
      const shiftInfo = {
        job,
        dept,
        background,
        source: "RDB",
        memberDateId: memberId,
        startTime,
        endTime,
        // 元のシフト識別用のID
        shiftId: `rdb_${memberId}_${startTime.getTime()}_${endTime.getTime()}_${job}`
      };
      
      // 該当する時間範囲内の各時間スロットに値を設定（同じシフトIDを持たせる）
      for (let j = 0; j < timeHeaders.length; j++) {
        const timeSlot = new Date(timeHeaders[j]);
        if (timeSlot >= startTime && timeSlot < endTime) {
          memberTimeMap.set(timeHeaders[j], shiftInfo);
        }
      }
    }
    
    return shiftsMap;
  }
  
  // マップを使用した重複検出と解決
  function detectConflictsWithMap(ganttMap, rdbMap, timeHeaders) {
    const validShiftsMap = new Map();
    const conflictShiftIds = new Set(); // 重複したシフトのIDを保存
    const allMemberIds = new Set([...ganttMap.keys(), ...rdbMap.keys()]);
  
    // 各メンバーについて処理
    for (const memberId of allMemberIds) {
      const ganttTimeMap = ganttMap.get(memberId) || new Map();
      const rdbTimeMap = rdbMap.get(memberId) || new Map();
  
      // このメンバーの有効なシフト情報を保持するマップ
      if (!validShiftsMap.has(memberId)) {
        validShiftsMap.set(memberId, new Map());
      }
      const memberValidMap = validShiftsMap.get(memberId);
  
      // 同一メンバー内の重複シフトを追跡するセット
      const conflictingShiftIds = new Set();
  
      // 各時間スロットの処理
      for (let i = 0; i < timeHeaders.length; i++) {
        const timeKey = timeHeaders[i];
        const ganttShift = ganttTimeMap.get(timeKey);
        const rdbShift = rdbTimeMap.get(timeKey);
  
        // 重複の検出
        if (ganttShift && rdbShift) {
          // 両方のソースに存在する場合
          if (ganttShift.job !== rdbShift.job || ganttShift.background !== rdbShift.background) {
            // 情報が異なる場合は両方のシフト全体をコンフリクトとしてマーク
            conflictingShiftIds.add(ganttShift.shiftId);
            conflictingShiftIds.add(rdbShift.shiftId);
          } else {
            // 情報が一致する場合はガントの情報を有効シフトに追加
            // ただし既にコンフリクトしているシフトの一部ならスキップ
            if (!conflictingShiftIds.has(ganttShift.shiftId)) {
              memberValidMap.set(timeKey, ganttShift);
            }
          }
        } else if (ganttShift) {
          // ガントチャートのみに存在（かつコンフリクトしていなければ）
          if (!conflictingShiftIds.has(ganttShift.shiftId)) {
            memberValidMap.set(timeKey, ganttShift);
          }
        } else if (rdbShift) {
          // RDBのみに存在（かつコンフリクトしていなければ）
          if (!conflictingShiftIds.has(rdbShift.shiftId)) {
            memberValidMap.set(timeKey, rdbShift);
          }
        }
      }
  
      // コンフリクトシフトIDを全体のセットに追加
      for (const id of conflictingShiftIds) {
        conflictShiftIds.add(id);
      }
    }
  
    // 重複シフトの実際のオブジェクトを収集
    const conflictShiftObjs = [];
  
    // ガントマップから重複シフトを収集
    for (const [memberId, timeMap] of ganttMap.entries()) {
      const processedShiftIds = new Set();
  
      for (const shiftInfo of timeMap.values()) {
        if (conflictShiftIds.has(shiftInfo.shiftId) && !processedShiftIds.has(shiftInfo.shiftId)) {
          conflictShiftObjs.push(shiftInfo);
          processedShiftIds.add(shiftInfo.shiftId);
        }
      }
    }
  
    // RDBマップから重複シフトを収集
    for (const [memberId, timeMap] of rdbMap.entries()) {
      const processedShiftIds = new Set();
  
      for (const shiftInfo of timeMap.values()) {
        if (conflictShiftIds.has(shiftInfo.shiftId) && !processedShiftIds.has(shiftInfo.shiftId)) {
          conflictShiftObjs.push(shiftInfo);
          processedShiftIds.add(shiftInfo.shiftId);
        }
      }
    }
  
    return { validShiftsMap, conflictShiftObjs };
  }
  