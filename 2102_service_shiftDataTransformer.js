function splitGanttData(ganttValues, ganttBgs) {
    const firstDataCol = GANTT_COL_INDEXES.firstData;
    const firstDataRow = GANTT_ROW_INDEXES.firstData;

    // シフトデータ部分
    const ganttShiftValues = ganttValues.slice(firstDataRow).map((row) => row.slice(firstDataCol));
    const ganttShiftBgs = ganttBgs.slice(firstDataRow).map((row) => row.slice(firstDataCol));

    // ヘッダー部分（「の形）
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

    // 削除前にtimeScaleとmemberDateIdの元のインデックスを取得
    const originalTimeRow = GANTT_ROW_INDEXES.timeScale;
    const originalMemberDateIdCol = GANTT_COL_INDEXES.memberDateId;
  
    // timescale,memberDateIdのリストを作成（元のインデックスを使用）
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
    
    const memberDateIdHeaders = ganttValues.slice(firstDataRow).map((row) => row[originalMemberDateIdCol]);
  
    return {
      ganttHeaderValues,
      ganttShiftValues,
      ganttHeaderBgs,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      // オフセット計算用の情報を追加
      firstDataColOffset: firstDataCol,
      firstDataRowOffset: firstDataRow,
    };
  }
  
  function convertObjsTo2dAry(
  validShiftsMap,
  conflictShiftObjs,
  timeHeaders,
  memberDateIdHeaders
) {
    // rdbDataとconflictDataのヘッダー行を追加
    const rdbData = [];
    const conflictData = [];
    
    // Mapからrdbデータを直接生成（中間変換なし）
    const processedShiftIds = new Set();
    
    // ganttData用のmemberMap（既に作成済み）
    const ganttValueMap = new Map();
    // 背景色用のmemberBgMap（新規追加）
    const ganttBgMap = new Map();
    
    // timeHeadersから最後の追加要素を除外（endTime計算用に追加されたもの）
    const originalTimeHeadersLength = timeHeaders.length - 1;
    
    // 各メンバーのシフト情報を処理
    for (const [memberId, timeMap] of validShiftsMap.entries()) {
      // 各時間スロットごとに処理
      for (const [timeKey, shiftInfo] of timeMap.entries()) {
                  // まだ処理していないシフトIDの場合のみrdbDataに追加
          if (!processedShiftIds.has(shiftInfo.shiftId)) {
            const rdbRow = getColumnOrder(RDB_COL_INDEXES).map(key => shiftInfo[key]);
            rdbData.push(rdbRow);
          processedShiftIds.add(shiftInfo.shiftId);
          
          // ganttData用のデータも準備（元の列数で初期化）
          if (!ganttValueMap.has(shiftInfo.memberDateId)) {
            ganttValueMap.set(shiftInfo.memberDateId, Array(originalTimeHeadersLength).fill(""));
            ganttBgMap.set(shiftInfo.memberDateId, Array(originalTimeHeadersLength).fill("#FFFFFF")); // 背景色の初期値は白
          }
          
          const timeRow = ganttValueMap.get(shiftInfo.memberDateId);
          const bgRow = ganttBgMap.get(shiftInfo.memberDateId);
          const startIndex = findTimeIndex(timeHeaders, shiftInfo.startTime);
          const endIndex = findTimeIndex(timeHeaders, shiftInfo.endTime);
          
          if (startIndex !== -1 && endIndex !== -1) {
            // 元の列数の範囲内でのみシフトデータを設定
            for (let i = startIndex; i < Math.min(endIndex, originalTimeHeadersLength); i++) {
              timeRow[i] = shiftInfo.job;
              // 背景色も設定
              bgRow[i] = shiftInfo.background || "#FFFFFF";
            }
          }
        }
      }
    }
  
    // 元のmemberDateIdHeadersの順序を保持してganttDataを生成（空白行も含む）
    const ganttValues = memberDateIdHeaders.map(memberId => {
      if (ganttValueMap.has(memberId)) {
        return ganttValueMap.get(memberId);
      } else {
        // 空白行の場合は空の配列を返す
        return Array(originalTimeHeadersLength).fill("");
      }
    });
    
    // 背景色も同様に元の順序を保持
    const ganttBgs = memberDateIdHeaders.map(memberId => {
      if (ganttBgMap.has(memberId)) {
        return ganttBgMap.get(memberId);
      } else {
        // 空白行の場合は白背景の配列を返す
        return Array(originalTimeHeadersLength).fill("#FFFFFF");
      }
    });
  
          // コンフリクトデータを処理（エラーデータは既に分離済み）
    conflictShiftObjs.forEach((shiftObj) => {
      const conflictRow = getColumnOrder(CONFLICT_COL_INDEXES).map((key) => shiftObj[key]);
      conflictData.push(conflictRow);
    });

    return {
      ganttValues,
      ganttBgs,
      rdbData,
      conflictData,
    };
  }
  
  // 時間ヘッダー配列から指定時間に最も近いインデックスを見つける
  function findTimeIndex(timeHeaders, time) {
    const timeStr = time.toISOString().slice(11, 16);
    for (let i = 0; i < timeHeaders.length; i++) {
      const headerTime = new Date(timeHeaders[i]).toISOString().slice(11, 16);
      if (headerTime === timeStr) {
        return i;
      }
    }
    return -1;
  }
  
  function mergeGanttData(
    ganttHeaderValues,
    ganttShiftValues,
    ganttHeaderBgs,
    ganttShiftBgs,
    firstDataColOffset,
    firstDataRowOffset
  ) {
    // オフセットを使用してfirstDataの位置を計算
    const firstDataCol = firstDataColOffset;
    const firstDataRow = firstDataRowOffset;
  
    // 結合後のデータを格納する配列
    const mergedValues = [];
    const mergedBgs = [];
  
    // 上部ヘッダー行を追加（そのまま）
    for (let i = 0; i < firstDataRow; i++) {
      mergedValues.push(ganttHeaderValues[i]);
      mergedBgs.push(ganttHeaderBgs[i]);
    }
  
    // 左側ヘッダー列とシフトデータを結合して追加
    for (let i = 0; i < ganttShiftValues.length; i++) {
      const headerRow = ganttHeaderValues[i + firstDataRow];
      const bgHeaderRow = ganttHeaderBgs[i + firstDataRow];
  
      mergedValues.push([...headerRow, ...ganttShiftValues[i]]);
      mergedBgs.push([...bgHeaderRow, ...ganttShiftBgs[i]]);
    }
  
    return {
      values: mergedValues,
      backgrounds: mergedBgs,
    };
  }