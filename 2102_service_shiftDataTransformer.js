
function splitGanttData(ganttValues, ganttBgs, ganttColManager, ganttRowManager) {
    const firstDataCol = ganttColManager.getColumnIndex("firstData");
    const firstDataRow = ganttRowManager.getColumnIndex("firstData");
  
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
  
    // ganttColManagerのorderを修正（firstDataより前の要素を削除）
    const originalColOrder = [...ganttColManager.config.order];
    const firstDataColIndex = ganttColManager.config.order.indexOf("firstData");
    const adjustedColOrder = ganttColManager.config.order.slice(firstDataColIndex);
  
    // ganttRowManagerのorderを修正（firstDataより前の要素を削除）
    const originalRowOrder = [...ganttRowManager.config.order];
    const firstDataRowIndex = ganttRowManager.config.order.indexOf("firstData");
    const adjustedRowOrder = ganttRowManager.config.order.slice(firstDataRowIndex);
  
    // 修正したorderで設定を更新
    ganttColManager.config.order = adjustedColOrder;
    ganttRowManager.config.order = adjustedRowOrder;
  
    // インデックスを再初期化
    ganttColManager.initializeIndexes();
    ganttRowManager.initializeIndexes();
  
    // timescale,memberDateIdのリストを作成
    const timeRow = ganttRowManager.getColumnIndex("timeScale");
    const memberDateIdCol = ganttColManager.getColumnIndex("memberDateId");
  
    const timeHeaders = ganttValues[timeRow].slice(firstDataCol);
    const memberDateIdHeaders = ganttValues.slice(firstDataRow).map((row) => row[memberDateIdCol]);
  
    return {
      ganttHeaderValues,
      ganttShiftValues,
      ganttHeaderBgs,
      ganttShiftBgs,
      timeHeaders,
      memberDateIdHeaders,
      originalColOrder,
      originalRowOrder,
    };
  }
  
  function convertObjsTo2dAry(
    validShiftsMap,
    conflictShiftObjs,
    timeHeaders,
    rdbColManager,
    ganttColManager,
    conflictColManager
  ) {
    // rdbDataとconflictDataのヘッダー行を追加
    const rdbData = [rdbColManager.config.order.slice()];
    const conflictData = [conflictColManager.config.order.slice()];
    
    // Mapからrdbデータを直接生成（中間変換なし）
    const processedShiftIds = new Set();
    
    // ganttData用のmemberMap（既に作成済み）
    const ganttValueMap = new Map();
    // 背景色用のmemberBgMap（新規追加）
    const ganttBgMap = new Map();
    
    // 各メンバーのシフト情報を処理
    for (const [memberId, timeMap] of validShiftsMap.entries()) {
      // 各時間スロットごとに処理
      for (const [timeKey, shiftInfo] of timeMap.entries()) {
        // まだ処理していないシフトIDの場合のみrdbDataに追加
        if (!processedShiftIds.has(shiftInfo.shiftId)) {
          const rdbRow = rdbColManager.config.order.map(key => shiftInfo[key]);
          rdbData.push(rdbRow);
          processedShiftIds.add(shiftInfo.shiftId);
          
          // ganttData用のデータも準備
          if (!ganttValueMap.has(shiftInfo.memberDateId)) {
            ganttValueMap.set(shiftInfo.memberDateId, Array(timeHeaders.length).fill(""));
            ganttBgMap.set(shiftInfo.memberDateId, Array(timeHeaders.length).fill("#FFFFFF")); // 背景色の初期値は白
          }
          
          const timeRow = ganttValueMap.get(shiftInfo.memberDateId);
          const bgRow = ganttBgMap.get(shiftInfo.memberDateId);
          const startIndex = findTimeIndex(timeHeaders, shiftInfo.startTime);
          const endIndex = findTimeIndex(timeHeaders, shiftInfo.endTime);
          
          if (startIndex !== -1 && endIndex !== -1) {
            for (let i = startIndex; i < endIndex; i++) {
              timeRow[i] = shiftInfo.job;
              // 背景色も設定
              bgRow[i] = shiftInfo.background || "#FFFFFF";
            }
          }
        }
      }
    }
  
    // マップから直接ganttDataに変換
    const ganttValues = Array.from(ganttValueMap.values());
    // 背景色の2次元配列も生成
    const ganttBgs = Array.from(ganttBgMap.values());
  
    // コンフリクトデータを処理
    conflictShiftObjs.forEach((shiftObj) => {
      const conflictRow = conflictColManager.config.order.map((key) => shiftObj[key]);
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
    ganttColManager,
    ganttRowManager,
    originalColOrder,
    originalRowOrder
  ) {
    const firstDataCol = ganttColManager.getColumnIndex("firstData");
    const firstDataRow = ganttRowManager.getColumnIndex("firstData");
  
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
  
    // managerのorderを元に戻す
    ganttColManager.config.order = originalColOrder;
    ganttRowManager.config.order = originalRowOrder;
  
    // インデックスを再初期化
    ganttColManager.initializeIndexes();
    ganttRowManager.initializeIndexes();
  
    return {
      values: mergedValues,
      backgrounds: mergedBgs,
    };
  }