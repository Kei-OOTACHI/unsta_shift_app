const MEMBER_DATA_SHEET_NAME = "メンバーリスト";

function sortMemberDataByHeaders(dataArray, headerOrder) {
  // 1行目を見出しとして取得
  const headers = dataArray[0];
  
  // 新しい配列の初期化
  const sortedArray = [];
  
  // 新しい見出し行を作成
  const newHeaders = headerOrder.map(header => {
    if (headers.includes(header)) {
      return header;
    } else {
      return ''; // 空白列を挿入
    }
  });
  
  // 新しい見出し行を追加
  sortedArray.push(newHeaders);
  
  // データ行を並べ替え
  for (let i = 1; i < dataArray.length; i++) {
    const row = dataArray[i];
    const newRow = headerOrder.map(header => {
      const index = headers.indexOf(header);
      if (index !== -1) {
        return row[index];
      } else {
        return ''; // 空白列を挿入
      }
    });
    sortedArray.push(newRow);
  }
  
  return sortedArray;
}

function duplicateMemberDataRows(dataArray, duplicateCount, insertBlankLine) {
  const resultArray = [];
  
  dataArray.forEach(row => {
    // 各行を指定された回数だけ複製
    for (let i = 0; i < duplicateCount; i++) {
      resultArray.push([...row]); // スプレッド演算子で配列をコピー
    }
    
    // 複製された行のまとまりの間に空白行を挿入
    if (insertBlankLine) {
      resultArray.push(new Array(row.length).fill(''));
    }
  });
  
  // 最後に追加された空白行を削除（必要であれば）
  if (insertBlankLine && resultArray.length > 0) {
    resultArray.pop();
  }
  
  return resultArray;
}