// 列・行インデックス定数オブジェクト
const RDB_COL_INDEXES = {
  job: 0,
  memberDateId: 1,
  dept: 2,
  startTime: 3,
  endTime: 4,
  background: 5,
  source: 6,
};

const GANTT_COL_INDEXES = {
  memberDateId: 12,
  firstData: 13,
};

const GANTT_ROW_INDEXES = {
  timeScale: 2,
  firstData: 3,
};

const CONFLICT_COL_INDEXES = {
  job: 0,
  memberDateId: 1,
  dept: 2,
  startTime: 3,
  endTime: 4,
  background: 5,
  source: 6,
};

// インデックスオブジェクトから列順序配列を生成する関数
function getColumnOrder(indexes) {
  // 最大のインデックス値を取得
  const maxIndex = Math.max(...Object.values(indexes));

  // インデックス値をキーとしてプロパティ名を格納するマップを作成
  const indexToKey = new Map();
  Object.entries(indexes).forEach(([key, index]) => {
    indexToKey.set(index, key);
  });

  // 0からmaxIndexまでの配列を生成し、欠番には""をセット
  return Array.from({ length: maxIndex + 1 }, (_, i) => {
    return indexToKey.has(i) ? indexToKey.get(i) : "";
  });
}

// 従来のスプレッドシート依存関数は削除済み
// 全てのインデックスは定数オブジェクトで管理
