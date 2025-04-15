// データベースの列設定（将来的にはフロントエンドから渡される）
const RDB_COLUMN_CONFIG = {
  columns: [
    {
      id: "job",
      label: "業務内容",
      type: "string",
      required: true,
      editable: true,
    },
    {
      id: "memberDateId",
      label: "メンバーID",
      type: "string",
      required: true,
      editable: false,
    },
    {
      id: "startTime",
      label: "開始時間",
      type: "datetime",
      required: true,
      editable: true,
    },
    {
      id: "endTime",
      label: "終了時間",
      type: "datetime",
      required: true,
      editable: true,
    },
    {
      id: "background",
      label: "背景色",
      type: "string",
      required: false,
      editable: true,
    },
    {
      id: "source",
      label: "データソース",
      type: "string",
      required: false,
      editable: false,
    },
  ],
  order: ["job", "memberDateId", "startTime", "endTime", "background", "source"],
};

// ガントチャートの列設定
const GANTT_COLUMN_CONFIG = {
  columns: [
    {
      id: "a",
      label: "A列",
      type: "string",
      required: false,
      editable: false,
    },
    {
      id: "memberDateId",
      label: "メンバーID",
      type: "string",
      required: true,
      editable: false,
    },
    {
      id: "firstData",
      label: "最初のデータ列",
      type: "string",
      required: false,
      editable: false,
    },
  ],
  order: ["a", "memberDateId", "firstData"],
};

// ガントチャートの行設定
const GANTT_ROW_CONFIG = {
  columns: [
    {
      id: "hour",
      label: "時間",
      type: "string",
      required: false,
      editable: false,
    },
    {
      id: "minute",
      label: "分",
      type: "string",
      required: false,
      editable: false,
    },
    {
      id: "timeScale",
      label: "時間ヘッダー",
      type: "string",
      required: false,
      editable: false,
    },
    {
      id: "firstData",
      label: "最初のデータ行",
      type: "string",
      required: false,
      editable: false,
    },
  ],
  order: ["hour", "minute", "timeScale", "firstData"],
};

// コンフリクト列設定
const CONFLICT_COLUMN_CONFIG = {
  columns: [
    {
      id: "memberDateId",
      label: "メンバーID",
      type: "string",
      required: true,
      editable: false,
    },
    {
      id: "job",
      label: "業務内容",
      type: "string",
      required: true,
      editable: false,
    },
    {
      id: "startTime",
      label: "開始時間",
      type: "datetime",
      required: true,
      editable: false,
    },
    {
      id: "endTime",
      label: "終了時間",
      type: "datetime",
      required: true,
      editable: false,
    },
    {
      id: "source",
      label: "データソース",
      type: "string",
      required: false,
      editable: false,
    },
  ],
  order: ["memberDateId", "job", "startTime", "endTime", "source"],
};

// 列管理クラス
class ColumnManager {
  constructor(config) {
    this.config = config;
    this.columnIndexes = {};
    this.initializeIndexes();
  }

  initializeIndexes() {
    this.config.order.forEach((columnId, index) => {
      this.columnIndexes[columnId] = index;
    });
  }

  getColumnIndex(columnId) {
    return this.columnIndexes[columnId];
  }

  getColumnConfig(columnId) {
    return this.config.columns.find((col) => col.id === columnId);
  }

  getValue(row, columnId) {
    return row[this.getColumnIndex(columnId)];
  }

  setValue(row, columnId, value) {
    const columnConfig = this.getColumnConfig(columnId);
    if (columnConfig && this.validateValue(value, columnConfig)) {
      row[this.getColumnIndex(columnId)] = value;
      return true;
    }
    return false;
  }

  validateValue(value, columnConfig) {
    if (columnConfig.required && !value) return false;
    // 型チェックなどの追加の検証ロジック
    return true;
  }
}

// 各設定のマネージャーを生成する関数
function createColumnManagers() {
  return {
    rdbCol: new ColumnManager(RDB_COLUMN_CONFIG),
    ganttCol: new ColumnManager(GANTT_COLUMN_CONFIG),
    ganttRow: new ColumnManager(GANTT_ROW_CONFIG),
    conflictCol: new ColumnManager(CONFLICT_COLUMN_CONFIG),
  };
}

// 既存の定数を新しい形式に変換する関数
function convertToNewFormat() {
  const managers = createColumnManagers();

  return {
    rdbColManager: managers.rdbCol,
    ganttColManager: managers.ganttCol,
    ganttRowManager: managers.ganttRow,
    conflictColManager: managers.conflictCol,
  };
}

// フロントエンド用の関数（将来的に使用）
function doGet(e) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <script>
          function sendColumnConfig() {
            const columnConfig = ${JSON.stringify(TEST_COLUMN_CONFIG)};
            google.script.run.processData(columnConfig);
          }
        </script>
      </head>
      <body>
        <button onclick="sendColumnConfig()">設定を送信</button>
      </body>
    </html>
  `);
}

// フロントエンドから設定を受け取る関数（将来的に使用）
function processData(columnConfig) {
  const columnManager = new ColumnManager(columnConfig);
  // 設定に基づいて処理を実行
  return columnManager;
}
