# unsta_shift_app
gannt_chart_app の改良版です。リファクタリングでフロントエンドとバックエンドを分けることに挑戦しました。

# 動的カスタムポップアップウィンドウ

Google Apps Script で動的にフォーム入力欄を生成し、カスタムポップアップウィンドウを表示するサンプルコードです。

## 機能

- 設定オブジェクトに基づいて動的にフォーム要素を生成
- 必須項目のバリデーション
- タイプに応じた入力値検証（string, number, email, url, date, time, tel など）
- 非同期処理でユーザー入力値を取得（Promise形式）
- CacheServiceを活用したデータ管理

## 使用方法

1. GASプロジェクトに以下のファイルを追加：
   - `CustomPopup.gs`
   - `CustomPopupTemplate.html`

2. 以下のようにフィールド設定を定義：

```javascript
const fieldConfigs = [
  {
    id: "job",           // フィールドのID（必須）
    label: "業務内容",    // 表示ラベル（必須）
    type: "string",      // 入力タイプ
    required: true,      // 必須かどうか
    value: ""            // 初期値（省略可）
  },
  {
    id: "memberDateId",
    label: "メンバーID",
    type: "string",
    required: true,
    value: "9901"
  },
  {
    id: "email",
    label: "メールアドレス",
    type: "email",       // メールアドレス検証
    required: true
  },
  {
    id: "website",
    label: "ウェブサイト",
    type: "url",         // URL検証
    required: false
  },
  {
    id: "age",
    label: "年齢",
    type: "number",      // 数値検証
    min: 18,             // 最小値（省略可）
    max: 100,            // 最大値（省略可）
    step: 1,             // 増減値（省略可）
    required: false
  },
  {
    id: "startDate",
    label: "開始日",
    type: "date",        // 日付検証
    required: false
  },
  {
    id: "phone",
    label: "電話番号",
    type: "tel",         // 電話番号検証
    required: false
  }
];
```

3. ダイアログを表示して結果を取得：

```javascript
// Promise形式で使用
showCustomInputDialog(fieldConfigs)
  .then(result => {
    console.log('入力されたデータ:', result); 
    // 結果は {job: "入力値", memberDateId: "9901", ...} のような形式
  })
  .catch(error => {
    console.error('エラー:', error);
  });
```

## サポートするフィールドタイプ

- `string` - 文字列入力（デフォルト）
- `number` - 数値入力（min, max, step属性をサポート）
- `email` - メールアドレス入力（形式検証あり）
- `url` - URL入力（形式検証あり）
- `datetime` - 日時入力
- `date` - 日付入力
- `time` - 時間入力
- `tel` - 電話番号入力（形式検証あり）

## 注意事項

- このコードを使用するには、GASプロジェクトでCacheServiceを有効にしてください
- `setInterval`は非同期処理のシミュレーションとして使用しています
- 実際の環境では適切なエラーハンドリングを追加してください
