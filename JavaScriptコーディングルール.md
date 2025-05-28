# JavaScript コーディングルール

## 1. 命名規則

### 変数・関数名
- **camelCase** を使用する
- 意味のある名前を付ける
- 略語は避ける

```javascript
// 良い例
const userName = 'John';
const calculateTotalPrice = () => {};

// 悪い例
const un = 'John';
const calc = () => {};
```

### 定数
- **UPPER_SNAKE_CASE** を使用する

```javascript
const MAX_RETRY_COUNT = 3;
const API_BASE_URL = 'https://api.example.com';
```

### クラス名
- **PascalCase** を使用する

```javascript
class UserManager {
  constructor() {}
}
```

## 2. インデントとフォーマット

### インデント
- **2スペース** または **4スペース** で統一
- タブは使用しない

### 行の長さ
- 1行は **80-120文字** 以内に収める

### セミコロン
- 文の終わりには必ずセミコロンを付ける

```javascript
const message = 'Hello World';
console.log(message);
```

## 3. 関数の設計

### 単一責任の原則
- 1つの関数は1つの責任のみを持つ

```javascript
// 良い例
const validateEmail = (email) => {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
};

const sendEmail = (email, message) => {
  // メール送信処理
};

// 悪い例
const validateAndSendEmail = (email, message) => {
  // バリデーションと送信を同時に行う
};
```

### 関数の長さ
- 1つの関数は **20-30行** 以内に収める
- 長くなる場合は分割を検討

### 引数の数
- 引数は **3個以下** に抑える
- 多い場合はオブジェクトを使用

```javascript
// 良い例
const createUser = ({ name, email, age, address }) => {
  // 処理
};

// 悪い例
const createUser = (name, email, age, street, city, zipCode) => {
  // 処理
};
```

## 4. コメント

### JSDoc形式
- 関数にはJSDoc形式でコメントを記述

```javascript
/**
 * ユーザー情報を取得する
 * @param {string} userId - ユーザーID
 * @returns {Promise<Object>} ユーザー情報
 */
const getUserInfo = async (userId) => {
  // 処理
};
```

### インラインコメント
- 複雑な処理には説明を追加
- なぜそうするのかを説明する

```javascript
// APIレスポンスのキャッシュを5分間保持
const CACHE_DURATION = 5 * 60 * 1000;
```

## 5. 変数宣言

### const/let の使い分け
- 再代入しない場合は **const** を使用
- 再代入する場合のみ **let** を使用
- **var** は使用しない

```javascript
// 良い例
const users = [];
let currentIndex = 0;

// 悪い例
var users = [];
var currentIndex = 0;
```

### 宣言のタイミング
- 変数は使用する直前で宣言する

## 6. エラーハンドリング

### try-catch の使用
- 非同期処理では適切なエラーハンドリングを行う

```javascript
const fetchUserData = async (userId) => {
  try {
    const response = await fetch(`/api/users/${userId}`);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    return await response.json();
  } catch (error) {
    console.error('ユーザーデータの取得に失敗:', error);
    throw error;
  }
};
```

## 7. オブジェクトと配列

### 分割代入の活用
```javascript
// 良い例
const { name, email } = user;
const [first, second] = items;

// 悪い例
const name = user.name;
const email = user.email;
```

### スプレッド演算子の活用
```javascript
// 配列のコピー
const newItems = [...items];

// オブジェクトのマージ
const updatedUser = { ...user, lastLogin: new Date() };
```

## 8. 条件分岐

### 早期リターン
- ネストを深くしない

```javascript
// 良い例
const processUser = (user) => {
  if (!user) {
    return null;
  }
  
  if (!user.isActive) {
    return null;
  }
  
  // メイン処理
  return processActiveUser(user);
};

// 悪い例
const processUser = (user) => {
  if (user) {
    if (user.isActive) {
      // メイン処理
      return processActiveUser(user);
    }
  }
  return null;
};
```

### 三項演算子の適切な使用
- 簡単な条件分岐のみに使用

```javascript
// 良い例
const status = isActive ? 'アクティブ' : '非アクティブ';

// 悪い例（複雑すぎる）
const message = user ? (user.isActive ? (user.isPremium ? 'プレミアムユーザー' : '一般ユーザー') : '非アクティブ') : 'ユーザーなし';
```

## 9. モジュール化

### import/export の使用
```javascript
// 名前付きエクスポート
export const validateEmail = (email) => {
  // 処理
};

// デフォルトエクスポート
export default class UserService {
  // クラス定義
}
```

### ファイル分割
- 機能ごとにファイルを分割
- 1ファイルは200-300行以内に収める

## 10. パフォーマンス

### 不要な処理の回避
```javascript
// 良い例
const users = data.users || [];

// 悪い例
const users = data.users ? data.users : [];
```

### ループの最適化
```javascript
// 良い例
const length = items.length;
for (let i = 0; i < length; i++) {
  // 処理
}

// 配列メソッドの活用
const activeUsers = users.filter(user => user.isActive);
```

---

## プロジェクト解析結果

### ✅ 実践済みのコーディングルール

#### 1. 命名規則
- **関数名**: camelCaseが一貫して使用されている
  - `buildFmMenu`, `setTimescale`, `updateGanttData`, `getMemberDataAndHeaders`
- **定数**: UPPER_SNAKE_CASEが適切に使用されている
  - `COL_HEADER_NAMES`, `REQUIRED_MEMBER_DATA_HEADERS`, `DIALOG_PROP_PREFIX`

#### 2. JSDocコメント
- ほとんどの関数でJSDoc形式のコメントが記述されている
- パラメータと戻り値の型が明記されている

#### 3. 変数宣言
- `const`と`let`が適切に使い分けられている
- `var`の使用は確認されていない

#### 4. エラーハンドリング
- try-catch文が適切に使用されている
- エラーメッセージが日本語で分かりやすく記述されている

### ❌ 改善が必要なコーディングルール

#### 1. 関数の長さ
- 一部の関数が推奨される20-30行を大幅に超えている
  - `createGanttChartsWithMemberData` (約100行)
  - `integrateShiftData` (約60行)

#### 2. 引数の数
- 一部の関数で引数が3個を超えている
  - `integrateShiftData` (10個の引数)
  - `createDeptGanttSheet` (7個の引数)

#### 3. ネストの深さ
- 一部の関数で条件分岐のネストが深くなっている

---

## このプロジェクト独自のコーディングルール

### 1. ファイル命名規則
- **数字プレフィックス + 機能分類 + 具体的な機能名**の形式
```
1000_template_ganttBuilder.js      // テンプレート関連
2000_service_memberDataUpdater.js  // サービス層
3000_controller_menu.js            // コントローラー層
9800_lib_copyMemberData.js         // ライブラリ
```

### 2. 機能分類による階層化
- **1000番台**: テンプレート関連
- **2000番台**: サービス層（ビジネスロジック）
- **3000番台**: コントローラー層（UI制御）
- **9000番台**: ライブラリ・ユーティリティ

### 3. 日本語コメント
- コメントは日本語で記述する
- エラーメッセージも日本語で統一

```javascript
// 良い例（このプロジェクトのスタイル）
// ガントチャートのヘッダーとシフトデータを分割
const { ganttHeaderValues, ganttShiftValues } = splitGanttData();

// メンバーIDをキーとするマップに変換
const memberDataMap = createMemberDataMap(memberData);
```

### 4. 定数オブジェクトの活用
- 関連する定数をオブジェクトでグループ化

```javascript
const COL_HEADER_NAMES = {
  DEPT: "dept",
  EMAIL: "email",
  MEMBER_ID: "memberId",
  MEMBER_DATE_ID: "memberDateId",
  DATE: "date"
};

const DEPARTMENTS = {
  A: "会場整備局",
  B: "参加対応局",
  C: "開発局",
  // ...
};
```

### 5. 関数名の動詞パターン
- **build**: UI要素やデータ構造の構築
- **get**: データの取得
- **create**: 新しいオブジェクトの生成
- **update**: 既存データの更新
- **process**: データの処理・変換
- **handle**: イベントやコールバックの処理

### 6. グローバル関数の明示
- ダイアログコールバック用のグローバル関数には明示的にコメントを記述

```javascript
// 時間軸設定ダイアログの送信処理 (グローバル関数)
function processTimescaleInput(formData, context) {
  // 処理
}
```

### 7. 分割代入の積極的活用
- オブジェクトの分割代入を多用してコードを簡潔に

```javascript
const {
  headers: ganttHeaders,
  headerRow,
  startCol,
  endCol,
} = getGanttHeaders(sheet, headerRangeA1, requiredHeaders);
```

### 8. Map/Setの活用
- データの重複チェックや高速検索にMap/Setを積極的に使用

```javascript
const validShiftsMap = new Map();
const conflictShiftIds = new Set();
```

## まとめ

これらのルールを守ることで：
- コードの可読性が向上
- メンテナンスが容易になる
- バグの発生率が減少
- チーム開発がスムーズになる

定期的にコードレビューを行い、これらのルールが守られているかチェックすることが重要です。 