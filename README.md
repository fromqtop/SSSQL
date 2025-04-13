# SSSQL - Google Spreadsheet SQL-like Query Library

SSSQL は Google スプレッドシート上のデータを、SQL ライクなクエリで柔軟に操作するための Google Apps Script ライブラリです。

---

## 📦 機能一覧

| メソッド | 概要 |
|---------|------|
| `select(sheet, query, options?)` | データの抽出・絞込み・整形 |
| `insert(sheet, record)` | 単一レコードの追加 |
| `bulkInsert(sheet, records)` | 複数レコードの一括追加 |
| `update(sheet, query)` | 条件に一致するレコードの更新 |
| `remove(sheet, query)` | 条件に一致するレコードの削除 |

---

## 使用例

### select
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const query = {
  columns: ["name", "age", "country"],
  where: {
    age: [">", "20"],
    country: ["=", "USA%"]
  }
};

const result = SSSQL.select(sheet, query);
console.log(result);

// [
//   { name: "Alice", age: 30, country: "USA" },
//   { name: "Bob", age: 25, country: "USA" }
// ]
```

#### queryオブジェクト
selectの条件等を指定する query オブジェクトには、下記のプロパティを指定可能です。

| プロパティ | 概要 | 例 |
|---------|------|------|
| `columns` | 取得するカラムを指定します。<br>省略時は全項目を取得します。 | `columns: ["name", "age", "country"]` |
| `where` | 行の抽出条件を指定します。<br>`{ 列名: ["比較演算子", 値], 列名: ["比較演算子", 値] ...}`の形式で条件を指定します。複数条件を指定した場合、すべての条件を満たす行が抽出されます。<br>使用できる比較演算子は後述。<br>where`および`whereOr`の両方を省略時は、全行が抽出されます。 | `where: { age: [">", 20], country: ["=", "USA" }` |
| `whereOr` | 行の抽出条件を指定します。<br>`{ 列名: ["比較演算子", 値], 列名: ["比較演算子", 値] ...}`の形式で条件を指定します。複数条件を指定した場合、いずれかの条件を満たす行が抽出されます。<br>使用できる比較演算子は後述。<br>where`および`whereOr`の両方を省略時は、全行が抽出されます。 | `whereOr: { age: [">", 20], country: ["=", "USA" }` |
| `groupBy` | データをグループ化・集計する際に指定します。<br>`[["グループ化項目1", "グループ化項目2" ...], { 出力列名1: ["集計列名", "集計関数"], 出力列名2: ["集計列名", "集計関数"] ...]の形式で指定します。` | `groupBy: [["job", "country"], { avg_salary: ["salary", "AVG"], max_salary: ["salary", "MAX"] }]` |
| `orderBy` | データをソートする際に指定します。<br>`{ 項目名1: "ソート順", 項目名2: "ソート順" ... }`の形式で指定します。 | `orderBy: { age: "ASC", name: "DESC" } |

#### optionsオブジェクト
selectには下記のオプションを指定可能です。

| オプション | 概要 | 例 |

##### 比較演算子
| 比較演算子 | 使用例 | 備考 |
| --------- | ------ | ---- |
| `=` | `age: ["=", 20]` |  |
| `<>` | `age: ["<>", 20]` |  |
| `>` | `age: [">", 20]` |  |
| `>=` | `age: [">=", 20]` |  |
| `<` | `age: ["<", 20]` |  |
| `<=` | `age: ["<=", 20]` |  |
| `BETWEEN` | `age: ["BETWEEN", [10, 30]]` |  |
| `NOT BETWEEN` | `age: ["NOT BETWEEN", [10, 30]]` |  |
| `IN` | `country: ["IN", ["JPN", "USA", "UK"]]` |  |
| `NOT IN` | `country: ["NOT IN", ["JPN", "USA", "UK"]]` |  |
| `LIKE` | `job: ["LIKE", "Sales%"]` | ワイルドカードとして下記を使用可能<br> `%` ・・・ 0文字以上の任意の文字列<br>`_` ・・・　任意の1文字 |
| `NOT LIKE` | `job: ["NOT LIKE", "Sales%"]` | ワイルドカードとして下記を使用可能<br> `%` ・・・ 0文字以上の任意の文字列<br>`_` ・・・　任意の1文字 |



### insert
```javascript
SSSQL.insert(sheet, {
  name: "Charlie",
  age: "28",
  country: "Canada"
});
```

### bulkInsert
```javascript
SSSQL.bulkInsert(sheet, [
  { name: "Dave", age: "35", country: "UK" },
  { name: "Eve", age: "27", country: "Germany" }
]);
```

### update
```javascript
SSSQL.update(sheet, {
  set: { phone: "090-1234-5678" },
  where: { id: "alice@example.com" }
});
```

### remove
```javascript
SSSQL.remove(sheet, {
  where: { id: "alice@example.com" }
});
```

