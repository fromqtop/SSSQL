# SSSQL - Google Spreadsheet SQL-like Query Library

SSSQL is a Google Apps Script library that allows you to flexibly manipulate data in Google Sheets using SQL-like queries.

---

## Methods

| Method | Return type | Brief description |
|--------|-------------|-------------------|
| select(sheet, query, options?) | Records | Get data from the sheet that matches specified conditions. |
| insert(sheet, record) | void | Insert a single row of data. |
| bulkInsert(sheet, records) | void | Insert multiple rows of data. |
| update(sheet, query) | void | Update data that meets specified conditions. |
| remove(sheet, query) | void | Delete data that meets specified conditions. |

---

## Detailed documentation
### select(sheet, query, options?)
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

// result
// [
//   { name: "Alice", age: 30, country: "USA" },
//   { name: "Bob", age: 25, country: "USA" }
// ]
```

#### Parameters

| Name | Description |
|------|-------------|
| sheet | The target sheet for selection. |
| query | Conditions for row selection, sorting, grouping, and other processing options. |
| options? | Options such as return format and row number retrieval. |

#### Details about the "query"

- **columns**

Specify the columns to be retrieved as an array.
If this parameter is omitted, all columns will be retrieved.

```javascript
const result = SSSQL.select(sheet, {
  columns: ["name", "age", "country"]
})
```

- **where**

Specify the row extraction conditions. When multiple conditions are provided, rows that meet all of the conditions will be extracted.
The available comparison operators are described later. If both where and whereOr are omitted, all rows will be extracted.

```javascript
const result = .select(sheet, {
  where: {
    age: [">", "20"],
    country: ["=", "USA%"]
  }
})
```

- **whereOr**

Specify the row extraction conditions. When multiple conditions are provided, rows that meet any of the conditions will be extracted.
The available comparison operators are described later. If both where and whereOr are omitted, all rows will be extracted.

```javascript
const result = SSSQL.select(sheet, {
  where: {
    age: [">", "20"],
    country: ["=", "USA%"]
  }
});
```

- **groupBy**

Specify when grouping and aggregating data.

```javascript
const result = SSSQL.select(sheet,
  groupBy: [
    ["job", "country"],
    { avg_salary: ["salary", "AVG"], max_salary: ["salary", "MAX"] }
  ]
);
// result
// [
//   { job: "Sales", country: "USA", avg_salary: 3500, max_salary: 7000 },
//   { job: "HR", country: "USA", avg_salary: 3800, max_salary: 8000 }
// ]
```

- **orderBy**

Specify when sorting data. The sort order can be specified as "ASC" or "DESC".

```javascript
const result = SSSQL.select(sheet, {
  orderBy: { age: "ASC", name: "DESC" }
});
```

#### optionsオブジェクト
selectには下記のオプションを指定可能です。

| オプション | 概要 | 例 |
| --------- | ---- | -- |
| `withRowNum` | シートの行番号(`ROWNUM`)も取得します。 | `options: { withRowNum: true }` |
| `asArray` | データを二次元配列として取得します。 | `options: { asArray: true }` |

---

### insert
#### 使用例
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const record ={
  name: "Charlie",
  age: "28",
  country: "Canada"
}

SSSQL.insert(sheet, record);
```

#### recordオブジェクト
挿入するデータをオブジェクトで指定します。（使用例参照）
指定しなかったカラムには `null` が設定されます。

---

### bulkInsert
#### 使用例
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const records ={[
  { name: "Dave", age: "35", country: "UK" },
  { name: "Eve", age: "27", country: "Germany" }
]};

SSSQL.bulkInsert(sheet, records);
```

#### recordsオブジェクト
挿入するデータをオブジェクトの配列で指定します。（使用例参照）
指定しなかったカラムには `null` が設定されます。

---

### update
#### 使用例
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const query = {
  set: { phone: "090-1234-5678" },
  where: { id: ["=", "alice@example.com"] }
};

SSSQL.update(sheet, query);
```

#### queryオブジェクト
| プロパティ | 概要 | 例 |
|---------|------|------|
| `set` | 更新する列と値を指定します。 | `set: { phone: "090-1234-5678" }` |
| `where`または`whereOr` | 更新対象行の条件を指定します。詳細は`select`メソッドの`where`・`whereOr`プロパティの解説を参照してください。| `where: { age: [">", 20], country: ["=", "USA" }` |


---

### remove
#### 使用例
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const query = {
  where: { id: "alice@example.com" }
}

SSSQL.remove(sheet, query);
```

## 補足
### 比較演算子
`where`, `whereOr` プロパティにおいて、利用できる比較演算子は下記のとおりです。

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
