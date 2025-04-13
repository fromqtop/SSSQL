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

