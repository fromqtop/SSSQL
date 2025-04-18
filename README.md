# SSSQL - Google Spreadsheet SQL-like Query Library

SSSQL is a Google Apps Script library that allows you to flexibly manipulate data in Google Sheets using SQL-like queries.

## How To Use
The Script ID is `1AvjKJfZQdNYRGue-UCOyCMhRtDrcgLt55VPbcfUcb9lz15ZZPJlAwrul`.

For instructions on how to use the library, please refer to the following link:<br>
[Add a Library to Your Script Project](https://developers.google.com/apps-script/guides/libraries#add_a_library_to_your_script_project).

---

## Methods

| Method | Brief description |
|--------|-------------------|
| select(sheet, query, options?) | Get data from the sheet that matches specified conditions. |
| insert(sheet, record) | Insert a single row of data. |
| bulkInsert(sheet, records) | Insert multiple rows of data. |
| update(sheet, query) | Update data that meets specified conditions. |
| remove(sheet, query) | Delete data that meets specified conditions. |

---

## Detailed documentation

### select(sheet, query, options?)

Get data from the sheet that matches specified conditions.

```javascript
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheet = ss.getSheetByName("customers");
 
 const query = {
   columns: ["name", "age", "country"],
   where: {
     age: [">", 20],
     country: ["=", "USA"]
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

Conditions for row selection, sorting, grouping, and other processing options.

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
const result = SSSQL.select(sheet, {
  where: {
    age: [">", 20],
    country: ["=", "USA"]
  }
})
```

- **whereOr**

Specify the row extraction conditions. When multiple conditions are provided, rows that meet any of the conditions will be extracted.
The available comparison operators are described later. If both where and whereOr are omitted, all rows will be extracted.

```javascript
const result = SSSQL.select(sheet, {
  whereOr: {
    age: [">", 20],
    country: ["=", "USA"]
  }
});
```

- **groupBy**

Specify when grouping and aggregating data.

```javascript
const result = SSSQL.select(sheet, {
  groupBy: [
    ["job", "country"],
    { avg_salary: ["salary", "AVG"], max_salary: ["salary", "MAX"] }
  ]
});
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

#### Details about the "options"

Options such as return format and row number retrieval.

| Name | Description |
|------|-------------|
| withRowNum | Includes the sheet's row number (`ROWNUM`) in the result. |
| asArray | Retrieves the data as a two-dimensional array. |

---

### insert(sheet, record)

Insert a single row of data.

```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const record = { 
  name: "Alice",
  age: 30,
  country: "USA"
};

const result = SSSQL.insert(sheet, record);
 
// result
// { name: "Alice", age: 30, country: "USA", job; null }
//
```

#### Parameters

| Name | Description |
|------|-------------|
| sheet | The target sheet for insert. |
| record | Specify the data to insert as an object (see usage example). Columns that are not specified will be set to null. |

---

### bulkInsert(sheet, records)

Insert multiple rows of data.

```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const result = SSSQL.bulkInsert(sheet, records);
 
// result
// [
//   { name: "Alice", age: 30, country: "USA", job; null },
//   { name: "Bob", age: 25, country: "USA", job; null }
// ]
```

#### Parameters

| Name | Description |
|------|-------------|
| sheet | The target sheet for insert. |
| records | Specify the data to insert as an array of objects (see usage example). Columns that are not specified will be set to null. |

---

### update(sheet, query)

Update data that meets specified conditions.

```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const query = {
  set: { phone: "090-1234-5678" },
  where: { id: ["=", "alice@example.com"] }
};

const result = SSSQL.update(sheet, query);

// result
// [
//   {
//     before: { id: "alice@example.com", name: "Alice", age: 30, country: "USA", phone: null },
//     after: { id: "alice@example.com", name: "Bob", age: 25, country: "USA", phone: "090-1234-5678" }
//   }
// ]
```

#### Parameters

| Name | Description |
|------|-------------|
| sheet | The target sheet for insert. |
| query | Specify the conditions for the rows to be updated. Use the set property to specify the columns and values to be updated.<br> As for specifying the rows to be updated, see the explanation of the `where` and `whereOr` properties in the `select` method. |

---

### remove(sheet, query)

Delete data that meets specified conditions.

```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName("customers");

const query = {
  where: { id: "alice@example.com" }
}

const result = SSSQL.remove(sheet, query);

// result
// [
//   { id: "alice@example.com", name: "Alice", age: 30, country: "USA" }
// ]
```

#### Parameters

| Name | Description |
|------|-------------|
| sheet | The target sheet for insert. |
| query | Specify the conditions for the rows to be deleted. As for specifying the rows to be updated, see the explanation of the `where` and `whereOr` properties in the `select` method. |

## Notes  

### Comparison Operators  

The following comparison operators can be used with the `where` and `whereOr` properties:

| Operator       | Example                                     | Notes                                                                 |
|----------------|---------------------------------------------|-----------------------------------------------------------------------|
| `=`            | `age: ["=", 20]`                            |                                                                       |
| `<>`           | `age: ["<>", 20]`                           |                                                                       |
| `>`            | `age: [">", 20]`                            |                                                                       |
| `>=`           | `age: [">=", 20]`                           |                                                                       |
| `<`            | `age: ["<", 20]`                            |                                                                       |
| `<=`           | `age: ["<=", 20]`                           |                                                                       |
| `BETWEEN`      | `age: ["BETWEEN", [10, 30]]`                |                                                                       |
| `NOT BETWEEN`  | `age: ["NOT BETWEEN", [10, 30]]`            |                                                                       |
| `IN`           | `country: ["IN", ["JPN", "USA", "UK"]]`     |                                                                       |
| `NOT IN`       | `country: ["NOT IN", ["JPN", "USA", "UK"]]` |                                                                       |
| `LIKE`         | `job: ["LIKE", "Sales%"]`                   | Wildcards:<br> `%` – any sequence of characters (including none)<br> `_` – any single character |
| `NOT LIKE`     | `job: ["NOT LIKE", "Sales%"]`               | Wildcards:<br> `%` – any sequence of characters (including none)<br> `_` – any single character |

## License

This project is licensed under the MIT License
