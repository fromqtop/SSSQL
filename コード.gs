/**
 * Get data from the sheet that matches specified conditions.
 * 
 * ### USAGE
 * ```javascript
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("customers");
 * 
 * const query = {
 *   columns: ["name", "age", "country"],
 *   where: {
 *     age: [">", 20],
 *     country: ["=", "USA"]
 *   }
 * };
 * 
 * const result = SSSQL.select(sheet, query);
 * 
 * // result
 * // [
 * //   { name: "Alice", age: 30, country: "USA" },
 * //   { name: "Bob", age: 25, country: "USA" }
 * // ]
 * ``` 
 */
function select(sheet, query, options) {
  // 表全体を取得
  let [columns, ...records] = sheet.getDataRange().getValues();

  // withRowNumオプション指定時　先頭に行番号を付与
  if (options?.withRowNum) {
    columns.unshift("ROWNUM");
    records.forEach((record, index) => record.unshift(index + 2));
  }

  // where・whereOr指定時　条件を満たすレコードに絞込み
  if (query?.hasOwnProperty("where")) {
    records = filterRecords_(query.where, "every", columns, records);
  } else if (query?.hasOwnProperty("whereOr")) {
    records = filterRecords_(query.whereOr, "some", columns, records);
  }

  // groupBy指定時　グループ化・集計
  if (query?.hasOwnProperty("groupBy")) {
    [columns, records] = groupRecords_(query.groupBy, columns, records);
  }

  // orderBy指定時　レコードを並び替え
  if (query?.hasOwnProperty("orderBy")) {
    records = orderRecords_(query.orderBy, columns, records)
  }

  // columns指定時　抽出するカラムを絞込み
  if (query?.hasOwnProperty("columns")) {
    records = selectColumns_(query.columns, columns, records);
    columns = query.columns;
  }

  // asArrayオプション指定時　抽出結果を二次元配列で返す。未指定時はオブジェクトの配列で返す。
  if (options?.asArray) {
    return [columns, ...records];
  } else {
    return arraysToObjects(columns, records);
  }
}

/**
 * Insert a single row of data.
 * 
 * ### USAGE
 * ```javascript
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("customers");

 * const record = { 
 *   name: "Alice",
 *   age: 30,
 *   country: "USA"
 * };
 * 
 * const result = SSSQL.insert(sheet, record);
 * 
 * // result
 * // { name: "Alice", age: 30, country: "USA", job; null }
 * //
 * ``` 
 */
function insert(sheet, record) {
  const columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const array = getRecordAryByRecordObj_(record, columns);
  sheet.appendRow(array);
  return arrayToObject(columns, array);
}

/**
 * Insert multiple rows of data.
 * 
 * ### USAGE
 * ```javascript
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("customers");
 * 
 * const records = [
 *   { name: "Alice", age: 30, country: "USA" },
 *   { name: "Bob", age: 25, country: "USA" }
 * ];
 * 
 * const result = SSSQL.bulkInsert(sheet, records);
 * 
 * // result
 * // [
 * //   { name: "Alice", age: 30, country: "USA", job; null },
 * //   { name: "Bob", age: 25, country: "USA", job; null }
 * // ]
 * ``` 
 */
function bulkInsert(sheet, records) {
  const columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const arrays = records.map(record => 
    getRecordAryByRecordObj_(record, columns)
  );
  sheet
    .getRange(sheet.getLastRow() + 1, 1, arrays.length, columns.length)
    .setValues(arrays);
  return arraysToObjects(columns, arrays);
}

/**
 * Update data that meets specified conditions.
 * 
 * ### USAGE
 * ```javascript
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("customers");
 * 
 * const query = {
 *   set: { phone: "090-1234-5678" },
 *   where: { id: ["=", "alice@example.com"] }
 * };
 * 
 * const result = SSSQL.update(sheet, query);
 * 
 * // result
 * // [
 * //   {
 * //     before: { id: "alice@example.com", name: "Alice", age: 30, country: "USA", phone: null },
 * //     after: { id: "alice@example.com", name: "Bob", age: 25, country: "USA", phone: "090-1234-5678" }
 * //   }
 * // ]
 * ``` 
 */
function update(sheet, query) {
  // 更新対象レコードを取得
  const records = getTargetRecord_(sheet, query, { withRowNum: true });

  // レコード更新
  const result = [];
  records.forEach(record => {
    const rownum = record.ROWNUM;
    delete record.ROWNUM;

    // setで指定したとおりレコードを更新
    const newRecord = Object.keys(query.set).reduce((acc, key) => {
      acc[key] = query.set[key];
      return acc;
    }, {...record})

    // 配列に変換
    const columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const array = getRecordAryByRecordObj_(newRecord, columns);
    
    // 更新実行
    sheet.getRange(rownum, 1, 1, columns.length).setValues([array]);
    
    // 戻り値をセット
    result.push({ before: record, after: newRecord });
  })

  return result;
}

/**
 * Delete data that meets specified conditions.
 * 
 * ### USAGE
 * ```javascript
 * const ss = SpreadsheetApp.getActiveSpreadsheet();
 * const sheet = ss.getSheetByName("customers");
 * 
 * const query = {
 *   where: { id: "alice@example.com" }
 * };
 * 
 * const result = SSSQL.remove(sheet, query);
 * 
 * // result
 * // [
 * //   { id: "alice@example.com", name: "Alice", age: 30, country: "USA" }
 * // ]
 * ``` 
 */
function remove(sheet, query) {
  // 更新対象レコードを取得
  const records = getTargetRecord_(sheet, query, { withRowNum: true });
  
  // レコードをROWNUMで降順ソート
  records.sort((a, b) => b.ROWNUM - a.ROWNUM);
  
  // 削除実行
  records.forEach(record => sheet.deleteRow(record.ROWNUM));

  return records.map(({ ROWNUM, ...rest }) => rest);
}

function filterRecords_(where, mode, columns, records) {
  const indexMap = getIndexMap_(columns);

  // whereオブジェクトを配列に展開
  const conditions = Object.entries(where).map(entry => {
    let column = entry[0];
    const [operator, criteria] = entry[1];

    // columnに$$を含む場合は$$の前までが正式カラム名（同一カラムに複数の条件を指定した場合に$$が付いている）
    const $$index = column.indexOf("$$");
    if ($$index !== -1) column = column.substring(0, $$index);

    return { column, operator, criteria };
  })

  // 条件を満たすレコードのみ抽出
  return records.filter(record => {
    // [mode]には every/some のいずれかが指定されている
    return conditions[mode](condition => {
      const { column, operator, criteria } = condition;
      const value = record[indexMap[column]];
      const regex = (operator === "LIKE" || operator === "NOT LIKE") 
        ? likeToRegex_(criteria)
        : null;
      return isValidValue_(value, operator, criteria, regex);
    });
  });
}

function isValidValue_(value, operator, criteria, regex) {
  // Date型の一致・不一致はオブジェクト同士の比較はNG
  if (criteria instanceof Date) {
    if (!(value instanceof Date)) return false;
    criteria = criteria.getTime();
    value = value.getTime();
  }

  switch (operator) {
    case "=":
      return value === criteria;
    case "<>":
      return value !== criteria;
    case ">":
      return value > criteria;
    case ">=":
      return value >= criteria;
    case "<":
      return value < criteria;
    case "<=":
      return value <= criteria;
    case "BETWEEN":
      return criteria[0] <= value && value <= criteria[1]; 
    case "NOT BETWEEN":
      return !(criteria[0] <= value && value <= criteria[1]); 
    case "IN":
      return criteria.includes(value);
    case "NOT IN":
      return !criteria.includes(value);
    case "LIKE":
      return regex.test(value);
    case "NOT LIKE":
      return !regex.test(value);
    default:
      throw new Error(`不正な演算子 "${operator}" が指定されました。`);
  }
}

function likeToRegex_(like) {
  pattern = like.replace(/%/g, ".*").replace(/_/g, ".");
  return new RegExp(`^${pattern}$`);
}

function orderRecords_(orderBy, columns, records) {
  const keys = Object.keys(orderBy);
  const sortCriteria = keys.map(key => {
    const index = columns.indexOf(key);
    if (index === -1) {
      throw new Error(`指定されたカラム "${key}" は存在しません。`);
    }
    return {
      order: orderBy[key],
      index
    };
  });

  return records.sort((a, b) => {
    for (const { order, index } of sortCriteria) {
      if (a[index] < b[index]) return order === "ASC" ? -1 : 1;
      if (a[index] > b[index]) return order === "ASC" ?  1 : -1;
    }
    return 0;
  });
}

function groupRecords_(groupBy, columns, records) {
  const [by, aggs] = groupBy;
  const indexMap = getIndexMap_(columns);

  const groups = [];
  for (const record of records) {
    const keys = by.map(column => record[indexMap[column]]);
    const keyString = JSON.stringify(keys);
    const groupIndex = groups.findIndex(group => group.keyString === keyString);
    groupIndex === -1
      ? groups.push({ keyString, keys, records: [record] })
      : groups[groupIndex].records.push(record);
  }

  function isNumber(value) {
    return typeof value === "number" && !isNaN(value);
  }

  function getNumericValues(records, index) {
    return records
      .map(record => record[index])
      .filter(value => isNumber(value));
  }

  for (const group of groups) {
    group.calculated = {};
    for (const as of Object.keys(aggs)) {
      const [column, aggType] = aggs[as];
      const index = indexMap[column];

      group.calculated[as] = (() => {
        switch (aggType) {
          case "COUNT":
            return group.records.filter(record => 
              record[index] !== ""
            ).length;
          case "SUM":
            return group.records.reduce((acc, record) => {
              const value = record[index];
              return isNumber(value) ? acc + value : acc;
            }, 0);
          case "AVG":
            const { count, sum } = group.records.reduce((acc, record) => {
              const value = record[index];
              if (isNumber(value)) {
                acc.count += 1;
                acc.sum += value;
              }
              return acc;
            }, { count: 0, sum: 0 });        
            return count !== 0 ? sum / count : null;  
          case "MIN":
            return Math.min(...getNumericValues(group.records, index));
          case "MAX":
            return Math.max(...getNumericValues(group.records, index));
        }
      })();
    }
  }

  const newColumns = [...by, ...Object.keys(aggs)];
  const newRecords = groups.map(group => [
    ...group.keys, 
    ...Object.values(group.calculated)
  ]);
  return [newColumns, newRecords];
}

function selectColumns_(targetColumns, columns, records) {
  const indexMap = getIndexMap_(columns);

  newRecords = records.map(record => 
    targetColumns.map(targetColumn => {
      if (!indexMap.hasOwnProperty(targetColumn)) throw new Error(`指定されたカラム "${column}" は存在しません。`);
      return record[indexMap[targetColumn]]
    })
  );
  return newRecords;
}

function getTargetRecord_(sheet, query) {
  const selectQuery = {};
  if (query.hasOwnProperty("where")) {
    selectQuery.where = query.where;
  } else if (query.hasOwnProperty("whereOr")) {
    selectQuery.whereOr = query.whereOr;
  }
  const records = select(sheet, selectQuery, { withRowNum: true });
  return records;
}

function getRecordAryByRecordObj_(obj, columns) {
  // カラム存在チェック
  for (const key in obj) {
    if (!columns.includes(key)) {
      throw new Error(`キー "${key}" はシートに存在しません。`);
    }
  }  

  // シート項目と同順の配列に変換
  return columns.map(column => {
    return (obj.hasOwnProperty(column))
      ? obj[column]
      : null;
  });
}

function getIndexMap_(columns) {
  return columns.reduce((acc, column, index) => {
    acc[column] = index;
    return acc;
  }, {});
}

function arraysToObjects(keys, records) {
  return records.map(record => 
    arrayToObject(keys, record)
  );
}

function arrayToObject(keys, record) {
  return record.reduce((acc, value, i) => {
    acc[keys[i]] = value;
    return acc;
  }, {})
}