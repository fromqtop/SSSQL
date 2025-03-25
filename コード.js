function test() {
  const spreadSheetId = "19Of_OoTfFRVX2uETDtt0R7OE79V4m43sqpi6hJAp-qM"
  const sheetName = "シート1"
    const spreadSheet = SpreadsheetApp.openById(spreadSheetId);
    const sheet = spreadSheet.getSheetByName("シート1");

    select({
      spreadSheetId: spreadSheetId,
      sheetName: sheetName,
      columns: ["name", "age"],
      where: ["name", "LIKE", "tak_sh%"]
    })
}

function select(q) {
  console.log(q)
  const spreadSheet = SpreadsheetApp.openById(q.spreadSheetId);
  const sheet = spreadSheet.getSheetByName(q.sheetName);
  let [columns, ...records] = sheet.getDataRange().getValues();

  // where
  if (q.hasOwnProperty("where")) {
    records = where_(columns, records, ...q.where, )
  }
  console.log(columns)
  console.log(records)
}

function where_(columns, records, column, operator, criteria) {
  const columnIndex = columns.indexOf(column);
  if (columnIndex === -1) {
    throw new Error(`指定されたカラム "${column}" は存在しません。`);
  }

  return records.filter(record => {
    const value = record[columnIndex];
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
      case "IN":
        return criteria.includes(value);
      case "LIKE":
        const pattern = criteria.replace(/%/g, ".*").replace(/_/g, ".");
        const regex = new RegExp(`^${pattern}$`);
        return regex.test(value);
      default:
        throw new Error(`不正な演算子 "${operator}" が指定されました。`);
    }
  });
}

function arraysToObjects(arrays) {
  const [keys, ...records] = arrays;
  return records.map(record => 
    record.reduce((acc, value, i) => {
      acc[keys[i]] = value;
      return acc;
    }, {})
  );
}