class SSSQL {
  constructor(spreadsheetId, sheetName, headerRow, dataStartRow, sharedState) {
    this.spreadsheetId = spreadsheetId;
    this.sheetName = sheetName;
    this.headerRow = headerRow !== undefined ? headerRow : 1;
    this.dataStartRow = dataStartRow !== undefined ? dataStartRow : this.headerRow + 1;

    // クローン間で共有される状態（どのクローンから書き換えても、全員に伝わる）
    this._sharedState = sharedState || {
      sheet: null,
      header: null,
      headerSet: null,
      rows: null // 行データのキャッシュ
    };

    // クエリの状態（クローンごとに個別。where/orderBy/groupBy/select/offset/useCache/aggregateで新しいクローンに設定される）
    this._condition = null;
    this._orderBy = null;
    this._useCache = false;
    this._offset = null;
    this._groupBy = null;
    this._selectCols = null;
    this._aggregateSpec = null;
    this._havingCondition = null;
  }

  // ---- 遅延取得プロパティ（クローン間で共有） ----
  get sheet() {
    if (!this._sharedState.sheet) {
      this._sharedState.sheet = SpreadsheetApp.openById(this.spreadsheetId).getSheetByName(this.sheetName);
      if (!this._sharedState.sheet) {
        throw new Error(`Sheet not found: ${this.sheetName}`);
      }
    }
    return this._sharedState.sheet;
  }

  get header() {
    if (!this._sharedState.header) {
      const lastCol = this.sheet.getLastColumn();
      const header = lastCol === 0 ? [] : this.sheet.getRange(this.headerRow, 1, 1, lastCol).getValues()[0];
      this._sharedState.header = header;
      this._sharedState.headerSet = new Set(header);
    }
    return this._sharedState.header;
  }

  get headerSet() {
    if (!this._sharedState.headerSet) this.header; // headerのgetter経由で一緒に構築される
    return this._sharedState.headerSet;
  }

  // ================================================================
  // クエリ組み立て系（呼ぶたびにクローンを返す。元のインスタンスは変更しない）
  // ================================================================

  /**
   * 検索条件を指定する。呼ぶたびに新しいクローンを返し、元のインスタンスには影響しない。
   * 複数回 where() を呼んだ場合は「上書き」になる（ANDにしたい場合は { AND: [...] } を明示する）。
   */
  where(condition) {
    const clone = this._clone();
    clone._condition = this._shallowCopy(condition);
    return clone;
  }

  /**
   * ソート順を指定する。呼ぶたびに新しいクローンを返す。
   * 例: db.orderBy({ age: "desc", name: "asc" })
   */
  orderBy(order) {
    const clone = this._clone();
    clone._orderBy = Object.entries(order).map(([column, direction]) => {
      const dir = typeof direction === "string" ? direction.toLowerCase() : direction;
      if (dir !== "asc" && dir !== "desc") {
        throw new Error(`Invalid direction: ${direction} (use "asc" or "desc")`);
      }
      return { column, direction: dir };
    });
    return clone;
  }

  /**
   * 取得する列を指定する（列を絞るだけで、実行はしない）。
   * groupBy() と同時に使用するとエラーになる（groupByで指定した列は自動的に結果へ含まれるため）。
   */
  select(...cols) {
    const clone = this._clone();
    clone._selectCols = cols;
    return clone;
  }

  /**
   * グループ化するカラムを指定する。呼ぶたびに新しいクローンを返す。
   * select() と同時には使用できない。集計したい場合は aggregate() を組み合わせる。
   */
  groupBy(...cols) {
    if (cols.length === 0) {
      throw new Error("groupBy() には少なくとも1つのカラム名を指定してください");
    }
    const clone = this._clone();
    clone._groupBy = cols;
    return clone;
  }

  /**
   * groupBy() と組み合わせて、グループごとに複数の集計値を同時に取得する。
   * spec の形式: { 出力キー名: { sum|avg|max|min: "カラム名" } | { count: true } }
   *
   * @example
   * db.groupBy("department").aggregate({
   *   totalScore: { sum: "score" },
   *   maxScore: { max: "score" },
   *   memberCount: { count: true }
   * }).all();
   */
  aggregate(spec) {
    const clone = this._clone();
    clone._aggregateSpec = this._shallowCopy(spec);
    return clone;
  }

  /**
   * groupBy() + aggregate() で計算した集計値に対する絞り込み条件を指定する。
   * where() と同じ演算子記法（比較演算子・BETWEEN・IN・LIKE・AND/OR）が使える。
   * aggregate() で指定した「出力キー名」を条件のキーとして使う。
   *
   * @example
   * db.groupBy("department").aggregate({ total: { sum: "score" } })
   *   .having({ total: [">=", 1000] }).all();
   */
  having(condition) {
    const clone = this._clone();
    clone._havingCondition = this._shallowCopy(condition);
    return clone;
  }

  /**
   * 取得開始位置（先頭から何件スキップするか）を指定する。呼ぶたびに新しいクローンを返す。
   * all()/first()/take() の結果にのみ適用される。
   */
  offset(n) {
    this._assertNonNegativeInteger(n, "offset");
    const clone = this._clone();
    clone._offset = n;
    return clone;
  }

  /**
   * 次の実行でキャッシュを利用する。呼ぶたびに新しいクローンを返す。
   */
  useCache() {
    const clone = this._clone();
    clone._useCache = true;
    return clone;
  }

  /**
   * キャッシュを破棄する。全クローンで共有している箱(_sharedState)を書き換えるので、
   * どのインスタンスから呼んでも、そこから生まれた／派生した全クローンに伝わる。
   */
  refreshCache() {
    this._sharedState.rows = null;
    return this;
  }

  // ================================================================
  // 実行系（行データを返す。offset が効く）
  // ================================================================

  /**
   * 条件・ソート順・グループ化に従って、全件を取得する。
   */
  all() {
    let rows = this._resultRows();
    if (this._offset) {
      rows = rows.slice(this._offset);
    }
    return rows;
  }

  /**
   * 条件・ソート順・グループ化に従って、最初の1件だけを取得する。
   * @returns {Object|null}
   */
  first() {
    let rows = this._resultRows();
    if (this._offset) {
      rows = rows.slice(this._offset);
    }
    return rows.length > 0 ? rows[0] : null;
  }

  /**
   * 条件・ソート順・グループ化に従って、先頭からn件を取得する。
   * @param {number} n
   */
  take(n) {
    this._assertNonNegativeInteger(n, "take");
    let rows = this._resultRows();
    if (this._offset) {
      rows = rows.slice(this._offset);
    }
    return rows.slice(0, n);
  }

  // ================================================================
  // 集計系（単体で完結。offset とは併用不可）
  // ================================================================

  /**
   * 条件に一致する行数を返す。
   * groupBy()と組み合わせた場合は、まずグループごとの件数を求め、
   * having()で絞り込んだ後の「グループの数」を返す。
   */
  count() {
    this._assertNoOffset("count");
    this._assertHavingRequiresGroupBy();
    this._assertNoAggregateSpec("count");
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("count");
      const groupCounts = this._groupedValues("count", rows => rows.length);
      return groupCounts.length;
    }
    return this._filteredRows().length;
  }

  /**
   * 指定したカラムの合計値を返す。数値でない値は無視する。
   * groupBy()と組み合わせた場合は、各グループのsumを求めた後、
   * having()で絞り込み、その「グループごとのsumの配列」をさらにsumする。
   */
  sum(column) {
    this._assertNoOffset("sum");
    this._assertHavingRequiresGroupBy();
    this._assertNoAggregateSpec("sum");
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("sum");
      const groupSums = this._groupedValues(`${column}_sum`, rows => {
        const values = this._extractNumeric(rows, column);
        return values.reduce((a, b) => a + b, 0);
      });
      return groupSums.reduce((a, b) => a + b, 0);
    }
    const values = this._extractNumeric(this._filteredRows(), column);
    return values.reduce((a, b) => a + b, 0);
  }

  /**
   * 指定したカラムの平均値を返す。数値でない値は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合は、各グループのavgを求めた後、
   * having()で絞り込み、その「グループごとのavgの配列」をさらにavgする（avg-of-avg）。
   */
  avg(column) {
    this._assertNoOffset("avg");
    this._assertHavingRequiresGroupBy();
    this._assertNoAggregateSpec("avg");
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("avg");
      const groupAvgs = this._groupedValues(`${column}_avg`, rows => {
        const values = this._extractNumeric(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
      });
      const numericGroupAvgs = groupAvgs.filter(v => typeof v === "number" && !isNaN(v));
      return numericGroupAvgs.length === 0 ? null : numericGroupAvgs.reduce((a, b) => a + b, 0) / numericGroupAvgs.length;
    }
    const values = this._extractNumeric(this._filteredRows(), column);
    return values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
  }

  /**
   * 指定したカラムの最大値を返す。数値・Date以外は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合は、各グループのmaxを求めた後、
   * having()で絞り込み、その「グループごとのmaxの配列」からさらにmaxを取る。
   */
  max(column) {
    this._assertNoOffset("max");
    this._assertHavingRequiresGroupBy();
    this._assertNoAggregateSpec("max");
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("max");
      const groupMaxes = this._groupedValues(`${column}_max`, rows => {
        const values = this._extractComparable(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => (b > a ? b : a));
      });
      const comparableGroupMaxes = groupMaxes.filter(v => (typeof v === "number" && !isNaN(v)) || v instanceof Date);
      return comparableGroupMaxes.length === 0 ? null : comparableGroupMaxes.reduce((a, b) => (b > a ? b : a));
    }
    const values = this._extractComparable(this._filteredRows(), column);
    return values.length === 0 ? null : values.reduce((a, b) => (b > a ? b : a));
  }

  /**
   * 指定したカラムの最小値を返す。数値・Date以外は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合は、各グループのminを求めた後、
   * having()で絞り込み、その「グループごとのminの配列」からさらにminを取る。
   */
  min(column) {
    this._assertNoOffset("min");
    this._assertHavingRequiresGroupBy();
    this._assertNoAggregateSpec("min");
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("min");
      const groupMins = this._groupedValues(`${column}_min`, rows => {
        const values = this._extractComparable(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => (b < a ? b : a));
      });
      const comparableGroupMins = groupMins.filter(v => (typeof v === "number" && !isNaN(v)) || v instanceof Date);
      return comparableGroupMins.length === 0 ? null : comparableGroupMins.reduce((a, b) => (b < a ? b : a));
    }
    const values = this._extractComparable(this._filteredRows(), column);
    return values.length === 0 ? null : values.reduce((a, b) => (b < a ? b : a));
  }

  /**
   * 指定したカラム（複数可）の組み合わせでユニークな値を取得する。
   * カラムを1つだけ指定した場合は値の配列、複数指定した場合はオブジェクトの配列を返す。
   * groupBy()/select()/aggregate() と組み合わせて呼ぶことはできない（それらの指定が無視されてしまうため）。
   */
  distinct(...cols) {
    this._assertNoOffset("distinct");
    if (this._groupBy) {
      throw new Error("distinct() は groupBy() と併用できません（distinctは単独で使用してください）");
    }
    if (this._selectCols) {
      throw new Error("distinct() は select() と併用できません（引数で指定したカラムがそのまま使われます）");
    }
    if (cols.length === 0) {
      throw new Error("distinct() には少なくとも1つのカラム名を指定してください");
    }
    cols.forEach(col => {
      if (!this.headerSet.has(col)) {
        throw new Error(`Unknown column: ${col}`);
      }
    });

    const filtered = this._filteredRows();
    const seen = new Set();
    const result = [];

    filtered.forEach(r => {
      const key = cols.map(c => String(r.data[c])).join("\u0000"); // 値の連結による偶発的な衝突を避けるための区切り文字
      if (!seen.has(key)) {
        seen.add(key);
        result.push(cols.length === 1 ? r.data[cols[0]] : this._pick(r.data, cols));
      }
    });

    return result;
  }

  /**
   * 条件に一致する行が1件以上あるかどうかを返す。
   */
  exists() {
    this._assertNoOffset("exists");
    return this._filteredRows().length > 0;
  }

  // ================================================================
  // 書き込み系
  // ================================================================

  /**
   * 1行または複数行を新規追加する。
   * 戻り値は「実際に書き込まれた内容」（未指定カラムの空欄補完後）。単一行ならオブジェクト、複数行なら配列。
   */
  insert(rowOrRows) {
    const isMultiple = Array.isArray(rowOrRows);
    const rows = isMultiple ? rowOrRows : [rowOrRows];

    if (rows.length === 0) {
      return isMultiple ? [] : undefined; // 空配列を渡された場合は何もしない
    }

    const rowsValues = rows.map(r => this._toRowValues(r));
    this._appendRows(rowsValues);
    this._sharedState.rows = null; // 書き込みしたのでキャッシュは破棄

    const afterRows = rowsValues.map(values => this._rowToDict(values));
    return isMultiple ? afterRows : afterRows[0];
  }

  /**
   * where()で指定した条件に一致する行を更新する。
   * where()を呼んでいない場合はエラーになる（誤って全件更新することを防ぐため）。
   * 意図的に全件を更新したい場合は updateAll() を使う。
   */
  update(values) {
    if (this._condition === null) {
      throw new Error("update() を実行するには where() で条件を指定してください。全件を更新したい場合は updateAll() を使用してください");
    }
    return this._performUpdate(values);
  }

  /**
   * where()の指定に関わらず、条件に一致する（あるいは全件の）行を更新する。
   * where()を呼んでいなければ全件が対象になる。誤操作に注意。
   */
  updateAll(values) {
    return this._performUpdate(values);
  }

  _performUpdate(values) {
    this._assertNoOffset("update");
    this._assertKnownColumns(values);

    const filtered = this._filteredRows(); // Before情報も兼ねる

    const changes = filtered.map(r => {
      const before = r.data;
      const after = { ...before, ...values };
      return { rowIndex: r.rowIndex, before, after };
    });

    const afterByRowIndex = new Map(
      changes.map(c => [c.rowIndex, this.header.map(key => c.after[key])])
    );

    const groups = this._groupConsecutiveRows(changes.map(c => c.rowIndex));
    groups.forEach(({ start, count }) => {
      const rowsValues = [];
      for (let i = 0; i < count; i++) {
        rowsValues.push(afterByRowIndex.get(start + i));
      }
      this.sheet.getRange(start, 1, count, this.header.length).setValues(rowsValues);
    });

    this._sharedState.rows = null; // 書き込みしたのでキャッシュは破棄

    return {
      count: changes.length,
      before: changes.map(c => c.before),
      after: changes.map(c => c.after)
    };
  }

  /**
   * where()で指定した条件に一致する行を削除する。
   * where()を呼んでいない場合はエラーになる（誤って全件削除することを防ぐため）。
   * 意図的に全件を削除したい場合は deleteAll() を使う。
   */
  delete() {
    if (this._condition === null) {
      throw new Error("delete() を実行するには where() で条件を指定してください。全件を削除したい場合は deleteAll() を使用してください");
    }
    return this._performDelete();
  }

  /**
   * where()の指定に関わらず、条件に一致する（あるいは全件の）行を削除する。
   * where()を呼んでいなければ全件が対象になる。誤操作に注意。
   */
  deleteAll() {
    return this._performDelete();
  }

  _performDelete() {
    this._assertNoOffset("delete");
    const filtered = this._filteredRows();

    const groups = this._groupConsecutiveRows(filtered.map(r => r.rowIndex));
    const sortedGroups = [...groups].sort((a, b) => b.start - a.start);
    sortedGroups.forEach(({ start, count }) => {
      this.sheet.deleteRows(start, count);
    });

    this._sharedState.rows = null; // 書き込みしたのでキャッシュは破棄

    return {
      count: filtered.length,
      deleted: filtered.map(r => r.data)
    };
  }

  /**
   * 条件に一致する行があれば更新し、無ければ「条件とvaluesをマージしたデータ」で新規追加する。
   * 存在確認・更新は where().first() / where().update() と同じ経路を使う。
   */
  upsert(condition, values) {
    this._assertKnownColumns(condition);
    this._assertKnownColumns(values);

    // useCache()を挟むことで、first()で読んだ内容をupdate()でも再利用し、
    // シートへのアクセスを1回にまとめる
    const queryable = this.useCache().where(condition);
    const existing = queryable.first();

    if (existing) {
      const result = queryable.update(values);
      return { action: "update", ...result };
    }

    const after = this.insert({ ...condition, ...values });
    return { action: "insert", count: 1, before: [], after: [after] };
  }

  // ================================================================
  // 内部処理
  // ================================================================

  /**
   * 現在のクエリの状態を引き継いだクローンを作る。
   * _sharedState は同じオブジェクトを参照する（クローン間で共有される、意図的な仕様）。
   * それ以外のクエリ状態（配列・オブジェクト）は浅くコピーし、
   * 親・子クローンが同じ配列/オブジェクトを参照して意図せず影響し合うことを防ぐ。
   */
  _clone() {
    const clone = new SSSQL(this.spreadsheetId, this.sheetName, this.headerRow, this.dataStartRow, this._sharedState);
    clone._condition = this._shallowCopy(this._condition);
    clone._orderBy = this._shallowCopy(this._orderBy);
    clone._useCache = this._useCache;
    clone._offset = this._offset;
    clone._groupBy = this._shallowCopy(this._groupBy);
    clone._selectCols = this._shallowCopy(this._selectCols);
    clone._aggregateSpec = this._shallowCopy(this._aggregateSpec);
    clone._havingCondition = this._shallowCopy(this._havingCondition);
    return clone;
  }

  /**
   * 値を浅くコピーする。配列は新しい配列に、オブジェクトは新しいオブジェクトにコピーする。
   * null/undefined/プリミティブ値はそのまま返す。ネストした中身まではコピーしない（浅いコピー）。
   */
  _shallowCopy(value) {
    if (Array.isArray(value)) return [...value];
    if (value && typeof value === "object") return { ...value };
    return value;
  }

  _assertKnownColumns(obj) {
    const unknownKeys = Object.keys(obj).filter(key => !this.headerSet.has(key));
    if (unknownKeys.length > 0) {
      throw new Error(`Unknown column(s): ${unknownKeys.join(", ")}`);
    }
  }

  /**
   * offset() が設定されている場合にエラーを投げる。
   * offset は all()/first()/take() の結果にのみ適用される仕様のため、
   * それ以外のメソッドで誤って使われた場合に、無視して動くのではなく気づけるようにする。
   */
  _assertNoOffset(methodName) {
    if (this._offset !== null) {
      throw new Error(`${methodName}() は offset() と併用できません（offsetは all()/first()/take() でのみ使用します）`);
    }
  }

  /**
   * having() が設定されているのに groupBy() が無い場合にエラーを投げる。
   */
  _assertHavingRequiresGroupBy() {
    if (this._havingCondition && !this._groupBy) {
      throw new Error("having() は groupBy() と組み合わせて使用してください");
    }
  }

  /**
   * orderBy() が設定されている状態で、groupBy() + count()/sum()/avg()/max()/min() を呼んだ場合にエラーを投げる。
   * これらはgroupBy()と組み合わせても最終的に単一の値を返すため、orderBy()の指定に意味が無くなるため。
   */
  _assertNoOrderByForScalarAggregate(methodName) {
    if (this._orderBy && this._orderBy.length > 0) {
      throw new Error(`${methodName}() は orderBy() と併用できません（groupBy()と組み合わせた場合、結果は単一の値になるため）`);
    }
  }

  /**
   * aggregate() が設定されている状態で、count()/sum()/avg()/max()/min() を呼んだ場合にエラーを投げる。
   * aggregate() は all()/first()/take() と組み合わせて使うものなので、
   * 単一集計メソッドと一緒に指定されていると、aggregate()の内容が黙って無視されてしまうため。
   */
  _assertNoAggregateSpec(methodName) {
    if (this._aggregateSpec) {
      throw new Error(`${methodName}() は aggregate() と併用できません（aggregateは all()/first()/take() と組み合わせて使用します）`);
    }
  }

  /**
   * offset()/take() に渡された値が「0以上の整数」であることを検証する。
   * 負数・小数・文字列などが渡された場合、sliceに予期しない挙動（負数は末尾から、
   * 文字列はNaN扱いになる等）をさせず、早期にエラーで気づけるようにする。
   */
  _assertNonNegativeInteger(n, methodName) {
    if (typeof n !== "number" || !Number.isInteger(n) || n < 0) {
      throw new Error(`${methodName}() には0以上の整数を指定してください（渡された値: ${JSON.stringify(n)}）`);
    }
  }

  _toRowValues(row) {
    this._assertKnownColumns(row);
    return this.header.map(key => (key in row ? row[key] : ""));
  }

  _appendRows(rowsValues) {
    const startRow = Math.max(this.sheet.getLastRow(), this.dataStartRow - 1) + 1;
    const numRows = rowsValues.length;
    const numCols = this.header.length;
    if (numCols === 0) return; // ヘッダーがない場合は書き込まない
    this.sheet.getRange(startRow, 1, numRows, numCols).setValues(rowsValues);
  }

  /**
   * all()/first()/take() が使う、実行結果の行データを組み立てる。
   * groupBy() が指定されていれば、グループキー（＋aggregate()の集計値）の配列を返す。
   * 指定されていなければ、通常の行データ（select()の列指定があれば絞る）を返す。
   */
  _resultRows() {
    if (this._aggregateSpec && !this._groupBy) {
      throw new Error("aggregate() は groupBy() と組み合わせて使用してください");
    }
    this._assertHavingRequiresGroupBy();

    if (this._groupBy) {
      if (this._selectCols) {
        throw new Error("groupBy() と select() は同時に使用できません。groupBy() で指定したカラムは自動的に結果へ含まれます");
      }
      return this._groupedAggregateRows();
    }

    const rows = this._filteredRows();
    return rows.map(r => this._selectCols ? this._pick(r.data, this._selectCols) : r.data);
  }

  /**
   * groupBy() + aggregate() の結果を組み立てる。
   * aggregate() が指定されていなければ、グループキーのみの配列を返す（distinctに近い）。
   */
  _groupedAggregateRows() {
    const groups = this._groupedRows();

    const result = groups.map(g => {
      if (!this._aggregateSpec) return { ...g.key };
      const row = { ...g.key };
      Object.entries(this._aggregateSpec).forEach(([outKey, spec]) => {
        row[outKey] = this._computeAggregate(g.rows, spec);
      });
      return row;
    });

    const allowedKeys = new Set([...this._groupBy, ...(this._aggregateSpec ? Object.keys(this._aggregateSpec) : [])]);
    return this._applyHavingAndOrder(result, allowedKeys);
  }

  /**
   * groupBy() + count()/sum()/avg()/max()/min() が使う。
   * 各グループに対して computeFn で集計値（outputKey）を計算し、having() で絞り込んだ後、
   * その「グループごとの集計値」だけを配列で返す（呼び出し元がさらに集計する）。
   * orderBy() はここでは適用しない（最終的に単一の値へ畳み込むため、順序に意味がない）。
   * @param {string} outputKey - having()の条件キーとして使える名前（例: "count", "age_sum"）
   * @param {(rows: Array) => *} computeFn - グループ内の行から集計値を計算する関数
   * @returns {Array} グループごとの集計値の配列（having()で絞り込んだ後）
   */
  _groupedValues(outputKey, computeFn) {
    const groups = this._groupedRows();
    const rows = groups.map(g => ({ ...g.key, [outputKey]: computeFn(g.rows) }));
    const allowedKeys = new Set([...this._groupBy, outputKey]);

    let filtered = rows;
    if (this._havingCondition) {
      this._assertHavingKeysValid(this._havingCondition, allowedKeys);
      filtered = filtered.filter(row => this._evaluate(row, this._havingCondition, false));
    }

    return filtered.map(row => row[outputKey]);
  }

  /**
   * groupBy()結果（プレーンなオブジェクトの配列）に having()/orderBy() を適用する共通処理。
   * キーはグループキーや集計の出力キーであり、シート上の実カラムではない場合があるため、
   * 列存在チェック（validateColumns）は行わない。ただし having() の条件キーは
   * allowedKeys（groupBy()のキー＋集計の出力キー）と照合し、タイポ等の未知キーを検出する。
   */
  _applyHavingAndOrder(rows, allowedKeys) {
    let result = rows;
    if (this._havingCondition) {
      this._assertHavingKeysValid(this._havingCondition, allowedKeys);
      result = result.filter(row => this._evaluate(row, this._havingCondition, false));
    }
    result = this._applyOrderBy(result, r => r, false);
    return result;
  }

  /**
   * having() の条件キー（AND/ORを再帰的に辿った先の各キー）が、
   * allowedKeys（groupBy()のキー＋aggregate()の出力キー）に含まれているか検証する。
   */
  _assertHavingKeysValid(condition, allowedKeys) {
    if (condition.AND) {
      condition.AND.forEach(sub => this._assertHavingKeysValid(sub, allowedKeys));
      return;
    }
    if (condition.OR) {
      condition.OR.forEach(sub => this._assertHavingKeysValid(sub, allowedKeys));
      return;
    }
    Object.keys(condition).forEach(key => {
      if (!allowedKeys.has(key)) {
        throw new Error(`Unknown column in having(): ${key} (having()で使えるのは groupBy() のキー、または aggregate() の出力キーのみです)`);
      }
    });
  }

  /**
   * aggregate() の1つの指定（例: { sum: "score" } や { count: true }）を計算する。
   */
  _computeAggregate(rows, spec) {
    const entries = Object.entries(spec);
    if (entries.length !== 1) {
      throw new Error(`aggregate() の各指定は1つの集計タイプを持つ必要があります: ${JSON.stringify(spec)}`);
    }
    const [type, column] = entries[0];

    switch (type) {
      case "count":
        return rows.length;
      case "sum": {
        const values = this._extractNumeric(rows, column);
        return values.reduce((a, b) => a + b, 0);
      }
      case "avg": {
        const values = this._extractNumeric(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
      }
      case "max": {
        const values = this._extractComparable(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => (b > a ? b : a));
      }
      case "min": {
        const values = this._extractComparable(rows, column);
        return values.length === 0 ? null : values.reduce((a, b) => (b < a ? b : a));
      }
      default:
        throw new Error(`Unknown aggregate type: ${type}`);
    }
  }

  /**
   * this._groupBy に従って、フィルタ・ソート済みの行を「グループキー」ごとにまとめる。
   * @returns {Array<{key: Object, rows: Array}>} グループの配列（最初に出現した順）
   */
  _groupedRows() {
    this._groupBy.forEach(col => {
      if (!this.headerSet.has(col)) {
        throw new Error(`Unknown column: ${col}`);
      }
    });

    const filtered = this._rawFilteredRows();
    const map = new Map();
    const order = [];

    filtered.forEach(r => {
      const mapKey = this._groupBy.map(c => String(r.data[c])).join("\u0000");
      if (!map.has(mapKey)) {
        map.set(mapKey, { key: this._pick(r.data, this._groupBy), rows: [] });
        order.push(mapKey);
      }
      map.get(mapKey).rows.push(r);
    });

    return order.map(k => map.get(k));
  }

  /**
   * 行の配列から、指定カラムの数値だけを抜き出す（sum/avg用）
   */
  _extractNumeric(rows, column) {
    return rows
      .map(r => r.data[column])
      .filter(v => typeof v === "number" && !isNaN(v));
  }

  /**
   * 行の配列から、指定カラムの数値またはDateだけを抜き出す（max/min用）
   */
  _extractComparable(rows, column) {
    return rows
      .map(r => r.data[column])
      .filter(v => (typeof v === "number" && !isNaN(v)) || v instanceof Date);
  }

  /**
   * 条件・ソート順を適用した行データを返す（offsetは適用しない）。
   */
  _filteredRows() {
    return this._applyOrderBy(this._rawFilteredRows(), r => r.data, true);
  }

  /**
   * 条件だけを適用した行データを返す（ソート・offsetは適用しない）。
   * groupBy()を使う場合、グループ化前の生データの並び順は結果に意味を持たない
   * （orderBy()は集計後の結果に適用されるべきもの）ため、こちらを使う。
   */
  _rawFilteredRows() {
    const rows = this._getRows();
    return rows.filter(r => this._evaluate(r.data, this._condition || {}));
  }

  /**
   * this._orderBy に従って items をソートする（汎用）。
   * getData: 各要素から「並び替えに使うデータオブジェクト」を取り出す関数
   *   - 通常の行（{rowIndex, data}）の場合は r => r.data
   *   - groupBy済みの結果（プレーンなオブジェクト）の場合は r => r（そのまま）
   * validateColumns: trueなら、指定カラムがシートのヘッダーに存在するかチェックする
   *   （groupBy結果はaggregate()の出力キーなど、シート上の実カラムでない場合があるためfalseにする）
   */
  _applyOrderBy(items, getData, validateColumns) {
    if (!this._orderBy || this._orderBy.length === 0) return items;

    if (validateColumns) {
      this._orderBy.forEach(({ column }) => {
        if (!this.headerSet.has(column)) {
          throw new Error(`Unknown column: ${column}`);
        }
      });
    }

    return [...items].sort((a, b) => {
      for (const { column, direction } of this._orderBy) {
        const sign = direction === "asc" ? 1 : -1;
        let va = getData(a)[column];
        let vb = getData(b)[column];

        // 空データ（空文字、null、undefined）は常に後ろに送る
        const vaIsEmpty = va === "" || va === null || va === undefined;
        const vbIsEmpty = vb === "" || vb === null || vb === undefined;
        if (vaIsEmpty && vbIsEmpty) continue;
        if (vaIsEmpty) return 1;
        if (vbIsEmpty) return -1;

        // 日付同士はミリ秒に変換してから比較する（_evaluateFieldと同じ正規化ロジックを使う）
        va = this._normalizeForComparison(va);
        vb = this._normalizeForComparison(vb);

        if (va === vb) continue; // 同値なら次の列で比較
        if (va < vb) return -1 * sign;
        if (va > vb) return 1 * sign;
      }
      return 0;
    });
  }

  _getRows() {
    if (this._useCache && this._sharedState.rows) {
      return this._sharedState.rows;
    }
    const rows = this._fetchAllRows();
    this._sharedState.rows = rows; // 常に保存する（useCacheの指定に関わらず）
    return rows;
  }

  _fetchAllRows() {
    const lastRow = this.sheet.getLastRow();
    const lastCol = this.sheet.getLastColumn();

    if (lastCol === 0) {
      // 完全な空シート: header関連も確定させておく（未確定のままにしない）
      this._sharedState.header = [];
      this._sharedState.headerSet = new Set();
      return [];
    }
    if (lastRow < this.dataStartRow) return [];

    if (!this._sharedState.header) {
      // ヘッダー未キャッシュ時: ヘッダー行〜データ最終行を1回のAPI呼び出しでまとめて取得
      const startRow = this.headerRow;
      const numRows = lastRow - startRow + 1;
      const allValues = this.sheet.getRange(startRow, 1, numRows, lastCol).getValues();

      const header = allValues[0];
      this._sharedState.header = header;
      this._sharedState.headerSet = new Set(header);

      const dataOffset = this.dataStartRow - startRow;
      const dataValues = allValues.slice(dataOffset);

      return dataValues.map((row, i) => ({
        rowIndex: this.dataStartRow + i,
        data: this._rowToDict(row)
      }));
    }

    // ヘッダーは既にキャッシュ済み → データだけ取得
    const numRows = lastRow - this.dataStartRow + 1;
    const values = this.sheet.getRange(this.dataStartRow, 1, numRows, lastCol).getValues();
    return values.map((row, i) => ({
      rowIndex: this.dataStartRow + i,
      data: this._rowToDict(row)
    }));
  }

  _rowToDict(row) {
    const dict = {};
    this.header.forEach((key, colIndex) => { dict[key] = row[colIndex]; });
    return dict;
  }

  /**
   * 条件オブジェクトを再帰的に評価する（AND/ORのネストに対応）。
   * validateColumns: trueなら、条件のキーがシートのヘッダーに存在するか検証する。
   *   having()で使う場合はfalseにする（キーがaggregate()の出力キーなど、実カラムでないことがあるため）。
   */
  _evaluate(row, condition, validateColumns = true) {
    if (condition.AND) {
      return condition.AND.every(sub => this._evaluate(row, sub, validateColumns));
    }
    if (condition.OR) {
      return condition.OR.some(sub => this._evaluate(row, sub, validateColumns));
    }
    return Object.entries(condition).every(([key, value]) => {
      if (validateColumns && !this.headerSet.has(key)) {
        throw new Error(`Unknown column: ${key}`);
      }
      return this._evaluateField(row[key], value);
    });
  }

  /**
   * 比較演算子で使うために値を正規化する。Dateインスタンスはミリ秒の数値に変換する。
   * それ以外の値はそのまま返す（文字列の自動Date変換など、暗黙の変換は行わない）。
   */
  _normalizeForComparison(value) {
    if (value instanceof Date) return value.getTime();
    return value;
  }

  _evaluateField(cellValue, value) {
    if (!Array.isArray(value)) {
      return cellValue === value; // 完全一致
    }
    const [operator, ...args] = value;
    const cv = this._normalizeForComparison(cellValue);

    switch (operator) {
      case "=":             return cellValue === args[0];
      case "<>":            return cellValue !== args[0];
      case ">":             return cv > this._normalizeForComparison(args[0]);
      case ">=":            return cv >= this._normalizeForComparison(args[0]);
      case "<":             return cv < this._normalizeForComparison(args[0]);
      case "<=":            return cv <= this._normalizeForComparison(args[0]);
      case "BETWEEN":       return cv >= this._normalizeForComparison(args[0]) && cv <= this._normalizeForComparison(args[1]);
      case "NOT BETWEEN":   return !(cv >= this._normalizeForComparison(args[0]) && cv <= this._normalizeForComparison(args[1]));
      case "IN":            return args[0].includes(cellValue);
      case "NOT IN":        return !args[0].includes(cellValue);
      case "LIKE":          return this._like(cellValue, args[0]);
      case "NOT LIKE":      return !this._like(cellValue, args[0]);
      default:
        throw new Error(`Unknown operator: ${operator}`);
    }
  }

  _like(cellValue, pattern) {
    // SQLのLIKEを模倣: % は任意文字列、_ は任意の1文字
    const escaped = pattern
      .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
      .replace(/%/g, ".*")
      .replace(/_/g, ".");
    return new RegExp(`^${escaped}$`).test(String(cellValue));
  }

  _pick(row, cols) {
    const result = {};
    cols.forEach(col => {
      if (!this.headerSet.has(col)) {
        throw new Error(`Unknown column: ${col}`);
      }
      result[col] = row[col];
    });
    return result;
  }

  /**
   * 行番号の配列を「連続する区間」ごとにグループ化する。
   * update/deleteで、連続する行をまとめて1回のAPI呼び出しで処理するために使う。
   * @returns {Array<{start: number, count: number}>}
   */
  _groupConsecutiveRows(rowIndexes) {
    const sorted = [...rowIndexes].sort((a, b) => a - b);
    const groups = [];

    sorted.forEach(rowIndex => {
      const last = groups[groups.length - 1];
      if (last && rowIndex === last.start + last.count) {
        last.count++;
      } else {
        groups.push({ start: rowIndex, count: 1 });
      }
    });

    return groups;
  }
}

// ---- ライブラリ経由で呼び出すためのファクトリ関数 ----
function createSSSQL(spreadsheetId, sheetName, headerRow, dataStartRow) {
  return new SSSQL(spreadsheetId, sheetName, headerRow, dataStartRow);
}
