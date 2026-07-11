class SSSQL {
  constructor(spreadsheetId, sheetName, headerRow, dataStartRow, sharedState) {
    this.spreadsheetId = spreadsheetId;
    this.sheetName = sheetName;
    this.headerRow = headerRow !== undefined ? headerRow : 1;
    this.dataStartRow = dataStartRow !== undefined ? dataStartRow : this.headerRow + 1;

    // クローン間で共有される状態（どのクローンから書き換えても、全員に伝わる）
    this._sharedState = sharedState || {
      sheet: null,
      sheetId: null, // シートの数値ID（DeleteDimensionRequestで使う。取得したらキャッシュする）
      header: null,
      headerSet: null,
      rows: { normal: null, sheetsApi: null } // 行データのキャッシュ（通常モードとSheets APIモードは別々に持つ。
                                                // 日付の型が異なる(Date/シリアル値)ため混在させない）
    };

    // クエリの状態（クローンごとに個別。where/orderBy/groupBy/select/offset/readCache/aggregateで新しいクローンに設定される）
    this._condition = null;
    this._orderBy = null;
    this._useCache = false;
    this._offset = null;
    this._groupBy = null;
    this._selectCols = null;
    this._aggregateSpec = null;
    this._havingCondition = null;
    this._useSheetsApi = false;
    this._useKeyScan = false;
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

  /**
   * シートの数値ID（gid）。DeleteDimensionRequest（行削除）で必要になる。
   * SpreadsheetApp経由（this.sheet.getSheetId()）で取得し、一度取得したらキャッシュする。
   * Sheets APIモードであっても、この値の取得にはSpreadsheetAppを使う
   * （Sheets APIのメタデータ取得より軽量で、追加のAPI呼び出しコストに大きな差がないため）。
   */
  get sheetId() {
    if (this._sharedState.sheetId === null) {
      this._sharedState.sheetId = this.sheet.getSheetId();
    }
    return this._sharedState.sheetId;
  }

  get header() {
    if (!this._sharedState.header) {
      if (this._useSheetsApi) {
        this._fetchHeaderViaSheetsApi();
      } else {
        const lastCol = this.sheet.getLastColumn();
        const header = lastCol === 0 ? [] : this.sheet.getRange(this.headerRow, 1, 1, lastCol).getValues()[0];
        this._sharedState.header = header;
        this._sharedState.headerSet = new Set(header);
      }
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
   * spec の形式: { 出力キー名: { sum|avg|max|min: "カラム名" } | { count: "*"|"カラム名" } }
   * count: "*" は全行数、count: "カラム名" はそのカラムが空文字・null・undefinedでない行数を返す。
   *
   * @example
   * db.groupBy("department").aggregate({
   *   totalScore: { sum: "score" },
   *   maxScore: { max: "score" },
   *   memberCount: { count: "*" },
   *   emailCount: { count: "email" }
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
   * 次の実行でキャッシュを参照する（読み込みは常に行われ、これはあくまで「参照するかどうか」の指定）。
   * 呼ぶたびに新しいクローンを返す。
   */
  readCache() {
    const clone = this._clone();
    clone._useCache = true;
    return clone;
  }

  /**
   * 次の読み取り実行で、Sheets API（高度なサービス）を使ってデータを取得する。
   * 呼ぶたびに新しいクローンを返す。事前にGASプロジェクトで
   * 「サービス」から Sheets API を有効化しておく必要がある。
   *
   * 現時点では読み取り（select/count/first/all/take等）のみ対応。
   * insert/update/delete は、このモードでも SpreadsheetApp 経由のまま動作する。
   *
   * 注意: 日付セルはJSの Date オブジェクトではなく、シリアル値（数値）として返る。
   * BETWEEN/orderBy/max/min などの大小比較は結果的に正しく動くが、
   * 見た目は日付ではなくただの数値になる。
   */
  useSheetsApi() {
    const clone = this._clone();
    clone._useSheetsApi = true;
    return clone;
  }

  /**
   * where()の条件に登場する列だけを先に読み込んで絞り込み、
   * 該当する行だけを取得する2段階の読み込み戦略に切り替える。
   * useSheetsApi() と組み合わせた場合のみ有効（Sheets APIのbatchGetで、
   * 飛び飛びの行でも1回のリクエストでまとめて取得できるため）。
   * useSheetsApi() を伴わない場合は無視され、通常通り動作する
   * （SpreadsheetApp には、複数の行範囲を1回でまとめて取得する手段が無いため）。
   * 呼ぶたびに新しいクローンを返す。
   */
  keyScan() {
    const clone = this._clone();
    clone._useKeyScan = true;
    return clone;
  }

  /**
   * キャッシュを破棄する。全クローンで共有している箱(_sharedState)を書き換えるので、
   * どのインスタンスから呼んでも、そこから生まれた／派生した全クローンに伝わる。
   */
  refreshCache() {
    this._sharedState.rows = { normal: null, sheetsApi: null };
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
   * groupBy()と組み合わせた場合は、having()で絞り込んだ後の「グループの数」を返す。
   */
  count() {
    this._assertNoOffset("count");
    this._assertHavingRequiresGroupBy();
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("count");
      const { rows } = this._groupedResultRows();
      return rows.length;
    }
    return this._filteredRows().length;
  }

  /**
   * 指定したカラムの合計値を返す。数値でない値は無視する。
   * groupBy()と組み合わせた場合、column には groupBy()のキー、または aggregate()の出力キー
   * （groupBy()で作られた「結果テーブル」に実際に存在するカラム）を指定する。
   * それらに存在しないカラムを指定するとエラーになる。
   */
  sum(column) {
    this._assertNoOffset("sum");
    this._assertHavingRequiresGroupBy();
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("sum");
      const { rows, allowedKeys } = this._groupedResultRows();
      if (!allowedKeys.has(column)) {
        throw new Error(`Unknown column: ${column} (groupBy()のキー、またはaggregate()の出力キーを指定してください)`);
      }
      const values = this._extractNumericPlain(rows, column);
      return values.reduce((a, b) => a + b, 0);
    }
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    const values = this._extractNumeric(this._filteredRows(), column);
    return values.reduce((a, b) => a + b, 0);
  }

  /**
   * 指定したカラムの平均値を返す。数値でない値は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合、column には groupBy()のキー、または aggregate()の出力キーを指定する。
   */
  avg(column) {
    this._assertNoOffset("avg");
    this._assertHavingRequiresGroupBy();
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("avg");
      const { rows, allowedKeys } = this._groupedResultRows();
      if (!allowedKeys.has(column)) {
        throw new Error(`Unknown column: ${column} (groupBy()のキー、またはaggregate()の出力キーを指定してください)`);
      }
      const values = this._extractNumericPlain(rows, column);
      return values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
    }
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    const values = this._extractNumeric(this._filteredRows(), column);
    return values.length === 0 ? null : values.reduce((a, b) => a + b, 0) / values.length;
  }

  /**
   * 指定したカラムの最大値を返す。数値・Date以外は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合、column には groupBy()のキー、または aggregate()の出力キーを指定する。
   */
  max(column) {
    this._assertNoOffset("max");
    this._assertHavingRequiresGroupBy();
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("max");
      const { rows, allowedKeys } = this._groupedResultRows();
      if (!allowedKeys.has(column)) {
        throw new Error(`Unknown column: ${column} (groupBy()のキー、またはaggregate()の出力キーを指定してください)`);
      }
      const values = this._extractComparablePlain(rows, column);
      return values.length === 0 ? null : values.reduce((a, b) => (b > a ? b : a));
    }
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
    }
    const values = this._extractComparable(this._filteredRows(), column);
    return values.length === 0 ? null : values.reduce((a, b) => (b > a ? b : a));
  }

  /**
   * 指定したカラムの最小値を返す。数値・Date以外は無視する。対象が0件の場合はnullを返す。
   * groupBy()と組み合わせた場合、column には groupBy()のキー、または aggregate()の出力キーを指定する。
   */
  min(column) {
    this._assertNoOffset("min");
    this._assertHavingRequiresGroupBy();
    if (this._groupBy) {
      this._assertNoOrderByForScalarAggregate("min");
      const { rows, allowedKeys } = this._groupedResultRows();
      if (!allowedKeys.has(column)) {
        throw new Error(`Unknown column: ${column} (groupBy()のキー、またはaggregate()の出力キーを指定してください)`);
      }
      const values = this._extractComparablePlain(rows, column);
      return values.length === 0 ? null : values.reduce((a, b) => (b < a ? b : a));
    }
    if (!this.headerSet.has(column)) {
      throw new Error(`Unknown column: ${column}`);
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
    this._sharedState.rows = { normal: null, sheetsApi: null }; // 書き込みしたのでキャッシュは破棄

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

    if (this._useSheetsApi) {
      this._writeGroupsViaSheetsApi(groups, afterByRowIndex);
    } else {
      groups.forEach(({ start, count }) => {
        const rowsValues = [];
        for (let i = 0; i < count; i++) {
          rowsValues.push(afterByRowIndex.get(start + i));
        }
        this.sheet.getRange(start, 1, count, this.header.length).setValues(rowsValues);
      });
    }

    this._sharedState.rows = { normal: null, sheetsApi: null }; // 書き込みしたのでキャッシュは破棄

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

    if (this._useSheetsApi) {
      this._deleteGroupsViaSheetsApi(groups);
    } else {
      const sortedGroups = [...groups].sort((a, b) => b.start - a.start);
      sortedGroups.forEach(({ start, count }) => {
        this.sheet.deleteRows(start, count);
      });
    }

    this._sharedState.rows = { normal: null, sheetsApi: null }; // 書き込みしたのでキャッシュは破棄

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

    // readCache()を挟むことで、first()で読んだ内容をupdate()でも再利用し、
    // シートへのアクセスを1回にまとめる
    const queryable = this.readCache().where(condition);
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
    clone._useSheetsApi = this._useSheetsApi;
    clone._useKeyScan = this._useKeyScan;
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
    if (this._useSheetsApi) {
      this._appendRowsViaSheetsApi(rowsValues);
      return;
    }
    const startRow = Math.max(this.sheet.getLastRow(), this.dataStartRow - 1) + 1;
    const numRows = rowsValues.length;
    const numCols = this.header.length;
    if (numCols === 0) return; // ヘッダーがない場合は書き込まない
    this.sheet.getRange(startRow, 1, numRows, numCols).setValues(rowsValues);
  }

  /**
   * Sheets API（高度なサービス）を使って行を追記する。this.sheet（SpreadsheetApp経由）には触れない。
   * Date は事前にISO文字列へ変換し、valueInputOption: "USER_ENTERED" で送る
   * （Sheets側で日付として自動認識されるようにするため）。
   */
  _appendRowsViaSheetsApi(rowsValues) {
    if (this.header.length === 0) return; // ヘッダーがない場合は書き込まない

    const values = rowsValues.map(row => row.map(v => this._toSheetsApiValue(v)));
    const resource = { values };

    try {
      Sheets.Spreadsheets.Values.append(
        resource,
        this.spreadsheetId,
        this.sheetName,
        { valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS" }
      );
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }
  }

  /**
   * 書き込む値をSheets API向けに変換する。Dateはユーザーが手入力したのと同じ形式（ISO文字列）にし、
   * valueInputOption: "USER_ENTERED" と組み合わせることで、Sheets側に日付として認識させる。
   */
  _toSheetsApiValue(value) {
    if (this._isDate(value)) {
      // "YYYY-MM-DD HH:mm:ss" 形式（USER_ENTEREDでSheetsが日付として解釈できる書式）
      const pad = n => String(n).padStart(2, "0");
      return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} `
        + `${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
    }
    return value;
  }

  /**
   * Sheets API（高度なサービス）を使って、連続する行のグループごとに値を書き込む（update用）。
   * this.sheet には触れない。1回の batchUpdate リクエストにまとめて送る。
   * @param {Array<{start: number, count: number}>} groups - _groupConsecutiveRows() の結果
   * @param {Map<number, Array>} afterByRowIndex - 行番号 -> 書き込む値（ヘッダー順の配列）
   */
  _writeGroupsViaSheetsApi(groups, afterByRowIndex) {
    if (groups.length === 0) return;
    const numCols = this.header.length;
    const lastColLetter = this._columnLetter(numCols);

    const data = groups.map(({ start, count }) => {
      const rowsValues = [];
      for (let i = 0; i < count; i++) {
        rowsValues.push(afterByRowIndex.get(start + i).map(v => this._toSheetsApiValue(v)));
      }
      const endRow = start + count - 1;
      return {
        range: `${this.sheetName}!A${start}:${lastColLetter}${endRow}`,
        values: rowsValues
      };
    });

    try {
      Sheets.Spreadsheets.Values.batchUpdate(
        { valueInputOption: "USER_ENTERED", data },
        this.spreadsheetId
      );
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }
  }

  /**
   * 1始まりの列番号を、A1記法の列文字（1→A, 26→Z, 27→AA, ...）に変換する。
   */
  _columnLetter(colNumber) {
    let n = colNumber;
    let letters = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  }

  /**
   * Sheets API（高度なサービス）を使って、連続する行のグループごとに行を削除する（delete用）。
   * DeleteDimensionRequestは spreadsheets.batchUpdate（Values.batchUpdateとは別のAPI）で行い、
   * シートの数値ID（this.sheetId）が必要になる。
   * 行削除は下の行が繰り上がるため、行番号が大きいグループから順にリクエストを並べる。
   * @param {Array<{start: number, count: number}>} groups - _groupConsecutiveRows() の結果
   */
  _deleteGroupsViaSheetsApi(groups) {
    if (groups.length === 0) return;

    // 行番号が大きい方から処理されるよう、リクエストの順番を降順にする
    const sortedGroups = [...groups].sort((a, b) => b.start - a.start);

    const requests = sortedGroups.map(({ start, count }) => ({
      deleteDimension: {
        range: {
          sheetId: this.sheetId,
          dimension: "ROWS",
          startIndex: start - 1,      // DeleteDimensionRequestは0始まり
          endIndex: start - 1 + count // 半開区間（endIndexは含まない）
        }
      }
    }));

    try {
      Sheets.Spreadsheets.batchUpdate({ requests }, this.spreadsheetId);
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }
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
   * groupBy() + aggregate() の「結果テーブル」を組み立てる（having()で絞り込み済み、orderBy()は未適用）。
   * 各行は { ...グループキー, ...aggregate()の出力キー（指定されていれば） } という形になる。
   * count()/sum()/avg()/max()/min() や、all()/first()/take()（groupBy時）の土台になる。
   * @returns {{rows: Object[], allowedKeys: Set<string>}}
   */
  _groupedResultRows() {
    const groups = this._groupedRows();

    const rows = groups.map(g => {
      const row = { ...g.key };
      if (this._aggregateSpec) {
        Object.entries(this._aggregateSpec).forEach(([outKey, spec]) => {
          row[outKey] = this._computeAggregate(g.rows, spec);
        });
      }
      return row;
    });

    const allowedKeys = new Set([...this._groupBy, ...(this._aggregateSpec ? Object.keys(this._aggregateSpec) : [])]);

    let filtered = rows;
    if (this._havingCondition) {
      this._assertHavingKeysValid(this._havingCondition, allowedKeys);
      filtered = filtered.filter(row => this._evaluate(row, this._havingCondition, false));
    }

    return { rows: filtered, allowedKeys };
  }

  /**
   * groupBy() + aggregate() の結果を、all()/first()/take() 用に組み立てる（orderBy()も適用する）。
   * aggregate() が指定されていなければ、グループキーのみの配列を返す（distinctに近い）。
   */
  _groupedAggregateRows() {
    const { rows, allowedKeys } = this._groupedResultRows();
    return this._applyOrderBy(rows, r => r, allowedKeys);
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
   * aggregate() の1つの指定（例: { sum: "score" } や { count: "*" }）を計算する。
   * count: "*" は全行数、count: "カラム名" はそのカラムが空文字・null・undefined でない行数を返す。
   */
  _computeAggregate(rows, spec) {
    const entries = Object.entries(spec);
    if (entries.length !== 1) {
      throw new Error(`aggregate() の各指定は1つの集計タイプを持つ必要があります: ${JSON.stringify(spec)}`);
    }
    const [type, column] = entries[0];

    switch (type) {
      case "count":
        if (column === "*") return rows.length;
        return rows.filter(r => !this._isEmptyValue(r.data[column])).length;
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
   * 値が「空」（空文字・null・undefined）かどうかを判定する。
   * orderBy()のソート時の空判定と同じ基準（0やfalseは「値あり」として扱う）。
   */
  _isEmptyValue(value) {
    return value === "" || value === null || value === undefined;
  }

  /**
   * 値がDateかどうかを判定する。`instanceof Date` は使わない。
   * このライブラリがGASの「ライブラリ」として使われる場合、呼び出し元のスクリプトと
   * ライブラリ自身は別々の実行コンテキスト（レルム）で動くため、呼び出し元で作られた
   * Dateオブジェクトは、ライブラリ側の `Date` コンストラクタとは別物とみなされ、
   * `instanceof Date` が false になってしまう。Object.prototype.toString を使えば、
   * コンストラクタの参照に依存せず、内部的な型タグだけで判定できるため、
   * レルムをまたいでも正しく動作する。
   */
  _isDate(value) {
    return Object.prototype.toString.call(value) === "[object Date]";
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
      .filter(v => (typeof v === "number" && !isNaN(v)) || this._isDate(v));
  }

  /**
   * プレーンなオブジェクトの配列（groupBy()の結果テーブルなど）から、
   * 指定カラムの数値だけを抜き出す（sum/avg用）。_extractNumericとの違いは
   * 行が {rowIndex, data} 形式ではなく、そのままのオブジェクトである点。
   */
  _extractNumericPlain(rows, column) {
    return rows
      .map(r => r[column])
      .filter(v => typeof v === "number" && !isNaN(v));
  }

  /**
   * プレーンなオブジェクトの配列から、指定カラムの数値またはDateだけを抜き出す（max/min用）。
   */
  _extractComparablePlain(rows, column) {
    return rows
      .map(r => r[column])
      .filter(v => (typeof v === "number" && !isNaN(v)) || this._isDate(v));
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
   * keyScan() + useSheetsApi() + where() が揃っている場合は、
   * 条件に使われている列だけを先に読み込んで絞り込む2段階の取得戦略を使う。
   */
  _rawFilteredRows() {
    if (this._useKeyScan && this._useSheetsApi && this._condition) {
      return this._keyScanFetch();
    }
    const rows = this._getRows();
    return rows.filter(r => this._evaluate(r.data, this._condition || {}));
  }

  /**
   * this._orderBy に従って items をソートする（汎用）。
   * getData: 各要素から「並び替えに使うデータオブジェクト」を取り出す関数
   *   - 通常の行（{rowIndex, data}）の場合は r => r.data
   *   - groupBy済みの結果（プレーンなオブジェクト）の場合は r => r（そのまま）
   * allowedKeys: 列名の検証に使うキー集合。
   *   - true を渡すと、シートのヘッダー（headerSet）に対して検証する（通常の行データ用）
   *   - Set を渡すと、そのSetに対して検証する（groupBy結果用。groupBy()のキー＋aggregate()の出力キー）
   *   - false/null/undefined を渡すと検証しない
   */
  _applyOrderBy(items, getData, allowedKeys) {
    if (!this._orderBy || this._orderBy.length === 0) return items;

    if (allowedKeys === true) {
      this._orderBy.forEach(({ column }) => {
        if (!this.headerSet.has(column)) {
          throw new Error(`Unknown column: ${column}`);
        }
      });
    } else if (allowedKeys instanceof Set) {
      this._orderBy.forEach(({ column }) => {
        if (!allowedKeys.has(column)) {
          throw new Error(`Unknown column in orderBy(): ${column} (groupBy()のキー、またはaggregate()の出力キーを指定してください)`);
        }
      });
    }

    return [...items].sort((a, b) => {
      for (const { column, direction } of this._orderBy) {
        const sign = direction === "asc" ? 1 : -1;
        let va = getData(a)[column];
        let vb = getData(b)[column];

        // 空データ（空文字、null、undefined）は常に後ろに送る
        const vaIsEmpty = this._isEmptyValue(va);
        const vbIsEmpty = this._isEmptyValue(vb);
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
    // 通常モードとSheets APIモードでは日付の型が異なる(Date/シリアル値)ため、
    // キャッシュはモードごとに別々のキー(normal/sheetsApi)で持つ
    const cacheKey = this._useSheetsApi ? "sheetsApi" : "normal";

    if (this._useCache && this._sharedState.rows[cacheKey]) {
      return this._sharedState.rows[cacheKey];
    }
    const rows = this._fetchAllRows();
    this._sharedState.rows[cacheKey] = rows; // 常に保存する（readCacheの指定に関わらず）
    return rows;
  }

  _fetchAllRows() {
    if (this._useSheetsApi) {
      return this._fetchAllRowsViaSheetsApi();
    }

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

  /**
   * Sheets API（高度なサービス）を使って、ヘッダー行だけを取得する。this.sheet には触れない。
   * insert()単体のように、全データではなくヘッダーだけが必要な場合に使う軽量版。
   */
  _fetchHeaderViaSheetsApi() {
    const range = `${this.sheetName}!${this.headerRow}:${this.headerRow}`;
    let response;
    try {
      response = Sheets.Spreadsheets.Values.get(this.spreadsheetId, range, {
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
      });
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }

    const header = (response.values && response.values[0]) || [];
    this._sharedState.header = header;
    this._sharedState.headerSet = new Set(header);
  }

  /**
   * Sheets API（高度なサービス）を使ってデータを取得する。this.sheet（SpreadsheetApp経由）には触れない。
   * 日付セルはDateオブジェクトではなく、シリアル値（数値）として返る点に注意。
   */
  _fetchAllRowsViaSheetsApi() {
    let response;
    try {
      response = Sheets.Spreadsheets.Values.get(this.spreadsheetId, this.sheetName, {
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
      });
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }

    const allValues = response.values || [];
    const headerIndex = this.headerRow - 1;

    if (allValues.length <= headerIndex) {
      // ヘッダー行にすら到達しない空シート
      this._sharedState.header = [];
      this._sharedState.headerSet = new Set();
      return [];
    }

    if (!this._sharedState.header) {
      const header = allValues[headerIndex];
      this._sharedState.header = header;
      this._sharedState.headerSet = new Set(header);
    }
    const header = this._sharedState.header;

    const dataStartIndex = this.dataStartRow - 1;
    if (allValues.length <= dataStartIndex) return [];

    const dataRows = allValues.slice(dataStartIndex);
    return dataRows.map((row, i) => ({
      rowIndex: this.dataStartRow + i,
      data: this._rowToDict(this._padRow(row, header.length))
    }));
  }

  /**
   * keyScan(): where()の条件に登場する列だけを先に読み込んで、条件に一致する行番号を特定し、
   * 該当する行だけを batchGet でまとめて取得する。範囲は常に開いた形式（"C3:C"）で取得する
   * （Sheets APIは範囲を閉じても末尾の空セルを省略して返すため、閉じることに意味が無い）。
   *
   * 条件に「空文字が一致しうるリーフ条件」が含まれる場合（例: {status: ""} や {status: ["<>", "active"]}）、
   * 列の末尾に空セルが連続していると、Sheets APIがその部分を省略して返すため、
   * APIが返した配列の長さだけでは本来の行数を正しく把握できない。そのため、その場合だけ
   * 正確な最終行（getLastRow()）を取得し、それを基準にループ回数を決めることで、
   * 省略された分の行を「空文字」として正しく扱う（安全な場合はこの追加取得は行わない）。
   */
  _keyScanFetch() {
    const columns = [...this._collectConditionColumns(this._condition, new Set())];
    columns.forEach(col => {
      if (!this.headerSet.has(col)) {
        throw new Error(`Unknown column: ${col}`);
      }
    });

    if (columns.length === 0) {
      // 条件がAND/ORだけで実質キーが無い等、特定できなければ通常取得にフォールバック
      const rows = this._getRows();
      return rows.filter(r => this._evaluate(r.data, this._condition || {}));
    }

    // 空文字が条件に一致しうる場合、末尾の空セル省略によって行数を見誤るリスクがあるため、
    // 正確な最終行を確認しておき、ループする行数の計算に使う
    // （範囲自体は開いた形式のままでよい。閉じてもAPI側の末尾省略は解消されないため、
    // 　「APIが返した長さ」ではなく「lastRowから計算した長さ」でループし、
    // 　足りない分は空文字として扱うことで正しさを保証する）
    const needsExactLastRow = this._hasEmptyValueRisk(this._condition);
    const lastRow = needsExactLastRow ? this.sheet.getLastRow() : null;

    const colRanges = columns.map(col => {
      const colIndex = this.header.indexOf(col) + 1;
      const letter = this._columnLetter(colIndex);
      return `${this.sheetName}!${letter}${this.dataStartRow}:${letter}`;
    });

    const colResponse = this._sheetsApiBatchGet(colRanges);
    const valueRanges = colResponse.valueRanges || [];
    const scannedLen = needsExactLastRow
      ? Math.max(0, lastRow - this.dataStartRow + 1)
      : Math.max(0, ...valueRanges.map(vr => (vr.values || []).length));

    const matchingRowIndexes = [];
    for (let i = 0; i < scannedLen; i++) {
      const partialData = {};
      columns.forEach((col, colIdx) => {
        const vals = valueRanges[colIdx].values || [];
        partialData[col] = (vals[i] && vals[i][0] !== undefined) ? vals[i][0] : "";
      });
      if (this._evaluate(partialData, this._condition)) {
        matchingRowIndexes.push(this.dataStartRow + i);
      }
    }

    if (matchingRowIndexes.length === 0) return [];

    const lastColLetter = this._columnLetter(this.header.length);
    const rowRanges = matchingRowIndexes.map(r => `${this.sheetName}!A${r}:${lastColLetter}${r}`);
    const rowResponse = this._sheetsApiBatchGet(rowRanges);
    const rowValueRanges = rowResponse.valueRanges || [];

    return matchingRowIndexes.map((rowIndex, i) => {
      const row = (rowValueRanges[i] && rowValueRanges[i].values && rowValueRanges[i].values[0]) || [];
      return {
        rowIndex,
        data: this._rowToDict(this._padRow(row, this.header.length))
      };
    });
  }

  /**
   * where()の条件（AND/ORのネスト可）に登場する列名をすべて集める。
   */
  _collectConditionColumns(condition, into) {
    if (condition.AND) {
      condition.AND.forEach(sub => this._collectConditionColumns(sub, into));
      return into;
    }
    if (condition.OR) {
      condition.OR.forEach(sub => this._collectConditionColumns(sub, into));
      return into;
    }
    Object.keys(condition).forEach(key => into.add(key));
    return into;
  }

  /**
   * keyScan()用: 条件（AND/ORのネスト可）の中に、
   * 「空文字がその条件に一致してしまうリーフ条件」が含まれているか判定する。
   * 含まれていれば、末尾の空セル省略によって行を見逃す可能性があるため、
   * 正確な最終行を確認してから閉じた範囲でスキャンする必要がある。
   */
  _hasEmptyValueRisk(condition) {
    if (condition.AND) {
      return condition.AND.some(sub => this._hasEmptyValueRisk(sub));
    }
    if (condition.OR) {
      return condition.OR.some(sub => this._hasEmptyValueRisk(sub));
    }
    return Object.values(condition).some(value => this._evaluateField("", value));
  }

  /**
   * Sheets API の batchGet を呼ぶ共通処理。
   */
  _sheetsApiBatchGet(ranges) {
    try {
      return Sheets.Spreadsheets.Values.batchGet(this.spreadsheetId, {
        ranges,
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
      });
    } catch (e) {
      throw new Error(
        "Sheets API の呼び出しに失敗しました。GASプロジェクトの「サービス」から Sheets API (高度なサービス) を有効化しているか確認してください。元のエラー: " + e.message
      );
    }
  }

  /**
   * Sheets APIは各行末尾の空セルを詰めて返す（配列の長さがヘッダーより短くなりうる）ため、
   * 不足分を空文字で埋めて、SpreadsheetAppのgetValues()と同じ「矩形」の形に揃える。
   */
  _padRow(row, length) {
    if (row.length >= length) return row;
    const padded = row.slice();
    while (padded.length < length) padded.push("");
    return padded;
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
    if (this._isDate(value)) return value.getTime();
    return value;
  }

  _evaluateField(cellValue, value) {
    if (!Array.isArray(value)) {
      // 完全一致の省略記法（{status: "active"} など）。"="演算子と同じ扱いにする
      return this._normalizeForComparison(cellValue) === this._normalizeForComparison(value);
    }
    const [operator, ...args] = value;
    const cv = this._normalizeForComparison(cellValue);

    switch (operator) {
      case "=":             return cv === this._normalizeForComparison(args[0]);
      case "<>":            return cv !== this._normalizeForComparison(args[0]);
      case ">":             return cv > this._normalizeForComparison(args[0]);
      case ">=":            return cv >= this._normalizeForComparison(args[0]);
      case "<":             return cv < this._normalizeForComparison(args[0]);
      case "<=":            return cv <= this._normalizeForComparison(args[0]);
      case "BETWEEN":       return cv >= this._normalizeForComparison(args[0]) && cv <= this._normalizeForComparison(args[1]);
      case "NOT BETWEEN":   return !(cv >= this._normalizeForComparison(args[0]) && cv <= this._normalizeForComparison(args[1]));
      case "IN":            return args[0].some(v => this._normalizeForComparison(v) === cv);
      case "NOT IN":        return !args[0].some(v => this._normalizeForComparison(v) === cv);
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
