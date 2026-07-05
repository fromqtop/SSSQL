# SSSQL

Google スプレッドシートを DB ライクに操作するための、GAS (Google Apps Script) 用ライブラリです。

シートの1つのタブを1つの「テーブル」として扱い、`where` / `orderBy` / `groupBy` / `select` などをメソッドチェーンでつないで、SQL に近い書き心地でスプレッドシートを読み書きできます。

```javascript
const db = SSSQL.createSSSQL("スプレッドシートID", "Users", 2);

const activeUsers = db.where({ status: "active" }).orderBy({ age: "desc" }).all();
```

## 目次

- [セットアップ](#セットアップ)
- [基本の使い方](#基本の使い方)
- [クエリ組み立て系メソッド](#クエリ組み立て系メソッド)
- [実行系メソッド](#実行系メソッド)
- [集計系メソッド](#集計系メソッド)
- [書き込み系メソッド](#書き込み系メソッド)
- [where の条件記法](#whereの条件記法)
- [groupBy / aggregate / having](#groupby--aggregate--having)
- [キャッシュ](#キャッシュ)
- [設計上の注意点](#設計上の注意点)
- [エラーになる組み合わせ一覧](#エラーになる組み合わせ一覧)

## セットアップ

1. GAS プロジェクトに `SSSQL.gs` の内容を貼り付けます。
2. ライブラリとして他プロジェクトから使う場合は、`createSSSQL` というファクトリ関数経由で呼び出します（GAS のライブラリは `class` を直接 public 化するのが不安定なため）。

```javascript
// ライブラリとして追加した場合
const db = SSSQL.createSSSQL(spreadsheetId, sheetName, headerRow, dataStartRow);

// 同一プロジェクト内で直接使う場合
const db = new SSSQL(spreadsheetId, sheetName, headerRow, dataStartRow);
```

### コンストラクタ引数

| 引数 | 必須 | デフォルト | 説明 |
|---|---|---|---|
| `spreadsheetId` | ✓ | - | 対象スプレッドシートのID |
| `sheetName` | ✓ | - | 対象シート（タブ）名 |
| `headerRow` | - | `1` | ヘッダー行の行番号（1始まり） |
| `dataStartRow` | - | `headerRow + 1` | データ開始行の行番号 |

```javascript
// ヘッダーが1行目、データが2行目から（デフォルト）
const db1 = SSSQL.createSSSQL(id, "Sheet1");

// ヘッダーが2行目、データは自動で3行目から
const db2 = SSSQL.createSSSQL(id, "Sheet1", 2);

// ヘッダーが2行目、データは間を空けて5行目から
const db3 = SSSQL.createSSSQL(id, "Sheet1", 2, 5);
```

## 基本の使い方

```javascript
const db = SSSQL.createSSSQL(spreadsheetId, "Users", 2);

// 全件取得
db.all();

// 条件を指定して取得
db.where({ status: "active" }).all();

// 並び替え
db.orderBy({ age: "desc" }).all();

// 列を絞る
db.select("id", "name").all();

// 1件だけ取得
db.where({ id: 5 }).first();

// 件数
db.where({ status: "active" }).count();

// 追加・更新・削除
db.insert({ id: 11, name: "山田次郎", status: "active" });
db.where({ id: 11 }).update({ status: "inactive" });
db.where({ id: 11 }).delete();
```

## クエリ組み立て系メソッド

これらは呼ぶたびに**新しいクローンを返し、元のインスタンスは一切変更しません**（immutable）。何度でも組み合わせて使えます。

### `where(condition)`

検索条件を指定します。詳しい条件の書き方は [where の条件記法](#whereの条件記法) を参照してください。複数回呼んだ場合は**上書き**になります（AND結合したい場合は `{ AND: [...] }` を明示します）。

```javascript
const q = db.where({ status: "active" });
q.first(); // 何度呼んでも同じ条件が使われる（消費されない）
q.first();
```

### `orderBy(order)`

並び替え順を指定します。複数カラムを指定する場合は、1つのオブジェクトにキーを追加した順が優先順位になります。

```javascript
db.orderBy({ age: "desc" }).all();
db.orderBy({ age: "desc", name: "asc" }).all(); // ageで降順、同じ年齢ならnameで昇順
```

### `select(...cols)`

取得する列を指定します（列を絞るだけで、実行はしません）。`groupBy()` と同時に使うとエラーになります。

```javascript
db.select("id", "name").all();
```

### `groupBy(...cols)`

グループ化するカラムを指定します。詳細は [groupBy / aggregate / having](#groupby--aggregate--having) を参照してください。

### `aggregate(spec)`

`groupBy()` と組み合わせて、グループごとの集計値を計算します。

### `having(condition)`

`groupBy()` + `aggregate()` の集計値に対する絞り込み条件です。`groupBy()` なしで使うとエラーになります。

### `offset(n)`

取得開始位置（先頭から何件スキップするか）を指定します。`all()` / `first()` / `take()` の結果にのみ適用されます。`n` は0以上の整数である必要があります。

```javascript
db.orderBy({ age: "asc" }).offset(10).take(5); // 11件目〜15件目
```

### `useCache()`

次の実行でキャッシュを利用します。詳細は [キャッシュ](#キャッシュ) を参照してください。

## 実行系メソッド

`offset()` が効くのはこの3つだけです。

### `all()`

条件・並び順・グループ化に従って、全件を取得します。

```javascript
db.where({ status: "active" }).all();
```

### `first()`

最初の1件だけを取得します。該当がなければ `null` を返します。

```javascript
db.where({ id: 5 }).first();
```

### `take(n)`

先頭からn件を取得します。`n` は0以上の整数である必要があります。

```javascript
db.orderBy({ age: "desc" }).take(3); // 年齢が高い上位3件
```

## 集計系メソッド

これらは単体で結果（数値や配列）を返します。`offset()` / `take()` とは併用できません。

### `count()`

条件に一致する行数を返します。`groupBy()` と組み合わせた場合は、条件に一致するグループの数を返します（詳細は [groupBy / aggregate / having](#groupby--aggregate--having) 参照）。

### `sum(column)` / `avg(column)` / `max(column)` / `min(column)`

指定カラムの合計・平均・最大・最小を返します。`sum` / `avg` は数値でない値を無視します。`max` / `min` は数値か `Date` 以外を無視します。対象が0件の場合、`avg` / `max` / `min` は `null` を返します。

```javascript
db.where({ status: "active" }).sum("score");
db.avg("age");
db.max("createdAt");
```

`groupBy()` と組み合わせた場合、**「グループ化した結果」に対して同じ集計関数をもう一度適用**します（詳細は [groupBy / aggregate / having](#groupby--aggregate--having) 参照）。単一の値が返り、`orderBy()` とは併用できません。

### `distinct(...cols)`

指定したカラム（複数可）の組み合わせでユニークな値を取得します。カラムを1つだけ指定した場合は値の配列、複数指定した場合はオブジェクトの配列を返します。`groupBy()` / `select()` と併用するとエラーになります。

```javascript
db.distinct("status");               // ["active", "inactive"]
db.distinct("status", "department"); // [{status:"active", department:"営業"}, ...]
```

### `exists()`

条件に一致する行が1件以上あるかどうかを `true`/`false` で返します。

```javascript
db.where({ email: "a@example.com" }).exists();
```

## 書き込み系メソッド

### `insert(rowOrRows)`

1行または複数行を新規追加します。ヘッダーに存在しないキーを渡すとエラーになります。未指定のカラムは空文字で埋められます。戻り値は「実際に書き込まれた内容」（空欄補完後）で、単一行ならオブジェクト、複数行なら配列を返します。

```javascript
db.insert({ id: 11, name: "山田次郎", status: "active" });
db.insert([
  { id: 12, name: "中島愛子" },
  { id: 13, name: "小林大輔" }
]);
```

### `where(condition).update(values)`

条件に一致する行を更新します。**`where()` を呼んでいない場合はエラーになります**（誤って全件更新することを防ぐため）。戻り値は `{ count, before, after }` です。

```javascript
const result = db.where({ status: "inactive" }).update({ status: "active" });
console.log(result.count, result.before, result.after);
```

意図的に全件更新したい場合は `updateAll(values)` を使います。

```javascript
db.updateAll({ status: "checked" }); // where()なしでも全件更新できる
```

### `where(condition).delete()`

条件に一致する行を削除します。**`where()` を呼んでいない場合はエラーになります**。戻り値は `{ count, deleted }` です。

```javascript
const result = db.where({ status: "inactive" }).delete();
```

意図的に全件削除したい場合は `deleteAll()` を使います。

```javascript
db.deleteAll(); // 注意: シートの全データが消えます
```

### `upsert(condition, values)`

条件に一致する行があれば更新し、無ければ「条件と `values` をマージしたデータ」で新規追加します。`condition` には完全一致の値のみ指定できます（比較演算子・AND/ORは使用できません）。

```javascript
db.upsert({ id: 5 }, { name: "田中太郎", status: "active" });
// 該当あり → update: { action: "update", count, before, after }
// 該当なし → insert: { action: "insert", count: 1, before: [], after: [...] }
```

## whereの条件記法

### 完全一致

```javascript
db.where({ status: "active" });          // status === "active"
db.where({ status: "active", age: 20 }); // 複数キーはAND
```

### 比較演算子

`[演算子, 値]` の配列で指定します。

| 演算子 | 意味 |
|---|---|
| `=` | 等しい |
| `<>` | 等しくない |
| `>` `>=` `<` `<=` | 比較 |
| `BETWEEN` / `NOT BETWEEN` | 範囲 |
| `IN` / `NOT IN` | 複数値のいずれか / いずれでもない |
| `LIKE` / `NOT LIKE` | 部分一致（`%` = 任意文字列、`_` = 任意の1文字） |

```javascript
db.where({ age: [">=", 20] });
db.where({ age: ["BETWEEN", 10, 100] });
db.where({ name: ["IN", ["田中太郎", "佐藤花子"]] });
db.where({ name: ["LIKE", "%太郎%"] });
```

`Date` 型のカラムに対する `>` `>=` `<` `<=` `BETWEEN` は、内部でミリ秒に変換して比較するため正しく動作します。`=` `<>` `IN` `NOT IN` は `Date` の値同士でも参照比較になるため非対応です（`BETWEEN` や `>=`/`<=` を使ってください）。

### AND / OR（ネスト可）

```javascript
db.where({
  OR: [
    { AND: [{ age: [">", 20] }, { status: "active" }] },
    { name: "田中太郎" }
  ]
});
// (age > 20 AND status = 'active') OR name = '田中太郎'
```

`where()` を複数回チェーンした場合も自動でAND結合されます。

```javascript
db.where({ age: [">", 20] }).where({ status: "active" });
```

## groupBy / aggregate / having

`groupBy()` には大きく2つの使い方があります。

- **`aggregate()` + `all()`/`first()`/`take()`**: グループごとの明細（グループキー＋集計値）を**配列**で見たい場合
- **`count()`/`sum()`/`avg()`/`max()`/`min()`**: グループ化した結果をさらに1段集計して、**単一の値**を得たい場合

### グループキーのみ取得（`distinct` に近い）

```javascript
db.groupBy("department").all();
// [{department: "営業"}, {department: "人事"}, ...]
```

### `aggregate()` で複数の集計を同時に取得（グループごとの明細を見たい場合）

複数の集計値を1回のクエリで得たい場合は `aggregate()` を使います。`aggregate()` は `groupBy()` と組み合わせて `all()` / `first()` / `take()` を呼んで実行し、**グループごとの配列**を返します。

```javascript
db.groupBy("department").aggregate({
  totalScore: { sum: "score" },
  maxScore: { max: "score" },
  memberCount: { count: true }
}).all();
// [{department: "営業", totalScore: 4500000, maxScore: 95, memberCount: 12}, ...]
```

### `count()`/`sum()`/`avg()`/`max()`/`min()` を `groupBy()` と組み合わせる（単一の値が欲しい場合）

これらのメソッドを `groupBy()` と組み合わせた場合、**「グループ化して求めた各グループの集計値」に対して、もう一度同じ集計関数を適用**し、最終的に単一の値を返します。`groupBy()` を使わない場合と同じく、常に「1つの値を返す」という性質は変わりません。

```javascript
// 部署でグループ化 → 部署の数を返す
db.groupBy("department").count();
// 3 (営業・人事・開発の3グループ)

// 部署ごとのscore合計を求め、さらにそれらを合計する
// (結果的に groupBy なしの sum("score") と同じ値になる)
db.groupBy("department").sum("score");

// 部署ごとのage平均を求め、さらにそれらを単純平均する（avg-of-avg）
// 部署の人数が均等でなければ、groupBy なしの avg("age") とは異なる値になる
db.groupBy("department").avg("age");

// 部署ごとのage最大値を求め、さらにその中の最大値を取る
// (結果的に groupBy なしの max("age") と同じ値になる)
db.groupBy("department").max("age");
db.groupBy("department").min("age");
```

`aggregate()` を設定した状態でこれらのメソッドを呼ぶとエラーになります（`aggregate()` は `all()`/`first()`/`take()` と組み合わせて使うものです）。また、最終的に単一の値を返す性質上、`orderBy()` との併用もエラーになります（`aggregate()` + `all()` の方は引き続き `orderBy()` を使えます）。

### `having()` で集計値を絞り込む

`having()` は「グループ化した結果」を絞り込みます。`aggregate()` + `all()` の場合はグループの明細を絞り込み、`count()`/`sum()` 等と組み合わせた場合は「絞り込んだ後のグループ」に対して集計します。

```javascript
// aggregate()+all(): 合計が1000以上の部署だけを一覧表示
db.groupBy("department").aggregate({ total: { sum: "score" } })
  .having({ total: [">=", 1000] })
  .all();

// count(): 人数が2人以上の部署が何個あるか
db.groupBy("department").having({ count: [">=", 2] }).count();
```

`having()` の条件キーは `groupBy()` のキー、または `aggregate()`/集計メソッドの出力キー（`count`、`カラム名_sum` など）である必要があります。それ以外のキー（タイポなど）を指定するとエラーになります。`groupBy()` なしで `having()` を使うこともエラーになります。

### `orderBy()` を集計結果に使う

`aggregate()` + `all()`/`first()`/`take()` の場合、`orderBy()` は「グループ化前の生データ」ではなく「集計後の結果（グループごとの行）」に対して適用されます。

```javascript
db.groupBy("department").aggregate({ total: { sum: "score" } })
  .orderBy({ total: "desc" })
  .all();
```

`count()`/`sum()`/`avg()`/`max()`/`min()` と `groupBy()` を組み合わせた場合は、結果が単一の値になるため `orderBy()` は併用できません（併用するとエラーになります）。

### `select()` は `groupBy()` と併用不可

`groupBy()` で指定したカラムは自動的に結果へ含まれるため、`select()` を同時に使うとエラーになります。

## キャッシュ

`useCache()` を呼ぶと、次に取得したデータをキャッシュに保存し、以降 `useCache()` 付きで呼んだ時にキャッシュを再利用します（シートへの読み込みを省略できます）。

```javascript
db.useCache().where({ status: "active" }).all();     // 1回目：取得してキャッシュに保存
db.useCache().where({ status: "inactive" }).count();  // 2回目：キャッシュを再利用（シートを読み直さない）
```

`useCache()` を付けなければ、常に最新の状態をシートから取得します（デフォルトの安全な挙動）。

`insert()` / `update()` / `delete()` / `upsert()` を実行すると、キャッシュは自動的に破棄されます。明示的に破棄したい場合は `refreshCache()` を呼びます。

**注意**: キャッシュは同じ `SSSQL` インスタンス（およびそこから作られたクローン）の間でのみ共有され、他プロセスや他の人がシートを直接編集した場合の変更には追従しません。同じスクリプト実行内で、同じデータに複数回アクセスする場合（集計・レポート生成など）に有効です。

## 設計上の注意点

- **immutable なクエリビルダー**: `where` / `orderBy` / `groupBy` / `select` / `aggregate` / `having` / `offset` / `useCache` は、呼ぶたびに新しいクローンを返します。元のインスタンスは変更されないため、同じクエリを何度でも再利用できます。
- **`offset` / `take` は `all()` / `first()` / `take()` 専用**: `count()` / `sum()` / `update()` / `delete()` などと併用するとエラーになります（`groupBy()` と組み合わせた集計系メソッドには効きます）。
- **`update()` / `delete()` は `where()` 必須**: 誤って全件を対象にしてしまう事故を防ぐため、`where()` を呼んでいない場合はエラーになります。全件操作したい場合は `updateAll()` / `deleteAll()` を使ってください。
- **API呼び出しの最適化**: ヘッダーとデータの取得は可能な限り1回のAPI呼び出しにまとめられます。`update` / `delete` で連続する行は、まとめて1回の書き込み/削除にまとめられます。

## エラーになる組み合わせ一覧

| 組み合わせ | 理由 |
|---|---|
| `groupBy()` + `select()` | groupByのキーは自動的に含まれるため |
| `groupBy()` + `distinct()` | distinctは単独で使う設計のため |
| `select()` + `distinct()` | distinctの引数がそのまま使われるため |
| `aggregate()` + `count()`/`sum()`/`avg()`/`max()`/`min()` | aggregateはall/first/takeと組み合わせるものであり、内容が無視されてしまうため |
| `groupBy()` + `orderBy()` + `count()`/`sum()`/`avg()`/`max()`/`min()` | これらはgroupByと組み合わせても最終的に単一の値を返すため、順序に意味がないため |
| `having()`（`groupBy()` なし） | havingはgroupByの集計結果に対する条件のため |
| `offset()`/`take()` の効果を期待した `count()`/`sum()`/`distinct()`/`exists()`/`update()`/`delete()` | offset/takeはall/first/take専用のため |
| `where()` なしの `update()`/`delete()` | 誤操作防止のため（`updateAll()`/`deleteAll()`を使う） |
| `offset()`/`take()` に負数・小数・文字列などを渡す | 0以上の整数のみ許可 |
| `where()`/`having()` で存在しないカラム名を指定 | タイポ等に気づけるようにするため |
