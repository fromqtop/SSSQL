# SSSQL

Google スプレッドシートを DB ライクに操作するための GAS (Google Apps Script) ライブラリです。

シートの1つのタブを1つのテーブルとして扱い、SQL に近い書き心地でメソッドチェーンを使って読み書きできます。

```javascript
const db = SSSQL.createSSSQL("スプレッドシートID", "Users", 2);

const activeUsers = db.where({ status: "active" }).orderBy({ age: "desc" }).all();
```

## セットアップ

`SSSQL.gs` の内容を GAS プロジェクトに貼り付けます。ライブラリとして他プロジェクトから使う場合は `createSSSQL` 経由で呼び出します。

```javascript
// ライブラリ経由
const db = SSSQL.createSSSQL(spreadsheetId, sheetName, headerRow, dataStartRow);

// 同一プロジェクト内
const db = new SSSQL(spreadsheetId, sheetName, headerRow, dataStartRow);
```

| 引数 | 必須 | デフォルト | 説明 |
|---|---|---|---|
| `spreadsheetId` | ✓ | - | 対象スプレッドシートのID |
| `sheetName` | ✓ | - | 対象シート名 |
| `headerRow` | | `1` | ヘッダー行の行番号 |
| `dataStartRow` | | `headerRow + 1` | データ開始行の行番号 |

## 基本の使い方

```javascript
const db = SSSQL.createSSSQL(spreadsheetId, "Users", 2);

db.all();                                        // 全件
db.where({ status: "active" }).all();            // 条件指定
db.orderBy({ age: "desc" }).all();                // 並び替え
db.select("id", "name").all();                    // 列を絞る
db.where({ id: 5 }).first();                      // 1件だけ
db.where({ status: "active" }).count();           // 件数

db.insert({ id: 11, name: "山田次郎", status: "active" });
db.where({ id: 11 }).update({ status: "inactive" });
db.where({ id: 11 }).delete();
```

`where` / `orderBy` / `select` などは呼ぶたびに新しいクエリを返すだけで、元の `db` は変化しません。同じクエリを変数に入れて何度でも再利用できます。

## クエリの組み立て

| メソッド | 説明 |
|---|---|
| `where(condition)` | 検索条件を指定する。複数回呼ぶと上書きされる（ANDにしたい場合は `{ AND: [...] }` を書く） |
| `orderBy(order)` | 並び替え。`{ age: "desc", name: "asc" }` のように複数カラム指定可 |
| `select(...cols)` | 取得する列を絞る |
| `offset(n)` | `orderBy` で並べ替えた後の結果から、先頭 `n` 件をスキップする（`all` / `first` / `take` にのみ効く） |
| `readCache()` | 直前に取得したデータのキャッシュを参照する |

### where の条件記法

```javascript
db.where({ status: "active" });                          // 完全一致
db.where({ age: [">=", 20] });                            // 比較演算子
db.where({ age: ["BETWEEN", 10, 100] });                  // 範囲
db.where({ name: ["IN", ["田中太郎", "佐藤花子"]] });       // 複数値
db.where({ name: ["LIKE", "%太郎%"] });                    // 部分一致
db.where({ OR: [{ status: "active" }, { name: "田中太郎" }] }); // AND/ORのネスト
```

使える演算子: `=` `<>` `>` `>=` `<` `<=` `BETWEEN` `NOT BETWEEN` `IN` `NOT IN` `LIKE` `NOT LIKE`

`Date` 型のカラムは、値としての比較が必要な演算子（`=` `<>` `>` `>=` `<` `<=` `BETWEEN` `NOT BETWEEN` `IN` `NOT IN`）すべてで正しく比較できます（別インスタンスでも同じ日時なら一致します）。

## データを取得する

```javascript
db.all();                    // 全件
db.first();                  // 最初の1件、無ければnull
db.take(5);                  // 先頭5件
db.orderBy({age:"desc"}).offset(10).take(5); // 11〜15件目
```

## 集計する

```javascript
db.count();                       // 件数
db.sum("score");                  // 合計（数値以外は無視）
db.avg("age");                    // 平均（対象が0件ならnull）
db.max("createdAt");              // 最大（数値・Dateのみ対象）
db.min("age");                    // 最小
db.distinct("status");            // ユニークな値の一覧
db.distinct("status", "dept");    // 複数カラムの組み合わせでユニーク
db.exists();                      // 該当があるかどうか
```

これらは `offset()` / `take()` と併用できません（`all`/`first`/`take` は別枠の「取得」メソッドです）。

## グループ集計（groupBy）

`groupBy` の使い方は2パターンあります。

**1. グループごとの明細を一覧で見る（`aggregate` + `all`/`first`/`take`）**

```javascript
db.groupBy("department").aggregate({
  totalScore: { sum: "score" },
  maxScore: { max: "score" },
  memberCount: { count: "*" },
  emailCount: { count: "email" }
}).all();
// [{department: "営業", totalScore: 4500000, maxScore: 95, memberCount: 12, emailCount: 10}, ...]
```

`count: "*"` はグループの全行数、`count: "カラム名"` はそのカラムが空文字・`null`・`undefined` でない行数を返します。

`aggregate` を付けなければ、グループキーだけの一覧になります（`distinct` に近い動き）。

**2. グループ化した結果をさらに集計して単一の値を得る（`count`/`sum`/`avg`/`max`/`min`）**

```javascript
db.groupBy("department").count(); // 部署の数

// aggregateで作った列を対象に、さらに集計できる
db.groupBy("department").aggregate({ totalScore: { sum: "score" } }).sum("totalScore");
db.groupBy("department").aggregate({ totalScore: { sum: "score" } }).avg("totalScore");
```

この場合、`sum`/`avg`/`max`/`min` に指定できるのは、`groupBy` のキー、または `aggregate` で作った出力キーだけです。元データの列（例: `score` そのもの）は指定できません。

**グループを絞り込む（`having`）**

```javascript
db.groupBy("department").aggregate({ total: { sum: "score" } })
  .having({ total: [">=", 1000] })
  .all();
```

`having` の条件キーも同様に、`groupBy` のキーか `aggregate` の出力キーのみ使えます。

## データを書き込む

```javascript
db.insert({ id: 11, name: "山田次郎" });
db.insert([{ id: 12, name: "中島愛子" }, { id: 13, name: "小林大輔" }]);
```

未指定のカラムは空文字で埋まります。戻り値は実際に書き込まれた内容です。

```javascript
const result = db.where({ status: "inactive" }).update({ status: "active" });
// { count, before, after }

const result = db.where({ status: "inactive" }).delete();
// { count, deleted }
```

`update` / `delete` は `where()` を呼んでいないとエラーになります。全件を対象にしたい場合は `updateAll()` / `deleteAll()` を使います。

```javascript
db.updateAll({ status: "checked" }); // シートの全行を更新
db.deleteAll();                      // シートの全データを削除
```

```javascript
db.upsert({ id: 5 }, { name: "田中太郎", status: "active" });
// 該当があれば更新、なければ新規追加
```

## キャッシュ

同じデータに何度もアクセスする処理（レポート生成、集計など）では、`readCache()` を使うとシートへの読み込みを1回にまとめられます。

```javascript
db.readCache().where({ status: "active" }).all();     // 1回目：取得
db.readCache().where({ status: "inactive" }).count();  // 2回目：さきほど取得した内容を再利用
```

`readCache()` を付けなければ常に最新の内容を取得します。`insert` / `update` / `delete` / `upsert` を行うとキャッシュは自動的に破棄されます。

## エラーになる組み合わせ

以下は誤用に気づけるよう、意図的にエラーにしています。

| 組み合わせ | 理由 |
|---|---|
| `where()` なしの `update()` / `delete()` | 誤って全件を対象にする事故を防ぐため |
| `groupBy()` + `select()` | groupByのキーは自動的に結果へ含まれるため |
| `groupBy()` + `distinct()` | distinctは単独で使う設計のため |
| `groupBy()` + `orderBy()` + `count()`/`sum()`/`avg()`/`max()`/`min()` | 最終的に単一の値になるため、並び順に意味がない |
| `having()`（`groupBy()` なし） | havingはgroupByの結果に対する条件のため |
| `offset()`/`take()` の効果を期待した集計系メソッドや `update()`/`delete()` | offset/takeは`all`/`first`/`take`専用 |
| 存在しないカラム名の指定 | タイポに気づけるようにするため |
