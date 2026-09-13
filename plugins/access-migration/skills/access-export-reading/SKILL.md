---
name: access-export-reading
description: MSAccess VCS Add-in が export した Access ソースツリー（export format 5.0）を読むときに使う。移植・棚卸しのためにフォルダ構成、読む順番、無視してよいファイルを示す。.form / .report / .macro / tbldefs / queries の .sql と .json が置かれたディレクトリを見つけたときに使用する。
---

# Access export（format 5.0）の読み方

このディレクトリ（`<db>.accdb.src/` のようなフォルダ）は
[msaccess-vcs-addin](https://github.com/joyfullservice/msaccess-vcs-addin) が Access データベースを
テキスト化して export したものです。**このプラグインは read-only の前工程を支援します**。
このフォルダを直接編集して Access に戻す作業（merge build 等）は対象外です。

## 0. 最初に読むもの

export 直下の `AGENTS.md` と `vcs-agent-docs/*.md` は **add-in が毎回の export で再生成する一次情報**です。
BOM / CRLF の扱い、`@Folder` 注釈によるファイル配置、merge build の手順など、
「このソースを編集して Access に戻す」ための正確な情報はそちらに書かれています。
この skill はそれを置き換えません。まずそちらを開いてください。

この skill が担当するのは、AGENTS.md には**書かれていない**、移植（他言語への書き換え）のための
読む順番と対応関係です。

## 1. 読む順番

データモデル → ロジック → 画面 → 帳票の順で読みます。先にモデルを固定しないと、後段の対応表が揺れます。

1. `tbldefs/` + `relations/`（テーブル定義とリレーション）
2. `queries/*.sql`（実行される SQL）
3. `forms/*.form` + `forms/*.cls`（画面レイアウトと code-behind）
4. `modules/`（標準モジュール・クラスモジュール）
5. `reports/*.report` + `reports/*.cls`
6. `macros/*.macro`

## 2. ファイル対応の早見表

```
<db>.accdb.src/
  AGENTS.md                 # add-in 生成。最初に読む
  vcs-agent-docs/*.md       # add-in 生成。種類別の詳細（vba-modules / forms-reports / queries / testing / troubleshooting）
  tbldefs/<name>.sql + .xml # テーブル定義
  relations/                # リレーション（関係が無い DB では出力されない）
  queries/<name>.sql        # 実行される SQL。唯一の真実
  queries/<name>.json       # Design View 情報・プロパティ（SQL の代わりにはならない）
  forms/<name>.form         # レイアウト
  forms/<name>.cls          # code-behind（VBA）。ロジックはここにある。@Folder 注釈もここだけ
  reports/<name>.report + .cls
  macros/<name>.macro       # 埋め込みマクロ
  modules/<name>.bas / .cls
  vcs-options.json          # export 設定
  vcs-index.idx             # バイナリ変更インデックス（読まない・触らない）
  logs/ , themes/ , .env    # 読まない
```

- `queries/*.sql` が SQL の唯一の真実。`.json` は Design View の見た目情報やプロパティで、SQL の代替ではない
- フォームやレポートのロジックは `.form` / `.report` 自体ではなく `.cls`（code-behind）にある
- 埋め込みマクロの本体は `.form` / `.report` の中、独立したものは `macros/*.macro`

## 3. 落とし穴

- **拡張子は export format のバージョンで変わる**。5.0.0 で `.bas` 一辺倒から `.form` / `.report` / `.macro` /
  `.sql`+`.json` に分離された。旧プロジェクトは今も `.bas` のままの場合があるので、
  読む前に実際にどの拡張子が存在するか確認する
- `@Folder("A.B")` 注釈でコンポーネントがサブディレクトリに散る（`modules/A/B/` のように）。
  同名ファイルが複数のツリーに存在しうるので、名前で検索するときはツリー全体を見る
- 全ファイル UTF-8 **BOM 付き** / **CRLF**。テキストツールで開き直すときに壊さないよう注意する
  （このプラグインは読むだけなので基本的に問題にならないが、差分取得時に留意する）

## 4. 読まない・触らない

`vcs-index.idx`（バイナリ）、`logs/`、`themes/`、`.env`（接続情報）、`.frx`（バイナリ OLE データ）、
`.thmx`（ZIP アーカイブ）。

## 5. これは read-only の前工程であること

この skill は棚卸し・設計のための読み方だけを扱います。import / merge build / RunVBA など
Access に書き戻す操作はこのプラグインの範囲外です（`access-to-csharp` skill も同様に read-only）。
