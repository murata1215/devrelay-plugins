# access-migration プラグイン仕様 v0.1.1 — Access 業務システムの移植を支援する Claude Code プラグイン

作成: 2026-09-14 ｜ 置き場所: `devrelay-plugins/doc/access-migration_spec_v0.1.1.md`（DevRelay プロジェクト devrelay-plugins、`cmtzl77ew02l9krskialm2k7m`）｜ 状態: 設計、未着手

---

## 1. 目的

Access（.accdb / .mdb）で作られた業務システムを他言語（第一目標は C#）へ移植する前工程を、Claude Code が
**読める・設計できる・検証できる**ようにするプラグイン。P1 で作った Capability 配布基盤の実用例の 1 つで、
特定の企業・システムに依存しない汎用品として `murata1215/devrelay-plugins`（索引名 `devrelay`）に置く。

DevRelay 側の関心は次の 3 点に限る。利用側（どの Access システムをどう移植するか）はこのドキュメントの範囲外。

1. Access のテキスト化と MCP は自作しない。**msaccess-vcs-addin v5**（export format 5.0、公開 API、MCP サーバー、AGENTS.md 自動生成）に乗る
2. プラグインは「読み方」「移植規約」「MCP の登録」を束ねるだけにし、Access を知るのはプラグインの中だけ（P1 の adapter 境界と同じ発想）
3. Plan フェーズで read-only MCP を使わせる経路（Settings → Allowed Tools、ユーザー × OS 単位）を前提に、必要なツール名を利用側に提示できるようにする

## 2. 前提と依存

| もの | 内容 |
|---|---|
| 実行環境 | Windows + Microsoft Access + DevRelay Agent。export 済みテキストの解析だけなら Linux Agent でも動く |
| add-in | msaccess-vcs-addin v5.0.1 以降（Releases の `Version_Control_v5.0.1.zip` → `Version Control.accda`）。導入・Export の GUI 操作は人間 |
| export 形式 | Export Format 5.0.0。`.form` / `.report` / `.macro`、queries は `.sql` + `.json`、tbldefs、relations、`AGENTS.md`、`vcs-options.json` |
| 検証用の題材 | Northwind Traders 2007 テンプレートの export（`access-sandbox` プロジェクト、Windows Agent DESKTOP-1E6SDOQ で export 済み、Windows 初の実機） |
| 配布 | 索引 `murata1215/devrelay-plugins` に entry を追加。利用側は Agent Settings（machine/user scope）または `.claude/settings.json`（project scope）で宣言 → Sync |
| Plan 許可 | MCP ツール名は `mcp__plugin_<plugin>_<server>__<tool>` の形。read-only のものだけを利用側が Allowed Tools に入れる（プラグインが自動で付けることはしない、P1 の方針どおり） |

## 3. プラグインの構成（案）

```
plugins/access-migration/
  .claude-plugin/plugin.json      # name: access-migration, version: 0.1.0
  README.md                       # 前提（add-in 導入、export 手順）と使い方
  skills/
    access-export-reading/SKILL.md   # export format 5.0 の読み方
    access-to-csharp/SKILL.md        # Access → C# 移植規約と成果物テンプレ
  commands/
    access-inventory.md              # /access-inventory: 棚卸し文書を生成する定型プロンプト（任意）
  .mcp.json                        # add-in の MCP サーバー登録（起動方式は §6 で確定してから）
```

### 3.1 skill: access-export-reading
- フォルダごとの意味と読む順番（tbldefs + relations → queries → forms → modules → reports）
- `.form` とレイアウト / VBA 分離ファイルの対応、埋め込みマクロの所在
- queries の `.sql`（実行される SQL）と `.json`（Design View 情報）の使い分け
- 無視してよいもの（`vcs-index.idx`、`logs/`、`themes/`、`.env`）
- export 側の `AGENTS.md` を最初に読むよう誘導する（add-in が生成する API / 読み方ガイド）

### 3.2 skill: access-to-csharp
- 対応表: テーブル → EF Core エンティティ、リレーション → ナビゲーション、選択クエリ → LINQ または生 SQL、アクションクエリ → サービス層メソッド、フォーム → 画面層（UI フレームワークは利用側が決める）、埋め込みマクロ / VBA → サービス層またはハンドラ
- そのまま移せないもの（`DoCmd`、`TempVars`、`Forms!...`、`DLookup` 系、Access 固有の SQL 方言）の扱い
- 成果物テンプレ: `inventory.md`（棚卸し）、`dependencies.md`（依存グラフ）、`csharp-mapping.md`（対応表）、`diff-test-plan.md`（差分テスト計画）

### 3.3 .mcp.json（add-in の MCP）

> v0.1.1 で訂正: 「権限は既定すべて off」は不正確だったため、P2-A の調査結果（`joyfullservice/msaccess-vcs-mcp`
> README の Security Model）に基づき既定値を書き分ける。

- MCP サーバー本体は add-in（VBA）とは別リポジトリの Python 製外部プロセス（`msaccess-vcs-mcp`）。
  add-in 側は HTTP コールバックで進捗を送る側であり、サーバーそのものではない
- 書き込み（import / rebuild / RunVBA）は **`ACCESS_VCS_DISABLE_WRITES` の既定が `false` のため既定で有効**。
  安全側に倒すには利用側 `.env` に `ACCESS_VCS_DISABLE_WRITES=true` を設定する必要がある
- VBA 実行は 2 段階で既定が異なる: `vcs_call_vba`（既存の public VBA 呼び出し）を許す
  `McpAllowCallVBA` は **既定 True**。`vcs_run_vba`（エージェント生成コードの任意実行）を許す
  `McpAllowRunVBA` は既定 Off で、Options 画面から人間が明示的に有効化する必要があり、
  MCP 側から `vcs_set_option` でも上げられない
- 既定 Off はこの 3 つ: **`McpAllowRunVBA`**（任意 VBA 実行）、**`McpAllowExecuteSQL`**（read-only SELECT）、
  **`McpAllowImport`**（source からの Object Import）
- プラグインが登録するのはサーバー定義のみ。権限は add-in 側の Options → MCP、またはサーバーの `.env` で人間が入れる
- Plan で使う候補は read-only 系（`vcs_get_version_info` / `vcs_list_objects` / `vcs_diff_database` /
  `vcs_check_vba_compiled` / `vcs_get_option` / `vcs_get_log` / `vcs_execute_sql`）。書き込み系
  （Import / RunVBA / rebuild 等）は v0.1 で扱わない

## 4. 実装サイクル計画（devrelay-plugins プロジェクト）

| サイクル | 種別 | 内容 | 完了条件 |
|---|---|---|---|
| P2-A | read-only 調査 | add-in の wiki「MCP and Automation」と Northwind export の `AGENTS.md` を読み、MCP サーバーの起動方式・ツール名・権限の対応を確定。`.mcp.json` の形を提案 | Plan に起動方式とツール名一覧が出る |
| P2-B | 実装 | プラグインの骨格（plugin.json、2 skill、README）と索引 entry を作成。version 0.1.0。CHANGELOG 1 行 | commit、索引で参照できる |
| P2-C | 実機検証（人間側） | Windows Agent で project scope 宣言 → Sync → `claude plugin list` で入っていること、Plan から skill が参照されることを確認 | Agent Settings の同期結果 installed 1 |
| P2-D | 実装 | `.mcp.json` を追加し version 0.2.0。利用側が Allowed Tools に入れるべき read-only ツール名を README に明記 | Plan で read-only MCP が呼べる（利用側で確認） |

Northwind export を devrelay-plugins のエージェントに読ませる必要がある場合は、`access-migration-sandbox` リポ（Northwind のみ、public/private どちらでも可）を ubuntu-prod に clone してパス参照する。

## 5. DevRelay core 側の関連残件

このプラグインとは独立だが、利用体験に効くもの。優先順は別途。

1. agents/windows の `server:config:update` 対応（Allowed Tools の即時反映。現状は再接続時のみ）
2. Settings → Allowed Tools の Save 時に既定値と一致する行を落として差分だけ保存する（既定からの削除が伝播しない凍結リスクの解消）
3. リポの CLAUDE.md にある「DB スキーマ変更は `migrate dev`」の記述を運用ルール（psql ALTER → `prisma generate`）に合わせて修正
4. P2 の他候補: codegraph プラグイン化、security-guidance の vendoring、typescript-lsp の効果確認（引継ぎ 2026-09-13 §5）

## 6. 未確認事項

- add-in の MCP サーバーの起動方式（Access プロセス内で待ち受けるのか、外部ブリッジか）と Claude Code への登録方法。P2-A で確定
- Plan フェーズで使う read-only ツールの正確な名前（`mcp__plugin_access-migration_<server>__<tool>`）
- project scope 宣言を Windows Agent で行ったときの prelaunch の挙動（Linux では E2E-A で確認済み、Windows は未検証）
- Export 直後の `AGENTS.md` にプラグインの skill と重複する内容がどれだけあるか（重複するなら skill は AGENTS.md への導線に寄せる）

## 7. 運用ルール（このプロジェクトの既存ルールを踏襲）

- Plan → 人間承認 → exec。AskUserQuestion 不使用。シェルは 1 コマンドずつ
- devrelay-plugins は CHANGELOG 1 行（CLAUDE.md のルール）。devlog は不要
- 索引リポは Web UI で手編集しない。すべて DevRelay プロジェクト devrelay-plugins のサイクルで変更する
- 特定企業・特定システムの名前、実データ、接続情報をプラグインに含めない

---

改訂履歴
- v0.1 2026-09-14 初版。利用側の指示書（access-sandbox）から DevRelay 側の関心事だけを分離。
- v0.1.1 2026-09-14 P2-A の調査結果を反映し §3.3 の既定値を訂正（MCP サーバーは別リポの外部ブリッジ。`ACCESS_VCS_DISABLE_WRITES` 既定 false・`McpAllowCallVBA` 既定 True、既定 Off は `McpAllowRunVBA` / `McpAllowExecuteSQL` / `McpAllowImport` の 3 つ）。
