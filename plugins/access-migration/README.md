# access-migration

Access（.accdb / .mdb）で作られた業務システムを他言語（第一目標 C#）へ移植する前工程を、
Claude Code が **読める・設計できる・検証できる** ようにするプラグイン。

Access のテキスト化と MCP は自作せず、
[msaccess-vcs-addin](https://github.com/joyfullservice/msaccess-vcs-addin) v5（export format 5.0、
公開 API、MCP サーバー、`AGENTS.md` 自動生成）に乗ります。Access を知るのはこのプラグインの中だけです。

## 前提

- Windows + Microsoft Access + msaccess-vcs-addin v5.0.1 以降（`Version Control.accda`）。
  add-in の導入と Export の GUI 操作は人間が行います
- **export 済みのテキストを読むだけなら Linux Agent でも動きます**（MCP を使う場合を除く）

## 入っているもの（v0.1.0）

| skill | 内容 |
|---|---|
| `access-export-reading` | export format 5.0 のフォルダ構成・読む順番・無視してよいファイルを示す |
| `access-to-csharp` | Access → C# の対応表、移せない Access 固有機能の扱い、成果物テンプレ |

v0.1.0 では `.mcp.json` は同梱していません（下記「MCP について」参照、v0.2.0 で追加予定）。

## 使い方

利用側で Agent Settings（machine/user scope）または `.claude/settings.json`（project scope）に
`access-migration@devrelay` を宣言し、Sync してください。

1. msaccess-vcs-addin で対象 Access データベースを export する（人間の GUI 操作）
2. export フォルダをエージェントに読ませ、`access-export-reading` skill の手順で棚卸しする
3. `access-to-csharp` skill の対応表に沿って移植設計（`inventory.md` 等）を作る

## MCP について（v0.2.0 予告）

msaccess-vcs-addin v5 の MCP は、add-in（VBA）本体とは**別リポジトリの Python 製外部プロセス**
[`joyfullservice/msaccess-vcs-mcp`](https://github.com/joyfullservice/msaccess-vcs-mcp) です。
add-in 側は HTTP コールバックで進捗を送る側であり、MCP サーバーそのものではありません。
**Windows 専用**（COM 自動化のため Access 実機が必要）です。

想定する登録形（v0.2.0 で `.mcp.json` として同梱予定）:

```json
{
  "mcpServers": {
    "msaccess-vcs-mcp": {
      "command": "uvx",
      "args": ["msaccess-vcs-mcp@latest"]
    }
  }
}
```

対象データベースはこの設定ではなく利用側プロジェクトの `.env`（`ACCESS_VCS_DATABASE=...`、
gitignore 対象）で指定します。プラグイン自体には接続情報を含みません。

利用側が Plan フェーズの Allowed Tools に入れることを想定する **read-only** ツール（接頭辞
`mcp__plugin_access-migration_msaccess-vcs-mcp__`）:

| ツール | 内容 |
|---|---|
| `vcs_get_version_info` | MCP / add-in / Access のバージョン確認 |
| `vcs_list_objects` | データベース内オブジェクト一覧 |
| `vcs_diff_database` | データベースとソースの差分 |
| `vcs_check_vba_compiled` | VBA のコンパイル状態確認のみ |
| `vcs_get_option` | add-in オプションの読み取り |
| `vcs_get_log` | Export / Build ログの読み取り |
| `vcs_execute_sql` | SELECT のみ実行（`McpAllowExecuteSQL` を有効にした場合。既定は Off） |

書き込み系（`vcs_import_*` / `vcs_rebuild_*` / `vcs_run_vba` / `vcs_set_option` 等）は
このプラグインでは扱いません。プラグインは MCP サーバーの登録のみを行い、
権限（`.env` の `ACCESS_VCS_DISABLE_WRITES` や add-in Options → MCP の各ゲート）は
利用側の人間が設定します（自動で付与しません）。

### 安全側の既定について

`msaccess-vcs-mcp` は `ACCESS_VCS_DISABLE_WRITES` の既定が `false` のため、
**書き込み（import / rebuild 等）は既定で有効**です。読むだけの用途では、利用側の `.env` に

```bash
ACCESS_VCS_DISABLE_WRITES=true
```

を設定することを推奨します。VBA 実行の権限は 2 段階あり、既存の public VBA 呼び出し
（`McpAllowCallVBA`）は既定 True、エージェント生成コードの任意実行（`McpAllowRunVBA`）は既定 Off です。
`McpAllowRunVBA` / `McpAllowExecuteSQL` / `McpAllowImport` は add-in の **Options → MCP** で
人間が明示的に有効化する必要があります（MCP 側から `vcs_set_option` で上げることはできません）。

## 出典

- [joyfullservice/msaccess-vcs-addin](https://github.com/joyfullservice/msaccess-vcs-addin)
  （Wiki: [MCP and Automation](https://github.com/joyfullservice/msaccess-vcs-addin/wiki/MCP-and-Automation)）
- [joyfullservice/msaccess-vcs-mcp](https://github.com/joyfullservice/msaccess-vcs-mcp)
