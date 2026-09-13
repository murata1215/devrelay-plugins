# Changelog

## 2026-09-14
- access-migration プラグイン v0.1.0 を追加（export format 5.0 の読み方と C# 移植規約の skill 2 つ。MCP サーバーは別リポ msaccess-vcs-mcp の外部ブリッジと確定し .mcp.json は v0.2.0 へ繰り延べ）

## 2026-09-13
- 公式プラグイン（commit-commands / pr-review-toolkit / context7 / typescript-lsp）を plugins/ に vendoring し、marketplace.json を全て ./plugins/<name> 参照に統一。security-guidance の git-subdir エントリを削除
