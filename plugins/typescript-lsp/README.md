# typescript-lsp

TypeScript/JavaScript language server for Claude Code, providing code intelligence
features like go-to-definition, find references, and error checking.

## Supported Extensions

`.ts`, `.tsx`, `.js`, `.jsx`, `.mts`, `.cts`, `.mjs`, `.cjs`

## Installation

Install the TypeScript language server globally via npm:

```bash
npm install -g typescript-language-server typescript
```

Or with yarn:

```bash
yarn global add typescript-language-server typescript
```

## More Information

- [typescript-language-server on npm](https://www.npmjs.com/package/typescript-language-server)
- [GitHub Repository](https://github.com/typescript-language-server/typescript-language-server)

## 取り込み元 (Provenance)

このプラグインの LSP 定義（`lspServers`）は `.claude-plugin/marketplace.json` の
`typescript-lsp` エントリに inline されています。参照元:

- リポジトリ: `anthropics/claude-plugins-official`
- 対象: `.claude-plugin/marketplace.json` 内の `typescript-lsp` エントリ
  （公式リポの `plugins/typescript-lsp/` ディレクトリには LICENSE と README.md しか無く、
  LSP 定義自体は marketplace.json のエントリに inline されているため）
- commit: `3deb821cb71ccfaaf2ffa9935e977df314ce5cd5`
