# AGENTS

このリポジトリで作業するエージェント向けのガイドです。

## プロジェクト概要

- シネマシティのチケット予約確認メールを Gmail から読み取り、未登録の上映を
  Google カレンダーへ追加する Google Apps Script (GAS) です。
- 実装は TypeScript、実行環境は Apps Script V8 です。
- `src/main.ts` が GAS のエントリーポイントと Gmail / Calendar 連携を担当します。
- `src/ticket.ts` がチケット型とメール本文を解析する純粋関数を担当します。
- `src/ticket.test.ts` が Vitest によるパーサーのユニットテストです。

## 重要なアーキテクチャ制約

- Apps Script V8 は ES Modules（`import` / `export`）をサポートしません。
- ビルド対象の `src/**/*.ts` では `import` / `export` を使わず、全ファイルで共有される
  グローバルスコープを前提に実装してください。
  - 例外はビルドから除外される `*.test.ts` です。テストでは Vitest の import を使えます。
- 別ファイルの型や関数は import せず、そのまま参照します。
- GAS やテストから名前で参照する関数は
  `Object.assign(globalThis, { functionName })` で公開します。
  - GAS のエントリーポイントを追加した場合は、公開対象にも追加してください。
  - 純粋関数をテストする場合は、対象ファイルをテストから副作用 import し、
    `globalThis` 経由で取得する既存パターンに合わせてください。
- グローバルスコープを共有するため、ファイルをまたいだトップレベル名の重複を
  避けてください。

## ビルドと配置

- `tsconfig.json` は厳格な型チェック専用で、出力を生成しません。
- `tsconfig.build.json` はテストを除外し、`src/` のコードを `dist/` へ出力します。
- `scripts/copy-manifest.mjs` がルートの `appsscript.json` を
  `dist/appsscript.json` へコピーします。
- `.clasp.json` の `rootDir` は `dist` です。`clasp push` の対象はビルド成果物です。
- `dist/` は生成物です。直接編集せず、必要な変更は `src/` または
  `appsscript.json` に加えてください。

## 実装時の注意

- GAS 固有 API の型は `@types/google-apps-script` を利用します。DOM API や
  Node.js API を GAS 実行コードへ持ち込まないでください。
- Apps Script のタイムゾーンは `appsscript.json` の `Asia/Tokyo` です。
  日時処理を変更する場合はローカル環境との差に注意してください。
- メール本文の解析は外部サービスに依存しない純粋関数として保ち、形式変更や
  境界条件にはユニットテストを追加してください。
- 現在のチケットメール解析は CRLF (`\r\n`) を含む本文形式を前提にしています。
  正規表現を変更する場合は、実際の改行形式を意識してください。
- Gmail / Calendar へのアクセスを伴う処理はユニットテストから分離してください。
- 既存の命名、コメント、oxfmt のフォーマットに合わせ、無関係な変更を混ぜないで
  ください。

## 依存関係

- `@google/clasp` の要件に合わせ、Node.js 20 以上を使用してください。
- ロックファイルは `pnpm-lock.yaml` です。依存関係を変更する場合は
  `package.json` とロックファイルを整合させてください。
- 依存関係の追加・更新は、タスクに必要な場合だけ行ってください。

## 検証コマンド（エージェントが実行可）

変更内容に応じて、次を実行してください。

```sh
npm run typecheck
npm test
npm run lint
npm run format:check
```

- テストの反復実行が必要な場合は `npm run test:watch` を利用できますが、通常の最終確認は
  終了する `npm test` を使ってください。
- `npm run format` はファイルを書き換えます。意図して整形する場合のみ実行し、その後に
  差分を確認してください。

## エージェントが実行しないコマンド

次のコマンドはユーザーに実行を委ねてください。

```sh
npm run build
npm run push
npm run deploy
```

- 特に `push` / `deploy` は Apps Script のリモートを更新します。
- `build` も `dist/` を書き換えるため、このリポジトリの運用方針では
  エージェントから自動実行しません。
- `clasp push`、`clasp deploy`、`clasp run` など、GAS プロジェクトや実サービスへ
  接続するコマンドも実行しないでください。

## 変更時のチェックリスト

1. GAS 実行対象に `import` / `export` を追加していないか確認する。
2. 公開が必要な関数を `globalThis` に登録したか確認する。
3. パーサー変更には正常系・不正入力・必要な境界条件のテストを追加する。
4. `dist/` を直接編集していないか確認する。
5. 型チェック、テスト、Lint、フォーマットチェックを実行する。
6. リモート更新やビルドが必要な場合は、実行コマンドをユーザーへ案内する。
