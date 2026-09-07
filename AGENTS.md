# AGENTS

このリポジトリで作業するエージェント向けのガイドです。利用者向けの導入と運用は
[README.md](README.md) にあります。

## プロジェクト概要とファイル配置

映画館のチケット予約確認メールを Gmail から読み取り、未登録の上映を Google
カレンダーへ登録する Google Apps Script です。実装は TypeScript、実行環境は
Apps Script の V8 ランタイムです。

```
.
├── src/
│   ├── main.ts            # GAS のエントリーポイントと Gmail / Calendar 連携
│   ├── logger.ts          # レベル付きロガー
│   ├── ticket.ts          # Ticket 型と、送信元アドレスによるパーサーの振り分け
│   ├── ticket-sources.ts  # 送信元アドレスとパーサーの登録一覧
│   ├── ticket-parser.ts   # パーサー共通の型と抽出ヘルパー
│   ├── cinemacity.ts      # 立川シネマシティ用パーサー
│   ├── musashinokan.ts    # 新宿武蔵野館用パーサー
│   └── *.test.ts          # Vitest のテスト
├── scripts/               # ビルド用スクリプト
├── appsscript.json        # Apps Script マニフェスト（ビルド時に dist/ へコピー）
├── tsconfig.json          # 型チェック専用（noEmit）
└── dist/                  # ビルド成果物（git 管理外・clasp のプッシュ元）
```

## アーキテクチャ制約

- Apps Script V8 は ES Modules をサポートしないため、ビルド時に esbuild が
  `src/main.ts` を起点に依存モジュールをすべて 1 つの `dist/main.js` へ
  バンドルします。`src/**/*.ts` は通常の TypeScript モジュールとして書き、
  ファイル間は `import` / `export` でやり取りします（`*.test.ts` も対象モジュールを
  直接 import します）。
- Apps Script の「実行する関数を選択」とトリガー設定 UI は、スクリプトを実行せず
  **トップレベルの `function` 宣言を静的にスキャン**して関数を列挙します。IIFE
  バンドルでは全宣言がラッパー関数の中に隠れるため、`scripts/build.mjs` が
  esbuild の `globalName` と `footer` を使い、バンドルの外側に委譲用の
  トップレベル関数を生成しています。
- **GAS から呼ばれるエントリーポイントを追加した場合は、`src/main.ts` で
  `export` するだけでなく `scripts/build.mjs` の `footer` にも委譲用の関数を
  追加してください。** それ以外の関数は `export` するだけで十分で、テストからも
  直接 `import` します。

## ビルドと配置

`pnpm build` は次の順に実行します。

1. `scripts/clean-dist.mjs` が `dist/` を削除する。前のバンドル構成で生成された
   ファイルが残らないようにするため。
2. `scripts/build.mjs` が esbuild で `src/main.ts` を `dist/main.js` へ
   バンドルする（minify・sourcemap なし）。
3. `scripts/copy-manifest.mjs` がルートの `appsscript.json` を
   `dist/appsscript.json` へコピーする。`clasp push` は `.clasp.json` の
   `rootDir`（`dist`）配下だけを送るため、マニフェストも成果物側に置く必要がある。

`dist/` は生成物です。直接編集せず、変更は `src/` または `appsscript.json` に
加えてください。

## 実装方針

- GAS 固有 API の型は `@types/google-apps-script` を使います。DOM API や Node.js
  API を GAS 実行コードへ持ち込まないでください。
- Apps Script のタイムゾーンは `appsscript.json` の `Asia/Tokyo` です。テストも
  `vitest.config.mts` で同じタイムゾーンに固定しています。
- ログは `logger.ts` の `Logger` を使います。実行開始時（`main` / `debugMain`）に
  一度だけ生成し、以降の処理へ明示的に渡します。詳細なトレースは `debug` に
  分類し、それ以外のレベルは常に出力されます。個人情報を含むメール本文は
  ログに出力しません。
- メール本文の解析は外部サービスに依存しない純粋関数として保ち、Gmail /
  Calendar へのアクセスを伴う処理から分離してください。
- **テストの意図はコメントではなく `it` の説明文に書いてください。** `describe` は
  対象の関数名なので英語、`it` は日本語です。パーサーの変更には正常系・不正入力・
  境界条件のテストを追加してください。
- **コメントは「知らずに触ると壊す理由」だけを書きます。** 何をしているかの
  言い換えや、このファイルにある説明の再掲は書かないでください。
- 既存の命名と oxfmt のフォーマットに合わせ、無関係な変更を混ぜないでください。

## 対応映画館の追加方法

`parseTicketBody`（`src/ticket.ts`）は、メールの送信元アドレスから登録済みの
`source` を一意に選び、そのパーサーだけを呼びます。複数のパーサーへ総当たりは
しないため、送信元アドレスと映画館は 1 対 1 で対応させてください。

1. `src/cinemacity.ts` を参考に、メール本文を `Ticket` へ変換するパーサーを
   新しいファイルへ追加します。共通の型と抽出ヘルパーは `src/ticket-parser.ts`
   にあります。
2. パーサーは次の契約に従います。本文がその映画館の形式でなければ `undefined` を
   返し、形式には一致したのに必須項目が欠けている場合は例外を投げます。前者は
   「別の映画館のメール」、後者は「メール形式が変わった可能性」で、呼び出し側が
   区別してログに残します。
3. `src/ticket-sources.ts` の `getTicketMailSources` に送信元アドレスとパーサーを
   追加します。**同じ送信元アドレスを複数の `source` へ登録しないでください**
   （`ticket-sources.test.ts` で重複がないことを検証しています）。

```ts
{
  mailAddresses: ["ticket@example-cinema.jp"],
  parseBody: parseExampleCinemaBody,
}
```

## 検証コマンド

変更内容に応じて実行してください。

```sh
pnpm typecheck
pnpm test
pnpm lint
pnpm format:check
pnpm build
```

- `pnpm build` / `pnpm clean` は `dist/` を書き換えるだけで、リモートには影響
  しません。バンドル結果を確認したい場合は実行して構いません。
- `pnpm test:watch` は終了しないので、最終確認には `pnpm test` を使ってください。
- `pnpm format` はファイルを書き換えます。意図して整形する場合のみ実行し、その後に
  差分を確認してください。

## 実行しないコマンド

次のコマンドは Apps Script のリモートや実サービスへ接続するため、ユーザーに実行を
委ねてください。

```sh
pnpm push
pnpm deploy
```

`clasp push` / `clasp deploy` / `clasp run` を直接呼ぶことも同様に避けてください。

## 依存関係

- `@google/clasp` の要件に合わせ、Node.js 20 以上を使用してください。
- パッケージマネージャは pnpm、ロックファイルは `pnpm-lock.yaml` です。依存関係を
  変更する場合は `package.json` とロックファイルを整合させてください。
- 依存関係の追加・更新は、タスクに必要な場合だけ行ってください。
