# google-cinema-schedule

映画館のチケット予約確認メールを Gmail から読み取り、未登録の上映を
Google カレンダーへ自動登録する Google Apps Script (GAS) です。

現在は立川シネマシティと新宿武蔵野館に対応しています。映画館ごとのメール送信元とパーサーを
登録する構成のため、他の映画館も追加できます。

Google Apps Script を **TypeScript** で記述し、`tsc` でビルドして
[`clasp`](https://github.com/google/clasp) でスクリプトの更新・デプロイを行います。

## 特徴

- TypeScript で記述し、`tsc` でプレーンな GAS 用スクリプト（`dist/`）へビルド
- `clasp` でプッシュ／デプロイ
- [Vitest](https://vitest.dev/) による純粋関数のユニットテスト
- [oxlint](https://oxc.rs/docs/guide/usage/linter) / [oxfmt](https://oxc.rs/) による Lint・フォーマット

## 前提

- Node.js（`package.json` の devDependencies を利用）
- 依存関係のインストール: `npm install`
- `clasp` へのログイン: `npm exec clasp login`
- `.clasp.json` の `scriptId` を自分の Apps Script プロジェクトの ID に変更
  （新規作成する場合は `npm exec clasp create-script`、既存の場合は
  `npm exec clasp clone-script <scriptId>` が利用できます）

## ディレクトリ構成

```
.
├── src/
│   ├── main.ts         # GAS のエントリーポイント（main）と Gmail/Calendar 連携
│   ├── ticket.ts       # 共通の Ticket 型とパーサー呼び出し
│   ├── ticket-sources.ts # メール送信元とパーサーの登録一覧
│   ├── cinemacity.ts   # 立川シネマシティ用パーサー
│   ├── musashinokan.ts # 新宿武蔵野館用パーサー
│   └── *.test.ts       # Vitest のテスト
├── appsscript.json     # Apps Script マニフェスト（ビルド時に dist/ へコピー）
├── scripts/
│   └── copy-manifest.mjs
├── tsconfig.json       # 型チェック用
├── tsconfig.build.json # ビルド用（dist/ へ出力）
├── .clasp.json         # clasp 設定（rootDir = dist）
└── dist/               # ビルド成果物（git 管理外・clasp のプッシュ元）
```

## コマンド

> [!IMPORTANT]
> `build` / `push` / `deploy` などの **更新系コマンドは各自で実行してください**。

| コマンド               | 内容                                           |
| ---------------------- | ---------------------------------------------- |
| `npm run typecheck`    | 型チェック（`tsc --noEmit`）                   |
| `npm test`             | テストを実行（Vitest）                         |
| `npm run test:watch`   | テストをウォッチ実行                           |
| `npm run lint`         | oxlint による Lint                             |
| `npm run format`       | oxfmt でフォーマット                           |
| `npm run format:check` | フォーマット差分のチェック                     |
| `npm run build`        | `dist/` へビルド（`tsc` + マニフェストコピー） |
| `npm run push`         | ビルドして `clasp push`（リモート更新）        |
| `npm run deploy`       | ビルドして `clasp deploy`（デプロイ）          |

## 使い方

1. 依存関係をインストールし、`clasp` にログインして `scriptId` を設定します。
2. `src/` を編集します。編集中は `npm run typecheck` / `npm test` / `npm run lint`
   で検証できます。
3. リモートへ反映する準備ができたら、ビルドしてプッシュします（各自で実行）。

   ```sh
   npm run push      # ビルド + clasp push
   # デプロイまで行う場合
   npm run deploy    # ビルド + clasp deploy
   ```

4. Apps Script エディタで `main` を選択して実行するか、`npm exec clasp run main`
   で実行すると、Gmail のチケットメールを読み取りカレンダーへ登録します。

## エントリーポイントの追加方法

Apps Script はすべてのファイルで **1 つのグローバルスコープ** を共有し、ES Modules
（`import` / `export`）をサポートしません。そのため本プロジェクトでは各ファイルを
グローバルスコープのスクリプトとして記述します。

- 関数は `import` なしで他ファイルから参照できます。
- エディタ／トリガー／`clasp run` から呼び出す関数は、`main.ts` の
  `Object.assign(globalThis, { ... })` に追加してください。
- テストしたい純粋関数は `ticket.ts` のように分離し、`Object.assign(globalThis, { ... })`
  で公開すると Vitest から読み込めます。

## 対応映画館の追加方法

1. `src/cinemacity.ts` を参考に、メール本文を `Ticket` へ変換する映画館別パーサーを
   新しいファイルへ追加します。対象外または不正な本文では `undefined` を返します。
2. パーサーを `Object.assign(globalThis, { parserName })` で公開します。
3. `src/ticket-sources.ts` の `getTicketMailSources` に、検索対象の送信元アドレスと
   パーサーを追加します。

```ts
{
  mailAddresses: ["ticket@example-cinema.jp"],
  parseBody: parseExampleCinemaBody,
}
```

GAS 実行対象のファイルでは `import` / `export` を使わず、映画館別のヘルパー関数名には
映画館固有の接頭辞を付けて、グローバルスコープ上の名前重複を避けてください。
