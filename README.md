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

### 実行日を指定してデバッグ実行する

`debugMain` を使用すると、任意の日を「実行日」とみなして `main` と同じ処理を実行できます。
検索対象の期間は `main` と同じ導出ロジックを通るため、その日に `main` を実行していた場合と
同じ結果を再現できます。

対象になるのは、**指定した実行日の前日 0 時以上、翌日 0 時未満**に受信したメールです。
たとえば `2026-09-05` を指定すると、9/4 0:00 〜 9/5 23:59:59 に届いたメールが対象になります。

実行日は `YYYY-MM-DD` 形式で指定します。日付は Apps Script のタイムゾーン（`Asia/Tokyo`）
として解釈されます。形式違いや実在しない日付を指定した場合は、別の日で実行してしまわない
ようエラーで停止します。

`clasp run` から実行する場合は、実行日を引数で渡します。

```sh
npm exec clasp run debugMain --params '["2026-09-05"]'
```

Apps Script エディタから実行する場合は、プロジェクトの Script Properties に次の値を
設定してから、`debugMain` を選択して実行します。引数と Script Property の両方がある場合は
引数が優先されます。

| プロパティ名           | 設定例       |
| ---------------------- | ------------ |
| `DEBUG_EXECUTION_DATE` | `2026-09-05` |

実行時には、指定が効いているかどうかを確認できるよう、仮想実行日と対象範囲が
`DEBUG_LOG_ENABLED` の設定に関係なく必ずログへ出力されます。

```
debugMain: 仮想実行日=2026-09-05 対象範囲=2026-09-04 00:00 以上 2026-09-06 00:00 未満
```

なお `main` はこのプロパティを参照せず、常に実際の日付で動きます。プロパティの消し忘れで
定期トリガーが特定の日に固定されてしまうことはありません。デバッグ実行で登録された不要な
予定は、カレンダーから削除してください。

### ログ

ログは `console.log` / `console.info` / `console.warn` / `console.error` を使い分けており、
GAS の V8 ランタイム上でそれぞれ Cloud Logging の DEBUG / INFO / WARNING / ERROR severity に
対応します。`Skip:` / `Registered:` のような実行結果のサマリや、送信元不明・解析失敗の
警告・エラーは常に出力されます。検索条件やメッセージごとの詳細なトレースは、Script
Properties の `DEBUG_LOG_ENABLED` に `true` / `1` / `yes` / `on`（大文字小文字は無視）の
いずれかを設定した場合のみ出力されます。

個人情報を含むメール本文はログに出力しません。代わりに `message.getId()` から組み立てた
Gmail への直リンクをログへ残しているので、必要であれば元メールをたどれます。

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

`parseTicketBody`（`src/ticket.ts`）は、メールの送信元アドレスをもとに登録済みの
`source` を一意に選び、そのパーサーだけを呼び出します。ある映画館の複数パーサーへ
総当たりすることはないため、送信元アドレスと映画館は 1 対 1 で対応させてください。

1. `src/cinemacity.ts` を参考に、メール本文を `Ticket` へ変換する映画館別パーサーを
   新しいファイルへ追加します。本文が想定の形式と一致しない場合は `undefined` を
   返し（呼び出し元が解析失敗としてログに残します）、必須項目が欠落・不正な場合は
   例外を投げます。
2. パーサーを `Object.assign(globalThis, { parserName })` で公開します。
3. `src/ticket-sources.ts` の `getTicketMailSources` に、検索対象の送信元アドレスと
   パーサーを追加します。**同じ送信元アドレスを複数の `source` に登録しないでください**
   （`ticket-sources.test.ts` で重複がないことを検証しています）。

```ts
{
  mailAddresses: ["ticket@example-cinema.jp"],
  parseBody: parseExampleCinemaBody,
}
```

GAS 実行対象のファイルでは `import` / `export` を使わず、映画館別のヘルパー関数名には
映画館固有の接頭辞を付けて、グローバルスコープ上の名前重複を避けてください。
