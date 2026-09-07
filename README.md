# google-cinema-schedule

映画館のチケット予約確認メールを Gmail から読み取り、まだ登録されていない上映を
Google カレンダーへ自動登録する Google Apps Script です。立川シネマシティと
新宿武蔵野館の予約メールに対応しています。

TypeScript で書き、esbuild で 1 つのスクリプトへバンドルして
[clasp](https://github.com/google/clasp) で Apps Script プロジェクトへ反映します。

## セットアップ

Node.js 20 以上と [pnpm](https://pnpm.io/) が必要です。

```sh
pnpm install
pnpm exec clasp login
```

次に、反映先となる Apps Script プロジェクトを用意します。新規に作る場合は
`create-script`、既存のプロジェクトを使う場合は `clone-script` です。

```sh
pnpm exec clasp create-script
# 既存プロジェクトを使う場合
pnpm exec clasp clone-script <scriptId>
```

どちらのコマンドも `.clasp.json` を生成します。このファイルは git 管理外なので、
生成後に `rootDir` をビルド成果物のディレクトリへ変更してください。

```json
{
  "scriptId": "<あなたの scriptId>",
  "rootDir": "dist"
}
```

## 実行

ビルドして Apps Script へ反映します。

```sh
pnpm push
```

反映したら、Apps Script エディタで `main` を選んで実行するか、時間主導トリガーを
設定するか、`pnpm exec clasp run main` を実行します。初回実行時に Gmail と
Google カレンダーへのアクセス承認を求められます。

デプロイまで行う場合は `pnpm deploy` を使います。

## 運用

### 実行日を指定したデバッグ実行

`debugMain` は、任意の日を「実行日」とみなして `main` と同じ処理を実行します。
実行日は `YYYY-MM-DD` 形式で、Apps Script のタイムゾーン（`Asia/Tokyo`）の日付として
解釈されます。

```sh
pnpm exec clasp run debugMain --params '["2026-09-05"]'
```

Apps Script エディタから実行する場合は、Script Properties の
`DEBUG_EXECUTION_DATE` に同じ形式で実行日を設定してください（引数があればそちらを
優先します）。

### ログ

既定では `Skip:` / `Registered:` のような実行結果のサマリと、送信元不明・解析失敗の
警告とエラーだけを出力します。Script Properties の `DEBUG_LOG_ENABLED` に `true` /
`1` / `yes` / `on` のいずれかを設定すると、検索条件やメッセージごとの詳細な
トレースも出力します。

メール本文はログに残しません。代わりに Gmail への直リンクを出力するので、必要なら
元メールをたどれます。

## 開発

| コマンド            | 内容                       |
| ------------------- | -------------------------- |
| `pnpm typecheck`    | 型チェック                 |
| `pnpm test`         | テストを実行               |
| `pnpm test:watch`   | テストをウォッチ実行       |
| `pnpm lint`         | oxlint による Lint         |
| `pnpm format`       | oxfmt でフォーマット       |
| `pnpm format:check` | フォーマット差分のチェック |
| `pnpm build`        | `dist/` へビルド           |
| `pnpm push`         | ビルドして `clasp push`    |
| `pnpm deploy`       | ビルドして `clasp deploy`  |

## 対応映画館を追加する

パーサーの追加手順は [AGENTS.md](AGENTS.md#対応映画館の追加方法) にあります。

## ライセンス

[MIT License](LICENSE)
