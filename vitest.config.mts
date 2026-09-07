import { defineConfig } from "vitest/config";

// Apps Script は `appsscript.json` の `Asia/Tokyo` で動く。テストも同じタイムゾーン
// に固定しないと、日時ロジックが実行マシンのローカルタイムで検証されてしまう。
export default defineConfig({
  test: {
    env: {
      TZ: "Asia/Tokyo",
    },
  },
});
