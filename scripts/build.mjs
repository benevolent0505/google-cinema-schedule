// Apps Script の「実行する関数を選択」とトリガー設定 UI は、スクリプトを実行せず
// トップレベルの `function` 宣言を静的にスキャンして関数を列挙する。IIFE バンドル
// では全宣言がラッパー関数の中に隠れてエントリーポイントが選べなくなるため、
// `globalName` で公開したエクスポートへ委譲するトップレベル関数を `footer` で
// 生やしている。エントリーポイントを追加したら `footer` にも追加すること。
import * as esbuild from "esbuild";

const entryPointsGlobalName = "CinemaSchedule";

await esbuild.build({
  entryPoints: ["src/main.ts"],
  outfile: "dist/main.js",
  bundle: true,
  format: "iife",
  target: "es2019",
  platform: "neutral",
  minify: false,
  sourcemap: false,
  charset: "utf8",
  globalName: entryPointsGlobalName,
  footer: {
    js: `
function main() {
  return ${entryPointsGlobalName}.main();
}

function debugMain(executionDate) {
  return ${entryPointsGlobalName}.debugMain(executionDate);
}
`,
  },
});

console.log("Bundled src/main.ts -> dist/main.js");
