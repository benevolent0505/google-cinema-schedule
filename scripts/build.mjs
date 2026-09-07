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
