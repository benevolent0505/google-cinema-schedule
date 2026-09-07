// Bundles `src/main.ts` (the Apps Script entry point) and every module it
// imports into a single `dist/main.js`.
//
// Apps Script's "Select function to run" dropdown and the trigger setup UI
// find functions by statically scanning the script for top-level `function`
// declarations - they don't execute the script or introspect `globalThis`.
// A plain esbuild IIFE bundle hides every declaration inside the wrapper
// function, so entry points would become invisible to both UIs even though
// they'd still work if invoked by name at runtime.
//
// `globalName` makes esbuild expose the entry module's exports as
// `ENTRY_POINTS_GLOBAL_NAME.<export>`, and the `footer` below adds real
// top-level `function` declarations that just delegate to them. Apps Script
// sees ordinary top-level functions; the implementation still lives in the
// bundle.
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
  // Keep Japanese log messages and comments literal instead of `\uXXXX`
  // escapes, so they stay readable in the Apps Script editor.
  charset: "utf8",
  globalName: entryPointsGlobalName,
  footer: {
    js: `
function main() {
  return ${entryPointsGlobalName}.main();
}

function debugMain(searchStartDateTime) {
  return ${entryPointsGlobalName}.debugMain(searchStartDateTime);
}
`,
  },
});

console.log("Bundled src/main.ts -> dist/main.js");
