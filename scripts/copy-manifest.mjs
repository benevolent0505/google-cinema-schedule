// Copies the Apps Script manifest into the build output so that `clasp push`
// (which uploads everything under `dist/`) publishes it alongside the code.
import { copyFileSync, mkdirSync } from "node:fs";

mkdirSync("dist", { recursive: true });
copyFileSync("appsscript.json", "dist/appsscript.json");
console.log("Copied appsscript.json -> dist/appsscript.json");
