import { copyFileSync, mkdirSync } from "node:fs";

mkdirSync("dist", { recursive: true });
copyFileSync("appsscript.json", "dist/appsscript.json");
console.log("Copied appsscript.json -> dist/appsscript.json");
