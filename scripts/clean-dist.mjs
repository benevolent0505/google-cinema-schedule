// Removes `dist/` before a build so stale output (e.g. leftover files from
// a previous bundle layout) never lingers alongside the new one.
import { rmSync } from "node:fs";

rmSync("dist", { recursive: true, force: true });
console.log("Removed dist/");
