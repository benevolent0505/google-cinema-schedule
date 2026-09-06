import { defineConfig } from "vitest/config";

// Apps Script always runs with the `Asia/Tokyo` timezone configured in
// `appsscript.json`. Pin the same timezone for tests so date/time logic
// (e.g. `formatMailSearchDate`, `main`'s "1 day before" calculation) is
// verified against the runtime environment instead of the host machine's
// local timezone.
export default defineConfig({
  test: {
    env: {
      TZ: "Asia/Tokyo",
    },
  },
});
