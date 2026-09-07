import { describe, expect, it } from "vitest";

import { getTicketMailSources } from "./ticket-sources";

describe("getTicketMailSources", () => {
  it("同じ送信元アドレスを複数の source に登録していない", () => {
    const sources = getTicketMailSources();
    const allAddresses = sources.flatMap((source) => source.mailAddresses);
    const uniqueAddresses = new Set(allAddresses.map((address) => address.toLowerCase()));

    expect(allAddresses).toHaveLength(uniqueAddresses.size);
  });
});
