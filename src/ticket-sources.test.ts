import { describe, expect, it } from "vitest";

import { getTicketMailSources } from "./ticket-sources";

describe("getTicketMailSources", () => {
  it("does not register the same sender address to more than one source", () => {
    const sources = getTicketMailSources();
    const allAddresses = sources.flatMap((source) => source.mailAddresses);
    const uniqueAddresses = new Set(allAddresses.map((address) => address.toLowerCase()));

    expect(allAddresses).toHaveLength(uniqueAddresses.size);
  });
});
