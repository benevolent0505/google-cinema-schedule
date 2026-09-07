import { describe, expect, it } from "vitest";

import "./ticket";
import "./cinemacity";
import "./musashinokan";
import "./ticket-sources";

type TicketMailSource = {
  mailAddresses: readonly string[];
  parseBody: (body: string) => unknown;
};

const getTicketMailSources = (
  globalThis as typeof globalThis & {
    getTicketMailSources: () => TicketMailSource[];
  }
).getTicketMailSources;

describe("getTicketMailSources", () => {
  it("does not register the same sender address to more than one source", () => {
    const sources = getTicketMailSources();
    const allAddresses = sources.flatMap((source) => source.mailAddresses);
    const uniqueAddresses = new Set(allAddresses.map((address) => address.toLowerCase()));

    expect(allAddresses).toHaveLength(uniqueAddresses.size);
  });
});
