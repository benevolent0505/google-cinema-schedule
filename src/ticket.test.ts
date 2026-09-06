import { describe, expect, it, vi } from "vitest";

import "./ticket";

type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

type TicketMailSource = {
  mailAddresses: readonly string[];
  parseBody: (body: string) => Ticket | undefined;
};

const parseTicketBody = (
  globalThis as typeof globalThis & {
    parseTicketBody: (body: string, sources: readonly TicketMailSource[]) => Ticket | undefined;
  }
).parseTicketBody;

const ticket: Ticket = {
  ticketNumber: "ticket-1",
  title: "テスト作品",
  startTime: new Date("2025-03-01T10:00:00+09:00"),
  endTime: new Date("2025-03-01T12:00:00+09:00"),
  theater: "テスト劇場",
  sheet: "A-1",
};

describe("parseTicketBody", () => {
  it("returns the first ticket parsed by a registered source", () => {
    const firstParser = vi.fn(() => undefined);
    const secondParser = vi.fn(() => ticket);
    const unusedParser = vi.fn(() => undefined);
    const sources: TicketMailSource[] = [
      { mailAddresses: ["first@example.com"], parseBody: firstParser },
      { mailAddresses: ["second@example.com"], parseBody: secondParser },
      { mailAddresses: ["unused@example.com"], parseBody: unusedParser },
    ];

    expect(parseTicketBody("mail body", sources)).toEqual(ticket);
    expect(firstParser).toHaveBeenCalledWith("mail body");
    expect(secondParser).toHaveBeenCalledWith("mail body");
    expect(unusedParser).not.toHaveBeenCalled();
  });

  it("returns undefined when none of the registered parsers match", () => {
    const sources: TicketMailSource[] = [
      {
        mailAddresses: ["cinema@example.com"],
        parseBody: () => undefined,
      },
    ];

    expect(parseTicketBody("unrelated body", sources)).toBeUndefined();
  });
});
