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

  it("reports a throwing source via onParseError and still tries the remaining sources", () => {
    const parseError = new Error("Missing required field: seats");
    const throwingSource: TicketMailSource = {
      mailAddresses: ["broken@example.com"],
      parseBody: vi.fn(() => {
        throw parseError;
      }),
    };
    const matchingSource: TicketMailSource = {
      mailAddresses: ["cinema@example.com"],
      parseBody: vi.fn(() => ticket),
    };
    const onParseError = vi.fn();

    const result = parseTicketBody("mail body", [throwingSource, matchingSource], onParseError);

    expect(result).toEqual(ticket);
    expect(onParseError).toHaveBeenCalledWith(throwingSource, parseError);
  });

  it("reports every throwing source and returns undefined when none parse successfully", () => {
    const firstError = new Error("Missing required field: ticketNumber");
    const secondError = new Error("Missing required field: seats");
    const sources: TicketMailSource[] = [
      {
        mailAddresses: ["first@example.com"],
        parseBody: () => {
          throw firstError;
        },
      },
      {
        mailAddresses: ["second@example.com"],
        parseBody: () => {
          throw secondError;
        },
      },
    ];
    const onParseError = vi.fn();

    expect(parseTicketBody("mail body", sources, onParseError)).toBeUndefined();
    expect(onParseError).toHaveBeenNthCalledWith(1, sources[0], firstError);
    expect(onParseError).toHaveBeenNthCalledWith(2, sources[1], secondError);
  });
});
