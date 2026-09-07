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

type TicketParseFailure =
  | { reason: "unknown_sender"; fromAddress: string }
  | { reason: "unrecognized_body"; source: TicketMailSource }
  | { reason: "error"; source: TicketMailSource; error: unknown };

const parseTicketBody = (
  globalThis as typeof globalThis & {
    parseTicketBody: (
      body: string,
      fromAddress: string,
      sources: readonly TicketMailSource[],
      onParseError?: (failure: TicketParseFailure) => void,
    ) => Ticket | undefined;
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
  it("calls only the parser registered to the sender address", () => {
    const otherParser = vi.fn(() => ticket);
    const matchingParser = vi.fn(() => ticket);
    const sources: TicketMailSource[] = [
      { mailAddresses: ["other@example.com"], parseBody: otherParser },
      { mailAddresses: ["cinema@example.com"], parseBody: matchingParser },
    ];

    expect(parseTicketBody("mail body", "cinema@example.com", sources)).toEqual(ticket);
    expect(matchingParser).toHaveBeenCalledWith("mail body");
    expect(otherParser).not.toHaveBeenCalled();
  });

  it("extracts the address from a display-name From header and ignores case", () => {
    const parser = vi.fn(() => ticket);
    const sources: TicketMailSource[] = [
      { mailAddresses: ["Cinema@Example.com"], parseBody: parser },
    ];

    expect(parseTicketBody("mail body", "映画館 <cinema@example.com>", sources)).toEqual(ticket);
    expect(parser).toHaveBeenCalledWith("mail body");
  });

  it("reports an unknown sender and calls no parser", () => {
    const parser = vi.fn(() => ticket);
    const sources: TicketMailSource[] = [
      { mailAddresses: ["cinema@example.com"], parseBody: parser },
    ];
    const onParseError = vi.fn();

    expect(
      parseTicketBody("mail body", "unknown@example.com", sources, onParseError),
    ).toBeUndefined();
    expect(parser).not.toHaveBeenCalled();
    expect(onParseError).toHaveBeenCalledWith({
      reason: "unknown_sender",
      fromAddress: "unknown@example.com",
    });
  });

  it("reports an unrecognized body when the sender's parser returns undefined", () => {
    const source: TicketMailSource = {
      mailAddresses: ["cinema@example.com"],
      parseBody: () => undefined,
    };
    const onParseError = vi.fn();

    expect(
      parseTicketBody("unrelated body", "cinema@example.com", [source], onParseError),
    ).toBeUndefined();
    expect(onParseError).toHaveBeenCalledWith({ reason: "unrecognized_body", source });
  });

  it("reports a thrown parse error", () => {
    const parseError = new Error("Missing required field: seats");
    const source: TicketMailSource = {
      mailAddresses: ["cinema@example.com"],
      parseBody: () => {
        throw parseError;
      },
    };
    const onParseError = vi.fn();

    expect(
      parseTicketBody("mail body", "cinema@example.com", [source], onParseError),
    ).toBeUndefined();
    expect(onParseError).toHaveBeenCalledWith({ reason: "error", source, error: parseError });
  });
});
