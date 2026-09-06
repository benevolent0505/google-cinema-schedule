import { afterEach, describe, expect, it, vi } from "vitest";

import "./ticket";
import "./cinemacity";
import "./musashinokan";
import "./ticket-sources";
import "./main";

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

const sampleBody = [
  "■チケット番号：12345",
  "■登録電話番号：09012345678（下4ケタのみでOK）",
  "",
  "君の名は。",
  "■上映時間",
  "2025年3月1日(土) 10:00 - 12:30",
  "■劇場",
  "シネマ・ツー/１階/a studio",
  "■座席",
  "[ A-10 ]",
  "■枚数",
  "1枚",
  "■合計金額",
  "1,000円(手数料込)",
  "",
].join("\r\n");

const { buildTicketMailSearchCriteria, main } = globalThis as typeof globalThis & {
  buildTicketMailSearchCriteria: (
    sources: readonly TicketMailSource[],
    newerThreshold: Date,
  ) => string;
  main: () => void;
};

describe("main", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("does not access Calendar when no tickets are found", () => {
    const search = vi.fn(() => []);
    const getDefaultCalendar = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });

    expect(() => main()).not.toThrow();
    expect(search).toHaveBeenCalledOnce();
    expect(search).toHaveBeenCalledWith(expect.stringContaining("from:ticket@cinemacity.co.jp"));
    expect(getDefaultCalendar).not.toHaveBeenCalled();
  });

  it("logs Skip when the ticket is already registered", () => {
    const getPlainBody = vi.fn(() => sampleBody);
    const getMessages = vi.fn(() => [{ getPlainBody }]);
    const search = vi.fn(() => [{ getMessages }]);
    const getTitle = vi.fn(() => "君の名は。");
    const getEvents = vi.fn(() => [{ getTitle }]);
    const createEvent = vi.fn();
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("Logger", { log });

    main();

    expect(log).toHaveBeenCalledWith("Skip: 君の名は。");
    expect(createEvent).not.toHaveBeenCalled();
  });
});

describe("buildTicketMailSearchCriteria", () => {
  it("combines and deduplicates sender addresses from all sources", () => {
    const parseBody = () => undefined;
    const sources: TicketMailSource[] = [
      {
        mailAddresses: ["first@example.com", "shared@example.com"],
        parseBody,
      },
      {
        mailAddresses: ["second@example.com", "shared@example.com"],
        parseBody,
      },
    ];

    expect(buildTicketMailSearchCriteria(sources, new Date("2025-03-01T00:00:00.000Z"))).toBe(
      "(from:first@example.com OR from:shared@example.com OR from:second@example.com) AND newer:2025-03-01",
    );
  });
});
