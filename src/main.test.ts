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

const { buildTicketMailSearchCriteria, debugMain, main } = globalThis as typeof globalThis & {
  buildTicketMailSearchCriteria: (
    sources: readonly TicketMailSource[],
    newerThreshold: Date,
  ) => string;
  debugMain: (searchStartDateTime?: string) => void;
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
    vi.stubGlobal("Logger", { log: vi.fn() });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    expect(() => main()).not.toThrow();
    expect(search).toHaveBeenCalledOnce();
    expect(search).toHaveBeenCalledWith(expect.stringContaining("from:ticket@cinemacity.co.jp"));
    expect(getDefaultCalendar).not.toHaveBeenCalled();
  });

  it("logs Skip when the ticket is already registered", () => {
    const getDate = vi.fn(() => new Date());
    const getPlainBody = vi.fn(() => sampleBody);
    const getMessages = vi.fn(() => [
      {
        getDate,
        getPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getTitle = vi.fn(() => "君の名は。");
    const getEvents = vi.fn(() => [{ getTitle }]);
    const createEvent = vi.fn();
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("Logger", { log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(log).toHaveBeenCalledWith("Skip: 君の名は。");
    expect(createEvent).not.toHaveBeenCalled();
  });

  it("emits fetchTickets debug logs when DEBUG_LOG_ENABLED is on", () => {
    const search = vi.fn(() => []);
    const getDefaultCalendar = vi.fn();
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("Logger", { log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({
        getProperty: (key: string) => (key === "DEBUG_LOG_ENABLED" ? "true" : null),
      }),
    });

    main();

    expect(log).toHaveBeenCalledWith(expect.stringContaining("fetchTickets: 検索条件"));
  });

  it("suppresses fetchTickets debug logs when DEBUG_LOG_ENABLED is off", () => {
    const search = vi.fn(() => []);
    const getDefaultCalendar = vi.fn();
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("Logger", { log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(log).not.toHaveBeenCalled();
  });
});

describe("debugMain", () => {
  it("uses the specified date and time as the mail search threshold", () => {
    const beforeThresholdGetPlainBody = vi.fn(() => sampleBody);
    const afterThresholdGetPlainBody = vi.fn(() => "not a ticket");
    const getMessages = vi.fn(() => [
      {
        getDate: () => new Date("2025-03-01T09:59:59+09:00"),
        getPlainBody: beforeThresholdGetPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
      },
      {
        getDate: () => new Date("2025-03-01T10:00:00+09:00"),
        getPlainBody: afterThresholdGetPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getDefaultCalendar = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("Logger", { log: vi.fn() });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    debugMain("2025-03-01T10:00:00+09:00");

    expect(search).toHaveBeenCalledWith(expect.stringContaining("newer:2025-03-01"));
    expect(beforeThresholdGetPlainBody).not.toHaveBeenCalled();
    expect(afterThresholdGetPlainBody).toHaveBeenCalledOnce();
    expect(getDefaultCalendar).not.toHaveBeenCalled();
  });

  it("reads the search threshold from Script Properties when no argument is given", () => {
    const search = vi.fn(() => []);
    const getProperty = vi.fn(() => "2025-03-01T10:00:00+09:00");
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    debugMain();

    expect(getProperty).toHaveBeenCalledWith("DEBUG_MAIL_SEARCH_START_DATETIME");
    expect(search).toHaveBeenCalledWith(expect.stringContaining("newer:2025-03-01"));
  });

  it("rejects an invalid search threshold", () => {
    expect(() => debugMain("invalid")).toThrow(
      "DEBUG_MAIL_SEARCH_START_DATETIME には有効な日時を指定してください: invalid",
    );
  });

  it("requires a search threshold when no argument or Script Property is set", () => {
    const getProperty = vi.fn(() => null);
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    expect(() => debugMain()).toThrow(
      "DEBUG_MAIL_SEARCH_START_DATETIME に検索開始日時を指定してください。",
    );
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
