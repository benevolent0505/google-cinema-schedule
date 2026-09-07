import { afterEach, describe, expect, it, vi } from "vitest";

import {
  buildTicketMailSearchCriteria,
  debugMain,
  dedupeTicketsByTicketNumber,
  main,
  parseExecutionDate,
  resolveMailSearchRange,
} from "./main";
import type { Ticket, TicketMailSource } from "./ticket";

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

describe("main", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("does not access Calendar when no tickets are found", () => {
    const search = vi.fn(() => []);
    const getDefaultCalendar = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    expect(() => main()).not.toThrow();
    expect(search).toHaveBeenCalledOnce();
    expect(search).toHaveBeenCalledWith(expect.stringContaining("from:ticket@cinemacity.co.jp"));
    expect(getDefaultCalendar).not.toHaveBeenCalled();
  });

  it("logs Skip when an event with the same ticket number already exists", () => {
    const getDate = vi.fn(() => new Date());
    const getPlainBody = vi.fn(() => sampleBody);
    const getMessages = vi.fn(() => [
      {
        getDate,
        getPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getTitle = vi.fn(() => "君の名は。");
    const getDescription = vi.fn(
      () => "劇場: 劇場\n座席: A-10\nチケット番号: 12345\n検索用キーワード: 映画館チケット",
    );
    const getEvents = vi.fn(() => [{ getTitle, getDescription }]);
    const createEvent = vi.fn();
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    const info = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, info });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(info).toHaveBeenCalledWith("Skip: 君の名は。");
    expect(createEvent).not.toHaveBeenCalled();
  });

  it("registers a ticket even when an existing event has the same title but a different ticket number", () => {
    // 同じ作品を別日にもう一度観た場合、タイトルだけで既存判定すると
    // 誤って登録がスキップされてしまう。チケット番号が異なれば登録される
    // ことを確認する回帰テスト。
    const getDate = vi.fn(() => new Date());
    const getPlainBody = vi.fn(() => sampleBody);
    const getMessages = vi.fn(() => [
      {
        getDate,
        getPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getTitle = vi.fn(() => "君の名は。");
    const getDescription = vi.fn(
      () => "劇場: 劇場\n座席: A-1\nチケット番号: 99999\n検索用キーワード: 映画館チケット",
    );
    const getEvents = vi.fn(() => [{ getTitle, getDescription }]);
    const createEvent = vi.fn(() => ({ getTitle: () => "君の名は。" }));
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    const info = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, info });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(info).not.toHaveBeenCalledWith("Skip: 君の名は。");
    expect(createEvent).toHaveBeenCalledOnce();
  });

  it("logs an error and still registers other tickets when a mail matches a sender's format but fails to parse", () => {
    // canParse は満たすが必須項目 (■座席) が欠けているメール。デバッグログの
    // 有効・無効に関係なくエラーが出て、他の正常なチケットの登録は継続される
    // ことを確認する。
    const brokenBody = sampleBody.replace("■座席\r\n[ A-10 ]\r\n", "");
    const getMessages = vi.fn(() => [
      {
        getDate: () => new Date(),
        getPlainBody: () => brokenBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getEvents = vi.fn(() => []);
    const createEvent = vi.fn();
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    const error = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, error });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(error).toHaveBeenCalledWith(
      expect.stringContaining("送信元=ticket@cinemacity.co.jp の解析に失敗しました"),
      expect.anything(),
    );
    expect(createEvent).not.toHaveBeenCalled();
  });

  it("logs the Gmail message link so the original mail can be found", () => {
    const getMessages = vi.fn(() => [
      {
        getDate: () => new Date(),
        getPlainBody: () => sampleBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getDefaultCalendar = vi.fn(() => ({
      createEvent: vi.fn(() => ({ getTitle: () => "君の名は。" })),
      getEvents: vi.fn(() => []),
    }));
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({
        getProperty: (key: string) => (key === "DEBUG_LOG_ENABLED" ? "true" : null),
      }),
    });

    main();

    expect(log).toHaveBeenCalledWith(
      expect.stringContaining("link=https://mail.google.com/mail/u/0/#all/message-1"),
    );
  });

  it("does not dump the mail body even when DEBUG_LOG_ENABLED is on", () => {
    const getMessages = vi.fn(() => [
      {
        getDate: () => new Date(),
        getPlainBody: () => sampleBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getDefaultCalendar = vi.fn(() => ({
      createEvent: vi.fn(() => ({ getTitle: () => "君の名は。" })),
      getEvents: vi.fn(() => []),
    }));
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({
        getProperty: (key: string) => (key === "DEBUG_LOG_ENABLED" ? "true" : null),
      }),
    });

    main();

    for (const call of log.mock.calls) {
      expect(String(call[0])).not.toContain("09012345678");
    }
  });

  it("emits fetchTickets debug logs when DEBUG_LOG_ENABLED is on", () => {
    const search = vi.fn(() => []);
    const getDefaultCalendar = vi.fn();
    const log = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("console", { ...console, log });
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
    vi.stubGlobal("console", { ...console, log });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(log).not.toHaveBeenCalled();
  });
});

describe("debugMain", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("targets only mails received within the specified execution date's range", () => {
    // 仮想実行日 2025-03-02 の対象範囲は 2025-03-01 00:00 以上 2025-03-03 00:00 未満。
    // 前後の境界にあるメールが除外されることを確認する。
    const beforeRangeGetPlainBody = vi.fn(() => "not a ticket");
    const inRangeGetPlainBody = vi.fn(() => sampleBody);
    const afterRangeGetPlainBody = vi.fn(() => "not a ticket");
    const getMessages = vi.fn(() => [
      {
        getDate: () => new Date("2025-02-28T23:59:59+09:00"),
        getPlainBody: beforeRangeGetPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-1",
      },
      {
        getDate: () => new Date("2025-03-01T00:00:00+09:00"),
        getPlainBody: inRangeGetPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-2",
      },
      {
        getDate: () => new Date("2025-03-03T00:00:00+09:00"),
        getPlainBody: afterRangeGetPlainBody,
        getFrom: () => "ticket@cinemacity.co.jp",
        getSubject: () => "予約確認",
        getId: () => "message-3",
      },
    ]);
    const search = vi.fn(() => [{ getMessages, getFirstMessageSubject: () => "予約確認" }]);
    const getDefaultCalendar = vi.fn(() => ({
      createEvent: vi.fn(() => ({ getTitle: () => "君の名は。" })),
      getEvents: vi.fn(() => []),
    }));
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    debugMain("2025-03-02");

    expect(search).toHaveBeenCalledWith(expect.stringContaining("newer:2025-03-01"));
    expect(beforeRangeGetPlainBody).not.toHaveBeenCalled();
    expect(inRangeGetPlainBody).toHaveBeenCalledOnce();
    expect(afterRangeGetPlainBody).not.toHaveBeenCalled();
  });

  it("logs the virtual execution date even when DEBUG_LOG_ENABLED is off", () => {
    // 指定した実行日が効いているかは実行ログでしか確認できないため、
    // デバッグログが無効でも出力されることを確認する。
    const search = vi.fn(() => []);
    const info = vi.fn();
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("console", { ...console, info });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    debugMain("2025-03-02");

    expect(info).toHaveBeenCalledWith(
      "debugMain: 仮想実行日=2025-03-02 対象範囲=2025-03-01 00:00 以上 2025-03-03 00:00 未満",
    );
  });

  it("reads the execution date from Script Properties when no argument is given", () => {
    const search = vi.fn(() => []);
    const getProperty = vi.fn((key: string) =>
      key === "DEBUG_EXECUTION_DATE" ? "2025-03-02" : null,
    );
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    debugMain();

    expect(getProperty).toHaveBeenCalledWith("DEBUG_EXECUTION_DATE");
    expect(search).toHaveBeenCalledWith(expect.stringContaining("newer:2025-03-01"));
  });

  it("prefers the argument over the Script Property", () => {
    const search = vi.fn(() => []);
    const getProperty = vi.fn((key: string) =>
      key === "DEBUG_EXECUTION_DATE" ? "2025-03-02" : null,
    );
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    debugMain("2025-04-10");

    expect(search).toHaveBeenCalledWith(expect.stringContaining("newer:2025-04-09"));
  });

  it("rejects an execution date that is not in YYYY-MM-DD format", () => {
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    expect(() => debugMain("2025-03-01T10:00:00+09:00")).toThrow(
      "DEBUG_EXECUTION_DATE には YYYY-MM-DD 形式で日付を指定してください: 2025-03-01T10:00:00+09:00",
    );
  });

  it("requires an execution date when no argument or Script Property is set", () => {
    const getProperty = vi.fn(() => null);
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    expect(() => debugMain()).toThrow(
      "DEBUG_EXECUTION_DATE に実行日を YYYY-MM-DD 形式で指定してください。",
    );
  });
});

describe("parseExecutionDate", () => {
  it("parses a date as midnight in the runtime timezone", () => {
    expect(parseExecutionDate("2025-03-02")).toEqual(new Date("2025-03-02T00:00:00+09:00"));
  });

  it("rejects a malformed date", () => {
    expect(() => parseExecutionDate("2025-3-2")).toThrow(
      "DEBUG_EXECUTION_DATE には YYYY-MM-DD 形式で日付を指定してください: 2025-3-2",
    );
  });

  it("rejects a date that does not exist", () => {
    // Date は 2025-02-30 を 3月へ繰り上げてしまうため、黙って別の日として
    // 実行しないことを確認する。
    expect(() => parseExecutionDate("2025-02-30")).toThrow(
      "DEBUG_EXECUTION_DATE には実在する日付を指定してください: 2025-02-30",
    );
  });
});

describe("resolveMailSearchRange", () => {
  it("spans from the previous day's midnight to the next day's midnight", () => {
    expect(resolveMailSearchRange(new Date("2025-03-02T13:45:00+09:00"))).toEqual({
      start: new Date("2025-03-01T00:00:00+09:00"),
      end: new Date("2025-03-03T00:00:00+09:00"),
    });
  });

  it("crosses a month boundary", () => {
    expect(resolveMailSearchRange(new Date("2025-03-01T00:00:00+09:00"))).toEqual({
      start: new Date("2025-02-28T00:00:00+09:00"),
      end: new Date("2025-03-02T00:00:00+09:00"),
    });
  });

  it("crosses a year boundary", () => {
    expect(resolveMailSearchRange(new Date("2025-01-01T00:00:00+09:00"))).toEqual({
      start: new Date("2024-12-31T00:00:00+09:00"),
      end: new Date("2025-01-02T00:00:00+09:00"),
    });
  });

  it("handles the end of a month", () => {
    expect(resolveMailSearchRange(new Date("2025-01-31T00:00:00+09:00"))).toEqual({
      start: new Date("2025-01-30T00:00:00+09:00"),
      end: new Date("2025-02-01T00:00:00+09:00"),
    });
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

describe("dedupeTicketsByTicketNumber", () => {
  const buildTicket = (overrides: Partial<Ticket>): Ticket => ({
    ticketNumber: "1",
    title: "テスト作品",
    startTime: new Date("2025-03-01T10:00:00+09:00"),
    endTime: new Date("2025-03-01T12:00:00+09:00"),
    theater: "テスト劇場",
    sheet: "A-1",
    ...overrides,
  });

  it("keeps the first ticket and drops later ones with the same ticket number", () => {
    // 予約確認メールの再送などで同じチケット番号のチケットが複数件
    // 取得されても、同一実行内で1件に絞り込まれることを確認する。
    const first = buildTicket({ ticketNumber: "12345", sheet: "A-1" });
    const resend = buildTicket({ ticketNumber: "12345", sheet: "A-1" });
    const other = buildTicket({ ticketNumber: "67890", sheet: "B-2" });

    expect(dedupeTicketsByTicketNumber([first, resend, other])).toEqual([first, other]);
  });

  it("returns an empty array unchanged", () => {
    expect(dedupeTicketsByTicketNumber([])).toEqual([]);
  });
});
