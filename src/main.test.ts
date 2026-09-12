import { afterEach, describe, expect, it, vi } from "vitest";

import {
  buildTicketMailSearchCriteria,
  debugMain,
  dedupeTicketsByTicketNumber,
  main,
  parseExecutionDate,
  resolveMailSearchRange,
} from "./main";
import type { FetchedTicket } from "./main";
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

  it("チケットが 1 件もない場合はカレンダーへアクセスしない", () => {
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

  it("同じチケット番号の予定が既にある場合は Skip をログに出す", () => {
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

  it("同じタイトルの予定があってもチケット番号が違えば登録する（同じ作品を別日に観た場合に誤ってスキップしない）", () => {
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

  it("予定の場所に映画館の正式名称を設定し、説明欄にスクリーンと元メールを記載する", () => {
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
    const getEvents = vi.fn(() => []);
    const createEvent = vi.fn(() => ({ getTitle: () => "君の名は。" }));
    const getDefaultCalendar = vi.fn(() => ({ createEvent, getEvents }));
    vi.stubGlobal("GmailApp", { search });
    vi.stubGlobal("CalendarApp", { getDefaultCalendar });
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    main();

    expect(createEvent).toHaveBeenCalledWith(
      "君の名は。",
      new Date("2025-03-01T10:00:00+09:00"),
      new Date("2025-03-01T12:30:00+09:00"),
      {
        description:
          "スクリーン: シネマ・ツー/１階/a studio\n座席: A-10\n元メール: https://mail.google.com/mail/u/0/#all/message-1\nチケット番号: 12345\n検索用キーワード: 映画館チケット",
        location: "シネマシティ シネマ・ツー",
      },
    );
  });

  it("送信元の形式に一致したのに解析に失敗した場合、エラーを出しつつ他のチケットの登録は続ける", () => {
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

  it("元メールをたどれるよう Gmail のメッセージリンクをログに出す", () => {
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

  it("DEBUG_LOG_ENABLED が有効でもメール本文はログに出さない", () => {
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

  it("DEBUG_LOG_ENABLED が有効なとき fetchTickets のデバッグログを出す", () => {
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

  it("DEBUG_LOG_ENABLED が無効なとき fetchTickets のデバッグログを抑制する", () => {
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

  it("指定した実行日の対象範囲に受信したメールだけを対象にし、前後の境界のメールは除外する", () => {
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

  it("指定が効いているか確認できるよう、DEBUG_LOG_ENABLED が無効でも仮想実行日をログに出す", () => {
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

  it("引数がない場合は Script Properties から実行日を読む", () => {
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

  it("引数を Script Property より優先する", () => {
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

  it("YYYY-MM-DD 形式でない実行日を拒否する", () => {
    vi.stubGlobal("PropertiesService", {
      getScriptProperties: () => ({ getProperty: () => null }),
    });

    expect(() => debugMain("2025-03-01T10:00:00+09:00")).toThrow(
      "DEBUG_EXECUTION_DATE には YYYY-MM-DD 形式で日付を指定してください: 2025-03-01T10:00:00+09:00",
    );
  });

  it("引数も Script Property もない場合は実行日を要求する", () => {
    const getProperty = vi.fn(() => null);
    const getScriptProperties = vi.fn(() => ({ getProperty }));
    vi.stubGlobal("PropertiesService", { getScriptProperties });

    expect(() => debugMain()).toThrow(
      "DEBUG_EXECUTION_DATE に実行日を YYYY-MM-DD 形式で指定してください。",
    );
  });
});

describe("parseExecutionDate", () => {
  it("実行環境のタイムゾーンにおけるその日の 0 時として解釈する", () => {
    expect(parseExecutionDate("2025-03-02")).toEqual(new Date("2025-03-02T00:00:00+09:00"));
  });

  it("形式違いの日付を拒否する", () => {
    expect(() => parseExecutionDate("2025-3-2")).toThrow(
      "DEBUG_EXECUTION_DATE には YYYY-MM-DD 形式で日付を指定してください: 2025-3-2",
    );
  });

  it("実在しない日付を拒否し、Date の繰り上げで黙って別の日として実行しない", () => {
    expect(() => parseExecutionDate("2025-02-30")).toThrow(
      "DEBUG_EXECUTION_DATE には実在する日付を指定してください: 2025-02-30",
    );
  });
});

describe("resolveMailSearchRange", () => {
  it("前日の 0 時から翌日の 0 時までを対象にする", () => {
    expect(resolveMailSearchRange(new Date("2025-03-02T13:45:00+09:00"))).toEqual({
      start: new Date("2025-03-01T00:00:00+09:00"),
      end: new Date("2025-03-03T00:00:00+09:00"),
    });
  });

  it("月をまたぐ実行日を扱う", () => {
    expect(resolveMailSearchRange(new Date("2025-03-01T00:00:00+09:00"))).toEqual({
      start: new Date("2025-02-28T00:00:00+09:00"),
      end: new Date("2025-03-02T00:00:00+09:00"),
    });
  });

  it("年をまたぐ実行日を扱う", () => {
    expect(resolveMailSearchRange(new Date("2025-01-01T00:00:00+09:00"))).toEqual({
      start: new Date("2024-12-31T00:00:00+09:00"),
      end: new Date("2025-01-02T00:00:00+09:00"),
    });
  });

  it("月末の実行日を扱う", () => {
    expect(resolveMailSearchRange(new Date("2025-01-31T00:00:00+09:00"))).toEqual({
      start: new Date("2025-01-30T00:00:00+09:00"),
      end: new Date("2025-02-01T00:00:00+09:00"),
    });
  });
});

describe("buildTicketMailSearchCriteria", () => {
  it("すべての source の送信元アドレスを重複なく結合する", () => {
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
    screen: "スクリーン1",
    sheet: "A-1",
    ...overrides,
  });

  const buildFetchedTicket = (
    ticketOverrides: Partial<Ticket>,
    receivedAt: string,
    messageLink: string,
  ): FetchedTicket => ({
    ticket: buildTicket(ticketOverrides),
    mail: {
      receivedAt: new Date(receivedAt),
      messageLink,
    },
  });

  it("同じチケット番号では最も新しいメールを残し、再送後の情報とリンクを採用する", () => {
    const older = buildFetchedTicket(
      { ticketNumber: "12345", sheet: "A-1" },
      "2025-03-01T09:00:00+09:00",
      "https://mail.test/older",
    );
    const newer = buildFetchedTicket(
      { ticketNumber: "12345", sheet: "A-2" },
      "2025-03-01T10:00:00+09:00",
      "https://mail.test/newer",
    );
    const other = buildFetchedTicket(
      { ticketNumber: "67890", sheet: "B-2" },
      "2025-03-01T08:00:00+09:00",
      "https://mail.test/other",
    );

    expect(dedupeTicketsByTicketNumber([older, newer, other])).toEqual([newer, other]);
  });

  it("メールの走査順にかかわらず同じチケット番号の最新メールを残す", () => {
    const newer = buildFetchedTicket(
      { ticketNumber: "12345", sheet: "A-2" },
      "2025-03-01T10:00:00+09:00",
      "https://mail.test/newer",
    );
    const older = buildFetchedTicket(
      { ticketNumber: "12345", sheet: "A-1" },
      "2025-03-01T09:00:00+09:00",
      "https://mail.test/older",
    );

    expect(dedupeTicketsByTicketNumber([newer, older])).toEqual([newer]);
  });

  it("空配列はそのまま返す", () => {
    expect(dedupeTicketsByTicketNumber([])).toEqual([]);
  });
});
