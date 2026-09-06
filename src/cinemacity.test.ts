import { describe, expect, it } from "vitest";

import "./ticket";
import "./cinemacity";

type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

const parseCinemaCityBody = (
  globalThis as typeof globalThis & {
    parseCinemaCityBody: (body: string) => Ticket | undefined;
  }
).parseCinemaCityBody;

const sampleBody = [
  "■チケット番号：12345",
  "■登録電話番号：09012345678（下4ケタのみでOK）",
  "",
  "君の名は。",
  "■上映時間",
  "2025年3月1日(土) 10:00 - 12:30",
  "■劇場",
  "シネマ・ツー/１階/a studio",
  "（ワン：高島屋右隣／ツー：モノレール下遊歩道沿）",
  "■座席",
  "[ A-10 ]",
  "■枚数",
  "1枚",
  "■合計金額",
  "1,000円(手数料込)",
  "",
].join("\r\n");

describe("parseCinemaCityBody", () => {
  it("parses a ticket confirmation email body", () => {
    expect(parseCinemaCityBody(sampleBody)).toEqual({
      ticketNumber: "12345",
      title: "君の名は。",
      startTime: new Date("2025-03-01T10:00:00+09:00"),
      endTime: new Date("2025-03-01T12:30:00+09:00"),
      theater: "シネマ・ツー/１階/a studio",
      sheet: "A-10",
    });
  });

  it("returns undefined for an unrelated body", () => {
    expect(parseCinemaCityBody("just a normal email")).toBeUndefined();
  });

  it("throws when a required field is missing", () => {
    // canParse は満たすが必須項目が欠けている場合、undefined ではなく例外を
    // 投げる。呼び出し側 (ticket.ts) がこれを「送信元の形式には一致したが解析に
    // 失敗した」ケースとして検知し、ログに残せるようにするため。
    const bodyWithoutSeats = sampleBody.replace("■座席\r\n[ A-10 ]\r\n", "");

    expect(() => parseCinemaCityBody(bodyWithoutSeats)).toThrow("Missing required field: ■座席");
  });
});
