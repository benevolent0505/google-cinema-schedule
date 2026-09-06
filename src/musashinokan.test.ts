import { describe, expect, it } from "vitest";

import "./ticket";
import "./musashinokan";

type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

const parseMusashinokanBody = (
  globalThis as typeof globalThis & {
    parseMusashinokanBody: (body: string) => Ticket | undefined;
  }
).parseMusashinokanBody;

const sampleBody = [
  "①予約番号：1234567890",
  "　QRコード：https://example.com",
  "②2026/12/31 00:00",
  "③タイトル",
  "④ｽｸﾘｰﾝ２",
  "⑤券種 1枚",
  "　合計 2,000",
  "⑥座席番号：Ａ－１",
  "",
].join("\r\n");

describe("parseMusashinokanBody", () => {
  it("parses a ticket confirmation email body", () => {
    expect(parseMusashinokanBody(sampleBody)).toEqual({
      ticketNumber: "1234567890",
      title: "タイトル",
      startTime: new Date("2026-12-31T00:00:00+09:00"),
      endTime: new Date("2026-12-31T00:00:00+09:00"),
      theater: "ｽｸﾘｰﾝ２",
      sheet: "Ａ－１",
    });
  });

  it("returns undefined for an unrelated body", () => {
    expect(parseMusashinokanBody("just a normal email")).toBeUndefined();
  });

  it("returns undefined when a required field is missing", () => {
    const bodyWithoutSeats = sampleBody.replace("⑥座席番号：Ａ－１\r\n", "");

    expect(parseMusashinokanBody(bodyWithoutSeats)).toBeUndefined();
  });
});
