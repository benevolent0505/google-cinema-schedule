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
  "①予約番号：1234567",
  "　QRコード：https://qrcode-url.test",
  "②2026/01/02 03:45",
  "③タイトル",
  "④ｽｸﾘｰﾝ１",
  "⑤券種",
  "　合計 0",
  "⑥座席番号：Ｅ－２",
  "",
].join("\r\n");

describe("parseMusashinokanBody", () => {
  it("parses a ticket confirmation email body", () => {
    expect(parseMusashinokanBody(sampleBody)).toEqual({
      ticketNumber: "1234567",
      title: "タイトル",
      startTime: new Date("2026-01-02T03:45:00+09:00"),
      endTime: new Date("2026-01-02T05:45:00+09:00"),
      theater: "ｽｸﾘｰﾝ１",
      sheet: "Ｅ－２",
    });
  });

  it("returns undefined for an unrelated body", () => {
    expect(parseMusashinokanBody("just a normal email")).toBeUndefined();
  });

  it("returns undefined when a required field is missing", () => {
    const bodyWithoutSeats = sampleBody.replace("⑥座席番号：Ｅ－２\r\n", "");

    expect(parseMusashinokanBody(bodyWithoutSeats)).toBeUndefined();
  });

  it("parses a full reservation email with surrounding boilerplate", () => {
    const fullBody = [
      "テスト太郎　様",
      "",
      "この度は、新宿武蔵野館のインターネットチケットをご利用いただき、誠にありがとうございます。",
      "",
      "お客様がご購入されましたチケットの情報は下記の通りです。",
      "発券をせずに下記のQRコードをご提示いただくことでそのままご入場いただけます。",
      "",
      "①予約番号：1234567",
      "　QRコード：https://qrcode-url.test",
      "②2026/01/02 03:45",
      "③タイトル",
      "④ｽｸﾘｰﾝ１",
      "⑤券種",
      "　合計 0",
      "⑥座席番号：Ｅ－２",
      "",
      "",
      "※ご購入されたチケットの変更、キャンセル、払戻しは一切いたしかねます。",
      "",
      "お問い合わせはこちら",
      "新宿武蔵野館",
      "TEL：03-3354-5670",
      "",
    ].join("\r\n");

    expect(parseMusashinokanBody(fullBody)).toEqual({
      ticketNumber: "1234567",
      title: "タイトル",
      startTime: new Date("2026-01-02T03:45:00+09:00"),
      endTime: new Date("2026-01-02T05:45:00+09:00"),
      theater: "ｽｸﾘｰﾝ１",
      sheet: "Ｅ－２",
    });
  });
});
