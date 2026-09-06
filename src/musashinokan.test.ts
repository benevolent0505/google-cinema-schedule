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

  it("throws when a required field's value is missing", () => {
    // canParse はラベルの有無しか見ないため、値だけを欠落させて
    // canParse を通過させつつ抽出を失敗させる。canParse を満たしたのに
    // 抽出に失敗した場合、undefined ではなく例外を投げる。呼び出し側
    // (ticket.ts) がこれを「送信元の形式には一致したが解析に失敗した」
    // ケースとして検知し、ログに残せるようにするため。
    const bodyWithoutSeatValue = sampleBody.replace("⑥座席番号：Ｅ－２", "⑥座席番号：");

    expect(() => parseMusashinokanBody(bodyWithoutSeatValue)).toThrow(
      "Missing value for field: ⑥座席番号：",
    );
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
