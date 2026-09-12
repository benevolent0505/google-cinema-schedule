import { describe, expect, it } from "vitest";

import { parseMusashinokanBody } from "./musashinokan";

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
  it("予約確認メールの本文をチケットとして解析する", () => {
    expect(parseMusashinokanBody(sampleBody)).toEqual({
      ticketNumber: "1234567",
      title: "タイトル",
      startTime: new Date("2026-01-02T03:45:00+09:00"),
      endTime: new Date("2026-01-02T05:45:00+09:00"),
      theater: "新宿武蔵野館",
      screen: "ｽｸﾘｰﾝ１",
      sheet: "Ｅ－２",
    });
  });

  it("無関係な本文には undefined を返す", () => {
    expect(parseMusashinokanBody("just a normal email")).toBeUndefined();
  });

  it("形式は一致するのに必須項目の値が空の場合、解析失敗として検知できるよう例外を投げる", () => {
    const bodyWithoutSeatValue = sampleBody.replace("⑥座席番号：Ｅ－２", "⑥座席番号：");

    expect(() => parseMusashinokanBody(bodyWithoutSeatValue)).toThrow(
      "Missing value for field: ⑥座席番号：",
    );
  });

  it("定型文に囲まれた実際の予約メールを解析する", () => {
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
      theater: "新宿武蔵野館",
      screen: "ｽｸﾘｰﾝ１",
      sheet: "Ｅ－２",
    });
  });
});
