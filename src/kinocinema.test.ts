import { describe, expect, it } from "vitest";

import { parseKinoCinemaBody } from "./kinocinema";

const sampleBody = [
  "kino cinéma 新宿",
  "テスト太郎 様 ",
  "ご予約いただきありがとうございます。 ",
  "下記流れに沿ってQRコードを提示してください。 ",
  "ご購入内容 購入番号 : 1234567",
  "タイトル",
  "IMAX",
  "2026/01/23(月) 12:34~",
  "スクリーン : 1    座席番号 : Ｅ－２",
  "券種 : 一般 合計金額2,000円",
  "レシートはこちら ",
  "<https://receipt-url.test>",
  "でご確認ください ",
].join("\r\n");

describe("parseKinoCinemaBody", () => {
  it("購入完了メールの本文をチケットとして解析する", () => {
    expect(parseKinoCinemaBody(sampleBody)).toEqual({
      ticketNumber: "1234567",
      title: "タイトル",
      startTime: new Date("2026-01-23T12:34:00+09:00"),
      endTime: new Date("2026-01-23T14:34:00+09:00"),
      theater: "kino cinema新宿",
      screen: "1",
      sheet: "Ｅ－２",
    });
  });

  it("同じ行に並ぶ複数座席をまとめる", () => {
    const bodyWithMultipleSeats = sampleBody.replace(
      "座席番号 : Ｅ－２",
      "座席番号 : Ｅ－２,Ｅ－３",
    );

    expect(parseKinoCinemaBody(bodyWithMultipleSeats)?.sheet).toBe("Ｅ－２, Ｅ－３");
  });

  it("無関係な本文には undefined を返す", () => {
    expect(parseKinoCinemaBody("just a normal email")).toBeUndefined();
  });

  it("形式は一致するのに座席の値が欠けている場合、解析失敗として検知できるよう例外を投げる", () => {
    const bodyWithoutSeatValue = sampleBody.replace(
      "スクリーン : 1    座席番号 : Ｅ－２",
      "スクリーン : 1    座席番号 :",
    );

    expect(() => parseKinoCinemaBody(bodyWithoutSeatValue)).toThrow(
      "Missing value for field: 座席番号",
    );
  });

  it("形式は一致するのに合計金額の値が欠けている場合、解析失敗として検知できるよう例外を投げる", () => {
    const bodyWithoutPrice = sampleBody.replace("合計金額2,000円", "合計金額円");

    expect(() => parseKinoCinemaBody(bodyWithoutPrice)).toThrow(
      "Missing value for field: 合計金額",
    );
  });
});
