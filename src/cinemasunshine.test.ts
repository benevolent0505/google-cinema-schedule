import { describe, expect, it } from "vitest";

import { parseCinemaSunshineBody } from "./cinemasunshine";

const sampleBody = [
  "[予約番号]",
  "1234567890",
  "",
  "[鑑賞日時]",
  "2026年12月31日(木) 00:00 - 23:59",
  "",
  "[作品名]",
  "タイトル",
  "",
  "[スクリーン名]",
  "シアター1",
  "",
  "[座席]",
  "ａ－１ 券種 ￥2,000",
  "",
  "[合計]",
  "￥2,000",
  "",
].join("\r\n");

describe("parseCinemaSunshineBody", () => {
  it("予約確認メールの本文をチケットとして解析する", () => {
    expect(parseCinemaSunshineBody(sampleBody)).toEqual({
      ticketNumber: "1234567890",
      title: "タイトル",
      startTime: new Date("2026-12-31T00:00:00+09:00"),
      endTime: new Date("2026-12-31T23:59:00+09:00"),
      theater: "シアター1",
      sheet: "ａ－１",
    });
  });

  it("複数座席の予約では券種や金額を除いた座席番号だけをまとめる", () => {
    const bodyWithMultipleSeats = sampleBody.replace(
      "ａ－１ 券種 ￥2,000",
      ["ａ－１ 一般 ￥2,000", "ａ－２ 学生 ￥1,500"].join("\r\n"),
    );

    expect(parseCinemaSunshineBody(bodyWithMultipleSeats)?.sheet).toBe("ａ－１, ａ－２");
  });

  it("無関係な本文には undefined を返す", () => {
    expect(parseCinemaSunshineBody("just a normal email")).toBeUndefined();
  });

  it("形式は一致するのに座席の値が欠けている場合、解析失敗として検知できるよう例外を投げる", () => {
    const bodyWithoutSeatValue = sampleBody.replace("ａ－１ 券種 ￥2,000", "");

    expect(() => parseCinemaSunshineBody(bodyWithoutSeatValue)).toThrow(
      "Missing value for field: [座席]",
    );
  });

  it("定型文に囲まれた実際の予約メールを解析する", () => {
    const fullBody = [
      "テスト太郎 様",
      "",
      "この度はグランドシネマサンシャイン池袋をご利用いただき誠にありがとうございます。",
      "ご予約の内容は下記の通りです。",
      "",
      "[予約番号]",
      "1234567890",
      "",
      "[鑑賞日時]",
      "2026年12月31日(木) 00:00 - 23:59",
      "",
      "[作品名]",
      "タイトル",
      "",
      "[スクリーン名]",
      "シアター1",
      "",
      "[座席]",
      "ａ－１ 券種 ￥2,000",
      "",
      "[合計]",
      "￥2,000",
      "",
      "※ご購入されたチケットの変更、キャンセル、払戻しはできません。",
      "",
    ].join("\r\n");

    expect(parseCinemaSunshineBody(fullBody)).toEqual({
      ticketNumber: "1234567890",
      title: "タイトル",
      startTime: new Date("2026-12-31T00:00:00+09:00"),
      endTime: new Date("2026-12-31T23:59:00+09:00"),
      theater: "シアター1",
      sheet: "ａ－１",
    });
  });
});
