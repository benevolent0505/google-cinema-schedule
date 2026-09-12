import { describe, expect, it } from "vitest";

import { parseCinemaCityBody } from "./cinemacity";

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
  it("予約確認メールの本文をチケットとして解析する", () => {
    expect(parseCinemaCityBody(sampleBody)).toEqual({
      ticketNumber: "12345",
      title: "君の名は。",
      startTime: new Date("2025-03-01T10:00:00+09:00"),
      endTime: new Date("2025-03-01T12:30:00+09:00"),
      theater: "シネマシティ シネマ・ツー",
      screen: "シネマ・ツー/１階/a studio",
      sheet: "A-10",
    });
  });

  it("シネマ・ワンの劇場欄から建物を区別した正式名称を設定する", () => {
    const cinemaOneBody = sampleBody.replace(
      "シネマ・ツー/１階/a studio",
      "シネマ・ワン/２階/cinema one",
    );

    expect(parseCinemaCityBody(cinemaOneBody)).toMatchObject({
      theater: "シネマシティ シネマ・ワン",
      screen: "シネマ・ワン/２階/cinema one",
    });
  });

  it("劇場欄からシネマ・ワン／シネマ・ツーを判定できない場合は例外を投げる", () => {
    const unknownTheaterBody = `${sampleBody.replace(
      "シネマ・ツー/１階/a studio",
      "不明な劇場/a studio",
    )}\nシネマ・ツー`;

    expect(() => parseCinemaCityBody(unknownTheaterBody)).toThrow(
      "Unknown Cinema City theater: 不明な劇場/a studio",
    );
  });

  it("無関係な本文には undefined を返す", () => {
    expect(parseCinemaCityBody("just a normal email")).toBeUndefined();
  });

  it("形式は一致するのに必須項目が欠けている場合、解析失敗として検知できるよう例外を投げる", () => {
    const bodyWithoutSeats = sampleBody.replace("■座席\r\n[ A-10 ]\r\n", "");

    expect(() => parseCinemaCityBody(bodyWithoutSeats)).toThrow("Missing required field: ■座席");
  });
});
