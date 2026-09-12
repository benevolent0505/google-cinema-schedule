import { createTicketParser, pad2, parsePrice, parseSeats, toLines } from "./ticket-parser";
import type { Reservation, Screening } from "./ticket-parser";

const KINOCINEMA_THEATER_NAME = "kino cinema新宿";

// kino cinéma 新宿の購入完了メールは上映開始時刻しか載せないため、開始時刻に既定の
// 上映時間を足して終了時刻を補完する。
const KINOCINEMA_DEFAULT_SCREENING_MINUTES = 120;

// 上映開始行「2026/01/23(月) 12:34~」。曜日と末尾のチルダは任意で受ける。
const KINOCINEMA_SCREENING_PATTERN =
  /(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\([^)]*\)|（[^）]*）)?\s*(\d{1,2}):(\d{2})/;

export const parseKinoCinemaBody = createTicketParser({
  canParse: kinoCinemaCanParse,
  parseReservation: kinoCinemaParseReservation,
});

function kinoCinemaCanParse(raw: string): boolean {
  return (
    raw.includes("購入番号") &&
    raw.includes("スクリーン") &&
    raw.includes("座席番号") &&
    raw.includes("合計金額")
  );
}

function kinoCinemaParseReservation(raw: string): Reservation {
  const lines = toLines(raw);
  const ticketNumber = kinoCinemaExtractTicketNumber(lines);
  const title = kinoCinemaExtractTitle(lines);
  const screening = kinoCinemaParseScreening(lines);
  const { screen, seats } = kinoCinemaExtractScreenAndSeats(lines);
  const totalPrice = kinoCinemaExtractTotalPrice(lines);

  return {
    theater: {
      name: KINOCINEMA_THEATER_NAME,
      screen,
    },
    movie: {
      title,
    },
    screening,
    ticketNumber,
    seats,
    totalPrice,
  };
}

function kinoCinemaFindLineIndex(lines: string[], keyword: string): number {
  const index = lines.findIndex((line) => line.includes(keyword));
  if (index === -1) {
    throw new Error(`Missing required field: ${keyword}`);
  }
  return index;
}

function kinoCinemaExtractTicketNumber(lines: string[]): string {
  const line = lines[kinoCinemaFindLineIndex(lines, "購入番号")];
  const match = line.match(/購入番号\s*[:：]\s*(\S+)/);
  if (!match) {
    throw new Error("Missing value for field: 購入番号");
  }
  return match[1];
}

// タイトルはラベルを持たず、購入番号の行の次に置かれる。上映フォーマットや上映開始行
// より前にある最初の非空行がタイトルなので、位置で拾う。
function kinoCinemaExtractTitle(lines: string[]): string {
  const purchaseIndex = kinoCinemaFindLineIndex(lines, "購入番号");
  const title = lines
    .slice(purchaseIndex + 1)
    .map((line) => line.trim())
    .find((line) => line !== "");
  if (title === undefined) {
    throw new Error("Missing required field: title");
  }
  return title;
}

function kinoCinemaParseScreening(lines: string[]): Screening {
  const line = lines.find((candidate) => KINOCINEMA_SCREENING_PATTERN.test(candidate));
  if (line === undefined) {
    throw new Error("Missing required field: screening time");
  }

  const match = line.match(KINOCINEMA_SCREENING_PATTERN);
  if (!match) {
    throw new Error(`Invalid screening time: ${line}`);
  }

  const [, year, month, day, hour, minute] = match;
  const start = new Date(`${year}-${pad2(month)}-${pad2(day)}T${pad2(hour)}:${minute}:00+09:00`);
  if (Number.isNaN(start.getTime())) {
    throw new Error(`Invalid screening time: ${line}`);
  }

  const end = new Date(start.getTime() + KINOCINEMA_DEFAULT_SCREENING_MINUTES * 60 * 1000);

  return { start, end };
}

// スクリーン番号と座席番号は同じ行に並ぶため、1 行から両方を取り出す。
function kinoCinemaExtractScreenAndSeats(lines: string[]): {
  screen: string;
  seats: string[];
} {
  const line = lines.find(
    (candidate) => candidate.includes("スクリーン") && candidate.includes("座席番号"),
  );
  if (line === undefined) {
    throw new Error("Missing required field: スクリーン");
  }

  const screenMatch = line.match(/スクリーン\s*[:：]\s*(.+?)\s+座席番号/);
  if (!screenMatch || screenMatch[1].trim() === "") {
    throw new Error("Missing value for field: スクリーン");
  }

  const seatsMatch = line.match(/座席番号\s*[:：]\s*(.+?)\s*$/);
  if (!seatsMatch) {
    throw new Error("Missing value for field: 座席番号");
  }

  return {
    screen: screenMatch[1].trim(),
    seats: parseSeats(seatsMatch[1], /[、,／/\s]+/),
  };
}

function kinoCinemaExtractTotalPrice(lines: string[]): number {
  const line = lines.find((candidate) => candidate.includes("合計金額"));
  if (line === undefined) {
    throw new Error("Missing required field: 合計金額");
  }

  const match = line.match(/合計金額\s*([\d,]+)\s*円/);
  if (!match) {
    throw new Error("Missing value for field: 合計金額");
  }

  return parsePrice(match[1]);
}
