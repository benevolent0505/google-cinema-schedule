/**
 * Ticket parser for 新宿武蔵野館 reservation emails.
 *
 * The implementation is ported from the `theather-mail-parser` project
 * (`src/parsers/musashinokan-parser.ts`). Reservation emails are sent from
 * `reserve@musashino.cineticket.jp` with the subject `インターネットチケット購入`.
 */
import { createTicketParser, pad2, parsePrice, parseSeats, toLines } from "./ticket-parser";
import type { Reservation, Screening } from "./ticket-parser";

const MUSASHINOKAN_THEATER_NAME = "新宿武蔵野館";

/**
 * 新宿武蔵野館の予約メールには上映終了時刻が含まれないため、開始時刻から
 * 既定の上映時間を加算して終了時刻を補完する。
 */
const MUSASHINOKAN_DEFAULT_SCREENING_MINUTES = 120;

const MUSASHINOKAN_LABELS = {
  ticketNumber: "①予約番号：",
  screeningTime: "②",
  title: "③",
  theater: "④",
  totalPrice: "合計",
  seats: "⑥座席番号：",
} as const;

/**
 * Parse a Gmail message body into a `Ticket`, or return `undefined` when the
 * body is not a recognizable 新宿武蔵野館 reservation email.
 *
 * When the body matches this cinema's format (`musashinokanCanParse`) but a
 * required field is missing or malformed, this throws instead of returning
 * `undefined` so the caller can tell "not this format" apart from "this
 * format, but failed to parse" and report the latter.
 */
export const parseMusashinokanBody = createTicketParser({
  canParse: musashinokanCanParse,
  parseReservation: musashinokanParseReservation,
});

function musashinokanCanParse(raw: string): boolean {
  return (
    raw.includes("①予約番号：") &&
    raw.includes("②") &&
    raw.includes("③") &&
    raw.includes("④") &&
    raw.includes("⑤") &&
    raw.includes("⑥座席番号：")
  );
}

function musashinokanParseReservation(raw: string): Reservation {
  const lines = toLines(raw);
  const ticketNumber = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.ticketNumber);
  const title = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.title);
  const theaterLocation = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.theater);
  const screening = musashinokanParseScreening(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.screeningTime),
  );
  const totalPrice = parsePrice(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.totalPrice),
  );
  const seats = parseSeats(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.seats),
    /[、,／/\s]+/,
  );

  return {
    theater: {
      name: MUSASHINOKAN_THEATER_NAME,
      location: theaterLocation,
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

function musashinokanExtractPrefixedValue(lines: string[], prefix: string): string {
  const line = lines.find((candidate) => candidate.trimStart().startsWith(prefix));
  if (line === undefined) {
    throw new Error(`Missing required field: ${prefix}`);
  }

  const value = line.trimStart().slice(prefix.length).trim();
  if (value === "") {
    throw new Error(`Missing value for field: ${prefix}`);
  }

  return value;
}

function musashinokanParseScreening(value: string): Screening {
  const match = value.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})$/);
  if (!match) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const [, year, month, day, hour, minute] = match;
  const start = new Date(`${year}-${pad2(month)}-${pad2(day)}T${pad2(hour)}:${minute}:00+09:00`);
  if (Number.isNaN(start.getTime())) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const end = new Date(start.getTime() + MUSASHINOKAN_DEFAULT_SCREENING_MINUTES * 60 * 1000);

  return { start, end };
}
