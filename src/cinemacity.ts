/**
 * Ticket parser for 立川シネマシティ reservation emails.
 *
 * The implementation is ported from the `theather-mail-parser` project
 * (`src/parsers/cinemacity-parser.ts`).
 */
import { createTicketParser, pad2, parsePrice, parseSeats, toLines } from "./ticket-parser";
import type { Reservation, Screening } from "./ticket-parser";

const CINEMACITY_THEATER_NAME = "立川シネマシティ";

const CINEMACITY_LABELS = {
  ticketNumber: "■チケット番号",
  phoneNumber: "■登録電話番号",
  screeningTime: "■上映時間",
  theater: "■劇場",
  seats: "■座席",
  totalPrice: "■合計金額",
} as const;

/**
 * Parse a Gmail message body into a `Ticket`, or return `undefined` when the
 * body is not a recognizable 立川シネマシティ reservation email.
 *
 * When the body matches this cinema's format (`cinemaCityCanParse`) but a
 * required field is missing or malformed, this throws instead of returning
 * `undefined` so the caller can tell "not this format" apart from "this
 * format, but failed to parse" and report the latter.
 */
export const parseCinemaCityBody = createTicketParser({
  canParse: cinemaCityCanParse,
  parseReservation: cinemaCityParseReservation,
});

function cinemaCityCanParse(raw: string): boolean {
  return (
    raw.includes(CINEMACITY_LABELS.ticketNumber) &&
    raw.includes(CINEMACITY_LABELS.screeningTime) &&
    raw.includes(CINEMACITY_LABELS.theater) &&
    /シネマ[・･]?(?:ワン|ツー)|cinema\s*(?:one|two)/i.test(raw)
  );
}

function cinemaCityParseReservation(raw: string): Reservation {
  const lines = toLines(raw);
  const ticketNumber = cinemaCityExtractInlineValue(raw, CINEMACITY_LABELS.ticketNumber);
  const title = cinemaCityExtractTitle(lines);
  const screening = cinemaCityParseScreening(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.screeningTime),
  );
  const theaterLocation = cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.theater);
  const seatsValue = cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.seats);
  const seats = parseSeats(seatsValue.replace(/[[\]]/g, ""), /[、,\s]+/);
  const totalPrice = parsePrice(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.totalPrice),
  );

  return {
    theater: {
      name: CINEMACITY_THEATER_NAME,
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

function cinemaCityExtractInlineValue(raw: string, label: string): string {
  const escapedLabel = cinemaCityEscapeRegExp(label);
  const match = raw.match(new RegExp(`^${escapedLabel}[：:](.+)$`, "m"));
  if (!match) {
    throw new Error(`Missing required field: ${label}`);
  }
  return match[1].trim();
}

function cinemaCityExtractValueAfterLabel(lines: string[], label: string): string {
  const labelIndex = lines.findIndex((line) => line.trim() === label);
  if (labelIndex === -1) {
    throw new Error(`Missing required field: ${label}`);
  }

  const value = lines.slice(labelIndex + 1).find((line) => line.trim() !== "");
  if (value === undefined) {
    throw new Error(`Missing value for field: ${label}`);
  }
  return value.trim();
}

function cinemaCityExtractTitle(lines: string[]): string {
  const phoneNumberIndex = lines.findIndex((line) =>
    line.trim().startsWith(CINEMACITY_LABELS.phoneNumber),
  );
  if (phoneNumberIndex === -1) {
    throw new Error(`Missing required field: ${CINEMACITY_LABELS.phoneNumber}`);
  }

  const title = lines
    .slice(phoneNumberIndex + 1)
    .map((line) => line.trim())
    .find((line) => line !== "" && !line.startsWith("■"));

  if (title === undefined) {
    throw new Error("Missing required field: title");
  }

  return title;
}

function cinemaCityParseScreening(value: string): Screening {
  const match = value.match(
    /^(\d{4})年(\d{1,2})月(\d{1,2})日(?:\([^)]*\)|（[^）]*）)?\s+(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/,
  );
  if (!match) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const [, year, month, day, startHour, startMinute, endHour, endMinute] = match;
  const date = `${year}-${pad2(month)}-${pad2(day)}`;
  const start = new Date(`${date}T${pad2(startHour)}:${startMinute}:00+09:00`);
  const end = new Date(`${date}T${pad2(endHour)}:${endMinute}:00+09:00`);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  return { start, end };
}

function cinemaCityEscapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
