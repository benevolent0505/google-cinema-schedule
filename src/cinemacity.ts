/**
 * Ticket parser for 立川シネマシティ reservation emails.
 *
 * The implementation is ported from the `theather-mail-parser` project
 * (`src/parsers/cinemacity-parser.ts`).
 */
import type { Ticket } from "./ticket";

type CinemaCityTheater = {
  name: string;
  location: string;
};

type CinemaCityMovie = {
  title: string;
};

type CinemaCityScreening = {
  start: Date;
  end?: Date;
};

type CinemaCityReservation = {
  theater: CinemaCityTheater;
  movie: CinemaCityMovie;
  screening: CinemaCityScreening;
  ticketNumber: string;
  phoneNumber?: string;
  seats: string[];
  ticketCount: number;
  totalPrice: number;
};

const CINEMACITY_THEATER_NAME = "立川シネマシティ";

const CINEMACITY_LABELS = {
  ticketNumber: "■チケット番号",
  phoneNumber: "■登録電話番号",
  screeningTime: "■上映時間",
  theater: "■劇場",
  seats: "■座席",
  ticketCount: "■枚数",
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
export function parseCinemaCityBody(body: string): Ticket | undefined {
  if (!cinemaCityCanParse(body)) {
    return undefined;
  }

  const reservation = cinemaCityParseReservation(body);

  return {
    ticketNumber: reservation.ticketNumber,
    title: reservation.movie.title,
    startTime: reservation.screening.start,
    endTime: reservation.screening.end ?? reservation.screening.start,
    theater: reservation.theater.location,
    sheet: reservation.seats.join(", "),
  };
}

function cinemaCityCanParse(raw: string): boolean {
  return (
    raw.includes(CINEMACITY_LABELS.ticketNumber) &&
    raw.includes(CINEMACITY_LABELS.screeningTime) &&
    raw.includes(CINEMACITY_LABELS.theater) &&
    /シネマ[・･]?(?:ワン|ツー)|cinema\s*(?:one|two)/i.test(raw)
  );
}

function cinemaCityParseReservation(raw: string): CinemaCityReservation {
  const lines = cinemaCityToLines(raw);
  const ticketNumber = cinemaCityExtractInlineValue(raw, CINEMACITY_LABELS.ticketNumber);
  const phoneNumber = cinemaCityStripParenthetical(
    cinemaCityExtractInlineValue(raw, CINEMACITY_LABELS.phoneNumber),
  );
  const title = cinemaCityExtractTitle(lines);
  const screening = cinemaCityParseScreening(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.screeningTime),
  );
  const theaterLocation = cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.theater);
  const seats = cinemaCityParseSeats(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.seats),
  );
  const ticketCount = cinemaCityParseTicketCount(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.ticketCount),
  );
  const totalPrice = cinemaCityParsePrice(
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
    phoneNumber,
    seats,
    ticketCount,
    totalPrice,
  };
}

function cinemaCityToLines(raw: string): string[] {
  return raw.replace(/\r\n?/g, "\n").split("\n");
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

function cinemaCityParseScreening(value: string): { start: Date; end: Date } {
  const match = value.match(
    /^(\d{4})年(\d{1,2})月(\d{1,2})日(?:\([^)]*\)|（[^）]*）)?\s+(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/,
  );
  if (!match) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const [, year, month, day, startHour, startMinute, endHour, endMinute] = match;
  const date = `${year}-${cinemaCityPad2(month)}-${cinemaCityPad2(day)}`;
  const start = new Date(`${date}T${cinemaCityPad2(startHour)}:${startMinute}:00+09:00`);
  const end = new Date(`${date}T${cinemaCityPad2(endHour)}:${endMinute}:00+09:00`);

  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  return { start, end };
}

function cinemaCityParseSeats(value: string): string[] {
  const withoutBrackets = value.replace(/\[/g, "").replace(/\]/g, "").trim();
  const seats = withoutBrackets
    .split(/[、,\s]+/)
    .map((seat) => seat.trim())
    .filter((seat) => seat !== "");

  if (seats.length === 0) {
    throw new Error(`Invalid seats: ${value}`);
  }

  return seats;
}

function cinemaCityParseTicketCount(value: string): number {
  const count = Number.parseInt(value.replace(/,/g, ""), 10);
  if (!Number.isFinite(count)) {
    throw new Error(`Invalid ticket count: ${value}`);
  }
  return count;
}

function cinemaCityParsePrice(value: string): number {
  const match = value.match(/[\d,]+/);
  if (!match) {
    throw new Error(`Invalid price: ${value}`);
  }

  const price = Number.parseInt(match[0].replace(/,/g, ""), 10);
  if (!Number.isFinite(price)) {
    throw new Error(`Invalid price: ${value}`);
  }
  return price;
}

function cinemaCityStripParenthetical(value: string): string {
  return value.replace(/（.*?）|\(.*?\)/g, "").trim();
}

function cinemaCityPad2(value: string): string {
  return value.padStart(2, "0");
}

function cinemaCityEscapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
