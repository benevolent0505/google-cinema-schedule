/**
 * Ticket parser for 新宿武蔵野館 reservation emails.
 *
 * The implementation is ported from the `theather-mail-parser` project
 * (`src/parsers/musashinokan-parser.ts`). Reservation emails are sent from
 * `reserve@musashino.cineticket.jp` with the subject `インターネットチケット購入`.
 */
type MusashinokanTheater = {
  name: string;
  location: string;
};

type MusashinokanMovie = {
  title: string;
};

type MusashinokanScreening = {
  start: Date;
  end?: Date;
};

type MusashinokanReservation = {
  theater: MusashinokanTheater;
  movie: MusashinokanMovie;
  screening: MusashinokanScreening;
  ticketNumber: string;
  seats: string[];
  ticketType: string;
  totalPrice: number;
};

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
  ticketType: "⑤",
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
function parseMusashinokanBody(body: string): Ticket | undefined {
  if (!musashinokanCanParse(body)) {
    return undefined;
  }

  const reservation = musashinokanParseReservation(body);

  return {
    ticketNumber: reservation.ticketNumber,
    title: reservation.movie.title,
    startTime: reservation.screening.start,
    endTime: reservation.screening.end ?? reservation.screening.start,
    theater: reservation.theater.location,
    sheet: reservation.seats.join(", "),
  };
}

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

function musashinokanParseReservation(raw: string): MusashinokanReservation {
  const lines = musashinokanToLines(raw);
  const ticketNumber = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.ticketNumber);
  const title = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.title);
  const theaterLocation = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.theater);
  const screening = musashinokanParseScreening(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.screeningTime),
  );
  const ticketType = musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.ticketType);
  const totalPrice = musashinokanParsePrice(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.totalPrice),
  );
  const seats = musashinokanParseSeats(
    musashinokanExtractPrefixedValue(lines, MUSASHINOKAN_LABELS.seats),
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
    ticketType,
    totalPrice,
  };
}

function musashinokanToLines(raw: string): string[] {
  return raw.replace(/\r\n?/g, "\n").split("\n");
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

function musashinokanParseScreening(value: string): { start: Date; end: Date } {
  const match = value.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})$/);
  if (!match) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const [, year, month, day, hour, minute] = match;
  const start = new Date(
    `${year}-${musashinokanPad2(month)}-${musashinokanPad2(day)}T${musashinokanPad2(hour)}:${minute}:00+09:00`,
  );
  if (Number.isNaN(start.getTime())) {
    throw new Error(`Invalid screening time: ${value}`);
  }

  const end = new Date(start.getTime() + MUSASHINOKAN_DEFAULT_SCREENING_MINUTES * 60 * 1000);

  return { start, end };
}

function musashinokanParsePrice(value: string): number {
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

function musashinokanParseSeats(value: string): string[] {
  const seats = value
    .split(/[、,／/\s]+/)
    .map((seat) => seat.trim())
    .filter((seat) => seat !== "");

  if (seats.length === 0) {
    throw new Error(`Invalid seats: ${value}`);
  }

  return seats;
}

function musashinokanPad2(value: string): string {
  return value.padStart(2, "0");
}

Object.assign(globalThis, { parseMusashinokanBody });
