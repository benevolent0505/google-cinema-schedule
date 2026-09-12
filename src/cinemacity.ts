import { createTicketParser, pad2, parsePrice, parseSeats, toLines } from "./ticket-parser";
import type { Reservation, Screening } from "./ticket-parser";

const CINEMACITY_THEATER_NAMES = {
  one: "シネマシティ シネマ・ワン",
  two: "シネマシティ シネマ・ツー",
} as const;

const CINEMACITY_LABELS = {
  ticketNumber: "■チケット番号",
  phoneNumber: "■登録電話番号",
  screeningTime: "■上映時間",
  theater: "■劇場",
  seats: "■座席",
  totalPrice: "■合計金額",
} as const;

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
  const theaterName = cinemaCityResolveTheaterName(theaterLocation);
  const seatsValue = cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.seats);
  const seats = parseSeats(seatsValue.replace(/[[\]]/g, ""), /[、,\s]+/);
  const totalPrice = parsePrice(
    cinemaCityExtractValueAfterLabel(lines, CINEMACITY_LABELS.totalPrice),
  );

  return {
    theater: {
      name: theaterName,
      screen: theaterLocation,
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

function cinemaCityResolveTheaterName(location: string): string {
  if (/シネマ[・･]?ワン|cinema\s*one/i.test(location)) {
    return CINEMACITY_THEATER_NAMES.one;
  }
  if (/シネマ[・･]?ツー|cinema\s*two/i.test(location)) {
    return CINEMACITY_THEATER_NAMES.two;
  }

  throw new Error(`Unknown Cinema City theater: ${location}`);
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
