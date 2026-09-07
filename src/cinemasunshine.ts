import { createTicketParser, pad2, parsePrice, toLines } from "./ticket-parser";
import type { Reservation, Screening } from "./ticket-parser";

const CINEMASUNSHINE_THEATER_NAME = "グランドシネマサンシャイン池袋";

const CINEMASUNSHINE_LABELS = {
  ticketNumber: "[予約番号]",
  screeningTime: "[鑑賞日時]",
  title: "[作品名]",
  theater: "[スクリーン名]",
  seats: "[座席]",
  totalPrice: "[合計]",
} as const;

export const parseCinemaSunshineBody = createTicketParser({
  canParse: cinemaSunshineCanParse,
  parseReservation: cinemaSunshineParseReservation,
});

function cinemaSunshineCanParse(raw: string): boolean {
  return Object.values(CINEMASUNSHINE_LABELS).every((label) => raw.includes(label));
}

function cinemaSunshineParseReservation(raw: string): Reservation {
  const lines = toLines(raw);
  const ticketNumber = cinemaSunshineExtractValueAfterLabel(
    lines,
    CINEMASUNSHINE_LABELS.ticketNumber,
  );
  const screening = cinemaSunshineParseScreening(
    cinemaSunshineExtractValueAfterLabel(lines, CINEMASUNSHINE_LABELS.screeningTime),
  );
  const title = cinemaSunshineExtractValueAfterLabel(lines, CINEMASUNSHINE_LABELS.title);
  const theaterLocation = cinemaSunshineExtractValueAfterLabel(
    lines,
    CINEMASUNSHINE_LABELS.theater,
  );
  const seats = cinemaSunshineParseSeats(
    cinemaSunshineExtractSectionValues(lines, CINEMASUNSHINE_LABELS.seats),
  );
  const totalPrice = parsePrice(
    cinemaSunshineExtractValueAfterLabel(lines, CINEMASUNSHINE_LABELS.totalPrice),
  );

  return {
    theater: {
      name: CINEMASUNSHINE_THEATER_NAME,
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

function cinemaSunshineExtractValueAfterLabel(lines: string[], label: string): string {
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

// 座席は「座席 券種 金額」の行がラベル直後に人数分並ぶ。空行が来たら座席欄の
// 終わり、次のラベル（[〜]）が来ても終わりと判断して、無関係な行を拾わない。
function cinemaSunshineExtractSectionValues(lines: string[], label: string): string[] {
  const labelIndex = lines.findIndex((line) => line.trim() === label);
  if (labelIndex === -1) {
    throw new Error(`Missing required field: ${label}`);
  }

  const values: string[] = [];
  for (const line of lines.slice(labelIndex + 1)) {
    const trimmed = line.trim();
    if (trimmed === "") {
      if (values.length > 0) {
        break;
      }
      continue;
    }
    if (/^\[[^\]]+\]$/.test(trimmed)) {
      break;
    }
    values.push(trimmed);
  }

  if (values.length === 0) {
    throw new Error(`Missing value for field: ${label}`);
  }

  return values;
}

function cinemaSunshineParseSeats(values: string[]): string[] {
  const seats = values
    .map((value) => value.split(/\s+/)[0])
    .map((seat) => seat.trim())
    .filter((seat) => seat !== "");

  if (seats.length === 0) {
    throw new Error(`Invalid seats: ${values.join("\n")}`);
  }

  return seats;
}

function cinemaSunshineParseScreening(value: string): Screening {
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
