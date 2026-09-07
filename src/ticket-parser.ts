import type { Ticket } from "./ticket";

export type Theater = {
  name: string;
  location: string;
};

export type Movie = {
  title: string;
};

export type Screening = {
  start: Date;
  end: Date;
};

export type Reservation = {
  theater: Theater;
  movie: Movie;
  screening: Screening;
  ticketNumber: string;
  seats: string[];
  totalPrice: number;
};

export function createTicketParser(config: {
  canParse: (raw: string) => boolean;
  parseReservation: (raw: string) => Reservation;
}): (body: string) => Ticket | undefined {
  return (body) => {
    if (!config.canParse(body)) {
      return undefined;
    }

    const reservation = config.parseReservation(body);

    return {
      ticketNumber: reservation.ticketNumber,
      title: reservation.movie.title,
      startTime: reservation.screening.start,
      endTime: reservation.screening.end,
      theater: reservation.theater.location,
      sheet: reservation.seats.join(", "),
    };
  };
}

export function toLines(raw: string): string[] {
  return raw.replace(/\r\n?/g, "\n").split("\n");
}

export function pad2(value: string): string {
  return value.padStart(2, "0");
}

export function parsePrice(value: string): number {
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

export function parseSeats(value: string, separator: RegExp): string[] {
  const seats = value
    .split(separator)
    .map((seat) => seat.trim())
    .filter((seat) => seat !== "");

  if (seats.length === 0) {
    throw new Error(`Invalid seats: ${value}`);
  }

  return seats;
}
