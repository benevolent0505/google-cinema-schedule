/**
 * Shared ticket types and parser dispatching.
 */
export type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

type TicketParser = (body: string) => Ticket | undefined;

export type TicketMailSource = {
  mailAddresses: readonly string[];
  parseBody: TicketParser;
};

export type TicketParseFailure =
  | { reason: "unknown_sender"; fromAddress: string }
  | { reason: "unrecognized_body"; source: TicketMailSource }
  | { reason: "error"; source: TicketMailSource; error: unknown };

function normalizeMailAddress(from: string): string {
  const match = from.match(/<([^<>]+)>/);
  const address = match ? match[1] : from;

  return address.trim().toLowerCase();
}

function findTicketSourceByFromAddress(
  sources: readonly TicketMailSource[],
  fromAddress: string,
): TicketMailSource | undefined {
  const normalized = normalizeMailAddress(fromAddress);

  return sources.find((source) =>
    source.mailAddresses.some((address) => address.toLowerCase() === normalized),
  );
}

/**
 * The sender address determines the parser exclusively - no fallback to
 * another parser once one is selected. A parser returning `undefined` and
 * one throwing are both reported via `onParseError`, since either means
 * this sender's mail no longer looks like what it used to.
 */
export function parseTicketBody(
  body: string,
  fromAddress: string,
  sources: readonly TicketMailSource[],
  onParseError?: (failure: TicketParseFailure) => void,
): Ticket | undefined {
  const source = findTicketSourceByFromAddress(sources, fromAddress);

  if (!source) {
    onParseError?.({ reason: "unknown_sender", fromAddress });
    return undefined;
  }

  try {
    const ticket = source.parseBody(body);

    if (!ticket) {
      onParseError?.({ reason: "unrecognized_body", source });
      return undefined;
    }

    return ticket;
  } catch (error) {
    onParseError?.({ reason: "error", source, error });
    return undefined;
  }
}
