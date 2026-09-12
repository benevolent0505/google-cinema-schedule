export type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  screen: string;
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
 * パーサーとの契約: 本文がその映画館の形式でなければ `undefined` を返し、形式には
 * 一致したのに必須項目が欠けている場合は例外を投げる。前者は「別の映画館のメール」、
 * 後者は「メール形式が変わった可能性」で、どちらも `onParseError` から区別して
 * 報告する。
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
