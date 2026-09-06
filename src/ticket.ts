/**
 * Shared ticket types and parser dispatching.
 *
 * Google Apps Script shares a single global scope across every source file,
 * so build-target files intentionally do not use `import` / `export`.
 */
type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

type TicketParser = (body: string) => Ticket | undefined;

type TicketMailSource = {
  mailAddresses: readonly string[];
  parseBody: TicketParser;
};

/**
 * Try the registered parsers in order and return the first parsed ticket.
 */
function parseTicketBody(body: string, sources: readonly TicketMailSource[]): Ticket | undefined {
  for (const source of sources) {
    const ticket = source.parseBody(body);

    if (ticket) {
      return ticket;
    }
  }

  return undefined;
}

Object.assign(globalThis, { parseTicketBody });
