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
 *
 * A parser returns `undefined` when the body does not look like its own
 * format, but may also throw when the body matches its format yet a
 * required field is missing or malformed. That case is reported via
 * `onParseError` (instead of being treated the same as "not this parser")
 * and the remaining parsers are still tried.
 */
function parseTicketBody(
  body: string,
  sources: readonly TicketMailSource[],
  onParseError?: (source: TicketMailSource, error: unknown) => void,
): Ticket | undefined {
  for (const source of sources) {
    try {
      const ticket = source.parseBody(body);

      if (ticket) {
        return ticket;
      }
    } catch (error) {
      onParseError?.(source, error);
    }
  }

  return undefined;
}

Object.assign(globalThis, { parseTicketBody });
