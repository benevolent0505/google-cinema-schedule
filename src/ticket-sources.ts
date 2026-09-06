/**
 * Registry of ticket email senders and their parsers.
 *
 * To support another cinema, add its parser in a separate source file and
 * register the sender addresses and parser in this list.
 */
function getTicketMailSources(): TicketMailSource[] {
  return [
    {
      mailAddresses: ["ticket@cinemacity.co.jp"],
      parseBody: parseCinemaCityBody,
    },
    {
      mailAddresses: ["reserve@musashino.cineticket.jp"],
      parseBody: parseMusashinokanBody,
    },
  ];
}

/**
 * Backward-compatible convenience parser using all registered sources.
 */
function parseBody(body: string): Ticket | undefined {
  return parseTicketBody(body, getTicketMailSources());
}

Object.assign(globalThis, { getTicketMailSources, parseBody });
