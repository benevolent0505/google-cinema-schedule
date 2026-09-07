/**
 * Registry of ticket email senders and their parsers.
 *
 * To support another cinema, add its parser in a separate source file and
 * register the sender addresses and parser in this list.
 */
import { parseCinemaCityBody } from "./cinemacity";
import { parseMusashinokanBody } from "./musashinokan";
import type { TicketMailSource } from "./ticket";

export function getTicketMailSources(): TicketMailSource[] {
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
