import { parseCinemaCityBody } from "./cinemacity";
import { parseCinemaSunshineBody } from "./cinemasunshine";
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
    {
      mailAddresses: ["noreply@ticket-cinemasunshine.com"],
      parseBody: parseCinemaSunshineBody,
    },
  ];
}
