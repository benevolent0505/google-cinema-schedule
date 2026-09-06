/**
 * Google Apps Script entry point.
 *
 * Select `main` in the Apps Script editor and run it, attach it to a
 * time-driven trigger, or call it with `clasp run main`. It reads recent
 * cinema ticket confirmation emails from Gmail and registers any missing
 * screenings into the default Google Calendar.
 *
 * Shared types, parsers, and source settings come from the other build-target
 * files. In Apps Script all files share one global scope, so no `import` is
 * needed.
 */
const calendarSearchKey = "映画館チケット";
const legacyCalendarSearchKeys = ["シネマシティ"];

function main(): void {
  // 実行日の1日前からのメールを取得する
  const now = new Date(Date.now());
  const newerThreshold = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 0, 0, 0);

  // メールからチケット情報を取得する
  const tickets = fetchTickets(getTicketMailSources(), newerThreshold);

  // 対象のチケットがない場合、空配列に対する reduce を避けて終了する
  if (tickets.length === 0) {
    return;
  }

  // チケット情報がカレンダーに登録されているか確認する
  const minStartTime = tickets
    .map((ticket) => ticket.startTime)
    .reduce((a, b) => (a.getTime() < b.getTime() ? a : b));
  const maxEndTime = tickets
    .map((ticket) => ticket.endTime)
    .reduce((a, b) => (a.getTime() > b.getTime() ? a : b));

  const existingEvents = fetchExistingEvents(minStartTime, maxEndTime);

  const willRegisterTickets = tickets.filter((ticket) => {
    const isExist = existingEvents.some((event) => {
      return event.getTitle().includes(ticket.title);
    });

    if (isExist) {
      Logger.log(`Skip: ${ticket.title}`);
    }

    return !isExist;
  });

  // 登録されていない場合はカレンダーに登録する;
  for (const ticket of willRegisterTickets) {
    const event = registerEvent(ticket);
    Logger.log(`Registered: ${event.getTitle()}`);
  }
}

/**
 * チケット情報を取得する
 */
function fetchTickets(sources: readonly TicketMailSource[], newerThreshold: Date): Ticket[] {
  if (sources.length === 0) {
    return [];
  }

  const threads = GmailApp.search(buildTicketMailSearchCriteria(sources, newerThreshold));

  let tickets: Ticket[] = [];

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const body = message.getPlainBody();
      const ticket = parseTicketBody(body, sources);

      if (ticket) {
        tickets = [...tickets, ticket];
      }
    }
  }

  return tickets;
}

function buildTicketMailSearchCriteria(
  sources: readonly TicketMailSource[],
  newerThreshold: Date,
): string {
  const mailAddresses = sources.flatMap((source) => source.mailAddresses);
  const uniqueMailAddresses = [...new Set(mailAddresses)];
  const mailAddressQuery = uniqueMailAddresses.map((address) => `from:${address}`).join(" OR ");
  const periodQuery = `newer:${newerThreshold.toISOString().slice(0, 10)}`;

  return `(${mailAddressQuery}) AND ${periodQuery}`;
}

/**
 * Google カレンダーのイベントを取得する
 */
function fetchExistingEvents(
  startTime: Date,
  endTime: Date,
): GoogleAppsScript.Calendar.CalendarEvent[] {
  const calendar = CalendarApp.getDefaultCalendar();

  return [calendarSearchKey, ...legacyCalendarSearchKeys].flatMap((search) =>
    calendar.getEvents(startTime, endTime, { search }),
  );
}

/**
 * イベントをカレンダーに登録する
 */
function registerEvent(ticket: Ticket): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getDefaultCalendar();

  const description = `劇場: ${ticket.theater}\n座席: ${ticket.sheet}\nチケット番号: ${ticket.ticketNumber}\n検索用キーワード: ${calendarSearchKey}`;
  const location = ticket.theater;

  const event = calendar.createEvent(ticket.title, ticket.startTime, ticket.endTime, {
    description,
    location,
  });

  return event;
}

// Register entry points on the global scope. Apps Script can already run a
// top-level function by name; listing them here documents the public surface
// and keeps linters from flagging them as "unused".
Object.assign(globalThis, { buildTicketMailSearchCriteria, main });
