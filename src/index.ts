const searchKey = "シネマシティ";

function main() {
  // 実行日の1日前からのメールを取得する
  const now = new Date(Date.now());
  const newerThreshold = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1,
    0,
    0,
    0,
  );

  // メールからチケット情報を取得する
  const tickets = fetchTickets(["ticket@cinemacity.co.jp"], newerThreshold);

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
function fetchTickets(mailAddresses: string[], newerThreshold: Date): Ticket[] {
  const mailAddressQuery = mailAddresses
    .map((address) => `from:${address}`)
    .join(" OR ");
  const periodQuery = `newer:${newerThreshold.toISOString().slice(0, 10)}`;
  const criteria = `(${mailAddressQuery}) AND ${periodQuery}`;

  const threads = GmailApp.search(criteria);

  let tickets: Ticket[] = [];

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const body = message.getPlainBody();
      const ticket = parseBody(body);

      if (ticket) {
        tickets = [...tickets, ticket];
      }
    }
  }

  return tickets;
}

type Ticket = {
  ticketNumber: string;
  title: string;
  startTime: Date;
  endTime: Date;
  theater: string;
  sheet: string;
};

function parseBody(body: string): Ticket | undefined {
  const parsed = body.match(
    /■チケット番号：(?<ticketNumber>\d+)\r\n■登録電話番号：\d+（下4ケタのみでOK）\r\n\r\n(?<title>.+)\r\n■上映時間\r\n(?<date>.+)\r\n■劇場 （ワン：高島屋右隣／ツー：モノレール下遊歩道沿）\r\n(?<theater>.+)\r\n■座席\r\n(?<sheet>(.+))\r\n/,
  )?.groups;

  if (!parsed) {
    return undefined;
  }

  const { ticketNumber, title, date, theater, sheet } = parsed;

  const parsedDate = date.match(
    /(?<year>\d+)年(?<month>\d+)月(?<day>\d+)日\(.\) (?<startHour>\d\d):(?<startMin>\d\d) - (?<endHour>\d\d):(?<endMin>\d\d)/,
  )?.groups;

  if (parsedDate) {
    const startTime = new Date(
      parseInt(parsedDate.year),
      parseInt(parsedDate.month) - 1,
      parseInt(parsedDate.day),
      parseInt(parsedDate.startHour),
      parseInt(parsedDate.startMin),
    );

    const endTime = new Date(
      parseInt(parsedDate.year),
      parseInt(parsedDate.month) - 1,
      parseInt(parsedDate.day),
      parseInt(parsedDate.endHour),
      parseInt(parsedDate.endMin),
    );

    return {
      ticketNumber,
      title,
      startTime,
      endTime,
      theater,
      sheet,
    };
  }
}

/**
 * Google カレンダーのイベントを取得する
 */
function fetchExistingEvents(
  startTime: Date,
  endTime: Date,
): GoogleAppsScript.Calendar.CalendarEvent[] {
  const calendar = CalendarApp.getDefaultCalendar();

  const events = calendar.getEvents(startTime, endTime, { search: searchKey });
  return events;
}

/**
 * イベントをカレンダーに登録する
 */
function registerEvent(
  ticket: Ticket,
): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getDefaultCalendar();

  const description = `劇場: ${ticket.theater}\n座席: ${ticket.sheet}\nチケット番号: ${ticket.ticketNumber}\n検索用キーワード: ${searchKey}`;
  const location = ticket.theater;

  const event = calendar.createEvent(
    ticket.title,
    ticket.startTime,
    ticket.endTime,
    {
      description,
      location,
    },
  );

  return event;
}
