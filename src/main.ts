import { getTicketMailSources } from "./ticket-sources";
import { parseTicketBody } from "./ticket";
import type { Ticket, TicketMailSource } from "./ticket";
import { createLogger } from "./logger";
import type { Logger } from "./logger";

const calendarSearchKey = "映画館チケット";
const legacyCalendarSearchKeys = ["シネマシティ"];
const debugExecutionDateProperty = "DEBUG_EXECUTION_DATE";
const debugLogEnabledProperty = "DEBUG_LOG_ENABLED";

export type MailSearchRange = {
  start: Date;
  end: Date;
};

function isDebugLogEnabled(): boolean {
  const value = PropertiesService.getScriptProperties().getProperty(debugLogEnabledProperty);

  if (value === null) {
    return false;
  }

  return ["true", "1", "yes", "on"].includes(value.trim().toLowerCase());
}

export function main(): void {
  const logger = createLogger(isDebugLogEnabled());

  runCinemaSchedule(resolveMailSearchRange(new Date(Date.now())), logger);
}

export function debugMain(executionDate?: string): void {
  const logger = createLogger(isDebugLogEnabled());

  const specifiedDate =
    executionDate ??
    PropertiesService.getScriptProperties().getProperty(debugExecutionDateProperty);

  if (!specifiedDate) {
    throw new Error(`${debugExecutionDateProperty} に実行日を YYYY-MM-DD 形式で指定してください。`);
  }

  const searchRange = resolveMailSearchRange(parseExecutionDate(specifiedDate));

  logger.info(
    `debugMain: 仮想実行日=${specifiedDate} 対象範囲=${formatDateTimeForLog(searchRange.start)} 以上 ${formatDateTimeForLog(searchRange.end)} 未満`,
  );

  runCinemaSchedule(searchRange, logger);
}

export function parseExecutionDate(value: string): Date {
  const matched = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);

  if (!matched) {
    throw new Error(
      `${debugExecutionDateProperty} には YYYY-MM-DD 形式で日付を指定してください: ${value}`,
    );
  }

  const year = Number(matched[1]);
  const month = Number(matched[2]);
  const day = Number(matched[3]);
  const parsed = new Date(year, month - 1, day, 0, 0, 0, 0);

  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    throw new Error(`${debugExecutionDateProperty} には実在する日付を指定してください: ${value}`);
  }

  return parsed;
}

export function resolveMailSearchRange(executionDate: Date): MailSearchRange {
  const year = executionDate.getFullYear();
  const month = executionDate.getMonth();
  const day = executionDate.getDate();

  return {
    start: new Date(year, month, day - 1, 0, 0, 0, 0),
    end: new Date(year, month, day + 1, 0, 0, 0, 0),
  };
}

function formatDateTimeForLog(date: Date): string {
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");

  return `${formatMailSearchDate(date)} ${hour}:${minute}`;
}

function runCinemaSchedule(searchRange: MailSearchRange, logger: Logger): void {
  const tickets = fetchTickets(getTicketMailSources(), searchRange, logger);

  if (tickets.length === 0) {
    return;
  }

  const minStartTime = tickets
    .map((ticket) => ticket.startTime)
    .reduce((a, b) => (a.getTime() < b.getTime() ? a : b));
  const maxEndTime = tickets
    .map((ticket) => ticket.endTime)
    .reduce((a, b) => (a.getTime() > b.getTime() ? a : b));

  const existingEvents = fetchExistingEvents(minStartTime, maxEndTime);

  const willRegisterTickets = tickets.filter((ticket) => {
    const isExist = existingEvents.some((event) => {
      return event.getDescription().includes(buildTicketNumberMarker(ticket.ticketNumber));
    });

    if (isExist) {
      logger.info(`Skip: ${ticket.title}`);
    }

    return !isExist;
  });

  for (const ticket of willRegisterTickets) {
    const event = registerEvent(ticket);
    logger.info(`Registered: ${event.getTitle()}`);
  }
}

function buildGmailMessageLink(messageId: string): string {
  return `https://mail.google.com/mail/u/0/#all/${messageId}`;
}

function fetchTickets(
  sources: readonly TicketMailSource[],
  searchRange: MailSearchRange,
  logger: Logger,
): Ticket[] {
  if (sources.length === 0) {
    logger.debug("fetchTickets: チケットメールの取得元が未設定のため終了します。");
    return [];
  }

  const searchCriteria = buildTicketMailSearchCriteria(sources, searchRange.start);
  logger.debug(`fetchTickets: 検索条件 = ${searchCriteria}`);

  const threads = GmailApp.search(searchCriteria);
  logger.debug(`fetchTickets: 検索スレッド数 = ${threads.length}`);

  let tickets: Ticket[] = [];

  for (const [threadIndex, thread] of threads.entries()) {
    const messages = thread.getMessages();
    logger.debug(
      `fetchTickets: スレッド[${threadIndex}] 件名="${thread.getFirstMessageSubject()}" メッセージ数=${messages.length}`,
    );

    for (const [messageIndex, message] of messages.entries()) {
      const messageDate = message.getDate();
      const fromAddress = message.getFrom();
      const messageLink = buildGmailMessageLink(message.getId());
      logger.debug(
        `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] from=${fromAddress} date=${messageDate.toISOString()} subject="${message.getSubject()}" link=${messageLink}`,
      );

      if (messageDate.getTime() < searchRange.start.getTime()) {
        logger.debug(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] は検索対象期間より前のためスキップします。`,
        );
        continue;
      }

      if (messageDate.getTime() >= searchRange.end.getTime()) {
        logger.debug(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] は検索対象期間より後のためスキップします。`,
        );
        continue;
      }

      const body = message.getPlainBody();
      const ticket = parseTicketBody(body, fromAddress, sources, (failure) => {
        if (failure.reason === "unknown_sender") {
          logger.warn(
            `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.fromAddress} は登録されていないためスキップします。 link=${messageLink}`,
          );
          return;
        }

        if (failure.reason === "unrecognized_body") {
          logger.warn(
            `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.source.mailAddresses.join(", ")} の本文が想定の形式と一致しませんでした。 link=${messageLink}`,
          );
          return;
        }

        logger.error(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.source.mailAddresses.join(", ")} の解析に失敗しました。 link=${messageLink}`,
          failure.error,
        );
      });

      if (ticket) {
        logger.debug(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 解析成功 title="${ticket.title}" start=${ticket.startTime.toISOString()} end=${ticket.endTime.toISOString()}`,
        );
        tickets = [...tickets, ticket];
      }
    }
  }

  logger.debug(`fetchTickets: 取得したチケット数 = ${tickets.length}`);

  const uniqueTickets = dedupeTicketsByTicketNumber(tickets);
  if (uniqueTickets.length !== tickets.length) {
    logger.debug(
      `fetchTickets: チケット番号の重複を除去しました 重複除去前=${tickets.length} 重複除去後=${uniqueTickets.length}`,
    );
  }

  return uniqueTickets;
}

export function dedupeTicketsByTicketNumber(tickets: readonly Ticket[]): Ticket[] {
  const seenTicketNumbers = new Set<string>();
  const uniqueTickets: Ticket[] = [];

  for (const ticket of tickets) {
    if (seenTicketNumbers.has(ticket.ticketNumber)) {
      continue;
    }

    seenTicketNumbers.add(ticket.ticketNumber);
    uniqueTickets.push(ticket);
  }

  return uniqueTickets;
}

export function buildTicketMailSearchCriteria(
  sources: readonly TicketMailSource[],
  searchStartDateTime: Date,
): string {
  const mailAddresses = sources.flatMap((source) => source.mailAddresses);
  const uniqueMailAddresses = [...new Set(mailAddresses)];
  const mailAddressQuery = uniqueMailAddresses.map((address) => `from:${address}`).join(" OR ");
  const periodQuery = `newer:${formatMailSearchDate(searchStartDateTime)}`;

  return `(${mailAddressQuery}) AND ${periodQuery}`;
}

function formatMailSearchDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function fetchExistingEvents(
  startTime: Date,
  endTime: Date,
): GoogleAppsScript.Calendar.CalendarEvent[] {
  const calendar = CalendarApp.getDefaultCalendar();

  return [calendarSearchKey, ...legacyCalendarSearchKeys].flatMap((search) =>
    calendar.getEvents(startTime, endTime, { search }),
  );
}

function registerEvent(ticket: Ticket): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getDefaultCalendar();

  const description = `劇場: ${ticket.theater}\n座席: ${ticket.sheet}\n${buildTicketNumberMarker(ticket.ticketNumber)}\n検索用キーワード: ${calendarSearchKey}`;
  const location = ticket.theater;

  const event = calendar.createEvent(ticket.title, ticket.startTime, ticket.endTime, {
    description,
    location,
  });

  return event;
}

function buildTicketNumberMarker(ticketNumber: string): string {
  return `チケット番号: ${ticketNumber}`;
}
