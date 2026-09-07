/**
 * Google Apps Script entry point.
 *
 * Select `main` in the Apps Script editor and run it, attach it to a
 * time-driven trigger, or call it with `clasp run main`. It reads recent
 * cinema ticket confirmation emails from Gmail and registers any missing
 * screenings into the default Google Calendar.
 *
 * This is the esbuild entry point; everything it depends on is pulled in via
 * `import` and bundled into a single `dist/main.js`. `scripts/build.mjs`
 * appends top-level `function main()` / `function debugMain()` wrappers
 * after the bundle so Apps Script's function picker and trigger UI (which
 * scan the script text for top-level declarations) can still find them.
 */
import { getTicketMailSources } from "./ticket-sources";
import { parseTicketBody } from "./ticket";
import type { Ticket, TicketMailSource } from "./ticket";

const calendarSearchKey = "映画館チケット";
const legacyCalendarSearchKeys = ["シネマシティ"];
const debugMailSearchStartDateTimeProperty = "DEBUG_MAIL_SEARCH_START_DATETIME";
const debugLogEnabledProperty = "DEBUG_LOG_ENABLED";

/**
 * Script Properties の `DEBUG_LOG_ENABLED` が有効値のときだけデバッグログを出力する。
 *
 * 有効値（大文字小文字は無視）: `true` / `1` / `yes` / `on`
 */
function isDebugLogEnabled(): boolean {
  const value = PropertiesService.getScriptProperties().getProperty(debugLogEnabledProperty);

  if (value === null) {
    return false;
  }

  return ["true", "1", "yes", "on"].includes(value.trim().toLowerCase());
}

/**
 * デバッグログを出力する。`DEBUG_LOG_ENABLED` が無効な場合は何もしない。
 */
function debugLog(message: string): void {
  if (isDebugLogEnabled()) {
    Logger.log(message);
  }
}

export function main(): void {
  // 実行日の1日前からのメールを取得する
  const now = new Date(Date.now());
  const searchStartDateTime = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1,
    0,
    0,
    0,
  );

  runCinemaSchedule(searchStartDateTime);
}

/**
 * 指定した日時以降のメールを対象に実行するデバッグ用エントリーポイント。
 *
 * `clasp run` から日時を引数で渡せるほか、Apps Script エディタから引数なしで
 * 実行する場合は Script Properties の `DEBUG_MAIL_SEARCH_START_DATETIME` を使用する。
 */
export function debugMain(searchStartDateTime?: string): void {
  const specifiedDateTime =
    searchStartDateTime ??
    PropertiesService.getScriptProperties().getProperty(debugMailSearchStartDateTimeProperty);

  if (!specifiedDateTime) {
    throw new Error(`${debugMailSearchStartDateTimeProperty} に検索開始日時を指定してください。`);
  }

  const parsedDateTime = new Date(specifiedDateTime);

  if (Number.isNaN(parsedDateTime.getTime())) {
    throw new Error(
      `${debugMailSearchStartDateTimeProperty} には有効な日時を指定してください: ${specifiedDateTime}`,
    );
  }

  runCinemaSchedule(parsedDateTime);
}

function runCinemaSchedule(searchStartDateTime: Date): void {
  // メールからチケット情報を取得する
  const tickets = fetchTickets(getTicketMailSources(), searchStartDateTime);

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
    // タイトルの部分一致ではなく、description に埋め込んだチケット番号で照合する。
    // タイトルだけで見ると、同じ作品を別日にもう一度観た場合や、短いタイトルが
    // 無関係な予定に一致してしまう場合に、登録すべきチケットを誤ってスキップして
    // しまう。
    const isExist = existingEvents.some((event) => {
      return event.getDescription().includes(buildTicketNumberMarker(ticket.ticketNumber));
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
function fetchTickets(sources: readonly TicketMailSource[], searchStartDateTime: Date): Ticket[] {
  if (sources.length === 0) {
    debugLog("fetchTickets: チケットメールの取得元が未設定のため終了します。");
    return [];
  }

  const searchCriteria = buildTicketMailSearchCriteria(sources, searchStartDateTime);
  debugLog(`fetchTickets: 検索条件 = ${searchCriteria}`);

  const threads = GmailApp.search(searchCriteria);
  debugLog(`fetchTickets: 検索スレッド数 = ${threads.length}`);

  let tickets: Ticket[] = [];

  for (const [threadIndex, thread] of threads.entries()) {
    const messages = thread.getMessages();
    debugLog(
      `fetchTickets: スレッド[${threadIndex}] 件名="${thread.getFirstMessageSubject()}" メッセージ数=${messages.length}`,
    );

    for (const [messageIndex, message] of messages.entries()) {
      const messageDate = message.getDate();
      const fromAddress = message.getFrom();
      debugLog(
        `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] from=${fromAddress} date=${messageDate.toISOString()} subject="${message.getSubject()}"`,
      );

      if (messageDate.getTime() < searchStartDateTime.getTime()) {
        debugLog(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] は検索開始日時より前のためスキップします。`,
        );
        continue;
      }

      const body = message.getPlainBody();
      debugLog(
        `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 本文 >>>\n${body}\n<<<`,
      );
      // メール仕様の変更に気づけるよう、デバッグフラグに関係なく常にログへ残す。
      const ticket = parseTicketBody(body, fromAddress, sources, (failure) => {
        if (failure.reason === "unknown_sender") {
          Logger.log(
            `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.fromAddress} は登録されていないためスキップします。`,
          );
          return;
        }

        if (failure.reason === "unrecognized_body") {
          Logger.log(
            `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.source.mailAddresses.join(", ")} の本文が想定の形式と一致しませんでした。`,
          );
          return;
        }

        const reason =
          failure.error instanceof Error ? failure.error.message : String(failure.error);
        Logger.log(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 送信元=${failure.source.mailAddresses.join(", ")} の解析に失敗しました: ${reason}`,
        );
      });

      if (ticket) {
        debugLog(
          `fetchTickets: スレッド[${threadIndex}] メッセージ[${messageIndex}] 解析成功 title="${ticket.title}" start=${ticket.startTime.toISOString()} end=${ticket.endTime.toISOString()}`,
        );
        tickets = [...tickets, ticket];
      }
    }
  }

  debugLog(`fetchTickets: 取得したチケット数 = ${tickets.length}`);

  // 予約確認メールの再送などで同じチケット番号が複数件取れることがあるため、
  // 同一実行内での重複登録を避ける。
  const uniqueTickets = dedupeTicketsByTicketNumber(tickets);
  if (uniqueTickets.length !== tickets.length) {
    debugLog(
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

  const description = `劇場: ${ticket.theater}\n座席: ${ticket.sheet}\n${buildTicketNumberMarker(ticket.ticketNumber)}\n検索用キーワード: ${calendarSearchKey}`;
  const location = ticket.theater;

  const event = calendar.createEvent(ticket.title, ticket.startTime, ticket.endTime, {
    description,
    location,
  });

  return event;
}

/**
 * カレンダーイベントの description に埋め込む、登録済みチケットの照合用マーカー。
 * `registerEvent` での登録時と `runCinemaSchedule` での重複判定の両方で使う。
 */
function buildTicketNumberMarker(ticketNumber: string): string {
  return `チケット番号: ${ticketNumber}`;
}
