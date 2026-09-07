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
import { createLogger } from "./logger";
import type { Logger } from "./logger";

const calendarSearchKey = "映画館チケット";
const legacyCalendarSearchKeys = ["シネマシティ"];
const debugExecutionDateProperty = "DEBUG_EXECUTION_DATE";
const debugLogEnabledProperty = "DEBUG_LOG_ENABLED";

/**
 * メールの検索対象期間。`start` 以上 `end` 未満の半開区間として扱う。
 */
export type MailSearchRange = {
  start: Date;
  end: Date;
};

/**
 * Script Properties の `DEBUG_LOG_ENABLED` が有効値かどうかを返す。デバッグ
 * レベルのログを出力するかどうかの閾値として、実行開始時に一度だけ読む。
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

export function main(): void {
  const logger = createLogger(isDebugLogEnabled());

  runCinemaSchedule(resolveMailSearchRange(new Date(Date.now())), logger);
}

/**
 * 実行日を指定して実行するデバッグ用エントリーポイント。
 *
 * 指定した日に `main` を実行したのと同じ検索対象期間で動くため、期間の導出
 * 自体も含めて挙動を再現できる。
 *
 * `clasp run` からは実行日を引数で渡せるほか、Apps Script エディタから引数なしで
 * 実行する場合は Script Properties の `DEBUG_EXECUTION_DATE` を使用する。どちらも
 * `YYYY-MM-DD` 形式で、Apps Script のタイムゾーン（`Asia/Tokyo`）の日付として
 * 解釈する。
 */
export function debugMain(executionDate?: string): void {
  const logger = createLogger(isDebugLogEnabled());

  const specifiedDate =
    executionDate ??
    PropertiesService.getScriptProperties().getProperty(debugExecutionDateProperty);

  if (!specifiedDate) {
    throw new Error(`${debugExecutionDateProperty} に実行日を YYYY-MM-DD 形式で指定してください。`);
  }

  const searchRange = resolveMailSearchRange(parseExecutionDate(specifiedDate));

  // 指定した実行日が効いているかどうかは実行ログからしか判断できないため、
  // DEBUG_LOG_ENABLED の設定に関係なく常に出力する。
  logger.info(
    `debugMain: 仮想実行日=${specifiedDate} 対象範囲=${formatDateTimeForLog(searchRange.start)} 以上 ${formatDateTimeForLog(searchRange.end)} 未満`,
  );

  runCinemaSchedule(searchRange, logger);
}

/**
 * `YYYY-MM-DD` 形式の文字列を、実行環境のタイムゾーン（Apps Script では
 * `Asia/Tokyo`）におけるその日の 0 時として解釈する。
 *
 * 形式違いと実在しない日付は、黙って別の日として実行してしまわないよう例外に
 * する。
 */
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

  // 2026-02-30 のような実在しない日付は Date が翌月へ繰り上げてしまうため、
  // 組み立てた結果が指定どおりかどうかで弾く。
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    throw new Error(`${debugExecutionDateProperty} には実在する日付を指定してください: ${value}`);
  }

  return parsed;
}

/**
 * 実行日からメールの検索対象期間を導出する。
 *
 * 前日の 0 時以上、翌日の 0 時未満。開始が前日の 0 時なのは、前日中に届いた
 * 予約確認メールを取りこぼさないため。終端を翌日の 0 時に置くのは、実行日
 * 当日に届いたメールまでを対象にするため。
 *
 * `main` は実際の現在日時を、`debugMain` は指定された実行日を渡す。どちらも
 * この関数だけを通るので、デバッグ実行で確認した期間の決まり方が本番実行でも
 * そのまま成り立つ。現在日時を渡す `main` では終端が未来になるため、実質的な
 * 上限としては働かない（受信済みメールの日時が未来になることはない）。
 */
export function resolveMailSearchRange(executionDate: Date): MailSearchRange {
  const year = executionDate.getFullYear();
  const month = executionDate.getMonth();
  const day = executionDate.getDate();

  return {
    start: new Date(year, month, day - 1, 0, 0, 0, 0),
    end: new Date(year, month, day + 1, 0, 0, 0, 0),
  };
}

/**
 * ログに載せる日時の整形。`toISOString` は UTC 表示になり Apps Script の
 * タイムゾーンとずれて読みにくいため、ローカルタイムのまま組み立てる。
 */
function formatDateTimeForLog(date: Date): string {
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");

  return `${formatMailSearchDate(date)} ${hour}:${minute}`;
}

function runCinemaSchedule(searchRange: MailSearchRange, logger: Logger): void {
  // メールからチケット情報を取得する
  const tickets = fetchTickets(getTicketMailSources(), searchRange, logger);

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
      logger.info(`Skip: ${ticket.title}`);
    }

    return !isExist;
  });

  // 登録されていない場合はカレンダーに登録する;
  for (const ticket of willRegisterTickets) {
    const event = registerEvent(ticket);
    logger.info(`Registered: ${event.getTitle()}`);
  }
}

/**
 * メールを Gmail 上で開くための直リンクを組み立てる。デバッグログに本文を
 * 出力する代わりに、これを載せて元メールをたどれるようにする。
 */
function buildGmailMessageLink(messageId: string): string {
  return `https://mail.google.com/mail/u/0/#all/${messageId}`;
}

/**
 * チケット情報を取得する
 */
function fetchTickets(
  sources: readonly TicketMailSource[],
  searchRange: MailSearchRange,
  logger: Logger,
): Ticket[] {
  if (sources.length === 0) {
    logger.debug("fetchTickets: チケットメールの取得元が未設定のため終了します。");
    return [];
  }

  // Gmail の検索クエリは日単位でしか絞れないため、ここでは開始日だけを渡して
  // 粗く絞り込み、期間の厳密な判定はメッセージごとの日時比較で行う。
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
      // メール仕様の変更に気づけるよう、デバッグフラグに関係なく常にログへ残す。
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

  // 予約確認メールの再送などで同じチケット番号が複数件取れることがあるため、
  // 同一実行内での重複登録を避ける。
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
