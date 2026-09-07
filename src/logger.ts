/**
 * レベル付きロガー。実行開始時（`main` / `debugMain`）に一度だけ生成し、
 * 以降の処理へ明示的に渡して使う。
 *
 * `console.log` / `console.info` / `console.warn` / `console.error` は GAS の
 * V8 ランタイム上でそれぞれ Cloud Logging の DEBUG / INFO / WARNING / ERROR
 * severity に対応するため、手動でレベルを表すプレフィックスを付ける必要はない。
 *
 * `debug` は `debugEnabled` が true のときだけ実際に出力する。詳細なトレース
 * ログ（検索条件・メッセージ本文のメタ情報など）はここに分類する。それ以外の
 * レベルは常に出力する。
 */
export type Logger = {
  debug: (message: string) => void;
  info: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, error: unknown) => void;
};

export function createLogger(debugEnabled: boolean): Logger {
  return {
    debug: (message) => {
      if (debugEnabled) {
        console.log(message);
      }
    },
    info: (message) => {
      console.info(message);
    },
    warn: (message) => {
      console.warn(message);
    },
    error: (message, error) => {
      console.error(message, error);
    },
  };
}
