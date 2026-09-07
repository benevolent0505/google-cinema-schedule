/**
 * `console.log` / `console.info` / `console.warn` / `console.error` は GAS の V8
 * ランタイム上でそれぞれ Cloud Logging の DEBUG / INFO / WARNING / ERROR severity
 * に対応するため、レベルを表すプレフィックスを手で付ける必要はない。
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
