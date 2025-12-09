export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface Logger {
  debug(message: string, ...optionalParams: unknown[]): void;
  info(message: string, ...optionalParams: unknown[]): void;
  warn(message: string, ...optionalParams: unknown[]): void;
  error(message: string, ...optionalParams: unknown[]): void;
}

let debugEnabled = false;

export function setDebugEnabled(enabled: boolean) {
  debugEnabled = enabled;
}

export function isDebugEnabled(): boolean {
  return debugEnabled;
}

function log(level: LogLevel, prefix: string, message: string, optionalParams: unknown[]) {
  const args: unknown[] = [`[RecipientCheck:${prefix}] ${message}`, ...optionalParams];
  switch (level) {
    case 'debug':
      if (debugEnabled) {
        console.debug(...args);
      }
      break;
    case 'info':
      console.info(...args);
      break;
    case 'warn':
      console.warn(...args);
      break;
    case 'error':
      console.error(...args);
      break;
    default:
      console.log(...args);
  }
}

export function createLogger(prefix: string): Logger {
  return {
    debug(message: string, ...optionalParams: unknown[]) {
      log('debug', prefix, message, optionalParams);
    },
    info(message: string, ...optionalParams: unknown[]) {
      log('info', prefix, message, optionalParams);
    },
    warn(message: string, ...optionalParams: unknown[]) {
      log('warn', prefix, message, optionalParams);
    },
    error(message: string, ...optionalParams: unknown[]) {
      log('error', prefix, message, optionalParams);
    }
  };
}
