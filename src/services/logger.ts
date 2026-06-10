type LogLevel = 'info' | 'warn' | 'error';

function write(level: LogLevel, message: string, details?: Record<string, unknown>) {
  const payload = {
    level,
    message,
    timestamp: new Date().toISOString(),
    ...(details || {}),
  };
  const line = JSON.stringify(payload);
  if (level === 'error') {
    console.error(line);
  } else if (level === 'warn') {
    console.warn(line);
  } else {
    console.log(line);
  }
}

export const logger = {
  info: (message: string, details?: Record<string, unknown>) => write('info', message, details),
  warn: (message: string, details?: Record<string, unknown>) => write('warn', message, details),
  error: (message: string, details?: Record<string, unknown>) => write('error', message, details),
};
