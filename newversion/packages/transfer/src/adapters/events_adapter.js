export function createEventsAdapter({ logger = console, emit } = {}) {
  return {
    info: (message, payload) => {
      logger?.info?.(message, payload);
      emit?.('transfer:info', { message, payload });
    },
    error: (message, payload) => {
      logger?.error?.(message, payload);
      emit?.('transfer:error', { message, payload });
    }
  };
}
