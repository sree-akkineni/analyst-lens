export type QuickTraceDialogChild = Readonly<{
  post(message: string): void;
  subscribe(listener: (message: string, origin?: string) => void): () => void;
}>;

export const createOfficeQuickTraceDialogChild = async (): Promise<QuickTraceDialogChild> => {
  await Office.onReady();
  const listeners = new Set<(message: string, origin?: string) => void>();
  await new Promise<void>((resolve, reject) => {
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived,
      (event) => {
        if (event.origin !== undefined && event.origin !== window.location.origin) return;
        listeners.forEach((listener) => listener(event.message, event.origin));
      },
      (result) => result.status === Office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(new Error(result.error.message)),
    );
  });

  return {
    post: (message) => Office.context.ui.messageParent(message, {
      targetOrigin: window.location.origin,
    }),
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
};
