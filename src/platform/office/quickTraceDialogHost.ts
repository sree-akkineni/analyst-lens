export type QuickTraceDialogSession = Readonly<{
  close(): void;
  post(message: string): void;
  subscribeClosed(listener: (error: number) => void): () => void;
  subscribeMessage(listener: (message: string, origin?: string) => void): () => void;
}>;

export type QuickTraceDialogHost = Readonly<{
  isSupported(): boolean;
  open(url: string): Promise<QuickTraceDialogSession>;
}>;

type DialogEvent = Readonly<{
  error?: number;
  message?: string;
  origin?: string;
}>;

type DialogHandle = Readonly<{
  addEventHandler(type: unknown, handler: (event: DialogEvent) => void): void;
  close(): void;
  messageChild(message: string, options: { targetOrigin: string }): void;
}>;

type DialogAsyncResult = Readonly<{
  error?: Readonly<{ code?: number; message?: string }>;
  status: unknown;
  value?: DialogHandle;
}>;

export type QuickTraceDialogRuntime = Readonly<{
  asyncStatusSucceeded: unknown;
  dialogEventReceived: unknown;
  dialogMessageReceived: unknown;
  displayDialogAsync(
    url: string,
    options: { height: number; promptBeforeOpen: boolean; width: number },
    callback: (result: DialogAsyncResult) => void,
  ): void;
  isDialogApi12Supported(): boolean;
}>;

const createOfficeRuntime = (): QuickTraceDialogRuntime => ({
  asyncStatusSucceeded: Office.AsyncResultStatus.Succeeded,
  dialogEventReceived: Office.EventType.DialogEventReceived,
  dialogMessageReceived: Office.EventType.DialogMessageReceived,
  displayDialogAsync: (url, options, callback) => {
    Office.context.ui.displayDialogAsync(url, options, (result) => {
      callback({
        ...(result.error === undefined ? {} : { error: result.error }),
        status: result.status,
        ...(result.value === undefined
          ? {}
          : { value: result.value as unknown as DialogHandle }),
      });
    });
  },
  isDialogApi12Supported: () =>
    Office.context.requirements.isSetSupported("DialogApi", "1.2"),
});

export const createQuickTraceDialogHost = (
  injectedRuntime?: QuickTraceDialogRuntime,
): QuickTraceDialogHost => {
  const runtime = injectedRuntime ?? createOfficeRuntime();
  let live: QuickTraceDialogSession | null = null;
  let opening: Promise<QuickTraceDialogSession> | null = null;

  const open = async (url: string): Promise<QuickTraceDialogSession> => {
    if (!runtime.isDialogApi12Supported()) {
      throw new Error("Quick Trace requires Office DialogApi 1.2");
    }
    if (live !== null) return live;
    if (opening !== null) return opening;

    opening = new Promise<QuickTraceDialogSession>((resolve, reject) => {
      const allowedOrigin = new URL(url).origin;
      runtime.displayDialogAsync(
        url,
        { height: 52, promptBeforeOpen: false, width: 36 },
        (result) => {
          if (result.status !== runtime.asyncStatusSucceeded || result.value === undefined) {
            opening = null;
            const code = result.error?.code;
            const message = result.error?.message ?? "Office could not open Quick Trace";
            reject(new Error(code === undefined ? message : `${message} (${code})`));
            return;
          }

          const dialog = result.value;
          const buffered: Array<Readonly<{ message: string; origin?: string }>> = [];
          const closedListeners = new Set<(error: number) => void>();
          const messageListeners = new Set<(message: string, origin?: string) => void>();
          let closed = false;

          const finish = (error: number) => {
            if (closed) return;
            closed = true;
            if (live === session) live = null;
            opening = null;
            const listeners = [...closedListeners];
            closedListeners.clear();
            messageListeners.clear();
            buffered.length = 0;
            listeners.forEach((listener) => listener(error));
          };

          const session: QuickTraceDialogSession = {
            close: () => {
              finish(0);
              dialog.close();
            },
            post: (message) => dialog.messageChild(message, { targetOrigin: allowedOrigin }),
            subscribeClosed(listener) {
              closedListeners.add(listener);
              return () => closedListeners.delete(listener);
            },
            subscribeMessage(listener) {
              messageListeners.add(listener);
              for (const event of buffered.splice(0)) {
                listener(event.message, event.origin);
              }
              return () => messageListeners.delete(listener);
            },
          };

          dialog.addEventHandler(runtime.dialogMessageReceived, (event) => {
            const message = event.message;
            if (message === undefined) return;
            if (event.origin !== undefined && event.origin !== allowedOrigin) return;
            if (messageListeners.size === 0) {
              buffered.push({
                message,
                ...(event.origin === undefined ? {} : { origin: event.origin }),
              });
              return;
            }
            messageListeners.forEach((listener) => listener(message, event.origin));
          });
          dialog.addEventHandler(runtime.dialogEventReceived, (event) => {
            finish(event.error ?? 0);
          });

          live = session;
          opening = null;
          resolve(session);
        },
      );
    });

    return opening;
  };

  return {
    isSupported: () => runtime.isDialogApi12Supported(),
    open,
  };
};
