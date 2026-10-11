import { afterEach, describe, expect, it, vi } from "vitest";
import { createOfficeQuickTraceDialogChild } from "../../src/platform/office/quickTraceDialogChild";
import { createQuickTraceDialogHost } from "../../src/platform/office/quickTraceDialogHost";

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("Office Quick Trace dialog host", () => {
  it("opens one movable Office dialog and reuses the live session", async () => {
    let opens = 0;
    const posted: Array<readonly [string, { targetOrigin: string }]> = [];
    const handlers = new Map<
      string,
      (arg: { error?: number; message?: string; origin?: string }) => void
    >();
    const dialog = {
      addEventHandler: (
        type: unknown,
        handler: (arg: { error?: number; message?: string; origin?: string }) => void,
      ) => handlers.set(String(type), handler),
      close: () => undefined,
      messageChild: (message: string, options: { targetOrigin: string }) => {
        posted.push([message, options]);
      },
    };
    const host = createQuickTraceDialogHost({
      asyncStatusSucceeded: "succeeded",
      dialogEventReceived: "dialog-event",
      dialogMessageReceived: "dialog-message",
      displayDialogAsync: (_url, options, callback) => {
        opens += 1;
        expect(options).toEqual({ height: 52, promptBeforeOpen: false, width: 36 });
        callback({ status: "succeeded", value: dialog });
      },
      isDialogApi12Supported: () => true,
    });

    const first = await host.open("https://127.0.0.1:43119/quick-trace.html");
    const second = await host.open("https://127.0.0.1:43119/quick-trace.html");
    const received: string[] = [];
    first.subscribeMessage((message) => received.push(message));
    first.post("host-state");
    handlers.get("dialog-message")?.({
      message: "trusted",
      origin: "https://127.0.0.1:43119",
    });
    handlers.get("dialog-message")?.({
      message: "untrusted",
      origin: "https://attacker.example",
    });

    expect(second).toBe(first);
    expect(opens).toBe(1);
    expect(posted).toEqual([[
      "host-state",
      { targetOrigin: "https://127.0.0.1:43119" },
    ]]);
    expect(received).toEqual(["trusted"]);
  });

  it("clears the live session after DialogEventReceived", async () => {
    let closed: ((arg: { error?: number }) => void) | undefined;
    let opens = 0;
    const host = createQuickTraceDialogHost({
      asyncStatusSucceeded: "succeeded",
      dialogEventReceived: "closed",
      dialogMessageReceived: "message",
      displayDialogAsync: (_url, _options, callback) => {
        opens += 1;
        callback({
          status: "succeeded",
          value: {
            addEventHandler: (type, handler) => {
              if (String(type) === "closed") closed = handler;
            },
            close: () => undefined,
            messageChild: () => undefined,
          },
        });
      },
      isDialogApi12Supported: () => true,
    });

    await host.open("https://127.0.0.1:43119/quick-trace.html");
    closed?.({ error: 12006 });
    await host.open("https://127.0.0.1:43119/quick-trace.html");

    expect(opens).toBe(2);
  });

  it("clears the live session when programmatic close emits no native event", async () => {
    let closes = 0;
    let opens = 0;
    const host = createQuickTraceDialogHost({
      asyncStatusSucceeded: "succeeded",
      dialogEventReceived: "closed",
      dialogMessageReceived: "message",
      displayDialogAsync: (_url, _options, callback) => {
        opens += 1;
        callback({
          status: "succeeded",
          value: {
            addEventHandler: () => undefined,
            close: () => {
              closes += 1;
            },
            messageChild: () => undefined,
          },
        });
      },
      isDialogApi12Supported: () => true,
    });

    const first = await host.open("https://127.0.0.1:43119/quick-trace.html");
    const closed = vi.fn();
    first.subscribeClosed(closed);
    first.close();
    const second = await host.open("https://127.0.0.1:43119/quick-trace.html");

    expect(second).not.toBe(first);
    expect(opens).toBe(2);
    expect(closes).toBe(1);
    expect(closed).toHaveBeenCalledOnce();
    expect(closed).toHaveBeenCalledWith(0);
  });

  it("rejects unsupported DialogApi 1.2 before opening", async () => {
    const host = createQuickTraceDialogHost({
      asyncStatusSucceeded: "succeeded",
      dialogEventReceived: "closed",
      dialogMessageReceived: "message",
      displayDialogAsync: () => {
        throw new Error("must not open");
      },
      isDialogApi12Supported: () => false,
    });

    expect(host.isSupported()).toBe(false);
    await expect(host.open("https://127.0.0.1:43119/quick-trace.html"))
      .rejects.toThrow("DialogApi 1.2");
  });
});

describe("Office Quick Trace dialog child", () => {
  it("posts only across the same-origin parent boundary", async () => {
    let parentHandler: ((event: { message: string; origin?: string }) => void) | undefined;
    const posted: Array<readonly [string, { targetOrigin: string }]> = [];
    vi.stubGlobal("Office", {
      AsyncResultStatus: { Succeeded: "succeeded" },
      EventType: { DialogParentMessageReceived: "parent-message" },
      context: {
        ui: {
          addHandlerAsync: (
            _type: string,
            handler: (event: { message: string; origin?: string }) => void,
            callback: (result: { status: string }) => void,
          ) => {
            parentHandler = handler;
            callback({ status: "succeeded" });
          },
          messageParent: (message: string, options: { targetOrigin: string }) => {
            posted.push([message, options]);
          },
        },
      },
      onReady: async () => undefined,
    });
    vi.stubGlobal("window", {
      location: { origin: "https://127.0.0.1:43119" },
    });
    const child = await createOfficeQuickTraceDialogChild();
    const received: string[] = [];
    child.subscribe((message) => received.push(message));

    child.post("dialog-ready");
    parentHandler?.({
      message: "trusted",
      origin: "https://127.0.0.1:43119",
    });
    parentHandler?.({
      message: "untrusted",
      origin: "https://attacker.example",
    });

    expect(posted).toEqual([[
      "dialog-ready",
      { targetOrigin: "https://127.0.0.1:43119" },
    ]]);
    expect(received).toEqual(["trusted"]);
  });
});
