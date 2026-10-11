import { describe, expect, it } from "vitest";
import {
  createOfficeCommandRunner,
  createOfficeGatewayProvider,
} from "../../src/runtime/officeGatewayBootstrap";

describe("Office gateway bootstrap", () => {
  it("does not start Office readiness until a command needs the gateway", () => {
    let readyCalls = 0;
    createOfficeCommandRunner({
      createGateway: () => ({ runWrite: async () => ({ ok: true }) }),
      ready: async () => {
        readyCalls += 1;
      },
    });

    expect(readyCalls).toBe(0);
  });

  it("does not construct the gateway until Office is ready", async () => {
    let releaseReady: (() => void) | undefined;
    const ready = new Promise<void>((resolve) => {
      releaseReady = resolve;
    });
    const created: string[] = [];
    const runner = createOfficeCommandRunner({
      createGateway: () => {
        created.push("gateway");
        return { runWrite: async ({ id }: { id: string }) => ({ code: id, message: id, ok: false as const }) };
      },
      ready: () => ready,
    });

    const result = runner("format.input");
    expect(created).toEqual([]);

    releaseReady?.();
    await expect(result).resolves.toMatchObject({ code: "format.input", ok: false });
    expect(created).toEqual(["gateway"]);
  });

  it("allows a later command to retry after Office readiness rejects", async () => {
    let readyCalls = 0;
    const runner = createOfficeCommandRunner({
      createGateway: () => ({ runWrite: async () => ({ ok: true }) }),
      ready: async () => {
        readyCalls += 1;
        if (readyCalls === 1) throw new Error("Office unavailable");
      },
    });

    await expect(runner("format.input")).rejects.toThrow("Office unavailable");
    await expect(runner("format.input")).resolves.toEqual({ ok: true });
    expect(readyCalls).toBe(2);
  });

  it("shares one lazy gateway between command and controller consumers", async () => {
    let creations = 0;
    const getGateway = createOfficeGatewayProvider({
      createGateway: () => {
        creations += 1;
        return { runWrite: async () => ({ ok: true }) };
      },
      ready: async () => undefined,
    });

    const [first, second] = await Promise.all([getGateway(), getGateway()]);
    expect(first).toBe(second);
    expect(creations).toBe(1);
  });
});
