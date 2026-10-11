import { describe, expect, it } from "vitest";
import { formatNumber, type SelectionWritePort } from "../../src/domain/commands/formatCommands";
import { createWriteExecutor } from "../../src/platform/office/writeExecutor";

const createPort = (options: { ignoreWrite?: boolean; preflightOk?: boolean } = {}) => {
  const calls: string[] = [];
  let numberFormat = "General";
  const port: SelectionWritePort = {
    async preflightSelection() {
      calls.push("preflight");
      return options.preflightOk === false
        ? { code: "PROTECTED", message: "Selection is protected", ok: false }
        : { ok: true };
    },
    async readSelectionNumberFormat() {
      calls.push("read");
      return {
        address: "$F$42",
        numberFormat,
        sheetId: "sheet-1",
        workbookId: "workbook-1",
      };
    },
    async restoreSelectionNumberFormat(snapshot) {
      calls.push("restore");
      numberFormat = snapshot.numberFormat;
    },
    async setSelectionNumberFormat(nextFormat) {
      calls.push("write");
      if (options.ignoreWrite !== true) numberFormat = nextFormat;
    },
  };
  return {
    calls,
    get numberFormat() {
      return numberFormat;
    },
    port,
  };
};

describe("verified write executor", () => {
  it("preflights, snapshots, writes, reads back, and reverts exactly once", async () => {
    const fake = createPort();
    const executor = createWriteExecutor(fake.port);

    expect(await executor.execute(formatNumber("0.0x"))).toEqual({ ok: true });
    expect(fake.numberFormat).toBe("0.0x");
    expect(fake.calls).toEqual(["read", "preflight", "write", "read"]);

    expect(await executor.revertLast()).toEqual({ ok: true });
    expect(fake.numberFormat).toBe("General");
    expect(await executor.revertLast()).toMatchObject({
      code: "NO_REVERT",
      ok: false,
    });
  });

  it("captures the target but does not write when its preflight fails", async () => {
    const fake = createPort({ preflightOk: false });
    const executor = createWriteExecutor(fake.port);

    await expect(executor.execute(formatNumber("0.0x"))).resolves.toMatchObject({
      code: "PROTECTED",
      ok: false,
    });
    expect(fake.calls).toEqual(["read", "preflight"]);
    expect(fake.numberFormat).toBe("General");
  });

  it("automatically restores and safety-locks after failed readback", async () => {
    const fake = createPort({ ignoreWrite: true });
    const executor = createWriteExecutor(fake.port);

    await expect(executor.execute(formatNumber("0.0x"))).resolves.toMatchObject({
      code: "WRITE_VERIFY_FAILED_REVERTED",
      ok: false,
    });
    expect(fake.numberFormat).toBe("General");
    expect(executor.isSafetyLocked()).toBe(true);
    await expect(executor.execute(formatNumber("0.00x"))).resolves.toMatchObject({
      code: "SAFETY_LOCKED",
      ok: false,
    });
  });

  it("blocks stale automatic revert after a successful structural mutation", async () => {
    const fake = createPort();
    const executor = createWriteExecutor(fake.port);
    const structuralExecutor = executor as typeof executor & {
      blockRevert(barrier: {
        commandId: string;
        inverseCommandId: string;
        message: string;
      }): void;
    };

    expect(await executor.execute(formatNumber("0.0x"))).toEqual({ ok: true });
    structuralExecutor.blockRevert?.({
      commandId: "structure.groupRows",
      inverseCommandId: "structure.ungroupRows",
      message: "Select the affected rows and run Ungroup rows once.",
    });

    await expect(executor.revertLast()).resolves.toEqual({
      code: "STRUCTURAL_REVERT_REQUIRES_INVERSE",
      message: "Select the affected rows and run Ungroup rows once.",
      ok: false,
    });
    expect(fake.numberFormat).toBe("0.0x");
  });
});
