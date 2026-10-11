import type { CommandResult } from "../../shared/types";
import {
  createRevertState,
  type RevertEntry,
  type WriteDescriptor,
} from "../../domain/commands/revert";

const failure = (code: string, message: string): CommandResult => ({
  code,
  message,
  ok: false,
});

export const createWriteExecutor = <Port>(port: Port) => {
  const revertState = createRevertState<Port>();
  let revertBarrier: Readonly<{
    commandId: string;
    inverseCommandId: string;
    message: string;
  }> | null = null;
  let safetyLocked = false;

  const attemptRevert = async (entry: RevertEntry<Port>) => {
    try {
      await entry.revert(port);
      return await entry.verify(port);
    } catch {
      return false;
    }
  };

  const execute = async <Snapshot>(
    descriptor: WriteDescriptor<Port, Snapshot>,
  ): Promise<CommandResult> => {
    if (safetyLocked) {
      return failure(
        "SAFETY_LOCKED",
        "Writes are paused until the previous verification failure is acknowledged",
      );
    }

    let snapshot: Snapshot;
    try {
      snapshot = await descriptor.capture(port);
    } catch {
      return failure("SNAPSHOT_FAILED", "The revert snapshot could not be captured");
    }

    let preflight: CommandResult;
    try {
      preflight = await descriptor.preflight(port, snapshot);
    } catch {
      return failure("PREFLIGHT_FAILED", "Write preflight could not be completed");
    }
    if (!preflight.ok) {
      return preflight;
    }

    const revertEntry: RevertEntry<Port> = {
      commandId: descriptor.id,
      revert: (targetPort) => descriptor.revert(targetPort, snapshot),
      verify: (targetPort) => descriptor.verifyRevert(targetPort, snapshot),
    };

    try {
      await descriptor.apply(port, snapshot);
      if (await descriptor.verify(port, snapshot)) {
        revertBarrier = null;
        revertState.record(revertEntry);
        return { ok: true };
      }
    } catch {
      // Any failure after apply begins is treated as a possibly partial write.
    }

    safetyLocked = true;
    return (await attemptRevert(revertEntry))
      ? failure(
          "WRITE_VERIFY_FAILED_REVERTED",
          "Write verification failed and the captured state was restored",
        )
      : failure(
          "WRITE_VERIFY_FAILED_REVERT_FAILED",
          "Write verification failed and the captured state could not be confirmed",
        );
  };

  const revertLast = async (): Promise<CommandResult> => {
    if (revertBarrier !== null) {
      return failure(
        "STRUCTURAL_REVERT_REQUIRES_INVERSE",
        revertBarrier.message,
      );
    }
    const entry = revertState.get();
    if (entry === null) {
      return failure("NO_REVERT", "There is no Analyst Lens action to revert");
    }
    if (await attemptRevert(entry)) {
      revertState.clear();
      return { ok: true };
    }
    safetyLocked = true;
    return failure(
      "REVERT_VERIFY_FAILED",
      "The prior state could not be restored and verified",
    );
  };

  return {
    acknowledgeSafetyLock() {
      safetyLocked = false;
    },
    blockRevert(barrier: Readonly<{
      commandId: string;
      inverseCommandId: string;
      message: string;
    }>) {
      revertBarrier = barrier;
    },
    execute,
    isSafetyLocked: () => safetyLocked,
    revertLast,
  };
};
