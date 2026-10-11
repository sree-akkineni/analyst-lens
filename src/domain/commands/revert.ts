import type { CommandId, CommandResult } from "../../shared/types";

export type WriteDescriptor<Port, Snapshot> = {
  apply(port: Port, snapshot: Snapshot): Promise<void>;
  capture(port: Port): Promise<Snapshot>;
  id: CommandId;
  preflight(port: Port, snapshot: Snapshot): Promise<CommandResult>;
  revert(port: Port, snapshot: Snapshot): Promise<void>;
  verify(port: Port, snapshot: Snapshot): Promise<boolean>;
  verifyRevert(port: Port, snapshot: Snapshot): Promise<boolean>;
};

export type RevertEntry<Port> = {
  commandId: CommandId;
  revert(port: Port): Promise<void>;
  verify(port: Port): Promise<boolean>;
};

export const createRevertState = <Port>() => {
  let last: RevertEntry<Port> | null = null;
  return {
    clear() {
      last = null;
    },
    get(): RevertEntry<Port> | null {
      return last;
    },
    record(entry: RevertEntry<Port>) {
      last = entry;
    },
  };
};
