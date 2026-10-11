import type {
  CellFacts,
  CellRef,
  CommandId,
  CommandResult,
  DependencyDirection,
  DependencyRef,
} from "../../shared/types";
import type { CapabilityMap } from "./capabilityMap";
import type { HealthWorkbookFacts } from "../../domain/health/rules";
import type { PersistedWatchItem } from "../../domain/watchlist/watchlist";

export type OfficeGatewayErrorCode =
  | "ABORTED"
  | "ACCESS_DENIED"
  | "INVALID_REFERENCE"
  | "NOT_FOUND"
  | "OFFICE_ERROR"
  | "STALE_REQUEST"
  | "UNSUPPORTED_OPERATION";

export class OfficeGatewayError extends Error {
  readonly code: OfficeGatewayErrorCode;

  constructor(code: OfficeGatewayErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "OfficeGatewayError";
    this.code = code;
  }
}

const codeMap: Readonly<Record<string, OfficeGatewayErrorCode>> = {
  AccessDenied: "ACCESS_DENIED",
  ApiNotFound: "UNSUPPORTED_OPERATION",
  InvalidObjectPath: "INVALID_REFERENCE",
  InvalidReference: "INVALID_REFERENCE",
  ItemNotFound: "NOT_FOUND",
  UnsupportedFeature: "UNSUPPORTED_OPERATION",
  UnsupportedOperation: "UNSUPPORTED_OPERATION",
};

const messageMap: Readonly<Record<OfficeGatewayErrorCode, string>> = {
  ABORTED: "The Office request was cancelled",
  ACCESS_DENIED: "Excel denied access to the requested workbook area",
  INVALID_REFERENCE: "The workbook reference is no longer valid",
  NOT_FOUND: "The requested workbook item was not found",
  OFFICE_ERROR: "Excel could not complete the requested operation",
  STALE_REQUEST: "A newer Office request superseded this result",
  UNSUPPORTED_OPERATION: "This Excel version does not support the requested operation",
};

export const normalizeOfficeError = (error: unknown): OfficeGatewayError => {
  if (error instanceof OfficeGatewayError) {
    return error;
  }
  if (
    typeof error === "object" &&
    error !== null &&
    "name" in error &&
    error.name === "AbortError"
  ) {
    return new OfficeGatewayError("ABORTED", messageMap.ABORTED, {
      cause: error,
    });
  }
  const sourceCode =
    typeof error === "object" && error !== null && "code" in error
      ? String(error.code)
      : "";
  const code = codeMap[sourceCode] ?? "OFFICE_ERROR";
  return new OfficeGatewayError(code, messageMap[code], { cause: error });
};

export type WorkbookWorksheetFacts = {
  id: string;
  name: string;
  visibility: string;
};

export type WorkbookFacts = {
  calculationMode: "automatic" | "automaticExceptTables" | "manual" | "unknown";
  id: string;
  name: string;
  worksheets: WorkbookWorksheetFacts[];
};

export type WriteRequest = Readonly<{ id: CommandId }>;
export type Unsubscribe = () => void;
export type NeighborDirection = "above" | "below" | "left" | "right";
export type NeighborCellFacts = Readonly<{
  direction: NeighborDirection;
  facts: CellFacts;
}>;

export interface OfficeGateway {
  getCapabilities(): CapabilityMap;
  getSelectionFacts(signal?: AbortSignal): Promise<CellFacts>;
  getCellFacts?(ref: CellRef, signal?: AbortSignal): Promise<CellFacts>;
  getCellsFacts?(
    refs: readonly CellRef[],
    signal?: AbortSignal,
  ): Promise<readonly (CellFacts | null)[]>;
  getDirectDependencies(
    ref: CellRef,
    direction: DependencyDirection,
    signal?: AbortSignal,
  ): Promise<DependencyRef[]>;
  getNeighborFacts?(ref: CellRef, signal?: AbortSignal): Promise<NeighborCellFacts[]>;
  getHealthWorkbookFacts?(signal?: AbortSignal): Promise<HealthWorkbookFacts>;
  loadWatchlistMetadata?(): Promise<readonly PersistedWatchItem[]>;
  saveWatchlistMetadata?(metadata: readonly PersistedWatchItem[]): Promise<void>;
  selectCell(ref: CellRef): Promise<void>;
  subscribeSelection(listener: () => void): Promise<Unsubscribe>;
  subscribeRecalculation(listener: () => void): Promise<Unsubscribe>;
  getWorkbookFacts(signal?: AbortSignal): Promise<WorkbookFacts>;
  runWrite(descriptor: WriteRequest): Promise<CommandResult>;
}
